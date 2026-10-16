#!/usr/bin/env python3
"""ntperfscan -- static worklist generator for the archived NT SDK/DDK sources.

This repository is an archive of Microsoft SDK/DDK media whose native build
environment is not available.  Auditing roughly half a million lines of driver
source therefore has to be done by narrowing from the outside in:

    inventory  ->  which subsystems exist and how large are they
    hotpaths   ->  which entry points recur (timers, DPCs, ISRs, StartIo)
    divloop    ->  where division/modulo appears inside a loop
    invariant  ->  where a call inside a loop cannot have changing inputs

The tool deliberately produces *worklists, not findings*.  A raw match is not
a defect: see README.md for the false positive modes these detectors exhibit,
all of which were observed on the tree in this repository.

Usage
    python3 ntperfscan.py <command> [root] [--json OUT]

Commands
    inventory    per-directory file, line and function counts
    hotpaths     recurring entry points and the registered callbacks
    divloop      division/modulo inside loop bodies
    invariant    calls inside loops with loop-invariant arguments
    all          every command, written to a JSON worklist

If root is omitted the NT 4.0 DDK source tree in this repository is used.
"""

import argparse
import json
import os
import re
import sys
from collections import Counter, defaultdict

SRC_EXT = ('.c', '.cpp')
SKIP_DIRS = ('bin', 'lib', 'hlp', 'doc', 'obj')

IDENT = r'[A-Za-z_][A-Za-z0-9_]*'

DEFAULT_ROOT = os.path.join(
    "Microsoft Windows NT 4.0 DDK (4.0.1381)",
    "winnt40.ddk", "ddk", "src")


# --------------------------------------------------------------------------
# lexing helpers
#
# Two transformations are needed before any structural analysis:
#   * comments and string/char literals must not participate in matching
#   * preprocessor directives must not participate in brace/paren matching,
#     because directives such as '#define PAR_PAPER_EMPTY(Status)' contain
#     unbalanced parentheses that otherwise corrupt every nested match
# Both preserve byte offsets so that reported positions stay accurate.
# --------------------------------------------------------------------------

def blank_preprocessor(src):
    out = list(src)
    for m in re.finditer(r'(?m)^[ \t]*#.*(?:\\\r?\n.*)*', src):
        for k in range(m.start(), m.end()):
            if out[k] != '\n':
                out[k] = ' '
    return ''.join(out)


def lex(src):
    """Blank out comments, string/char literals and preprocessor lines."""
    src = blank_preprocessor(src)
    out = list(src)
    i, n = 0, len(src)
    while i < n:
        c = src[i]
        if c == '/' and i + 1 < n and src[i + 1] == '*':
            j = src.find('*/', i + 2)
            j = n if j < 0 else j + 2
            for k in range(i, j):
                if out[k] != '\n':
                    out[k] = ' '
            i = j
        elif c == '/' and i + 1 < n and src[i + 1] == '/':
            j = src.find('\n', i)
            j = n if j < 0 else j
            for k in range(i, j):
                out[k] = ' '
            i = j
        elif c in '"\'':
            q = c
            j = i + 1
            while j < n:
                if src[j] == '\\':
                    j += 2
                    continue
                if src[j] == q:
                    j += 1
                    break
                if src[j] == '\n':
                    break
                j += 1
            for k in range(i, j):
                if out[k] != '\n':
                    out[k] = ' '
            i = j
        else:
            i += 1
    return ''.join(out)


def line_of(s, off):
    return s.count('\n', 0, off) + 1


def match_pair(src, i, open_ch, close_ch):
    depth = 0
    while i < len(src):
        if src[i] == open_ch:
            depth += 1
        elif src[i] == close_ch:
            depth -= 1
            if depth == 0:
                return i
        i += 1
    return len(src) - 1


def source_files(root):
    for dp, dn, fn in os.walk(root):
        dn[:] = [d for d in dn if d.lower() not in SKIP_DIRS]
        for f in sorted(fn):
            if f.lower().endswith(SRC_EXT):
                yield os.path.join(dp, f)


def read(path):
    try:
        with open(path, 'r', errors='replace') as fh:
            return fh.read()
    except OSError:
        return None


C_KEYWORDS = frozenset('''if else for while do switch case default return
    break continue goto sizeof typedef struct union enum static const extern
    register volatile inline signed unsigned void char short int long float
    double asm __asm __inline __declspec cdecl __cdecl stdcall __stdcall
    fastcall __fastcall pascal __pascal fortran see_above'''.split())

# A definition may put the return type on the previous line (the DDK kernel
# style) or on the same line (the video/user-mode style).  Match a run of
# type tokens at the start of a line and take the identifier that immediately
# precedes '('.  Statements ending in ';' are rejected by the '{' test.
HEAD = re.compile(r'(?m)^[ \t]*(?:' + IDENT + r'[ \t]+|\*[ \t]*)*?(' + IDENT + r')[ \t]*\(')


def functions(raw, src):
    """Top-level function definitions as (name, body_start, body_end)."""
    out = []
    for m in HEAD.finditer(src):
        name = m.group(1)
        if name in C_KEYWORDS:
            continue
        head = src[:m.start()]
        if head.count('{') - head.count('}') != 0:
            continue
        p = m.end() - 1
        rp = match_pair(src, p, '(', ')')
        j = rp + 1
        while j < len(src) and src[j] in ' \t\r\n':
            j += 1
        if j < len(src) and src[j] == '{':
            e = match_pair(src, j, '{', '}')
            out.append((name, j, e, line_of(src, m.start())))
    return out


def loops_in(src, start, end):
    """Loops inside src[start:end] as (loop_offset, body_start, body_end)."""
    res = []
    for m in re.finditer(r'\b(for|while)\b', src[start:end]):
        o = start + m.start()
        p = src.find('(', o)
        if p < 0 or p - o > 8:
            continue
        rp = match_pair(src, p, '(', ')')
        j = rp + 1
        while j < len(src) and src[j] in ' \t\r\n':
            j += 1
        if j < len(src) and src[j] == '{':
            e = match_pair(src, j, '{', '}')
        else:
            e = src.find(';', j)
            if e < 0:
                continue
        res.append((o, j, e))
    return res


# --------------------------------------------------------------------------
# commands
# --------------------------------------------------------------------------

def cmd_inventory(root, args):
    per_dir = defaultdict(lambda: [0, 0, 0])   # files, lines, functions
    per_ext = Counter()
    for path in source_files(root):
        raw = read(path)
        if raw is None:
            continue
        src = lex(raw)
        rel = os.path.relpath(path, root)
        key = os.sep.join(rel.split(os.sep)[:2])
        d = per_dir[key]
        d[0] += 1
        d[1] += raw.count('\n')
        d[2] += len(functions(raw, src))
        per_ext[os.path.splitext(path)[1].lower()] += 1
    rows = sorted(per_dir.items(), key=lambda kv: -kv[1][1])
    print('%-46s %6s %10s %8s' % ('subsystem', 'files', 'lines', 'funcs'))
    for k, v in rows:
        print('%-46s %6d %10d %8d' % (k[:46], v[0], v[1], v[2]))
    print('%-46s %6d %10d %8d' % ('TOTAL', sum(v[0] for _, v in rows),
                                  sum(v[1] for _, v in rows),
                                  sum(v[2] for _, v in rows)))
    return {'directories': {k: {'files': v[0], 'lines': v[1], 'functions': v[2]}
                            for k, v in rows},
            'extensions': dict(per_ext)}


# Registration sites that imply recurring execution, and the symbol they name.
HOTPATTERNS = (
    ('timer', r'IoInitializeTimer\s*\(\s*[^,]+,\s*(' + IDENT + r')'),
    ('startio', r'DriverObject->DriverStartIo\s*=\s*(' + IDENT + r')'),
    ('dpc', r'IoInitializeDpcRequest\s*\(\s*[^,]+,\s*(' + IDENT + r')'),
    ('dpc', r'KeInitializeDpc\s*\(\s*&?[^,]+,\s*(' + IDENT + r')'),
    # IoConnectInterrupt(InterruptObject, ServiceRoutine, ServiceContext, ...)
    ('isr', r'IoConnectInterrupt\s*\(\s*[^,]+,\s*&?\s*(' + IDENT + r')\s*,'),
)


def cmd_hotpaths(root, args):
    found = defaultdict(list)
    for path in source_files(root):
        raw = read(path)
        if raw is None:
            continue
        src = lex(raw)
        rel = os.path.relpath(path, root)
        for kind, pat in HOTPATTERNS:
            for m in re.finditer(pat, src):
                found[kind].append({
                    'kind': kind, 'callback': m.group(1), 'file': rel,
                    'line': line_of(src, m.start()),
                })
    for kind in sorted(found):
        print('\n== %s handlers (%d) ==' % (kind, len(found[kind])))
        for h in found[kind]:
            print('  %-28s %s:%d' % (h['callback'], h['file'], h['line']))
    return {'hotpaths': {k: v for k, v in found.items()}}


def _constant_expression(rhs):
    """True for expressions built only from literals, e.g. '( 60 * 75 )'."""
    return bool(re.fullmatch(r'[\s0-9a-fA-FxX()+\-*/%]+', rhs)) and \
        not re.search(r'[A-Za-z_]', rhs.replace('0x', '').replace('X', ''))


DIVOP = re.compile(r'(?<![/%*+\-<>=!])(/|%)(?![/%*=])')
LITERAL = re.compile(r'^\s*\(*\s*(0[xX][0-9a-fA-F]+|\d+)[uUlL]*\s*\)*\s*$')
# sizeof(...) and ALL_CAPS macro names are compile time constants; anything
# else that reaches the report is a genuine variable or struct member.
CONSTISH = re.compile(r'^\s*\(*\s*(sizeof\s*\(|[A-Z][A-Z0-9_]*)\b')
LIKELY_MACRO = re.compile(r'^[A-Z][A-Z0-9_]*$')
# characters that end an operand expression at parenthesis depth 0
OPERAND_STOP = ';,)}]'
OPERAND_OP = '+-*/%&|^<>=!?:'


def _operand(body, at):
    """Text of the divisor starting just after the operator."""
    k = at
    while k < len(body) and body[k] in ' \t\r\n':
        k += 1
    if k < len(body) and body[k] in OPERAND_STOP:
        return ''
    end, depth = k, 0
    while end < len(body):
        ch = body[end]
        if ch in '([':
            depth += 1
        elif ch in ')]':
            if depth == 0:
                break
            depth -= 1
        elif depth == 0 and ch in OPERAND_STOP:
            break
        elif depth == 0 and ch in ' \t\r\n':
            # a whitespace run followed by another operator ends the operand
            j = end
            while j < len(body) and body[j] in ' \t\r\n':
                j += 1
            if j < len(body) and body[j] in OPERAND_OP + '{}':
                break
            end = j
            continue
        end += 1
    return body[k:end]


def cmd_divloop(root, args):
    hits = []
    for path in source_files(root):
        raw = read(path)
        if raw is None:
            continue
        src = lex(raw)
        rel = os.path.relpath(path, root)
        for name, bs, be, _ in functions(raw, src):
            for o, ls, le in loops_in(src, bs, be):
                body = src[ls:le]
                if len(body) < 20:
                    continue
                for m in DIVOP.finditer(body):
                    rhs = _operand(body, m.end())
                    if not rhs:
                        continue
                    ls_line = body.rfind('\n', 0, m.start()) + 1
                    le_line = body.find('\n', m.start())
                    ln = body[ls_line:le_line if le_line > 0 else len(body)]
                    if re.search(r'\b(float|double|FLOAT|DOUBLE)\b', ln):
                        continue
                    if LITERAL.match(rhs):
                        continue          # folded by any optimising compiler
                    if _constant_expression(rhs):
                        continue          # e.g. ( 60 * 75 )
                    macroish = bool(LIKELY_MACRO.match(rhs)) or \
                        bool(CONSTISH.match(rhs))
                    hits.append({
                        'file': rel, 'func': name,
                        'loop_line': line_of(src, o),
                        'op_line': line_of(src, ls + m.start()),
                        'op': m.group(1), 'divisor': rhs[:48],
                        'macro_or_sizeof': macroish,
                        'text': ' '.join(ln.split())[:120],
                    })
    variable = [h for h in hits if not h['macro_or_sizeof']]
    print('division/modulo inside loops: %d sites' % len(hits))
    print('  divisor is a macro or sizeof() (not a defect): %d' % (len(hits) - len(variable)))
    print('  divisor is a variable or member (review these): %d' % len(variable))
    for h in sorted(variable, key=lambda h: h['func']):
        print('  %-38s %-24s L%-6d %s %s'
              % (os.path.basename(h['file']), h['func'][:24], h['op_line'],
                 h['op'], h['text'][:64]))
    return {'divloop': hits}


ASSIGN = re.compile(r'(?:^|[;{}\(,])\s*\*?\s*(' + IDENT +
                    r')(?:\s*->\s*' + IDENT + r'|\s*\.\s*' + IDENT + r')*'
                    r'(?:\[[^\]]*\])?\s*(?:[-+*/%&|^]?=|\+\+|--)')
INCDEC = re.compile(r'(\+\+|--)\s*(' + IDENT + r')|\b(' + IDENT + r')\s*(\+\+|--)')
ADDR = re.compile(r'&\s*(' + IDENT + r')')

KEYWORDS = set('''if for while switch return sizeof else do case break continue
    goto TRUE FALSE NULL VOID ULONG UCHAR USHORT BOOLEAN NTSTATUS PIRP PVOID
    char int long short unsigned signed void float double struct union enum
    static const extern register volatile typedef inline asm __asm
    defined'''.split())

ACCESSOR = re.compile(r'^(IoGet|KeGet|MmGet|ExGet|ObGet|CONTAINING_RECORD|'
                      r'ABS|MIN|MAX|NT_SUCCESS|ASSERT|DebugPrint|DbgPrint|'
                      r'IF_LOG|IF_LOUD|ParDump)')

CALL = re.compile(r'\b(' + IDENT + r')\s*\(')


def _modified(body):
    s = set()
    for m in ASSIGN.finditer(body):
        s.add(m.group(1))
    for m in INCDEC.finditer(body):
        s.add(m.group(2) or m.group(3))
    for m in ADDR.finditer(body):
        s.add(m.group(1))
    return s


def cmd_invariant(root, args):
    hits = []
    for path in source_files(root):
        raw = read(path)
        if raw is None:
            continue
        src = lex(raw)
        rel = os.path.relpath(path, root)
        for name, bs, be, _ in functions(raw, src):
            for o, ls, le in loops_in(src, bs, be):
                body = src[ls:le]
                if len(body) < 60:
                    continue
                mods = _modified(body)
                seen = set()
                for m in CALL.finditer(body):
                    callee = m.group(1)
                    if callee in seen or callee in KEYWORDS:
                        continue
                    seen.add(callee)
                    if ACCESSOR.match(callee):
                        continue
                    p = ls + m.end() - 1
                    rp = match_pair(src, p, '(', ')')
                    args = src[p + 1:rp]
                    ids = {i for i in re.findall(r'\b' + IDENT + r'\b', args)
                           if i not in KEYWORDS and i != callee}
                    if not ids or (ids & mods):
                        continue
                    hits.append({
                        'file': rel, 'func': name,
                        'loop_line': line_of(src, o),
                        'call_line': line_of(src, p),
                        'callee': callee,
                        'args': ' '.join(args.split())[:100],
                    })
    print('calls inside loops whose arguments never change: %d candidates' % len(hits))
    for c, n in Counter(h['callee'] for h in hits).most_common(25):
        print('  %-34s %d' % (c, n))
    print('\nNOTE: a call with invariant arguments is only redundant if the')
    print('callee is a pure function of those arguments.  Read each candidate.')
    return {'invariant': hits}


COMMANDS = {
    'inventory': cmd_inventory,
    'hotpaths': cmd_hotpaths,
    'divloop': cmd_divloop,
    'invariant': cmd_invariant,
}


def main(argv):
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument('command', choices=sorted(COMMANDS) + ['all'])
    ap.add_argument('root', nargs='?', default=DEFAULT_ROOT)
    ap.add_argument('--json', metavar='OUT')
    args = ap.parse_args(argv)

    if not os.path.isdir(args.root):
        ap.error('not a directory: %s' % args.root)

    result = {}
    names = sorted(COMMANDS) if args.command == 'all' else [args.command]
    for name in names:
        if args.command == 'all':
            print('\n' + '=' * 72 + '\n== %s\n' % name.upper() + '=' * 72)
        result.update(COMMANDS[name](args.root, args))

    if args.json:
        with open(args.json, 'w') as fh:
            json.dump(result, fh, indent=1)
        print('\nwrote %s' % args.json)
    return 0


if __name__ == '__main__':
    sys.exit(main(sys.argv[1:]))
