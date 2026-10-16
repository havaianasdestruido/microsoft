# Performance audit tooling

This directory supports a source-level performance audit of the archived
Microsoft SDK/DDK sources that make up this repository.  The media in this
repository target Windows NT 4 / Windows NT 5 (Windows 2000) class systems.

## Why tooling rather than a profiler

The native build environment is not present:

* the DDK components build through `$(NTMAKEENV)\makefile.def` with the
  Microsoft C/C++ compiler of the period, which is not available here;
* the tree is an archive of installation media (ISO contents plus extracted
  binaries, help files and `.pfm`/`.ifi` printer font resources), so there is
  no running NT system to profile;
* the checkout has no CI, no test harness and no recorded performance baseline.

So the audit is done by narrowing: inventory the subsystems, enumerate the
entry points that recur, and then read the strongest candidates by hand.  A
match from any detector below is a *worklist item*, never a finding.  Every
entry in `docs/performance/NT5-PERFORMANCE-AUDIT.md` was confirmed by reading
the code and, where the routine is self-contained, by executing an extracted
copy on the host.

## ntperfscan.py

Pure Python 3, no dependencies.  The default root is the NT 4.0 DDK source
tree in this repository.

```
python3 tools/perf-audit/ntperfscan.py inventory
python3 tools/perf-audit/ntperfscan.py hotpaths
python3 tools/perf-audit/ntperfscan.py divloop
python3 tools/perf-audit/ntperfscan.py invariant
python3 tools/perf-audit/ntperfscan.py all --json /tmp/worklist.json
```

| command | what it produces |
| --- | --- |
| `inventory` | per-subsystem file, line and function counts |
| `hotpaths` | the callbacks registered with `IoInitializeTimer`, `DriverStartIo`, `IoInitializeDpcRequest`/`KeInitializeDpc` and `IoConnectInterrupt` — the entry points that recur by construction |
| `divloop` | integer division or modulo inside a loop body, split into constant divisors (folded by the compiler, not a defect) and variable or struct-member divisors |
| `invariant` | calls inside a loop whose argument identifiers are never assigned in that loop |

All commands share one lexer.  Two properties of the input forced its design:

1. **Preprocessor lines must be blanked before brace/paren matching.**  A
   directive such as `#define PAR_PAPER_EMPTY(Status)` contains parentheses
   that are not part of the C block structure; leaving them in corrupts every
   nested match that follows, which is how an earlier version of this tool
   silently attributed whole functions to the wrong file region.
2. **Function headers come in two shapes.**  Kernel drivers put the return
   type on its own line (`ULONG` / `DrvDitherColor(`) while the video and
   user-mode components put it on the same line (`ULONG DrvDitherColor(`).
   Matching only the first shape found 8 functions in a 87,000-line
   directory; matching both found 5,205 functions tree-wide.

Offsets are preserved through both transformations so reported line numbers
are accurate.

## Known false positive modes

These were all observed on this tree.  They are the reason a detector hit is
not a finding.

* **Divisor is a `#define` or `sizeof`.**  `% MAX_COMMAND_CORRELATOR`,
  `% FIFOSIZE`, `/ sizeof(KEYBOARD_INPUT_DATA)` are compile-time constants and
  are strength-reduced to multiply/shift like any literal.  `divloop`
  classifies these separately and they are not defects.  An earlier version of
  the detector reported `sizeof(...)` divisors as variable because it only
  recognised numeric literals.
* **Loop where the hardware, not the CPU, sets the iteration cost.**  A
  Centronics strobe loop, a 1284 nibble handshake, an NE2000 remote-DMA wait
  or a floppy recalibration loop is dominated by port I/O or media time.
  Removing instructions from such a loop changes nothing measurable.
* **One-shot calibration.**  `ParCheckBusyDelay` and `ParPolling` divide to
  size a retry budget exactly once per device or per invocation, guarded by a
  latch.  The division is not per iteration even though it appears inside one.
* **Debug-only code.**  `DebugPrint`, `DbgPrint`, `ParDump`, `I8xPrint`,
  `FloppyDump`, `IF_LOG` and `ASSERT` compile out of free builds.  Their
  arguments are excluded from the `invariant` count for that reason.
* **Error and teardown paths.**  A loop that frees a list of resources, or
  that runs only after a failure, is not a hot path however expensive its body
  looks.
* **Chained operators.**  `currentSector/sectorsPerTrack/numberOfHeads`
  is reported as one operand; the tool does not model C operator precedence.
  Divide-by-constant that the reader can see is constant may still appear in
  the variable list — `readtoc.c`'s `( 60 * 75 )` was such a case before
  constant expressions were folded.

## verify-ascii85.c

An executable check for the one change made to the driver sources, described
in `docs/performance/NT5-PERFORMANCE-AUDIT.md` finding `PERF-NT5-0001`.  It
transcribes the pre-change and post-change bodies of
`FilterConvertToASCII85()` from the PostScript driver's `FILTER.C` and:

1. compares them over every one of the 2^32 inputs (`--full`) or a large
   sample (default), including the all-zero `z` case with the zero
   optimisation both on and off;
2. decodes the emitted digits independently and compares against the input;
3. reads `FILTER.C` and asserts the pre-change construct is gone and the
   post-change construct is present, so the harness cannot drift from the file
   it validates;
4. times both forms over a page-raster sized buffer.

```
cc -O2 -o /tmp/verify tools/perf-audit/verify-ascii85.c
/tmp/verify --full          # run from the repository root
```

This validates the arithmetic and the cost of the conversion.  It does **not**
validate that the NT 4.0 driver builds or runs: that requires the original
toolchain and an NT 4/NT 5 target.
