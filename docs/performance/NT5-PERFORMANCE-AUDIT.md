# Performance audit — recurring execution paths in the archived NT 5 era sources

| | |
| --- | --- |
| Status | In progress. One defect found and corrected; several candidates documented as unresolved |
| Scope | The Windows NT 4.0 DDK source tree in this repository, plus the neighbouring SDK/DDK media where relevant |
| Target platform | Windows NT 5 (Windows 2000) class systems, which inherit these NT 4 drivers |
| Tooling | `tools/perf-audit/` |
| Native build | **Not available.** See "What could not be established" |

This report is written so that the investigation can be resumed without
repeating the searches that produced it. Findings are numbered `PERF-NT5-nnnn`
and classified:

* **Category A** — the source establishes avoidable redundant work or an
  avoidable algorithmic cost; a safe correction can proceed with focused
  validation.
* **Category B** — the source suggests a real bottleneck but the impact needs
  profiling or workload information that is not available here.
* **Category C** — the benefit is uncertain or the change carries behavioural,
  compatibility or concurrency risk.

---

## 1. What is actually in this repository

Measured on the checkout, not from a file listing tool's summary:

| | Files | Lines |
| --- | --- | --- |
| Everything in the working tree | 35,131 | — |
| `.c` / `.cpp` | 5,270 | 3,176,803 |
| `.h` 4,109 + `.inc` 593 + `.asm` 1,139 | 5,841 | — |
| Binary payload (`.dll` 1,733, `.exe` 1,517, `.sys` 48, `.cab` 443) | ~3,700 | — |
| Printer font resources (`.pfm` 1,718, `.ifi` 1,007) | 2,725 | — |

Source is distributed across the media trees:

| Tree | `.c`/`.cpp`/header/asm files |
| --- | --- |
| Microsoft Windows NT 4.0 SDK (1996) | 4,131 |
| Windows Chicago — Build 331 SDK | 2,604 |
| **Microsoft Windows NT 4.0 DDK (4.0.1381)** | **1,438** |
| Microsoft Windows 95 DDK | 1,282 |
| Microsoft Windows 2.11 (286) OEM BAK | 705 |
| Microsoft Windows 2.03 DDK / SDK | 428 |
| others (Chicago beta, MSDN, Windows 1.x) | ~380 |

The `.iso` and `.img` archives this repository mirrors are **not present in the
checkout** (0 of each); the media has already been extracted.  Consequently
`extract_isos.bat`, `extract_imgs.bat`, `clear_isos.ps1` and `clear_imgs.ps1`
have no inputs here.

### The audited subsystem

The NT 4.0 DDK source tree
(`Microsoft Windows NT 4.0 DDK (4.0.1381)/winnt40.ddk/ddk/src`) is the largest
coherent body of kernel-mode code in the repository and the direct ancestor of
the NT 5 drivers, so it was chosen as the primary target:

| | |
| --- | --- |
| Files | 780 |
| Lines | 681,344 |
| Functions | 5,205 |

Largest subsystems by line count (from `ntperfscan.py inventory`):

| Subsystem | Files | Lines | Functions |
| --- | --- | --- | --- |
| video/displays | 109 | 87,412 | 748 |
| storage/tape | 178 | 55,817 | 271 |
| input/fareast | 57 | 47,768 | 452 |
| video/miniport | 14 | 37,630 | 129 |
| print/psprint | 34 | 26,855 | 254 |
| storage/class | 11 | 26,585 | 145 |
| storage/disk | 7 | 22,384 | 92 |
| storage/miniport | 6 | 22,190 | 123 |
| mmedia/vidcap | 37 | 20,750 | 326 |
| comm/serial | 14 | 20,414 | 119 |
| network/tdi | 20 | 17,217 | 137 |

---

## 2. Method

Because no profiler can run against this code, the audit narrows from the
outside in and treats every mechanical match as a worklist item rather than a
finding:

1. **Inventory** — `ntperfscan.py inventory` establishes subsystem boundaries
   and sizes, and the build targets (`SOURCES` / `MAKEFILE` per component).
2. **Recurring entry points** — `ntperfscan.py hotpaths` enumerates the
   callbacks that recur *by construction*: those registered with
   `IoInitializeTimer`, `DriverObject->DriverStartIo`, `IoInitializeDpcRequest`
   / `KeInitializeDpc` and `IoConnectInterrupt`. On this tree: **35 DPCs, 12
   ISRs, 10 StartIo routines, 4 timer callbacks.**
3. **Candidate generation** — detectors for the cost classes that the brief
   calls out: allocation inside loops, repeated lock acquisition, repeated
   traversal, loop-invariant work (`invariant`), and integer division or
   modulo inside loops (`divloop`).
4. **Caller and lifecycle tracing** — each candidate is traced to its callers
   to establish whether it is on a hot path, how often it runs, and whether it
   is latched or amortised. Functions containing loops were *not* assumed to
   be hot: the longest functions in the tree (`S3FindAdapter`, 76,968 bytes;
   `AtDiskDeferredProcedure`, 54,686; `ScsiCdRomStartIo`, 52,417) are
   state machines that run per request or per initialisation, not per byte.
5. **Direct reading** — the strong candidates were read in full and either
   confirmed or dismissed with a reason (section 5).
6. **Executable validation where the routine is self-contained** — see
   section 7.

### Why the mechanical detectors are not the answer

Known false positive modes, all encountered on this tree and recorded in
`tools/perf-audit/README.md`, include: divisors that are `#define`s or
`sizeof(...)` (strength-reduced like any literal); loops whose iteration cost
is set by hardware rather than the CPU (Centronics strobe, 1284 nibble
handshake, NE2000 remote-DMA wait); one-shot calibration guarded by a latch
(`ParCheckBusyDelay`, `ParPolling`); debug-only code that compiles out of free
builds; and error or teardown paths.

---

## 3. Recurring execution paths on this tree

| Entry point class | Count | Notes |
| --- | --- | --- |
| DPC routines | 35 | serial (14), parallel port (4), floppy (3 variants), AT disk, i8042, TDI, sound, video capture |
| ISRs | 12 | parallel port, AT disk, floppy (3 variants), sound, video capture, generic stat |
| StartIo routines | 10 | parallel port, AT disk, i8042, inport, keyboard/mouse class |
| `IoInitializeTimer` callbacks | 4 | `CdRomTickHandler` (1 Hz, cdrom), `AtDiskTick` (atdisk), `ParTimerRoutine` (1 Hz, parallel port), `Async2SimulatedDPC` (sample) |
| `KeSetTimer` sites | 25 | serial read/write timeout machinery, i8042 command timers, TDI request timeouts, WAVE device check |

The parallel-port and serial drivers carry by far the densest collection of
timers and DPCs, and the input drivers (i8042prt, kbdclass, mouclass) carry
the highest-frequency event load (a mouse delivers 100–200 reports per second
continuously). Both were examined; neither yielded a Category A defect
(section 5).

---

## 4. Findings

### PERF-NT5-0001 — four hardware divides per four bytes of ASCII85 print data — Category A — **FIXED**

| | |
| --- | --- |
| Subsystem | print/psprint/pscript (PostScript printer driver, `TARGETNAME=pscript`, `TARGETTYPE=GDI_DRIVER`) |
| Location | `.../ddk/src/print/psprint/pscript/FILTER.C`, `FilterConvertToASCII85()` |
| Call path | `BITBLT.C:769` and `BITBLT.C:1425` → `FilterInit(..., FilterGenerateFlags(pdev))` → `FILTER.C:199` sets `FILTER_FLAG_RLE | FILTER_FLAG_ASCII85` when the device is a PostScript Level 2 device and `PSDEVMODE_COMPRESSBMP` is set → `FilterInit` chains `FilterWriteRLE` → `FilterWriteASCII85` → `FilterConvertToASCII85` per 4 bytes. `TYPE1.C:1359` also initialises the filter with `FILTER_FLAG_ASCII85` for embedded Type 1 fonts. |
| Also reached by | `FilterGenerateFilterProc()` emits `/ASCII85Decode filter` so the printer reverses the encoding |

**Original operation**

```c
for ( i = 0, pConst = ConstAscii85;
        i < 4  ;
    i++, pConst++ ) {
    dwTmp    = dwNewNum / *pConst;      /* 32-bit integer division */
    dwNewNum -= (*pConst * dwTmp);
    *pbConvert++ = (BYTE) dwTmp + '!';
}
```

`ConstAscii85` was declared `DWORD ConstAscii85[]` — a **non-constant, writable,
externally visible array**. The divisor therefore had to be re-loaded from
memory on each iteration, which prevents the translator from replacing the
division with a multiply/shift sequence. The loop performs four 32-bit integer
divisions per four input bytes: **one hardware divide per byte of raster data**.

**Why the inefficiency is established**

* The four divisors are the fixed powers of 85 (`85^4`, `85^3`, `85^2`, `85`)
  written as literals two lines above their use in the same file, so the
  division is by a compile-time constant that the source merely obscures.
* A tree-wide search for `ConstAscii85`, `52200625` and `614125` finds this
  file only, so no sibling implementation exists to compare against.
* The array is never written after initialisation (two references in the whole
  repository: the definition at line 28 and the read at line 120), which makes
  the substitution exactly equivalent.

**Complexity / cost.** The conversion is Θ(n) either way; the defect is a
constant factor of four hardware divisions per four bytes. On the x86 parts
this code targeted, a 32-bit `div` is a multi-cycle microcoded instruction
(tens of cycles), against a few cycles for the multiply/shift sequence a
constant divisor allows. It runs for every byte of every raster image in every
print job on a PostScript device with bitmap compression enabled.

**Proposed and implemented correction.** Declare the table `static const` and
index it with constant subscripts, unrolling the loop. The arithmetic is
unchanged; the loop was unrolled because a pointer walk over a table cannot be
folded. The now-unused `pConst` and `i` locals were removed so that the
component still compiles warning-clean.

**Correctness, concurrency, compatibility.** Pure integer arithmetic, no shared
state, no IRQL or locking implications. `ConstAscii85` becomes `static const`,
so it loses external linkage — confirmed safe: no header declares it, no `.def`
exports it, and nothing else in the repository names it. The routine is called
from one place in the same translation unit, in the user-mode print path.

**Validation performed.** Exhaustive and executable; see section 7. Over all
2^32 inputs the new form is byte-for-byte identical to the old, including the
returned length and the all-zero `z` case.

**Remaining uncertainty.** The measured 1.52× is a host-compiler measurement of
the *arithmetic*, not a measurement of the NT 4.0 driver. The MSVC of the
period was not available, so whether it strength-reduces `ConstAscii85[0]` with
the same result is an expectation, not a measurement. The change is neutral if
it does not: the divisibility by a constant is visible to any conforming
translator from a `static const` array with constant subscripts, and at worst
the four memory loads that the old form already performed remain.

---

### PERF-NT5-0002 — a large transfer is materialised as N IRPs up front — Category B — not changed

| | |
| --- | --- |
| Subsystem | storage/class/class (the shared SCSI class driver) |
| Location | `.../ddk/src/storage/class/class/CLASS.C`, `ScsiClassSplitRequest()` |
| Call path | `ScsiClassReadWrite` → `ScsiClassSplitRequest` when the request exceeds the port driver's maximum transfer length |

**Observation.** `irpCount = (transferByteCount + MaximumBytes - 1) / MaximumBytes`
and the routine then loops `IoAllocateIrp(DeviceObject->StackSize, FALSE)`
`irpCount` times, issuing every partial request before allocating the next one
is needed. A 16 MB unbuffered read against a 64 KB maximum transfer length
becomes 256 IRPs allocated in one pass at dispatch level, followed by 256
outstanding child requests fanned in through `ScsiClassIoCompleteAssociated`.

**Why it is only Category B.** The batch allocation is deliberate and is
documented in the function: the comment states that asynchronous partial
transfers are an optimisation for the successful case, and the fan-in counter
in `nextIrpStack->Parameters.Others.Argument1` depends on the total being known
up front. Whether it is a bottleneck depends on the port driver's
`MaximumTransferLength` on real hardware, on non-paged pool pressure on a
memory-constrained NT 4 system, and on how often the layer is handed transfers
above the split threshold — none of which can be measured here.

**What is missing to promote or dismiss it.** The port driver's maximum
transfer length as actually negotiated; a profile of `IoAllocateIrp` cost in the
disk path; and a reading of how NT 5's storage stack changed this contract.

**Why it was not changed.** Restructuring the fan-in would alter completion
ordering, cancellation and error propagation for multi-IRP requests — exactly
the kind of behavioural risk that must not be taken without a build and a test
suite.

---

### PERF-NT5-0003 — loop-invariant `KeQueryTimeIncrement()` inside busy-wait poll loops — Category B — evidence incomplete

| | |
| --- | --- |
| Subsystem | comm (parallel-port class driver and its polling twin) |
| Locations | `comm/parclass/PARCLASS.C`: `ParInitializeDevice` L1382, `ParWriteOutData` L2190, `ParTerminateNibbleMode` L2361 and L2383, `ParEnterNibbleMode` L2446 onwards, `ParNibbleModeRead` L2565 and L2590; `comm/pollpar/PARDRVR.C`: `ParWriteOutData` L1968 |

**Observation.** Each of these sites evaluates `KeQueryTimeIncrement()` inside a
spin or poll loop whose other work is `KeQueryTickCount()`, a port read and an
indirect call. The value is compared against a deadline that was itself
computed with a single `KeQueryTimeIncrement()` call earlier in the same
function, which suggests the author already treated it as invariant. The
nibble-mode sites are the most frequent: they poll two handshake transitions
per nibble, and a nibble is four bits.

**Why it is Category B, not A.** Removing repeated work is only provably safe
if the value cannot change across the loop. `KeQueryTimeIncrement` reports the
system time increment derived from the HAL, which is expected to be fixed for
the life of the system, but the kernel-side implementation is not in this
repository and could not be checked. Nor is the share of time these loops
actually consume knowable without a running system.

**What is missing.** Confirmation from the NT 4.0 kernel/HAL sources or headers
that the time increment is written once during initialisation and never again;
and a measurement of how long the spin phases last on real hardware.

**Note that the handshake polling itself is not the defect.** IEEE 1284 nibble
mode has the peripheral driving the clock, so the host must poll; replacing
that with an interrupt-driven design would be a Category C redesign, not an
optimisation.

---

### PERF-NT5-0004 — comment and code disagree on the raised-IRQL batch size — Category C — not changed

| | |
| --- | --- |
| Location | `comm/parclass/PARCLASS.C`, `ParWriteOutData()`, immediately before the `PushSomeBytes` label |

The comment states "We never want to be at raised irql from more than
200 microseconds, so we will do no more than 100 bytes at a time", while the
code sets `loopNumber = 512`, then raises to `DISPATCH_LEVEL` and strobes that
many bytes. The choice trades system-wide dispatch latency against parallel
port throughput, and the two constants cannot be reconciled without knowing the
printer and machine this was tuned against. Recorded rather than changed.

---

### Investigated and dismissed

These were traced, read and rejected. They are recorded so a later session does
not re-open them.

| Candidate | Location | Why it is not a defect |
| --- | --- | --- |
| `ParCheckBusyDelay` calibrates with four `KeQueryPerformanceCounter` calls | `comm/parclass/PARCLASS.C` L1711–1717 | Two timing pairs (one around `GetStatus`, one bare) measure the operation against its own call overhead. The calibration is latched by `BusyDelayDetermined` and bounded by `numberOfBusyChecks = 9` per device, not per write. |
| `ParPolling` divides to size a retry budget inside loops | `comm/intpar/PARDRVR.C` L2596 | Guarded by `if (callToIsr == 1)`, so it executes once per invocation, not per iteration. |
| NE2000 copies a packet with a per-iteration DMA setup and wait | `network/NE2000/CARD.C` `CardCopyUp` (`EightBitSlot` path) | Marked "NE2000 PCMCIA CHANGE!!!" — a hardware workaround, not an oversight. The `ISR_DMA_DONE` wait is bounded by `0xFFFF` with `NdisStallExecution(4)` between polls. |
| `Sonic`/`Lance`/`IBMTOK` receive and transmit interrupt processing | `network/*/INTERRUP.C` | Single pass over the descriptor ring per interrupt; the per-packet copy is the hardware transfer itself. |
| Floppy CHS arithmetic: 21 division/modulo sites in the request loop | `storage/disk/floppy/FLOPPY.C` `FlReadWrite`, `FlFormat` (+ the two Far East copies) | Real, and the divisors (`sectorsPerTrack`, `numberOfHeads`) are variables so no strength reduction applies — but the loops advance per track, and per-iteration cost is dominated by media rotation and seek time. `currentSector/sectorsPerTrack` is needlessly recomputed on the following line; it is not worth a behavioural-risk-free-looking change that cannot be measured. |
| `I8042KeyboardIsrDpc`, `I8042MouseIsrDpc`, `InportIsrDpc`, `SerialMouseIsrDpc` divide inside a loop | `input/.../{KBD,MOU,INP,SERM}CMN.C` | The divisor is `sizeof(KEYBOARD_INPUT_DATA)` / `sizeof(MOUSE_INPUT_DATA)` — a compile-time constant, strength-reduced. These DPCs also mark the highest-frequency paths in the tree and were read in full; no avoidable work was found. |
| `IbmtokHandleSrbSsb` takes a modulo inside its command loop | `network/ibmtok/INTERRUP.C` L1332, L1398, L1467 | `MAX_COMMAND_CORRELATOR` is a `#define`; the correlator ring wraps by constant. |
| `GetTabbedLineFromBuffer` takes two modulos by a runtime variable | `print/WINPRINT/TEXT.C` L528, L593–594 | Executed once per TAB character and once per destination-buffer overflow, not per character; the surrounding loop copies one byte per iteration and the modulus is negligible beside the space-fill it sizes. |
| `vidSetDefaultPalette` divides by `(NUM_GRAY-1)` per palette entry; `HW_BuildDefaultXlate` divides by `((UINT) 0x8000 / 64)` | `mmedia/vidcap/{bravado,spigot}/...` | `NUM_GRAY` is a `#define` and `0x8000/64` is a constant expression; both are one-shot table builds, 256 and 32768 iterations once per device. |
| `bInitializeNonPlanar` / `bInitializePlanar` divide by the scan-line delta | `video/displays/*/BANK.C` | Inside the bank-walking `while (TRUE)`, which iterates once per video bank (a few hundred) at mode set. |
| `bMilEnableBrushCache` divides and takes a modulo per entry | `video/displays/*/BRUSH.C` L376–377 | Runs once per brush-cache allocation over a cache of tens of entries. |
| `VgaQueryAvailableModes` divides by loop variables | `video/miniport/{cirrus,et4000}/MODESET.C` | Mode enumeration runs at display change, not per frame. |
| `MCDCLIP.C` divides per polygon vertex during clipping | `video/displays` | The 3D software clipper; the divisions are the interpolation parameter and are intrinsic to the algorithm. |
| `FilterWriteRLE`, the stage immediately upstream of the fixed one | `print/psprint/pscript/FILTER.C` | Clean single-pass run-length encoder: the inner `while` consumes its run and the outer loop advances past it, so total work is Θ(n) with no division. |
| Allocation or free inside a loop, 52 flagged sites | tree-wide | Each was read: unload and teardown paths (`ScsiCdWriterUnload`, `SoundCleanup`, `PacketUnload`), error paths, or enumeration that legitimately allocates per element (`SoundEnumSubkeys`). No site allocates per byte or per packet. |
| Extraction and cleanup scripts | `extract_*.bat`, `clear_*.ps1` | No `.iso` or `.img` files exist in the checkout, so the loops have no inputs. `for /R` enumerates live, so a `.iso` produced by extraction inside the tree could be re-processed — but cmd.exe is not available here to confirm the semantics, and file deletion is I/O bound. Left unchanged. |

---

## 5. Changes made

| Commit | Change |
| --- | --- |
| `cea45317` | `perf: use constant divisors in ASCII85 encoding` — `FILTER.C`: `ConstAscii85` becomes `static const`, the loop is unrolled with constant subscripts, unused locals removed. 28 insertions, 10 deletions, one file. |
| `8f6ea1ce` | `test: add executable equivalence and cost harness for ASCII85 conversion` — `tools/perf-audit/verify-ascii85.c`. |
| `67212bd0` | `add: static worklist tooling for the recurring-path audit` — `tools/perf-audit/ntperfscan.py`, `README.md`. |

No other source file was modified. Nothing was changed in the SDK trees, in
media payload, or in unrelated documentation.

---

## 6. Validation performed, and what it does not prove

`tools/perf-audit/verify-ascii85.c` transcribes the pre-change and post-change
bodies of `FilterConvertToASCII85()` and validates the change without the NT
toolchain:

| Check | Result |
| --- | --- |
| Byte-identical output over **all 2^32 inputs** (`--full`) | `checked=4294967296  mismatches=0` |
| Returned length and the all-zero `z` special case, zero-optimisation on and off | identical |
| Independent base-85 decode of the emitted digits compared with the input | `roundtrip_failures=0` |
| Source fingerprint of `FILTER.C` (pre-change construct absent, post-change present) | 5 of 5 |
| Cost over 64 MiB of raster data, host compiler at `-O2` | 0.402 s → 0.265 s (1.52×) |
| Disassembly of the two forms (`objdump`) | before: `divq (%rdi,%rcx,8)` in a 4-iteration loop, 35 instructions. after: 0 `div`, 4 `imul`, 79 instructions |

**What this does not establish.**

* It does **not** establish that the NT 4.0 `pscript` driver compiles. The DDK
  builds through `$(NTMAKEENV)\makefile.def` with the MSVC of the period; that
  environment is not present, and no other component in this repository can be
  built here either — there is no CI, no test harness and no recorded baseline.
* It does **not** establish runtime behaviour. The 1.52× figure is the cost of
  the conversion routine on the host, which is evidence about the algorithm and
  about what a compiler that folds `static const` arrays with constant
  subscripts will do. It is not a driver, print-job or system measurement.
* The equivalence result is stronger than a compile check: it is exhaustive over
  the input domain, so it holds for every input the shipped function can be
  given. It still assumes the transcription in the harness matches the file,
  which the fingerprint check guards.

---

## 7. What remains unaudited

Stated plainly so that no reader mistakes this for a completed
repository-wide optimisation.

* **Roughly half the tree was read only mechanically.** The NT 4.0 DDK source
  (681,344 lines) is the only tree examined at function granularity. The NT 4.0
  SDK (4,131 source files), the Chicago/Windows 95 DDK trees and the Windows 2.x
  and 1.x material were inventoried but not audited; the SDK is largely sample
  and tool code, which makes it lower value but not worthless.
* **The `invariant` detector leaves 1,310 candidates.** They are a worklist, not
  a verdict; each needs the callee read to decide whether it is a pure function
  of its arguments.
* **49 variable-divisor loop sites remain in the `divloop` worklist.**
  Twenty-seven are the floppy CHS sites dismissed above (`FlReadWrite` and
  `FlFormat` across the three copies of the driver, 18 + 9). The remaining 22
  have been sized but not all read to conclusion: `MCDCLIP.C` 5, `BANK.C` 4,
  `MODESET.C` 4, `TEXT.C` 3, `PARDRVR.C` 2 (`ParPolling`, PERF-NT5-0003),
  `HARDWARE.C` 2 (`CalcSampleRate`), `BRUSH.C` 2.
* **The higher-level media trees** (`.dll`, `.sys`, `.exe`, `.cab` payloads,
  `.pfm`/`.ifi` printer fonts) are binaries and were not disassembled.
* **No subsystem has been measured.** Not one number in this report is a
  runtime measurement of the software under audit.

### Suggested next steps, in priority order

1. Read the remaining variable-divisor sites to closure, starting with
   `HARDWARE.C:CalcSampleRate` and `readtoc.c:ReadTableOfContents`, both of
   which divide per element of a data structure rather than per hardware event.
2. Resolve PERF-NT5-0003 by checking an NT 4.0 kernel/HAL header or source for
   the lifetime of the time increment; if it is fixed, the hoist becomes a
   Category A change across eight sites in two drivers.
3. Apply the `divloop` and `invariant` detectors to the NT 4.0 SDK `MSTOOLS`
   tree and triage what they find, which extends coverage beyond kernel mode.
4. Re-establish whether any component of this repository can be built and run
   under emulation with a period-appropriate compiler; that is the only route
   from source-level evidence to a baseline.

---

## 8. Reproduction

```
# inventory, hot entry points, and worklists
python3 tools/perf-audit/ntperfscan.py inventory
python3 tools/perf-audit/ntperfscan.py hotpaths
python3 tools/perf-audit/ntperfscan.py divloop
python3 tools/perf-audit/ntperfscan.py all --json /tmp/worklist.json

# executable validation of the change in cea45317
cc -O2 -o /tmp/verify tools/perf-audit/verify-ascii85.c
/tmp/verify --full        # ~93 s, run from the repository root
```

`ntperfscan.py` is pure Python 3 with no dependencies. It defaults to the
NT 4.0 DDK source tree; pass another root to widen the audit.
