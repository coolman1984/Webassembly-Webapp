# Testing an Excel pipeline: real Excel, real limits, real runtime

## Generate fixtures through the same Excel COM layer, not by hand

Hand-crafted `.xlsx` files (written with a pure-Python library, or copied once and never regenerated) can't
reproduce the actual quirks that only exist on the real automation surface: how `Find()` behaves, how
`UsedRange` gets inflated, how error cells and the 1904 date system actually round-trip, how `.xlsb`'s binary
format differs from `.xlsx`. Build fixtures by writing them out through a *second*, independent Excel COM
session (the test's own, separate from the one the code under test drives) - this is a genuine end-to-end test
of the real surface, both ways.

## Fix the meaning, fuzz the storage

Design one small, fully hand-computed "oracle" dataset (a handful of records whose expected output you work out
independently, e.g. with plain `datetime` arithmetic written separately from the code under test - never by
calling the code's own helper functions to produce the "expected" answer). Then generate many *storage*
variations of exactly that same data and assert every variation produces the identical output:

- shuffled column order
- extra, unrelated "junk" columns inserted at random positions
- noisy headers: case, extra/irregular whitespace, embedded newlines, invisible Unicode characters
- title/banner rows above the real header row
- the workbook saved under a different sheet name, or with extra hidden sheets present
- the 1904 date system
- ID/model codes with case and whitespace noise, or stored as a numeric-looking value vs. text
- the same logical file saved as both `.xlsx` and `.xlsb` (if both are accepted inputs)

A test failure here means "the real logic doesn't survive realistic messiness," which is a far more valuable
signal than "does this parse one clean file."

## Randomised combinations, not just the named cases

On top of the named storage variations, run a modest number of *randomised combinations* of them together
(pick a random seed, turn a random subset of the fuzz options on, run the same oracle assertions). This finds
interactions between fuzz dimensions that no single named test case would - e.g., a bug that only appears when
columns are shuffled *and* the workbook uses the 1904 date system at once.

## Validation tests need to prove the *right* rejection, fast

For every "this file should be rejected" test, assert both the specific, correct error message (not just "it
raised") and that it happened quickly - especially for scenarios like a password-protected file, a stale
external link, or a file already open elsewhere, where the naive implementation's failure mode is "hangs
indefinitely" rather than "fails". A slow-but-eventually-correct rejection is a regression even if the assertion
on the message text still passes, so bound the elapsed time explicitly in the test.

## Stress test to the platform's actual limit, not a round number

Test against the real hard ceiling that matters - for Excel, a sheet filled to its literal maximum of
1,048,576 rows - rather than an arbitrary "big enough" figure. Also test a size representative of the largest
real file the tool will actually see, generated once and cached on disk between test runs (regenerating a
500,000-row workbook through Excel COM takes real minutes; don't pay that cost on every run). Record wall-clock
time and peak memory (via the OS process-memory API, not an estimate) for both extraction and any subsequent
data-layer write, so there's a real number to compare against next time, not just a pass/fail.

## Test the thing the user will actually run

Run the full suite through the *bundled* runtime (`runtime\python.exe -I`), not the development interpreter -
packaging mistakes (a missing DLL, a wrong `.pth` entry, a stripped-out module the tool turns out to still need)
only show up there. Also test the service layer as an actual separate process launched the way the `.bat` file
launches it (not by importing its module in-process), so process-lifecycle bugs (port already in use, a crashed
previous run's orphan, concurrent requests) are exercised for real.

## Recognise "blocked" vs "slow" while debugging a stalled test run

If a long-running background test process seems to have stopped progressing, don't assume it's simply slow.
Check whether it's actually consuming CPU time relative to elapsed wall-clock time - a process that's used only
a few seconds of CPU over several minutes of wall time is blocked (on I/O, a lock, or - as in gotcha #5 of
`gotchas.md` - an unanswerable dialog), not computing. If the blocked process owns a GUI-capable child process
(like Excel), check whether that child has a live window handle and whether it reports itself "Responding" -
`Responding: True` with a real window handle, combined with near-zero CPU on your own process, is the specific
signature of "waiting on a dialog nothing can click," and points straight at the fix (a dedicated short
watchdog around exactly that call) rather than a longer general-purpose timeout.
