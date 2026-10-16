@echo off
rem ⚡ Bolt: Replaced 'dir /b /s' with native 'for /R' loop to avoid spawning a subshell
rem and corrected loop variable syntax from %f to %%f for batch file execution.
rem ⚡ Bolt: Appended '> nul' to suppress slow console I/O during extraction
for /R %%f in (*.iso) do (
    "C:\Program Files\7-Zip\7z.exe" x "%%f" -o"%%~dpf%%~nf" > nul
)
