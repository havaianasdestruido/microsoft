@echo off
rem ⚡ Bolt: Appended '> nul' to suppress slow console I/O during extraction
for /R %%f in (*.img) do (
    "C:\Program Files\7-Zip\7z.exe" x "%%f" -o"%%~dpf%%~nf" > nul
)
pause
