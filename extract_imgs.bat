@echo off
for /R %%f in (*.img) do (
    "C:\Program Files\7-Zip\7z.exe" x "%%f" -o"%%~dpf%%~nf"
)
pause
