for /f "usebackq delims=" %f in (`dir /b /s *.iso`) do "C:\Program Files\7-Zip\7z.exe" x "%f" -o"%~pf%~nf"
