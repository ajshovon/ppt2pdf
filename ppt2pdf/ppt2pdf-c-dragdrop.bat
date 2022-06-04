@echo off

title ppt2pdf
echo Working...
echo Please wait...
set files=%*

REM Set the correct path to your ppt2pdf.py here
python "C:\utils\ppt2pdf\ppt2pdf.py" -c %files%
