@echo off
rem Installs the Clip Chopper Labels panel into Premiere Pro (Windows).
set EXT=%APPDATA%\Adobe\CEP\extensions\com.daksh.cliplabels
mkdir "%EXT%\CSXS" "%EXT%\jsx" 2>nul
copy /Y "%~dp0premiere_panel\CSXS\manifest.xml" "%EXT%\CSXS\" >nul
copy /Y "%~dp0premiere_panel\index.html" "%EXT%\" >nul
copy /Y "%~dp0add_labels.jsx" "%EXT%\jsx\" >nul
for %%v in (10 11 12 13 14) do reg add HKCU\SOFTWARE\Adobe\CSXS.%%v /v PlayerDebugMode /t REG_SZ /d 1 /f >nul
echo Installed. Restart Premiere, then: Window ^> Extensions (Legacy) ^> Clip Chopper Labels
pause
