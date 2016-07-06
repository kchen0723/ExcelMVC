@echo off
if not exist bin\Release\net35\*.nupkg goto error
for /f "tokens=*" %%a in ('dir bin\Release\net35\*.nupkg /B /O:N ^| findstr /v /i "symbols\.nupkg$"') do set recentPackage=%%a
echo Most recent nupkg file in bin\Release\net35 directory:
echo %recentPackage%

for /f "tokens=*" %%a in ('dir bin\Release\net35\*.nupkg /B /O:N') do set recentSymbolsPackage=%%a
echo Most recent nupkg symbols file in bin\Release\net35 directory:
echo %recentSymbolsPackage%

set /p apikey=Enter apikey on nuget.org:
..\packages\NuGet.CommandLine.2.8.0\tools\NuGet.exe push bin\Release\net35\%recentPackage% -apikey %apikey%
..\packages\NuGet.CommandLine.2.8.0\tools\NuGet.exe push -source http://nuget.gw.symbolsource.org/Public/NuGet bin\Release\net35\%recentSymbolsPackage% -apikey %apikey%
goto end
:error
echo Error, no nupkg file has been found in output directory. Aborted.
:end
pause
