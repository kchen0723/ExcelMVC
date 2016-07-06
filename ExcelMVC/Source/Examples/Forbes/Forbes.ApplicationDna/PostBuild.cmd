REM %1 "$(ProjectDir)"
REM %2 "$(TargetDir)"
REM %3 "pack" to DNA pack

pushd .

REM -----------------------------------------------------------------------------------------------
REM copy the workbook and data file the target directory

copy "%~1..\Forbes.Models\Forbes.csv" "%~2*.*"
copy "%~1..\Forbes.Views\Forbes2000.xlsx" "%~2*.*"

REM -----------------------------------------------------------------------------------------------
REM copy DNA pack files to the target directory

copy "%~1..\packages\Excel-DNA.0.32.0\tools\ExcelDnaPack.exe" "%~2"
copy "%~1..\packages\Excel-DNA.0.32.0\tools\ExcelDna.Integration.dll" "%~2"
copy "%~1..\packages\Excel-DNA.0.32.0\tools\ExcelDna.xll" "%~2Forbes.ApplicationDna.xll"
copy "%~1..\packages\Excel-DNA.0.32.0\tools\ExcelDna64.xll" "%~2Forbes.ApplicationDna (x64).xll"

copy "Forbes.ApplicationDna.dll.config" "Forbes.ApplicationDna.xll.config"
copy "Forbes.ApplicationDna.dll.config" "Forbes.ApplicationDna (x64).xll.config"
copy "Forbes.ApplicationDna.dna" "Forbes.ApplicationDna (x64).dna"

if "%~3" == "" (
	goto :eof
)

REM -----------------------------------------------------------------------------------------------
REM pack x86 addin

ExcelDnaPack.exe "Forbes.ApplicationDna.dna" /Y
del "Forbes.ApplicationDna.xll"
rename "Forbes.ApplicationDna-packed.xll" "Forbes.ApplicationDna.xll"

REM -----------------------------------------------------------------------------------------------
REM pack x64 addin

ExcelDnaPack.exe "Forbes.ApplicationDna (x64).dna" /Y
del "Forbes.ApplicationDna (x64).xll"
rename "Forbes.ApplicationDna (x64)-packed.xll" "Forbes.ApplicationDna (x64).xll"

REM -----------------------------------------------------------------------------------------------
REM clean up unwanted files
del "*.dna"
del "*.exe"
del "*.pdb"
del "*.dll"
del "*.xml"
del "*.dll.config"

popd