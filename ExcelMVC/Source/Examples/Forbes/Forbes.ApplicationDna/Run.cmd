pushd "%~dp0"

set addin="Forbes.ApplicationDna.xll"

if exist "C:\Program Files (x86)\." (

if exist "C:\Program Files\Microsoft Office\Office15\Excel.Exe" (
  set addin="Forbes.ApplicationDna (x64).xll"

))

START EXCEL /x %addin% "Forbes2000.xlsx"
popd
