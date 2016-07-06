pushd "%~dp0"

set addin="ExcelMvc.Addin.xll"
if exist "C:\Program Files (x86)\." (

if exist "C:\Program Files\Microsoft Office\Office15\." (
	set addin="ExcelMvc.Addin (x64).xll"
))

START EXCEL /x %addin% "Forbes2000.xlsx"

popd
