REM %1 "$(ProjectDir)"
REM %2 "$(TargetDir)"

pushd .

-----------------------------------------------------------------------------------------------
REM copy the workbook and data file the target directory

copy "%~1..\Forbes.Models\Forbes.csv" "%~2*.*"
copy "%~1..\Forbes.Views\Forbes2000.xlsx" "%~2*.*"

popd