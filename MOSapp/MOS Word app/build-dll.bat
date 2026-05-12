@echo off
echo Building 30 DLL files...


REM Build Group1 DLLs
for /L %%i in (1,1,10) do (
    echo Building ExcelChecker1_%%i.dll...
    "C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\Roslyn\csc.exe" ^
     /target:library /out:Libraries\Group1\ExcelChecker1_%%i.dll ^
     /reference:System.dll /reference:System.Core.dll ^
     /reference:System.IO.dll ^
     /reference:System.Runtime.InteropServices.dll ^
     /reference:"C:\Windows\assembly\GAC_MSIL\Microsoft.Office.Interop.Excel\15.0.0.0__71e9bce111e9429c\Microsoft.Office.Interop.Excel.dll" ^
     /langversion:7.3 ^
     Libraries\Group1\ExcelChecker1_%%i.cs
     
    if errorlevel 1 (
        echo Error building ExcelChecker1_%%i.dll
    ) else (
        echo Success: ExcelChecker1_%%i.dll
        if not exist "bin\Debug\Libraries\Group1" mkdir "bin\Debug\Libraries\Group1"
        copy "Libraries\Group1\ExcelChecker1_%%i.dll" "bin\Debug\Libraries\Group1\"
    )
)

REM Build Group2 DLLs
for /L %%i in (1,1,10) do (
    echo Building ExcelChecker2_%%i.dll...
    "C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\Roslyn\csc.exe" ^
     /target:library /out:Libraries\Group2\ExcelChecker2_%%i.dll ^
     /reference:System.dll /reference:System.Core.dll ^
     /reference:System.IO.dll ^
     /reference:System.Runtime.InteropServices.dll ^
     /reference:"C:\Windows\assembly\GAC_MSIL\Microsoft.Office.Interop.Excel\15.0.0.0__71e9bce111e9429c\Microsoft.Office.Interop.Excel.dll" ^
     /langversion:7.3 ^
     Libraries\Group2\ExcelChecker2_%%i.cs
     
    if errorlevel 1 (
        echo Error building ExcelChecker2_%%i.dll
    ) else (
        echo Success: ExcelChecker2_%%i.dll
        if not exist "bin\Debug\Libraries\Group2" mkdir "bin\Debug\Libraries\Group2"
        copy "Libraries\Group2\ExcelChecker2_%%i.dll" "bin\Debug\Libraries\Group2\"
    )
)


REM Build Group3 DLLs
for /L %%i in (1,1,10) do (
    echo Building ExcelChecker3_%%i.dll...
    "C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\Roslyn\csc.exe" ^
     /target:library /out:Libraries\Group3\ExcelChecker3_%%i.dll ^
     /reference:System.dll /reference:System.Core.dll ^
     /reference:System.IO.dll ^
     /reference:System.Runtime.InteropServices.dll ^
     /reference:"C:\Windows\assembly\GAC_MSIL\Microsoft.Office.Interop.Excel\15.0.0.0__71e9bce111e9429c\Microsoft.Office.Interop.Excel.dll" ^
     /langversion:7.3 ^
     Libraries\Group3\ExcelChecker3_%%i.cs
     
    if errorlevel 1 (
        echo Error building ExcelChecker3_%%i.dll
    ) else (
        echo Success: ExcelChecker3_%%i.dll
        if not exist "bin\Debug\Libraries\Group3" mkdir "bin\Debug\Libraries\Group3"
        copy "Libraries\Group3\ExcelChecker3_%%i.dll" "bin\Debug\Libraries\Group3\"
    )
)

echo All DLLs built successfully!
pause