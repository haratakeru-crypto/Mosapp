@echo off
for /L %%i in (1,1,10) do (
    echo using System;> "Libraries\Group3\ExcelChecker3_%%i.cs"
    echo.>> "Libraries\Group3\ExcelChecker3_%%i.cs"
    echo namespace Libraries.Group3>> "Libraries\Group3\ExcelChecker3_%%i.cs"
    echo {>> "Libraries\Group3\ExcelChecker3_%%i.cs"
    echo     public class ExcelChecker3_%%i>> "Libraries\Group3\ExcelChecker3_%%i.cs"
    echo     {>> "Libraries\Group3\ExcelChecker3_%%i.cs"
    echo         public bool CheckExcel^(string filePath^)>> "Libraries\Group3\ExcelChecker3_%%i.cs"
    echo         {>> "Libraries\Group3\ExcelChecker3_%%i.cs"
    echo             return false;>> "Libraries\Group3\ExcelChecker3_%%i.cs"
    echo         }>> "Libraries\Group3\ExcelChecker3_%%i.cs"
    echo     }>> "Libraries\Group3\ExcelChecker3_%%i.cs"
    echo }>> "Libraries\Group3\ExcelChecker3_%%i.cs"
)