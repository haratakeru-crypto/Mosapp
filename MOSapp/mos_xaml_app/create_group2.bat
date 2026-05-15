@echo off
for /L %%i in (3,1,10) do (
    echo using System;> "Libraries\Group2\ExcelChecker2_%%i.cs"
    echo.>> "Libraries\Group2\ExcelChecker2_%%i.cs"
    echo namespace Libraries.Group2>> "Libraries\Group2\ExcelChecker2_%%i.cs"
    echo {>> "Libraries\Group2\ExcelChecker2_%%i.cs"
    echo     public class ExcelChecker2_%%i>> "Libraries\Group2\ExcelChecker2_%%i.cs"
    echo     {>> "Libraries\Group2\ExcelChecker2_%%i.cs"
    echo         public bool CheckExcel^(string filePath^)>> "Libraries\Group2\ExcelChecker2_%%i.cs"
    echo         {>> "Libraries\Group2\ExcelChecker2_%%i.cs"
    echo             return false;>> "Libraries\Group2\ExcelChecker2_%%i.cs"
    echo         }>> "Libraries\Group2\ExcelChecker2_%%i.cs"
    echo     }>> "Libraries\Group2\ExcelChecker2_%%i.cs"
    echo }>> "Libraries\Group2\ExcelChecker2_%%i.cs"
)