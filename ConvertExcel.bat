@echo off
title Convert Excel Tool
set CheckReferencePath=dotnet ConvertExcel\ConvertExcel.dll
set ExcelFolderPath=%cd%

%CheckReferencePath% %ExcelFolderPath%

pause