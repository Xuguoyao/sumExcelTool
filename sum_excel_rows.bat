@echo off
:: 提示用户输入参数
set /p MODIFY_COLUMN=请输入要修改的列字母（如A, B, C等）: 
set /p START_ROW=请输入开始的行号(如1,2,3等): 

:: 调用cscript并传递参数给VBScript脚本
cscript //NoLogo "./SumRows.vbs" "%MODIFY_COLUMN%" "%START_ROW%"

:: 可选：暂停批处理文件，以便查看脚本输出
pause
