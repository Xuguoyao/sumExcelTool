@echo off
:: ��ʾ�û��������
set /p MODIFY_COLUMN=������Ҫ�޸ĵ�����ĸ����A, B, C�ȣ�: 
set /p START_ROW=�����뿪ʼ���к�(��1,2,3��): 

:: ����cscript�����ݲ�����VBScript�ű�
cscript //NoLogo "./SumRows.vbs" "%MODIFY_COLUMN%" "%START_ROW%"

:: ��ѡ����ͣ�������ļ����Ա�鿴�ű����
pause
