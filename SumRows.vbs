Option Explicit

Dim xlApp, xlBook, xlSheet
Dim totalAmount, filePath, sheetName, amountColumn, startRow
Dim objFSO, objFolder, objFile, fileEnum
Dim row, cellValue

' ��ȡ�����в���
Dim args
Set args = WScript.Arguments
If args.Count < 2 Then
    WScript.Echo "Usage: SumRows.vbs <modify_column> <start_row>"
    WScript.Quit
End If

' ��ʼ��ExcelӦ�ó������
Set xlApp = CreateObject("Excel.Application")
xlApp.Visible = False ' ����ʾExcel����

' ��ʼ���ܺͱ���
totalAmount = 0

' ���ù��������ƺͽ���У����磬B�д���ڶ��У�
' sheetName = "Sheet1"
amountColumn =  args(0)
startRow =  args(1)

' ���ð���Excel�ļ����ļ���·��
Dim folderPath
folderPath = "./" ' �޸�Ϊ����ļ���·��

' �����ļ�ϵͳ����
Set objFSO = CreateObject("Scripting.FileSystemObject")

' ����ļ����Ƿ����
If objFSO.FolderExists(folderPath) Then
    ' ��ȡ�ļ����е��ļ�ö����
    Set fileEnum = objFSO.GetFolder(folderPath).Files
    
    ' �����ļ����е�ÿ���ļ�
    For Each objFile In fileEnum
    
        ' ����ļ��Ƿ�ΪExcel�ļ���.xlsx��.xls��
        If LCase(objFSO.GetExtensionName(objFile.Path)) = "xlsx" Or LCase(objFSO.GetExtensionName(objFile.Path)) = "xls" Or LCase(objFSO.GetExtensionName(objFile.Path)) = "csv" Then
        
    		WScript.Echo " files: " & objFile
            ' ��Excel�ļ�
            Set xlBook = xlApp.Workbooks.Open(objFile.Path)
            ' ѡ������
            Set xlSheet = xlBook.Sheets(1)
            
            ' �����������е�ÿһ�У��ӵڶ��п�ʼ�������һ���Ǳ����У�
            For row = startRow To xlSheet.UsedRange.Rows.Count
                ' ��ȡ����е�ֵ
                cellValue = xlSheet.Cells(row, amountColumn).Value
                
                ' ���ֵ�Ƿ�Ϊ���֣����ۼӵ��ܺ���
                If IsNumeric(cellValue) Then
                    totalAmount = totalAmount + CDbl(cellValue)
                End If
            Next
            
            ' �رչ����������������
            xlBook.Close False
        End If
    Next
    
    ' ����ܺ�
    WScript.Echo "Total amount from all files: " & totalAmount
Else
    WScript.Echo "The specified folder does not exist."
End If

' ������Դ
Set xlSheet = Nothing
Set xlBook = Nothing
xlApp.Quit
Set xlApp = Nothing
Set objFSO = Nothing
Set fileEnum = Nothing
