Option Explicit

Dim xlApp, xlBook, xlSheet
Dim totalAmount, filePath, sheetName, amountColumn, startRow
Dim objFSO, objFolder, objFile, fileEnum
Dim row, cellValue

' 获取命令行参数
Dim args
Set args = WScript.Arguments
If args.Count < 2 Then
    WScript.Echo "Usage: SumRows.vbs <modify_column> <start_row>"
    WScript.Quit
End If

' 初始化Excel应用程序对象
Set xlApp = CreateObject("Excel.Application")
xlApp.Visible = False ' 不显示Excel界面

' 初始化总和变量
totalAmount = 0

' 设置工作表名称和金额列（例如，B列代表第二列）
' sheetName = "Sheet1"
amountColumn =  args(0)
startRow =  args(1)

' 设置包含Excel文件的文件夹路径
Dim folderPath
folderPath = "./" ' 修改为你的文件夹路径

' 创建文件系统对象
Set objFSO = CreateObject("Scripting.FileSystemObject")

' 检查文件夹是否存在
If objFSO.FolderExists(folderPath) Then
    ' 获取文件夹中的文件枚举器
    Set fileEnum = objFSO.GetFolder(folderPath).Files
    
    ' 遍历文件夹中的每个文件
    For Each objFile In fileEnum
    
        ' 检查文件是否为Excel文件（.xlsx或.xls）
        If LCase(objFSO.GetExtensionName(objFile.Path)) = "xlsx" Or LCase(objFSO.GetExtensionName(objFile.Path)) = "xls" Or LCase(objFSO.GetExtensionName(objFile.Path)) = "csv" Then
        
    		WScript.Echo " files: " & objFile
            ' 打开Excel文件
            Set xlBook = xlApp.Workbooks.Open(objFile.Path)
            ' 选择工作表
            Set xlSheet = xlBook.Sheets(1)
            
            ' 遍历工作表中的每一行，从第二行开始（假设第一行是标题行）
            For row = startRow To xlSheet.UsedRange.Rows.Count
                ' 读取金额列的值
                cellValue = xlSheet.Cells(row, amountColumn).Value
                
                ' 检查值是否为数字，并累加到总和中
                If IsNumeric(cellValue) Then
                    totalAmount = totalAmount + CDbl(cellValue)
                End If
            Next
            
            ' 关闭工作簿，不保存更改
            xlBook.Close False
        End If
    Next
    
    ' 输出总和
    WScript.Echo "Total amount from all files: " & totalAmount
Else
    WScript.Echo "The specified folder does not exist."
End If

' 清理资源
Set xlSheet = Nothing
Set xlBook = Nothing
xlApp.Quit
Set xlApp = Nothing
Set objFSO = Nothing
Set fileEnum = Nothing
