'For Grape
Sub Export_PPT_Charts_To_Excel_In_Single_Sheet()
    '声明变量
    Dim excelApp As Excel.Application '创建 Excel 应用程序对象
    Dim excelWorkbook As Excel.Workbook '创建 Excel 工作簿对象
    Dim excelWorksheet As Excel.Worksheet '创建 Excel 工作表对象
    Dim chartData As PowerPoint.Chart '创建 PowerPoint 图表对象
    Dim pptShape As Shape '创建 PowerPoint 形状对象
    Dim pptSlide As PowerPoint.Slide '创建 PowerPoint 幻灯片对象
    Dim chartRow As Long '用于计数导出的图表数量
    Dim errorMsg As String '错误信息
    Dim errorCount As Long '错误数量

    On Error GoTo ErrorHandler

    Set excelApp = CreateObject("Excel.Application") '创建新的 Excel 实例
    excelApp.Visible = True '将 Excel 应用程序设置为可见，以便查看导出的结果
    Set excelWorkbook = excelApp.Workbooks.Add '打开 Excel 工作簿
    Set excelWorksheet = excelWorkbook.Sheets(1) '创建新的工作表

    chartRow = 1 '将导出的图表计数器设置为 1

    For Each pptSlide In ActivePresentation.Slides '遍历 PowerPoint 中的所有幻灯片
        For Each pptShape In pptSlide.Shapes '遍历当前幻灯片中的所有形状
            If pptShape.HasChart Then '如果当前形状包含图表
                Set chartData = pptShape.Chart '获取当前选定的 Chart 数据
                chartData.chartData.Activate '激活图表数据工作簿
                chartData.chartData.Workbook.Sheets(1).UsedRange.Copy '将图表数据复制到剪贴板中
                excelWorksheet.Range("A" & chartRow).PasteSpecial xlPasteValues '将图表数据粘贴到 Excel 工作表中
                '将导出的图表计数器更新为下一个图表的行数
                chartRow = chartRow + chartData.chartData.Workbook.Sheets(1).UsedRange.Find("*", LookIn:=xlValues, searchorder:=xlByRows, searchdirection:=xlPrevious).Row + 2
                chartData.chartData.Workbook.Close False '关闭图表数据工作簿，不保存更改
            End If
        Next pptShape
    Next pptSlide

    excelWorksheet.Columns.AutoFit
    excelWorksheet.Rows.AutoFit

ErrorHandler:
    errorCount = errorCount + 1
    'errorMsg = "Error #" & Err.Number & " occurred" & vbCrLf & "Description: " & Err.Description & vbCrLf & "In procedure ExportChartDataToExcel, line: " & Erl
    'MsgBox errorMsg, vbExclamation + vbOKOnly, "Error"
    'Dim userResponse As VbMsgBoxResult
    'userResponse = MsgBox("Do you want to continue running the macro?", vbQuestion + vbYesNo, "Continue?")
    If errorCount < 10 Then
        Resume
    Else
        Exit Sub
    End If
End Sub