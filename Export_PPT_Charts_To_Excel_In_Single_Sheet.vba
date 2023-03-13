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
    Dim unneededSlides() As Integer '不需要的幻灯片
    Dim neededSlides() As Integer '需要的幻灯片，如果有值，则不判断{unneededSlides}
    Dim flagTemp As Boolean

    ReDim unneededSlides(-1 To -1)
    ReDim neededSlides(-1 To -1)

    'flagTemp = ArrayPush(unneededSlides, 14) '第14页不需要
    'flagTemp = ArrayRangePush(unneededSlides, 1, 7) '第1到7页不需要

    On Error GoTo ErrorHandler

    Set excelApp = CreateObject("Excel.Application") '创建新的 Excel 实例
    excelApp.Visible = True '将 Excel 应用程序设置为可见，以便查看导出的结果
    Set excelWorkbook = excelApp.Workbooks.Add '打开 Excel 工作簿
    Set excelWorksheet = excelWorkbook.Sheets(1) '创建新的工作表

    chartRow = 1 '将导出的图表计数器设置为 1

    For Each pptSlide In ActivePresentation.Slides '遍历 PowerPoint 中的所有幻灯片
        If UBound(neededSlides) = -1 Then
            If IsValueInArray(pptSlide.SlideIndex, unneededSlides) Then '判断unneededSlides内是否含有当前页
                GoTo NextIteration
            End If
        Else
            If IsValueInArray(pptSlide.SlideIndex, neededSlides) Then '判断unneededSlides内是否含有当前页
                GoTo NextIteration
            End If
        End If
        For Each pptShape In pptSlide.Shapes '遍历当前幻灯片中的所有形状
            If pptShape.HasChart Then '如果当前形状包含图表
                Set chartData = pptShape.Chart '获取当前选定的 Chart 数据

                excelWorksheet.Cells(chartRow, 1).MergeArea.Value = pptSlide.SlideIndex '添加页码

                '方法一
                chartData.chartData.ActivateChartDataWindow '激活图表数据工作簿
                chartData.chartData.Workbook.Sheets(1).UsedRange.Copy '将图表数据复制到剪贴板中
                excelWorksheet.Range("B" & chartRow).PasteSpecial xlPasteValues '将图表数据粘贴到 Excel 工作表中

                '方法二
                'For rowTemp = 0 To chartData.chartData.Workbook.Sheets(1).UsedRange.Rows.Count
                '    For colTemp = 1 To chartData.chartData.Workbook.Sheets(1).UsedRange.Columns.Count
                '        excelWorksheet.Cells(rowTemp + chartRow, colTemp + 1).MergeArea.Value = chartData.chartData.Workbook.Sheets(1).Cells(rowTemp + 1, colTemp).Value
                '        '颜色
                '        'excelWorksheet.Cells(rowTemp + row, colTemp + 1).MergeArea.Interior.Color = pptShape.Table.Cell(rowTemp, colTemp).Shape.Fill.ForeColor.RGB
                '    Next colTemp
                'Next rowTemp

                '将导出的图表计数器更新为下一个图表的行数
                chartRow = chartRow + chartData.chartData.Workbook.Sheets(1).UsedRange.Find("*", LookIn:=xlValues, searchorder:=xlByRows, searchdirection:=xlPrevious).Row + 2
                chartData.chartData.Workbook.Close False '关闭图表数据工作簿，不保存更改
            End If
        Next pptShape
NextIteration:
    Next pptSlide
    excelWorksheet.Columns.AutoFit
    excelWorksheet.Rows.AutoFit

ErrorHandler:
    errorCount = errorCount + 1
    'errorMsg = "Error #" & Err.Number & " occurred" & vbCrLf & "Description: " & Err.Description & vbCrLf & "In procedure ExportChartDataToExcel, line: " & Erl
    'MsgBox errorMsg, vbExclamation + vbOKOnly, "Error"
    'Dim userResponse As VbMsgBoxResult
    'userResponse = MsgBox("Do you want to continue running the macro?", vbQuestion + vbYesNo, "Continue?")
    If errorCount < 100 Then
        Resume
    Else
        Exit Sub
    End If

End Sub

Function IsValueInArray(ByVal searchValue As Variant, ByVal arr As Variant) As Boolean
    Dim i As Long
    For i = LBound(arr) To UBound(arr)
        If arr(i) = searchValue Then
            IsValueInArray = True
            Exit Function
        End If
    Next i
    IsValueInArray = False
End Function

Function ArrayPush(ByRef arr As Variant, ByVal val As Long) As Boolean
    Dim arrLength As Long

    arrLength = UBound(arr)

    ReDim Preserve arr(arrLength + 1)
    arr(arrLength + 1) = val
End Function

Function ArrayRangePush(ByRef arr As Variant, ByVal val0 As Long, ByVal val1 As Long) As Boolean
    Dim arrLength As Long
    Dim rangeLength As Long

    arrLength = UBound(arr)
    rangeLength = val1 - val0

    ReDim Preserve arr(arrLength + 1 + rangeLength)

    For i = 0 To rangeLength
        arr(arrLength + i + 1) = val0 + i
    Next i
End Function
