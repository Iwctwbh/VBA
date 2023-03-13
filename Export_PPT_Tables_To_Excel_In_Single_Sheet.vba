'For Grape
Sub Export_PPT_Tables_To_Excel_In_Single_Sheet()
    Dim pptSlide As Slide '创建 PowerPoint 幻灯片对象
    Dim pptShape As Shape '创建 PowerPoint 形状对象
    Dim excelApp As Object '创建 Excel 应用程序对象
    Dim excelWorkbook As Object '创建 Excel 工作簿对象
    Dim excelWorksheet As Object '创建 Excel 工作表对象
    Dim row As Long
    Dim rowTemp As Long
    Dim col As Long
    Dim flag As Boolean
    Dim unneededSlides() As Integer '不需要的幻灯片
    Dim neededSlides() As Integer '需要的幻灯片，如果有值，则不判断{unneededSlides}
    Dim flagTemp As Boolean

    ReDim unneededSlides(-1 To -1)
    ReDim neededSlides(-1 To -1)

    'flagTemp = ArrayPush(unneededSlides, 14) '第14页不需要
    'flagTemp = ArrayRangePush(unneededSlides, 1, 7) '第1到7页不需要

    row = 0

    '创建一个新的 Excel 应用程序
    Set excelApp = CreateObject("Excel.Application")
    excelApp.Visible = True '可以看到导出的结果

    '创建一个新的 Excel 工作簿
    Set excelWorkbook = excelApp.Workbooks.Add
    '获取第一个 Excel 工作表
    Set excelWorksheet = excelWorkbook.Sheets(1)

    '循环遍历每个幻灯片
    For Each pptSlide In ActivePresentation.Slides
        If UBound(neededSlides) = -1 Then
            If IsValueInArray(pptSlide.SlideIndex, unneededSlides) Then '判断unneededSlides内是否含有当前页
                GoTo NextIteration
            End If
        Else
            If IsValueInArray(pptSlide.SlideIndex, neededSlides) Then '判断unneededSlides内是否含有当前页
                GoTo NextIteration
            End If
        End If

        '循环遍历每个窗格
        For Each pptShape In pptSlide.Shapes
            If pptShape.HasTable Then
                rowTemp = 1

                excelWorksheet.Cells(rowTemp + row, 1).Value = "Slide" & CStr(pptSlide.SlideIndex)

                '判断原本表格中是否有换行符，有则保留并合并相应单元格
                flag = False
                If InStr(1, pptShape.Table.Cell(rowTemp, 1).Shape.TextFrame.TextRange.Text, Chr(11)) > 0 Then
                    flag = True
                End If
                If flag Then
                    For colTemp = 1 To pptShape.Table.Columns.Count
                        excelWorksheet.Range(excelWorksheet.Cells(rowTemp + row, colTemp + 1), excelWorksheet.Cells(rowTemp + row + 1, colTemp + 1)).Merge
                    Next colTemp
                    row = row + 1
                End If

                '将表格数据复制到 Excel 工作表中
                For rowTemp = 1 To pptShape.Table.Rows.Count
                    For colTemp = 1 To pptShape.Table.Columns.Count
                        excelWorksheet.Cells(rowTemp + row, colTemp + 1).MergeArea.Value = pptShape.Table.Cell(rowTemp, colTemp).Shape.TextFrame.TextRange.Text
                        '颜色
                        excelWorksheet.Cells(rowTemp + row, colTemp + 1).MergeArea.Interior.Color = pptShape.Table.Cell(rowTemp, colTemp).Shape.Fill.ForeColor.RGB
                    Next colTemp
                Next rowTemp
                '累加每个Table的行数，末尾为空行数
                row = row + pptShape.Table.Rows.Count + 2
            End If
        Next pptShape
NextIteration:
    Next pptSlide
    '列宽自适应
    excelWorksheet.Columns.AutoFit
    excelWorksheet.Rows.AutoFit
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
