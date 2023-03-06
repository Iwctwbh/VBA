Sub ExportTablesToExcel()
    Dim pptSlide As Slide
    Dim pptShape As Shape
    Dim excelApp As Object
    Dim excelWorkbook As Object
    Dim excelWorksheet As Object
    Dim row As Long
    Dim rowTemp As Long
    Dim col As Long

    row = 0

    '创建一个新的 Excel 应用程序
    Set excelApp = CreateObject("Excel.Application")
    excelApp.Visible = True '可以看到导出的结果

    '创建一个新的 Excel 工作簿
    Set excelWorkbook = excelApp.Workbooks.Add
    '创建一个新的 Excel 工作表
    Set excelWorksheet = excelWorkbook.Sheets.Add

    '循环遍历每个幻灯片
    For Each pptSlide In ActivePresentation.Slides
        '循环遍历每个表格
        For Each pptShape In pptSlide.Shapes
            If pptShape.HasTable Then
                rowTemp = 1

                excelWorksheet.Cells(rowTemp + row, 1).Value = "Slide" & CStr(pptSlide.SlideIndex)

                '将表格数据复制到 Excel 工作表中
                For rowTemp = 1 To pptShape.Table.Rows.Count
                    For col = 2 To pptShape.Table.Columns.Count
                        excelWorksheet.Cells(rowTemp + row, col).Value = pptShape.Table.Cell(rowTemp, col).Shape.TextFrame.TextRange.Text
                    Next col
                Next rowTemp
                row = row + pptShape.Table.Rows.Count + 1
            End If
        Next pptShape
    Next pptSlide
End Sub

