On Error Resume Next
Dim docPath, pdfPath, doc2Pdf, objDoc, wdFormatPdf, state

' *********************************** '
' word转pdf
' 利用本地word导出为pdf的功能来实现
'
' 状态码说明:
' 1 没有word
' 2 没有word转pdf
' 3 打开word失败
' 0 成功
' *********************************** '

pdfPath = "E:\\新建测试doc.pdf"
docPath = "E:\\test.docx"
wdFormatPdf = 17

err.clear
Set doc2Pdf = CreateObject("Word.Application")

If err.number <> 0 Then
  state = 1
Else
  err.clear
  Set objDoc = doc2Pdf.Documents.Open(docPath)
  If err.number <> 0 Then
    state = 3
  Else
    err.clear
    objDoc.SaveAs pdfPath, wdFormatPdf
    objDoc.Close()
    doc2Pdf.Quit
    If err.number <> 0 Then
      state = 2
    Else
      state = 0
    End If
  End If
End IF

msgBox(state)
  