Sub BatchConvertDocToPDF()
    Dim folderPath As String
    Dim outputPath As String
    Dim fileName As String
    Dim doc As Document

    ' Input folder (change this)
    folderPath = "path\of\word\documents\"

    ' Output folder (change this)
    outputPath = "path\where\you\want\converted\PDF\"

    fileName = Dir(folderPath & "*.doc*")

    While fileName <> ""
        Set doc = Documents.Open(folderPath & fileName)
        
        doc.ExportAsFixedFormat _
            OutputFileName:=outputPath & Replace(fileName, ".docx", ".pdf"), _
            ExportFormat:=wdExportFormatPDF
        
        doc.Close False
        fileName = Dir
    Wend

    MsgBox "Conversion Completed!"
End Sub
