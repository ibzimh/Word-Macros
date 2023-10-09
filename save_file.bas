Function TestDate() As String
   Dim dtToday As String
   dtToday = Format(Now(), "ddmmyy hhmm")
   TestDate = dtToday
End Function

Sub save_to_portfolio()
'
' save_to_portfolio Macro
'
'
    dateString = TestDate
    outputPath = "C:\Users\path_to_folder" & dateString & " file_name.pdf"
    
    ActiveDocument.ExportAsFixedFormat OutputFileName:= _
        outputPath _
        , ExportFormat:=wdExportFormatPDF, OpenAfterExport:=True, OptimizeFor:= _
        wdExportOptimizeForPrint, Range:=wdExportAllDocument, From:=1, To:=1, _
        Item:=wdExportDocumentContent, IncludeDocProps:=True, KeepIRM:=True, _
        CreateBookmarks:=wdExportCreateNoBookmarks, DocStructureTags:=True, _
        BitmapMissingFonts:=True, UseISO19005_1:=False
    ChangeFileOpenDirectory _
        "C:\Users\path_to_folder\"
End Sub
Sub save_as_pdf()
'
' save_as_pdf Macro
'
'
    ChangeFileOpenDirectory _
        "C:\Users\path_to_folder\"
    ActiveDocument.ExportAsFixedFormat OutputFileName:= _
        "C:\Users\path_to_file" _
        , ExportFormat:=wdExportFormatPDF, OpenAfterExport:=True, OptimizeFor:= _
        wdExportOptimizeForPrint, Range:=wdExportAllDocument, From:=1, To:=1, _
        Item:=wdExportDocumentContent, IncludeDocProps:=True, KeepIRM:=True, _
        CreateBookmarks:=wdExportCreateNoBookmarks, DocStructureTags:=True, _
        BitmapMissingFonts:=True, UseISO19005_1:=False
End Sub
