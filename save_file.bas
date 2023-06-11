Attribute VB_Name = "NewMacros"
Sub save_as_pdf()
Attribute save_as_pdf.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.save_as_pdf"
'
' Save the file as a pdf
'
    ActiveDocument.ExportAsFixedFormat OutputFileName:= _
        "C:\file_path\file_name.pdf" _
        , ExportFormat:=wdExportFormatPDF, OpenAfterExport:=True, OptimizeFor:= _
        wdExportOptimizeForPrint, Range:=wdExportAllDocument, From:=1, To:=1, _
        Item:=wdExportDocumentContent, IncludeDocProps:=True, KeepIRM:=True, _
        CreateBookmarks:=wdExportCreateNoBookmarks, DocStructureTags:=True, _
        BitmapMissingFonts:=True, UseISO19005_1:=False
    ChangeFileOpenDirectory _
        "C:\file_path\"
End Sub
Sub save_to_portfolio()
Attribute save_to_portfolio.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.save_to_portfolio"
'
' Save the file with a different name as a copy AND as a pdf
'
    ChangeFileOpenDirectory _
        "C:\file_path\"
    ActiveDocument.SaveAs2 FileName:= _
        "file_path\file_name.docx" _
        , FileFormat:=wdFormatXMLDocument, LockComments:=False, Password:="", _
        AddToRecentFiles:=True, WritePassword:="", ReadOnlyRecommended:=False, _
        EmbedTrueTypeFonts:=False, SaveNativePictureFormat:=False, SaveFormsData _
        :=False, SaveAsAOCELetter:=False, CompatibilityMode:=15
    ActiveDocument.ExportAsFixedFormat OutputFileName:= _
        "C:\file_path\file_name.pdf" _
        , ExportFormat:=wdExportFormatPDF, OpenAfterExport:=True, OptimizeFor:= _
        wdExportOptimizeForPrint, Range:=wdExportAllDocument, From:=1, To:=1, _
        Item:=wdExportDocumentContent, IncludeDocProps:=True, KeepIRM:=True, _
        CreateBookmarks:=wdExportCreateNoBookmarks, DocStructureTags:=True, _
        BitmapMissingFonts:=True, UseISO19005_1:=False
End Sub
