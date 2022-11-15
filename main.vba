Sub main()

    Dim Ligne As Long
    
        For Ligne = 1 To 50
            If Ligne Mod 2 = 0 Then Rows(Ligne).Interior.ColorIndex = 15 Else Rows(Ligne).Interior.ColorIndex = 16
        Next
    FormatPolice
    
    
    MsgBox "Finish! the document saved in document"
    
End Sub
Sub SimplePrintToPDF()

ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:="demo.pdf", Quality:=xlQualityStandard, _
  IncludeDocProperties:=False, IgnorePrintAreas:=False, OpenAfterPublish:=True

End Sub

Sub FormatPolice()
    Dim Cellule As Range
    For Each Cellule In Selection
    With Selection.Font
        .Name = "Arial"
        .Size = 12
    End With
    Next Cellule
End Sub
