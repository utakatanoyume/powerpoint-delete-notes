Attribute VB_Name = "DeleteAllNotes"
Sub DeleteAllNotes()
    Dim sld As Slide
    Dim shp As Shape

    For Each sld In ActivePresentation.Slides
        For Each shp In sld.NotesPage.Shapes
            If shp.Type = msoPlaceholder Then
                If shp.PlaceholderFormat.Type = ppPlaceholderBody Then
                    shp.TextFrame.TextRange.Text = ""
                End If
            End If
        Next shp
    Next sld

    MsgBox "すべてのノートを削除しました。"
End Sub
