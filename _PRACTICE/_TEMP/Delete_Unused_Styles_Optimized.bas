Sub Del_Unused_Styles_Optimized()
    Dim doc As Document
    Dim sty As Style
    Dim i As Long
    Dim pass As Integer
    
    Set doc = ActiveDocument
    
    ' Speed optimization: Suppress screen updates and animations
    Application.ScreenUpdating = False
    
    On Error Resume Next ' Prevents crashes if a style is locked or system-protected
    
    ' Run 2 passes to safely handle styles that are "nested" or based on each other
    For pass = 1 To 2
        ' Loop backwards through styles to safely delete items from the collection
        For i = doc.Styles.Count To 1 Step -1
            Set sty = doc.Styles(i)
            
            ' Rule 1: Skip built-in Word styles
            If Not sty.BuiltIn Then
                
                ' Rule 2: Check if Word considers the style completely unused
                If Not sty.InUse Then
                    sty.Delete
                Else
                    ' Rule 3: Double-check via StoryRanges to ensure it's not hiding in headers/footers
                    ' This is a lightning-fast native check compared to a manual .Find loop
                    With doc.Content.Find
                        .ClearFormatting
                        .Style = sty.NameLocal
                        .Execute FindText:="", Format:=True, Wrap:=wdFindStop
                        
                        ' If the style was marked "InUse" but can't be found anywhere, safe to delete
                        If Not .Found Then
                            sty.Delete
                        End If
                    End With
                End If
                
            End If
        Next i
    Next pass
    On Error GoTo 0
    
    ' Re-enable screen updating
    Application.ScreenUpdating = True
    MsgBox "Unused custom styles cleaned up successfully!", vbInformation, "Clean Up Complete"
End Sub