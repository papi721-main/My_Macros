Attribute VB_Name = "Module1"
Sub ChangeContinuousBreakToNextPageBreak()
  Dim nSectionNum As Integer
  Dim objDoc As Document
 
  Application.ScreenUpdating = False
 
  Set objDoc = ActiveDocument
  nSectionNum = objDoc.Sections.Count

  For nSectionNum = nSectionNum To 1 Step -1
    With objDoc.Sections(nSectionNum).Range.PageSetup
      If .SectionStart = wdSectionContinuous Then
        .SectionStart = wdSectionNewPage
      End If
    End With
  Next nSectionNum
 
  Application.ScreenUpdating = True
End Sub
