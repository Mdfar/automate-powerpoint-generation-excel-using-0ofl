Attribute VB_Name = "PPT_Orchestrator" ' Staqlt Corporate Automation - Excel to PowerPoint Engine ' Requires reference to: Microsoft PowerPoint 16.0 Object Library

Public Sub GenerateEntityReports() Dim pptApp As Object Dim pptPres As Object Dim pptSlide As Object Dim wsData As Worksheet Dim lastRow As Long, i As Long Dim templatePath As String

Set wsData = ThisWorkbook.Sheets("EntityData")
lastRow = wsData.Cells(wsData.Rows.Count, "A").End(xlUp).Row
templatePath = ThisWorkbook.Path & "\Template.potx"

' Initialize PowerPoint
On Error Resume Next
Set pptApp = GetObject(class:="PowerPoint.Application")
If pptApp Is Nothing Then Set pptApp = CreateObject(class:="PowerPoint.Application")
On Error GoTo 0

pptApp.Visible = True
Set pptPres = pptApp.Presentations.Open(templatePath, Untitled:=True)

' Optimization
Application.ScreenUpdating = False

' Loop through entities (+300 rows)
For i = 2 To lastRow
    ' Duplicate the template slide (Slide 1)
    Set pptSlide = pptPres.Slides(1).Duplicate
    pptSlide.MoveTo (pptPres.Slides.Count)
    
    ' Map Excel Columns to PPT Placeholders
    ' Assumes placeholders are named in the Selection Pane
    With pptSlide.Shapes
        .Item("EntityName").TextFrame.TextRange.Text = wsData.Cells(i, 1).Value
        .Item("Status").TextFrame.TextRange.Text = wsData.Cells(i, 2).Value
        .Item("KPI_Metric").TextFrame.TextRange.Text = wsData.Cells(i, 3).Value
        ' ... Continue mapping for 44 columns ...
    End With
Next i

' Remove the initial template slide
pptPres.Slides(1).Delete

Application.ScreenUpdating = True
MsgBox "Report Generation Complete: " & (lastRow - 1) & " Slides Created.", vbInformation


End Sub