Attribute VB_Name = "UI_Handlers"

Sub ShowAutomationDashboard() ' Triggers the generation from a Worksheet Button Dim resp As VbMsgBoxResult resp = MsgBox("Start Automated PPT Generation for " & _ (Sheets("EntityData").Cells(Rows.Count, 1).End(xlUp).Row - 1) & _ " entities?", vbQuestion + vbYesNo, "Staqlt Automator")

If resp = vbYes Then
    Call GenerateEntityReports
End If


End Sub