Attribute VB_Name = "MTA_Snapshot"
'==============================================================================
' OPTIONAL macro for Advanced_Milestone_Trend_Analysis (save the workbook as .xlsm)
'
' Why VBA is (optionally) useful: Excel formulas cannot write new rows into a
' table. Without VBA, the Reporting History form builds a validated,
' paste-ready block that the user copies and pastes as values below
' tblMilestoneData. This macro does that step in one click.
'
' What it does:
'   * reads the "ready to post" block (Reporting History!O10:S34)
'   * re-checks that each Milestone ID + Reporting Date is not already recorded
'   * appends each row to tblMilestoneData (history is never overwritten)
'   * clears the yellow input cells (G10:I34) on success
'
' Install: Alt+F11 > File > Import File > MTA_Snapshot.bas, then
'          Insert > Shape on Reporting History > right-click > Assign Macro > PostSnapshot
'==============================================================================
Option Explicit

Public Sub PostSnapshot()
    Dim wsRH As Worksheet, lo As ListObject, r As Long, added As Long, skipped As Long
    Dim mid As String, repDate As Variant, fc As Variant, act As Variant, cmt As Variant
    Dim newRow As ListRow, msg As String

    Set wsRH = ThisWorkbook.Worksheets("Reporting History")
    Set lo = ThisWorkbook.Worksheets("Milestone Data").ListObjects("tblMilestoneData")

    Application.Calculate
    If Not IsDate(wsRH.Range("C5").Value) Then
        MsgBox "Enter a valid reporting date in C5 first.", vbExclamation, "Post snapshot"
        Exit Sub
    End If

    Application.ScreenUpdating = False
    For r = 10 To 34
        mid = CStr(wsRH.Cells(r, "O").Value)
        If Len(mid) > 0 Then
            repDate = wsRH.Cells(r, "P").Value
            fc = wsRH.Cells(r, "Q").Value
            act = wsRH.Cells(r, "R").Value
            cmt = wsRH.Cells(r, "S").Value
            If ObservationExists(lo, mid, CDate(repDate)) Then
                skipped = skipped + 1
                msg = msg & vbCrLf & "  " & mid & " - already recorded for " & Format(repDate, "dd-mmm-yyyy")
            ElseIf Not IsDate(fc) Then
                skipped = skipped + 1
                msg = msg & vbCrLf & "  " & mid & " - invalid forecast date"
            Else
                Set newRow = lo.ListRows.Add
                With newRow.Range
                    .Cells(1, lo.ListColumns("Milestone ID").Index).Value = mid
                    .Cells(1, lo.ListColumns("Reporting Date").Index).Value = CDate(repDate)
                    .Cells(1, lo.ListColumns("Forecast Completion Date").Index).Value = CDate(fc)
                    If IsDate(act) Then .Cells(1, lo.ListColumns("Actual Completion Date").Index).Value = CDate(act)
                    If Len(CStr(cmt)) > 0 Then .Cells(1, lo.ListColumns("Comments").Index).Value = CStr(cmt)
                End With
                added = added + 1
            End If
        End If
    Next r

    If added > 0 And skipped = 0 Then wsRH.Range("G10:I34").ClearContents
    Application.ScreenUpdating = True
    Application.Calculate

    MsgBox added & " observation(s) posted to tblMilestoneData." & _
           IIf(skipped > 0, vbCrLf & skipped & " row(s) skipped:" & msg & vbCrLf & "Inputs were kept so you can correct them.", ""), _
           IIf(skipped > 0, vbExclamation, vbInformation), "Post snapshot"
End Sub

Private Function ObservationExists(lo As ListObject, mid As String, repDate As Date) As Boolean
    ObservationExists = Application.WorksheetFunction.CountIfs( _
        lo.ListColumns("Milestone ID").DataBodyRange, mid, _
        lo.ListColumns("Reporting Date").DataBodyRange, repDate) > 0
End Function
