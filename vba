Option Explicit

Const SCORE_THRESHOLD As Integer = 5
Const KW_SHEET        As String  = "Keywords_Reference"
Const DATA_SHEET      As String  = "Classifier"
Const KW_START_ROW    As Integer = 5
Const DATA_START_ROW  As Integer = 5
Const COL_NAME        As Integer = 2   ' B = Party Name
Const COL_COUNTRY     As Integer = 3   ' C = Country Code
Const COL_VBA_TAG     As Integer = 5   ' E = VBA Result
Const COL_KW_TEXT     As Integer = 2   ' B on KW sheet
Const COL_KW_COUNTRY  As Integer = 4   ' D on KW sheet
Const COL_KW_MATCH    As Integer = 6   ' F on KW sheet
Const COL_KW_WEIGHT   As Integer = 7   ' G on KW sheet

'==============================================================
' MAIN: Classify all names
'==============================================================
Sub ClassifyAllNames()
    Dim wsData As Worksheet, wsKW As Worksheet
    Dim lastRow As Long, lastKW As Long, i As Long
    Dim result As String, score As Integer, matched As String
    Dim startTime As Single

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    On Error GoTo ErrorHandler

    Set wsData = ThisWorkbook.Sheets(DATA_SHEET)
    Set wsKW   = ThisWorkbook.Sheets(KW_SHEET)
    lastRow    = wsData.Cells(wsData.Rows.Count, COL_NAME).End(xlUp).Row
    lastKW     = wsKW.Cells(wsKW.Rows.Count, COL_KW_TEXT).End(xlUp).Row
    startTime  = Timer

    If lastRow < DATA_START_ROW Then
        MsgBox "No data found in Classifier sheet.", vbExclamation
        GoTo Cleanup
    End If

    For i = DATA_START_ROW To lastRow
        Dim partyName As String, countryCode As String
        partyName   = Trim(wsData.Cells(i, COL_NAME).Value)
        countryCode = UCase(Trim(wsData.Cells(i, COL_COUNTRY).Value))

        If partyName = "" Then
            wsData.Cells(i, COL_VBA_TAG).Value = ""
            wsData.Cells(i, COL_VBA_TAG).Interior.Color = RGB(240, 240, 240)
            wsData.Cells(i, COL_VBA_TAG).Font.Bold = False
        Else
            Call ClassifySingleName(wsKW, lastKW, partyName, countryCode, result, score, matched)
            wsData.Cells(i, COL_VBA_TAG).Value = result

            If result = "COMPANY" Then
                wsData.Cells(i, COL_VBA_TAG).Interior.Color = RGB(198, 239, 206) ' Green
                wsData.Cells(i, COL_VBA_TAG).Font.Color     = RGB(0, 97, 0)
                wsData.Cells(i, COL_VBA_TAG).Font.Bold      = True
            Else
                wsData.Cells(i, COL_VBA_TAG).Interior.Color = RGB(221, 235, 247) ' Blue
                wsData.Cells(i, COL_VBA_TAG).Font.Color     = RGB(31, 56, 100)
                wsData.Cells(i, COL_VBA_TAG).Font.Bold      = True
            End If
        End If

        If i Mod 200 = 0 Then
            Application.StatusBar = "Classified " & (i - DATA_START_ROW + 1) & _
                                    " of " & (lastRow - DATA_START_ROW + 1)
            DoEvents
        End If
    Next i

    MsgBox "Done! " & (lastRow - DATA_START_ROW + 1) & " records in " & _
           Format(Timer - startTime, "0.0") & "s", vbInformation, "MDM Classifier"

Cleanup:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.StatusBar = False
    Exit Sub
ErrorHandler:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    MsgBox "Error: " & Err.Description, vbCritical
End Sub


'==============================================================
' CORE ENGINE: Score one name vs. all keywords
'==============================================================
Sub ClassifySingleName(wsKW As Worksheet, lastKW As Long, _
                       partyName As String, countryCode As String, _
                       ByRef result As String, ByRef totalScore As Integer, _
                       ByRef topKeyword As String)

    Dim r As Long, kwText As String, kwUpper As String
    Dim kwCountry As String, matchType As String
    Dim kwWeight As Integer, isMatched As Boolean
    Dim topWeight As Integer, nameUpper As String

    totalScore = 0 : topKeyword = "" : topWeight = 0
    nameUpper = UCase(Trim(partyName))

    For r = KW_START_ROW To lastKW
        kwText = Trim(wsKW.Cells(r, COL_KW_TEXT).Value)
        If kwText = "" Then GoTo NextKW

        kwCountry = UCase(Trim(wsKW.Cells(r, COL_KW_COUNTRY).Value))
        matchType = UCase(Trim(wsKW.Cells(r, COL_KW_MATCH).Value))
        kwWeight  = Val(wsKW.Cells(r, COL_KW_WEIGHT).Value)
        kwUpper   = UCase(kwText)

        ' Apply country filter
        If countryCode <> "" And countryCode <> "ALL" Then
            If kwCountry <> "ALL" And kwCountry <> "" And kwCountry <> countryCode Then
                GoTo NextKW
            End If
        End If

        isMatched = False
        Select Case matchType
            Case "ENDS_WITH"
                If Right(RTrim(nameUpper), Len(kwUpper)) = kwUpper Then isMatched = True
                If InStr(nameUpper, " " & kwUpper) > 0 Then
                    Dim p As Integer: p = InStr(nameUpper, " " & kwUpper)
                    If p + Len(kwUpper) >= Len(RTrim(nameUpper)) Then isMatched = True
                End If
            Case "STARTS_WITH"
                If Left(nameUpper, Len(kwUpper)) = kwUpper Then isMatched = True
            Case "CONTAINS"
                If InStr(nameUpper, kwUpper) > 0 Then isMatched = True
            Case Else ' WHOLE_WORD
                If InStr(" " & nameUpper & " ", " " & kwUpper & " ") > 0 Then isMatched = True
                Dim nd As String: nd = Replace(kwUpper, ".", "")
                If Len(nd) > 1 Then
                    If InStr(" " & nameUpper & " ", " " & nd & " ") > 0 Then isMatched = True
                End If
        End Select

        If isMatched Then
            totalScore = totalScore + kwWeight
            If kwWeight > topWeight Then topWeight = kwWeight : topKeyword = kwText
        End If
NextKW:
    Next r

    result = IIf(totalScore >= SCORE_THRESHOLD, "COMPANY", "INDIVIDUAL")
End Sub


'==============================================================
' TEST: Test a single name interactively
'==============================================================
Sub TestSingleName()
    Dim wsKW As Worksheet, lastKW As Long
    Dim testName As String, country As String
    Dim result As String, score As Integer, matched As String

    testName = InputBox("Enter party name to test:", "Test Classifier")
    If testName = "" Then Exit Sub
    country = InputBox("Country code (blank=ALL)  e.g. MX BR ES PT US", "Country")

    Set wsKW = ThisWorkbook.Sheets(KW_SHEET)
    lastKW = wsKW.Cells(wsKW.Rows.Count, COL_KW_TEXT).End(xlUp).Row
    Call ClassifySingleName(wsKW, lastKW, testName, UCase(country), result, score, matched)

    MsgBox "Name   : " & testName & vbNewLine & _
           "Result : " & result & vbNewLine & _
           "Score  : " & score & " (min=" & SCORE_THRESHOLD & ")" & vbNewLine & _
           "Hit    : " & IIf(matched = "", "None", matched), vbInformation, "Result"
End Sub


'==============================================================
' UTILITY: Clear column E
'==============================================================
Sub ClearClassifications()
    If MsgBox("Clear all Column E results?", vbYesNo) = vbNo Then Exit Sub
    Dim wsData As Worksheet: Set wsData = ThisWorkbook.Sheets(DATA_SHEET)
    Dim lastRow As Long
    lastRow = wsData.Cells(wsData.Rows.Count, COL_NAME).End(xlUp).Row
    If lastRow >= DATA_START_ROW Then
        With wsData.Range(wsData.Cells(DATA_START_ROW, COL_VBA_TAG), _
                          wsData.Cells(lastRow, COL_VBA_TAG))
            .ClearContents
            .Interior.ColorIndex = xlNone
            .Font.Bold = False
            .Font.ColorIndex = xlAutomatic
        End With
    End If
    MsgBox "Cleared.", vbInformation
End Sub


'==============================================================
' EXPORT: Values-only sheet for Oracle reload
'==============================================================
Sub ExportResults()
    Dim wsData As Worksheet, wsExp As Worksheet
    Set wsData = ThisWorkbook.Sheets(DATA_SHEET)
    Dim lastRow As Long
    lastRow = wsData.Cells(wsData.Rows.Count, COL_NAME).End(xlUp).Row

    On Error Resume Next: Set wsExp = ThisWorkbook.Sheets("Export_Results"): On Error GoTo 0
    If wsExp Is Nothing Then
        Set wsExp = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        wsExp.Name = "Export_Results"
    Else
        wsExp.Cells.Clear
    End If

    wsExp.Range("A1:G1").Value = Array("Row","Party_Name","Country","Formula_Tag","VBA_Tag","Score","Matched_KW")
    wsExp.Range("A1:G1").Font.Bold = True
    wsExp.Range("A1:G1").Interior.Color = RGB(47, 117, 181)
    wsExp.Range("A1:G1").Font.Color = RGB(255, 255, 255)

    Dim i As Long, er As Long: er = 2
    For i = DATA_START_ROW To lastRow
        If wsData.Cells(i, COL_NAME).Value <> "" Then
            wsExp.Cells(er, 1) = er - 1
            wsExp.Cells(er, 2) = wsData.Cells(i, COL_NAME).Value
            wsExp.Cells(er, 3) = wsData.Cells(i, COL_COUNTRY).Value
            wsExp.Cells(er, 4) = wsData.Cells(i, 4).Value   ' Formula result
            wsExp.Cells(er, 5) = wsData.Cells(i, COL_VBA_TAG).Value
            wsExp.Cells(er, 6) = wsData.Cells(i, 6).Value   ' Score
            wsExp.Cells(er, 7) = wsData.Cells(i, 7).Value   ' Keyword
            er = er + 1
        End If
    Next i
    wsExp.Columns("A:G").AutoFit
    MsgBox (er - 2) & " records exported to 'Export_Results'.", vbInformation
End Sub


'==============================================================
' STATS: Show summary counts
'==============================================================
Sub ShowStats()
    Dim wsData As Worksheet: Set wsData = ThisWorkbook.Sheets(DATA_SHEET)
    Dim lastRow As Long
    lastRow = wsData.Cells(wsData.Rows.Count, COL_NAME).End(xlUp).Row
    Dim co As Long, ind As Long, unc As Long, i As Long
    For i = DATA_START_ROW To lastRow
        If wsData.Cells(i, COL_NAME).Value <> "" Then
            Select Case wsData.Cells(i, COL_VBA_TAG).Value
                Case "COMPANY":    co  = co + 1
                Case "INDIVIDUAL": ind = ind + 1
                Case Else:         unc = unc + 1
            End Select
        End If
    Next i
    Dim t As Long: t = co + ind + unc
    MsgBox "COMPANY     : " & co  & " (" & Format(co  / IIf(t=0,1,t)*100,"0.0") & "%)" & vbNewLine & _
           "INDIVIDUAL  : " & ind & " (" & Format(ind / IIf(t=0,1,t)*100,"0.0") & "%)" & vbNewLine & _
           "Unclassified: " & unc & vbNewLine & "Total: " & t, vbInformation, "Stats"
End Sub
```

---

## Available Macros (Alt+F8)

| Macro | What it does |
|---|---|
| `ClassifyAllNames` | Batch classify all names, colors Column E |
| `TestSingleName` | Interactive popup test for one name |
| `ClearClassifications` | Reset Column E |
| `ExportResults` | Creates Oracle-ready values-only export sheet |
| `ShowStats` | COMPANY vs INDIVIDUAL summary counts |

## Oracle Reload Flow
```
Oracle 30K records → Export CSV → Paste in Column B 
→ Run ClassifyAllNames → ExportResults sheet 
→ Copy/paste values back → SQL Loader / External Table into Oracle
