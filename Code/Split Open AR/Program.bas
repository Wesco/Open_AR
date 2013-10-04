Attribute VB_Name = "Program"
Option Explicit

Sub Main()
          Dim OARPath As String
          Dim OldOARFile As String
          Dim OldARWkbk As Workbook
          Dim NewARFile As String
          Dim NewARWkbk As Workbook
          Dim FilePath As String
          Dim OAR As Workbook
          Dim br As String
          Dim os_name As String
          Dim TotalRows As Long
          Dim TotalCols As Integer
          Dim s As Worksheet
          Dim i As Long
          Dim j As Long

          'Prompt the user for the Open AR master file
10        MsgBox "Select the open ar file for your branch."
20        FilePath = Application.GetOpenFilename()

          'If a file was selected open it
30        If FilePath <> "False" Then
40            Set OAR = Workbooks.Open(FilePath)
50        Else
60            MsgBox "Macro canceled."
70            Exit Sub
80        End If

90        On Error GoTo UID_ERR
100       For j = 1 To OAR.Worksheets.Count
110           Set s = OAR.Worksheets(j)
              'Sheets with outside sales names are yellow, claim sheets are red
              '6 = Yellow, 3 = Red
120           If s.Tab.ColorIndex = 6 Then
130               s.AutoFilterMode = False
140               TotalCols = s.Columns(Columns.Count).End(xlToLeft).Column
150               br = s.Cells(2, FindColumn("br", s.Range(s.Cells(1, 1), s.Cells(1, TotalCols)))).Value
160               os_name = s.Cells(2, FindColumn("os_name", s.Range(s.Cells(1, 1), s.Cells(1, TotalCols)))).Value

170               InsertUID WS:=s

180               OARPath = "\\7938-HP02\Shared\" & br & " Open AR\" & UCase(os_name) & "\"
190               If Not FolderExists(OARPath) Then
200                   RecMkDir OARPath
210               End If

220               For i = 0 To 120
230                   OldOARFile = s.Name & " " & Format(Date - i, "yyyy-mm-dd") & ".xlsx"
240                   If FileExists(OARPath & OldOARFile) Then
250                       Set OldARWkbk = Workbooks.Open(OARPath & OldOARFile)
260                       Exit For
270                   End If
280               Next

290               If TypeName(OldARWkbk) <> "Nothing" Then
300                   InsertUID OldARWkbk.Sheets(s.Name)
310                   InsertNotes s, OldARWkbk
320                   OldARWkbk.Saved = True
330                   OldARWkbk.Close
340               End If

350               Set NewARWkbk = Workbooks.Add
360               Application.DisplayAlerts = False
370               NewARWkbk.Sheets(3).Delete
380               NewARWkbk.Sheets(2).Delete
390               Application.DisplayAlerts = True
400               NewARWkbk.Sheets(1).Name = s.Name

410               TotalRows = s.Rows(Rows.Count).End(xlUp).Row + 2
420               TotalCols = s.Columns(Columns.Count).End(xlToLeft).Column
430               s.Range(s.Cells(1, 1), s.Cells(TotalRows, TotalCols)).Copy Destination:=NewARWkbk.Sheets(1).Range("A1")

440               NewARFile = s.Name & " " & Format(Date, "yyyy-mm-dd") & ".xlsx"
450               NewARWkbk.SaveAs OARPath & NewARFile, xlOpenXMLWorkbook
460               NewARWkbk.Close
470           End If
480       Next
490       On Error GoTo 0
500       OAR.Saved = True
510       OAR.Close
520       Exit Sub

UID_ERR:
530       Application.DisplayAlerts = False
540       If Err.Number = CustErr.COLNOTFOUND Then
550           MsgBox "Column '" & Err.Description & "' could not be found on '" & s.Name & "'"
560       Else
570           MsgBox Err.Description, vbOKOnly, Err.Source & " Ln# " & Erl
580       End If

590       If TypeName(OAR) <> "Nothing" Then
600           On Error Resume Next
610           OAR.Saved = True
620           OAR.Close
630           On Error GoTo 0
640       End If

650       If TypeName(NewARWkbk) <> "Nothing" Then
660           On Error Resume Next
670           NewARWkbk.Saved = True
680           NewARWkbk.Close
690           On Error GoTo 0
700       End If

710       If TypeName(OldARWkbk) <> "Nothing" Then
720           On Error Resume Next
730           OldARWkbk.Saved = True
740           OldARWkbk.Close
750           On Error GoTo 0
760       End If
770       Application.DisplayAlerts = True
End Sub

Sub InsertUID(WS As Worksheet)
10  Dim TotalCols As Integer: TotalCols = WS.Columns(Columns.Count).End(xlToLeft).Column
20  Dim TotalRows As Long: TotalRows = WS.Rows(Rows.Count).End(xlUp).Row
30  Dim HeaderRow As Range: Set HeaderRow = WS.Range(WS.Cells(1, 1), WS.Cells(1, TotalCols))
40  Dim inv_col As Integer: inv_col = FindColumn("inv", HeaderRow) + 1
50  Dim mfr_col As Integer: mfr_col = FindColumn("mfr", HeaderRow) + 1
60  Dim itm_col As Integer: itm_col = FindColumn("item", HeaderRow) + 1
70  Dim sls_col As Integer: sls_col = FindColumn("sales", HeaderRow) + 1

80  WS.Columns(1).Insert
90  WS.Range("A1").Value = "UID"

100 With WS.Range("A2:A" & TotalRows)
        'inv + mfr + item + sales
110     .Formula = "=""=""""""&" & _
                   Cells(2, inv_col).Address(False, False) & " & " & _
                   Cells(2, mfr_col).Address(False, False) & " & " & _
                   Cells(2, itm_col).Address(False, False) & " & " & _
                   Cells(2, sls_col).Address(False, False) & _
                   "&"""""""""
120     .Value = .Value
130 End With
End Sub

Sub InsertNotes(CurSheet As Worksheet, OldBook As Workbook)
    Dim note_lookup As String
    Dim note1_col As Integer
    Dim note2_col As Integer
    Dim TotalCols As Integer
    Dim TotalRows As Long

10  TotalCols = OldBook.Sheets(CurSheet.Name).Columns(Columns.Count).End(xlToLeft).Column

20  On Error Resume Next
30  With OldBook.Sheets(CurSheet.Name)
40      note1_col = FindColumn("note 1", .Range(.Cells(1, 1), .Cells(1, TotalCols)))
50      note2_col = FindColumn("note 2", .Range(.Cells(1, 1), .Cells(1, TotalCols)))
60  End With
70  On Error GoTo 0

80  TotalCols = CurSheet.Columns(Columns.Count).End(xlToLeft).Column + 1
90  TotalRows = CurSheet.Rows(Rows.Count).End(xlUp).Row

100 If note1_col > 0 Then
110     note_lookup = "VLOOKUP(A2,'[" & OldBook.Name & "]" & CurSheet.Name & "'!A:ZZ," & note1_col & ",FALSE)"
120     note_lookup = "=IFERROR(IF(" & note_lookup & "=0, """", " & note_lookup & "),"""")"

130     With CurSheet
140         .Cells(1, TotalCols).Value = "note 1"
150         With .Range(.Cells(2, TotalCols), .Cells(TotalRows, TotalCols))
160             .Formula = note_lookup
170             .NumberFormat = "@"
180             .Value = .Value
190         End With
200     End With
210 Else
220     CurSheet.Cells(1, TotalCols).Value = "note 1"
230 End If

240 TotalCols = TotalCols + 1

250 If note2_col > 0 Then
260     note_lookup = "VLOOKUP(A2,'[" & OldBook.Name & "]" & CurSheet.Name & "'!A:ZZ," & note2_col & ",FALSE)"
270     note_lookup = "=IFERROR(IF(" & note_lookup & "=0, """", " & note_lookup & "),"""")"

280     With CurSheet
290         .Cells(1, TotalCols).Value = "note 2"
300         With .Range(.Cells(2, TotalCols), .Cells(TotalRows, TotalCols))
310             .Formula = note_lookup
320             .NumberFormat = "@"
330             .Value = .Value
340         End With
350     End With
360 Else
370     CurSheet.Cells(1, TotalCols).Value = "note 2"
380 End If

390 CurSheet.Columns(1).Delete
End Sub
