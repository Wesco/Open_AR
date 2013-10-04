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

90        Application.DisplayAlerts = False
100       Application.ScreenUpdating = False

110       On Error GoTo UID_ERR
120       For j = 1 To OAR.Worksheets.Count
130           Set s = OAR.Worksheets(j)
              'Sheets with outside sales names are yellow, claim sheets are red
              '6 = Yellow, 3 = Red
140           If s.Tab.ColorIndex = 6 Then
150               s.AutoFilterMode = False
160               TotalCols = s.Columns(Columns.Count).End(xlToLeft).Column
170               br = s.Cells(2, FindColumn("br", s.Range(s.Cells(1, 1), s.Cells(1, TotalCols)))).Value
180               os_name = s.Cells(2, FindColumn("os_name", s.Range(s.Cells(1, 1), s.Cells(1, TotalCols)))).Value

190               InsertUID WS:=s

200               OARPath = "\\7938-HP02\Shared\" & br & " Open AR\" & UCase(os_name) & "\"
210               If Not FolderExists(OARPath) Then
220                   RecMkDir OARPath
230               End If

240               For i = 0 To 120
250                   OldOARFile = s.Name & " " & Format(Date - i, "yyyy-mm-dd") & ".xlsx"
260                   If FileExists(OARPath & OldOARFile) Then
270                       Set OldARWkbk = Workbooks.Open(OARPath & OldOARFile)
280                       Exit For
290                   End If
300               Next

310               If TypeName(OldARWkbk) <> "Nothing" Then
320                   InsertUID OldARWkbk.Sheets(s.Name)
330                   InsertNotes s, OldARWkbk
340                   OldARWkbk.Saved = True
350                   OldARWkbk.Close
360               End If

370               Set NewARWkbk = Workbooks.Add
380               Application.DisplayAlerts = False
390               NewARWkbk.Sheets(3).Delete
400               NewARWkbk.Sheets(2).Delete
410               Application.DisplayAlerts = True
420               NewARWkbk.Sheets(1).Name = s.Name

430               TotalRows = s.Rows(Rows.Count).End(xlUp).Row + 2
440               TotalCols = s.Columns(Columns.Count).End(xlToLeft).Column
450               s.Range(s.Cells(1, 1), s.Cells(TotalRows, TotalCols)).Copy Destination:=NewARWkbk.Sheets(1).Range("A1")

460               NewARFile = s.Name & " " & Format(Date, "yyyy-mm-dd") & ".xlsx"
470               NewARWkbk.SaveAs OARPath & NewARFile, xlOpenXMLWorkbook
480               NewARWkbk.Close
490           End If
500       Next
510       On Error GoTo 0
520       OAR.Saved = True
530       OAR.Close
540       Application.DisplayAlerts = True
550       Application.ScreenUpdating = True
560       Exit Sub

UID_ERR:
570       Application.DisplayAlerts = False
580       If Err.Number = CustErr.COLNOTFOUND Then
590           MsgBox "Column '" & Err.Description & "' could not be found on '" & s.Name & "'"
600       Else
610           MsgBox Err.Description, vbOKOnly, Err.Source & " Ln# " & Erl
620       End If

630       If TypeName(OAR) <> "Nothing" Then
640           On Error Resume Next
650           OAR.Saved = True
660           OAR.Close
670           On Error GoTo 0
680       End If

690       If TypeName(NewARWkbk) <> "Nothing" Then
700           On Error Resume Next
710           NewARWkbk.Saved = True
720           NewARWkbk.Close
730           On Error GoTo 0
740       End If

750       If TypeName(OldARWkbk) <> "Nothing" Then
760           On Error Resume Next
770           OldARWkbk.Saved = True
780           OldARWkbk.Close
790           On Error GoTo 0
800       End If
810       Application.DisplayAlerts = True
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
