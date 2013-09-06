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
100       For Each s In OAR.Sheets
              'Sheets with outside sales names are yellow, claim sheets are red
              '6 = Yellow, 3 = Red
110           If s.Tab.ColorIndex = 6 Then
120               s.AutoFilterMode = False
130               TotalCols = s.Columns(Columns.Count).End(xlToLeft).Column
140               br = s.Cells(2, FindColumn("br", s.Range(s.Cells(1, 1), s.Cells(1, TotalCols)))).Value
150               os_name = s.Cells(2, FindColumn("os_name", s.Range(s.Cells(1, 1), s.Cells(1, TotalCols)))).Value

160               InsertUID WS:=s

170               OARPath = "\\7938-HP02\Shared\" & br & " Open AR\" & UCase(os_name) & "\"
180               If Not FolderExists(OARPath) Then
190                   RecMkDir OARPath
200               End If

210               For i = 0 To 120
220                   OldOARFile = s.Name & " " & Format(Date - i, "yyyy-mm-dd") & ".xlsx"
230                   If FileExists(OARPath & OldOARFile) Then
240                       Set OldARWkbk = Workbooks.Open(OARPath & OldOARFile)
250                       Exit For
260                   End If
270               Next

280               If TypeName(OldARWkbk) <> "Nothing" Then
290                   InsertUID OldARWkbk.Sheets(s.Name)
300                   InsertNotes s, OldARWkbk
310                   OldARWkbk.Saved = True
320                   OldARWkbk.Close
330               End If

340               Set NewARWkbk = Workbooks.Add
350               Application.DisplayAlerts = False
360               NewARWkbk.Sheets(3).Delete
370               NewARWkbk.Sheets(2).Delete
380               Application.DisplayAlerts = True
390               NewARWkbk.Sheets(1).Name = s.Name

400               TotalRows = s.Rows(Rows.Count).End(xlUp).Row + 2
410               TotalCols = s.Columns(Columns.Count).End(xlToLeft).Column
420               s.Range(s.Cells(1, 1), s.Cells(TotalRows, TotalCols)).Copy Destination:=NewARWkbk.Sheets(1).Range("A1")

430               NewARFile = s.Name & " " & Format(Date, "yyyy-mm-dd") & ".xlsx"
440               NewARWkbk.SaveAs OARPath & NewARFile, xlOpenXMLWorkbook
450               NewARWkbk.Close
460           End If
470       Next
480       On Error GoTo 0
490       OAR.Saved = True
500       OAR.Close
510       Exit Sub

UID_ERR:
520       If Err.Number = CustErr.COLNOTFOUND Then
530           MsgBox "Column '" & Err.Description & "' could not be found on '" & s.Name & "'"
540       Else
550           MsgBox Err.Description, vbOKOnly, Err.Source & " Ln# " & Erl
560       End If

570       If TypeName(OAR) <> "Nothing" Then
580           OAR.Saved = True
590           OAR.Close
600       End If

610       If TypeName(NewARWkbk) <> "Nothing" Then
620           NewARWkbk.Saved = True
630           NewARWkbk.Close
640       End If

650       If TypeName(OldARWkbk) <> "Nothing" Then
660           OldARWkbk.Saved = True
670           OldARWkbk.Close
680       End If

End Sub

Sub InsertUID(WS As Worksheet)
10        Dim TotalCols As Integer: TotalCols = WS.Columns(Columns.Count).End(xlToLeft).Column
20        Dim TotalRows As Long: TotalRows = WS.Rows(Rows.Count).End(xlUp).Row
30        Dim HeaderRow As Range: Set HeaderRow = WS.Range(WS.Cells(1, 1), WS.Cells(1, TotalCols))
40        Dim inv_col As Integer: inv_col = FindColumn("inv", HeaderRow) + 1
50        Dim mfr_col As Integer: mfr_col = FindColumn("mfr", HeaderRow) + 1
60        Dim itm_col As Integer: itm_col = FindColumn("item", HeaderRow) + 1
70        Dim sls_col As Integer: sls_col = FindColumn("sales", HeaderRow) + 1

80        WS.Columns(1).Insert
90        WS.Range("A1").Value = "UID"

100       With WS.Range("A2:A" & TotalRows)
              'inv + mfr + item + sales
110           .Formula = "=""=""""""&" & _
                         Cells(2, inv_col).Address(False, False) & " & " & _
                         Cells(2, mfr_col).Address(False, False) & " & " & _
                         Cells(2, itm_col).Address(False, False) & " & " & _
                         Cells(2, sls_col).Address(False, False) & _
                         "&"""""""""
120           .Value = .Value
130       End With
End Sub

Sub InsertNotes(CurSheet As Worksheet, OldBook As Workbook)
          Dim note_lookup As String
          Dim note1_col As Integer
          Dim note2_col As Integer
          Dim TotalCols As Integer
          Dim TotalRows As Long

10        TotalCols = OldBook.Sheets(CurSheet.Name).Columns(Columns.Count).End(xlToLeft).Column

20        On Error Resume Next
30        With OldBook.Sheets(CurSheet.Name)
40            note1_col = FindColumn("note 1", .Range(.Cells(1, 1), .Cells(1, TotalCols)))
50            note2_col = FindColumn("note 2", .Range(.Cells(1, 1), .Cells(1, TotalCols)))
60        End With
70        On Error GoTo 0

80        TotalCols = CurSheet.Columns(Columns.Count).End(xlToLeft).Column + 1
90        TotalRows = CurSheet.Rows(Rows.Count).End(xlUp).Row

100       If note1_col > 0 Then
110           note_lookup = "VLOOKUP(A2,'[" & OldBook.Name & "]" & CurSheet.Name & "'!A:ZZ," & note1_col & ",FALSE)"
120           note_lookup = "=IFERROR(IF(" & note_lookup & "=0, """", " & note_lookup & "),"""")"

130           With CurSheet
140               .Cells(1, TotalCols).Value = "note 1"
150               With .Range(.Cells(2, TotalCols), .Cells(TotalRows, TotalCols))
160                   .Formula = note_lookup
170                   .NumberFormat = "@"
180                   .Value = .Value
190               End With
200           End With
210       Else
220           CurSheet.Cells(1, TotalCols).Value = "note 1"
230       End If

240       TotalCols = TotalCols + 1

250       If note2_col > 0 Then
260           note_lookup = "VLOOKUP(A2,'[" & OldBook.Name & "]" & CurSheet.Name & "'!A:ZZ," & note2_col & ",FALSE)"
270           note_lookup = "=IFERROR(IF(" & note_lookup & "=0, """", " & note_lookup & "),"""")"

280           With CurSheet
290               .Cells(1, TotalCols).Value = "note 2"
300               With .Range(.Cells(2, TotalCols), .Cells(TotalRows, TotalCols))
310                   .Formula = note_lookup
320                   .NumberFormat = "@"
330                   .Value = .Value
340               End With
350           End With
360       Else
370           CurSheet.Cells(1, TotalCols).Value = "note 2"
380       End If

390       CurSheet.Columns(1).Delete
End Sub
