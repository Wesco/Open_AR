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
    MsgBox "Select the open ar file for your branch."
    FilePath = Application.GetOpenFilename()

    'If a file was selected open it
    If FilePath <> "False" Then
        Set OAR = Workbooks.Open(FilePath)
    Else
        MsgBox "Macro canceled."
        Exit Sub
    End If

    On Error GoTo UID_ERR
    For Each s In OAR.Sheets
        'Sheets with outside sales names are yellow, claim sheets are red
        '6 = Yellow, 3 = Red
        If s.Tab.ColorIndex = 6 Then
            s.AutoFilterMode = False
            TotalCols = s.Columns(Columns.Count).End(xlToLeft).Column
            br = s.Cells(2, FindColumn("br", s.Range(s.Cells(1, 1), s.Cells(1, TotalCols)))).Value
            os_name = s.Cells(2, FindColumn("os_name", s.Range(s.Cells(1, 1), s.Cells(1, TotalCols)))).Value

            InsertUID WS:=s

            OARPath = "\\7938-HP02\Shared\" & br & " Open AR\" & UCase(os_name) & "\"
            For i = 0 To 120
                OldOARFile = s.Name & " " & Format(Date - i, "yyyy-mm-dd") & ".xlsx"
                If FileExists(OARPath & OldOARFile) Then
                    Set OldARWkbk = Workbooks.Open(OARPath & OldOARFile)
                    Exit For
                End If
            Next

            If TypeName(OldARWkbk) <> "Nothing" Then
                InsertUID OldARWkbk.Sheets(s.Name)
                InsertNotes s, OldARWkbk
                OldARWkbk.Saved = True
                OldARWkbk.Close
            End If

            Set NewARWkbk = Workbooks.Add
            Application.DisplayAlerts = False
            NewARWkbk.Sheets(3).Delete
            NewARWkbk.Sheets(2).Delete
            Application.DisplayAlerts = True
            NewARWkbk.Sheets(1).Name = s.Name
            
            TotalRows = s.Rows(Rows.Count).End(xlUp).Row + 2
            TotalCols = s.Columns(Columns.Count).End(xlToLeft).Column
            s.Range(s.Cells(1, 1), s.Cells(TotalRows, TotalCols)).Copy Destination:=NewARWkbk.Sheets(1).Range("A1")
            
            NewARFile = s.Name & " " & Format(Date, "yyyy-mm-dd") & ".xlsx"
            NewARWkbk.SaveAs OARPath & NewARFile, xlOpenXMLWorkbook
            NewARWkbk.Close
        End If
    Next
    On Error GoTo 0
    OAR.Saved = True
    OAR.Close
    Exit Sub

UID_ERR:
    If Err.Number = CustErr.COLNOTFOUND Then
        MsgBox "Column '" & Err.Description & "' could not be found on '" & s.Name & "'"
    Else
        MsgBox Err.Description, vbOKOnly, Err.Source
    End If
    'OAR.Saved = True
    'OAR.Close
End Sub

Sub InsertUID(WS As Worksheet)
    Dim TotalCols As Integer: TotalCols = WS.Columns(Columns.Count).End(xlToLeft).Column
    Dim TotalRows As Long: TotalRows = WS.Rows(Rows.Count).End(xlUp).Row
    Dim HeaderRow As Range: Set HeaderRow = WS.Range(WS.Cells(1, 1), WS.Cells(1, TotalCols))
    Dim inv_col As Integer: inv_col = FindColumn("inv", HeaderRow) + 1
    Dim mfr_col As Integer: mfr_col = FindColumn("mfr", HeaderRow) + 1
    Dim itm_col As Integer: itm_col = FindColumn("item", HeaderRow) + 1
    Dim sls_col As Integer: sls_col = FindColumn("sales", HeaderRow) + 1

    WS.Columns(1).Insert
    WS.Range("A1").Value = "UID"

    With WS.Range("A2:A" & TotalRows)
        'inv + mfr + item + sales
        .Formula = "=""=""""""&" & _
                   Cells(2, inv_col).Address(False, False) & " & " & _
                   Cells(2, mfr_col).Address(False, False) & " & " & _
                   Cells(2, itm_col).Address(False, False) & " & " & _
                   Cells(2, sls_col).Address(False, False) & _
                   "&"""""""""
        .Value = .Value
    End With
End Sub

Sub InsertNotes(CurSheet As Worksheet, OldBook As Workbook)
    Dim note_lookup As String
    Dim note1_col As Integer
    Dim note2_col As Integer
    Dim TotalCols As Integer
    Dim TotalRows As Long

    TotalCols = OldBook.Sheets(CurSheet.Name).Columns(Columns.Count).End(xlToLeft).Column

    On Error Resume Next
    With OldBook.Sheets(CurSheet.Name)
        note1_col = FindColumn("note 1", .Range(.Cells(1, 1), .Cells(1, TotalCols)))
        note2_col = FindColumn("note 2", .Range(.Cells(1, 1), .Cells(1, TotalCols)))
    End With
    On Error GoTo 0

    TotalCols = CurSheet.Columns(Columns.Count).End(xlToLeft).Column + 1
    TotalRows = CurSheet.Rows(Rows.Count).End(xlUp).Row

    If note1_col > 0 Then
        note_lookup = "VLOOKUP(A2,'[" & OldBook.Name & "]" & CurSheet.Name & "'!A:ZZ," & note1_col & ",FALSE)"
        note_lookup = "=IFERROR(IF(" & note_lookup & "=0, """", " & note_lookup & "),"""")"

        With CurSheet
            .Cells(1, TotalCols).Value = "note 1"
            With .Range(.Cells(2, TotalCols), .Cells(TotalRows, TotalCols))
                .Formula = note_lookup
                .NumberFormat = "@"
                .Value = .Value
            End With
        End With
    Else
        CurSheet.Cells(1, TotalCols).Value = "note 1"
    End If

    TotalCols = TotalCols + 1

    If note2_col > 0 Then
        note_lookup = "VLOOKUP(A2,'[" & OldBook.Name & "]" & CurSheet.Name & "'!A:ZZ," & note2_col & ",FALSE)"
        note_lookup = "=IFERROR(IF(" & note_lookup & "=0, """", " & note_lookup & "),"""")"

        With CurSheet
            .Cells(1, TotalCols).Value = "note 2"
            With .Range(.Cells(2, TotalCols), .Cells(TotalRows, TotalCols))
                .Formula = note_lookup
                .NumberFormat = "@"
                .Value = .Value
            End With
        End With
    Else
        CurSheet.Cells(1, TotalCols).Value = "note 2"
    End If

    CurSheet.Columns(1).Delete
End Sub
