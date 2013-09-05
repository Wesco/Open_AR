Attribute VB_Name = "Program"
Option Explicit

Sub Main()
    Dim FilePath As String
    Dim OAR As Workbook
    Dim s As Worksheet
    Dim TotalRows As Long
    Dim TotalCols As Integer

    Dim os_col As Integer
    Dim br_col As Integer

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

    For Each s In OAR.Sheets
        'Sheets with outside sales names are yellow, claim sheets are red
        '6 = Yellow, 3 = Red
        If s.Tab.ColorIndex = 6 Then
            s.AutoFilterMode = False
            TotalRows = s.Rows(s.Columns(1).Rows.Count).End(xlUp).Row
            TotalCols = s.Columns(s.Columns.Count).End(xlToLeft).Column

            On Error GoTo UID_ERR
            InsertUID WS:=s
            On Error GoTo 0
        End If
    Next

    OAR.Close
    Exit Sub

UID_ERR:
    If Err.Number = CustErr.COLNOTFOUND Then
        MsgBox "Column '" & Err.Description & "' could not be found on '" & s.Name & "'"
    Else
        MsgBox Err.Description, vbOKOnly, Err.Source
    End If
    OAR.Saved = True
    OAR.Close
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
