Attribute VB_Name = "AHF_Imports"
Option Explicit

'---------------------------------------------------------------------------------------
' Proc  : Sub ImportGaps
' Date  : 12/12/2012
' Desc  : Imports gaps to the workbook containing this macro.
' Ex    : ImportGaps
'---------------------------------------------------------------------------------------
Sub ImportGaps()
          Dim sPath As String     'Gaps file path
          Dim sName As String     'Gaps Sheet Name
          Dim iCounter As Long    'Counter to decrement the date
          Dim iRows As Long       'Total number of rows
          Dim dt As Date          'Date for gaps file name and path
          Dim Result As VbMsgBoxResult    'Yes/No to proceed with old gaps file if current one isn't found
          Dim Gaps As Worksheet           'The sheet named gaps if it exists, else this = nothing
          Dim StartTime As Double         'The time this function was started
          Dim FileFound As Boolean        'Indicates whether or not gaps was found

10        StartTime = Timer
20        FileFound = False

          'This error is bypassed so you can determine whether or not the sheet exists
30        On Error GoTo CREATE_GAPS
40        Set Gaps = ThisWorkbook.Sheets("Gaps")
50        On Error GoTo 0

60        Application.DisplayAlerts = False

          'Find gaps
70        For iCounter = 0 To 15
80            dt = Date - iCounter
90            sPath = "\\br3615gaps\gaps\3615 Gaps Download\" & Format(dt, "yyyy") & "\"
100           sName = "3615 " & Format(dt, "yyyy-mm-dd") & ".xlsx"
110           If FileExists(sPath & sName) Then
120               FileFound = True
130               Exit For
140           End If
150       Next

          'Make sure Gaps file was found
160       If FileFound = True Then
170           If dt <> Date Then
180               Result = MsgBox( _
                           Prompt:="Gaps from " & Format(dt, "mmm dd, yyyy") & " was found." & vbCrLf & "Would you like to continue?", _
                           Buttons:=vbYesNo, _
                           Title:="Gaps not up to date")
190           End If

200           If Result <> vbNo Then
210               If ThisWorkbook.Sheets("Gaps").Range("A1").Value <> "" Then
220                   Gaps.Cells.Delete
230               End If

240               Workbooks.Open sPath & sName
250               ActiveSheet.UsedRange.Copy Destination:=ThisWorkbook.Sheets("Gaps").Range("A1")
260               ActiveWorkbook.Close

270               Sheets("Gaps").Select
280               iRows = ActiveSheet.UsedRange.Rows.Count
290               Columns(1).EntireColumn.Insert
300               Range("A1").Value = "SIM"
310               Range("A2").Formula = "=C2&D2"
320               Range("A2").AutoFill Destination:=Range(Cells(2, 1), Cells(iRows, 1))
330               Range(Cells(2, 1), Cells(iRows, 1)).Value = Range(Cells(2, 1), Cells(iRows, 1)).Value
340           Else
350               Err.Raise 18, "ImportGaps", "Import canceled"
360           End If
370       Else
380           Err.Raise 53, "ImportGaps", "Gaps could not be found."
390       End If

400       Application.DisplayAlerts = True
410       Exit Sub

CREATE_GAPS:
420       ThisWorkbook.Sheets.Add After:=Sheets(ThisWorkbook.Sheets.Count)
430       ActiveSheet.Name = "Gaps"
440       Resume

End Sub

'---------------------------------------------------------------------------------------
' Proc : UserImportFile
' Date : 1/29/2013
' Desc : Prompts the user to select a file for import
'---------------------------------------------------------------------------------------
Sub UserImportFile(DestRange As Range, Optional DelFile As Boolean = False, Optional ShowAllData As Boolean = False, Optional SourceSheet As String = "", Optional FileFilter = "")
          Dim File As String              'Full path to user selected file
          Dim FileDate As String          'Date the file was last modified
          Dim OldDispAlert As Boolean     'Original state of Application.DisplayAlerts

10        OldDispAlert = Application.DisplayAlerts
20        File = Application.GetOpenFilename(FileFilter)

30        Application.DisplayAlerts = False
40        If File <> "False" Then
50            FileDate = Format(FileDateTime(File), "mm/dd/yy")
60            Workbooks.Open File
70            If SourceSheet = "" Then SourceSheet = ActiveSheet.Name
80            If ShowAllData = True Then
90                On Error Resume Next
100               ActiveSheet.AutoFilter.ShowAllData
110               ActiveSheet.UsedRange.Columns.Hidden = False
120               ActiveSheet.UsedRange.Rows.Hidden = False
130               On Error GoTo 0
140           End If
150           Sheets(SourceSheet).UsedRange.Copy Destination:=DestRange
160           ActiveWorkbook.Close
170           ThisWorkbook.Activate

180           If DelFile = True Then
190               DeleteFile File
200           End If
210       Else
220           Err.Raise 18
230       End If
240       Application.DisplayAlerts = OldDispAlert
End Sub

'---------------------------------------------------------------------------------------
' Proc : Import117byISN
' Date : 4/10/2013
' Desc : Imports the most recent 117 report for the specified sales number
'---------------------------------------------------------------------------------------
Sub Import117byISN(RepType As ReportType, Destination As Range, Optional ByVal ISN As String = "", Optional Cancel As Boolean = False)
          Dim sPath As String
          Dim FileName As String

10        If ISN = "" And Cancel = False Then
20            ISN = InputBox("Inside Sales Number:", "Please enter the ISN#")
30        Else
40            If ISN = "" Then
50                Err.Raise 53
60            End If
70        End If

80        If ISN <> "" Then
90            Select Case RepType
                  Case ReportType.DS:
100                   FileName = "3615 " & Format(Date, "m-dd-yy") & " DSORDERS.xlsx"

110               Case ReportType.BO:
120                   FileName = "3615 " & Format(Date, "m-dd-yy") & " BACKORDERS.xlsx"

130               Case ReportType.ALL
140                   FileName = "3615 " & Format(Date, "m-dd-yy") & " ALLORDERS.xlsx"
150           End Select

160           sPath = "\\br3615gaps\gaps\3615 117 Report\ByInsideSalesNumber\" & ISN & "\" & FileName

170           If FileExists(sPath) Then
180               Workbooks.Open sPath
190               ActiveSheet.UsedRange.Copy Destination:=Destination
200               Application.DisplayAlerts = False
210               ActiveWorkbook.Close
220               Application.DisplayAlerts = True
230           Else
240               MsgBox Prompt:=ReportTypeText(RepType) & " report not found.", Title:="Error 53"
250           End If
260       Else
270           Err.Raise 18
280       End If
End Sub

'---------------------------------------------------------------------------------------
' Proc : Import473
' Date : 4/11/2013
' Desc : Imports a 473 report from the current day
'---------------------------------------------------------------------------------------
Sub Import473(Destination As Range, Optional Branch As String = "3615")
          Dim sPath As String
          Dim FileName As String
          Dim AlertStatus As Boolean

10        FileName = "473 " & Format(Date, "m-dd-yy") & ".xlsx"
20        sPath = "\\br3615gaps\gaps\" & Branch & " 473 Download\" & FileName
30        AlertStatus = Application.DisplayAlerts

40        If FileExists(sPath) Then
50            Workbooks.Open sPath
60            ActiveSheet.UsedRange.Copy Destination:=Destination

70            Application.DisplayAlerts = False
80            ActiveWorkbook.Close
90            Application.DisplayAlerts = AlertStatus
100       Else
110           MsgBox Prompt:="473 report not found."
120           Err.Raise 18
130       End If
End Sub

'---------------------------------------------------------------------------------------
' Proc : ImportSupplierContacts
' Date : 4/22/2013
' Desc : Imports the supplier contact master list
'---------------------------------------------------------------------------------------
Sub ImportSupplierContacts(Destination As Range)
          Const sPath As String = "\\br3615gaps\gaps\Contacts\Supplier Contact Master.xlsx"
          Dim PrevDispAlerts As Boolean

10        PrevDispAlerts = Application.DisplayAlerts

20        Workbooks.Open sPath
30        ActiveSheet.UsedRange.Copy Destination:=Destination

40        Application.DisplayAlerts = False
50        ActiveWorkbook.Close
60        Application.DisplayAlerts = PrevDispAlerts
End Sub
