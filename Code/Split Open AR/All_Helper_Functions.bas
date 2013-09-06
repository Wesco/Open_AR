Attribute VB_Name = "All_Helper_Functions"
Option Explicit

'List of error codes
Enum Errors
    USER_INTERRUPT = 18
    FILE_NOT_FOUND = 53
    FILE_ALREADY_OPEN = 55
    FILE_ALREADY_EXISTS = 58
    DISK_FULL = 63
    PERMISSION_DENIED = 70
    PATH_FILE_ACCESS_ERROR = 75
    PATH_NOT_FOUND = 76
    ORBJECT_OR_WITH_BLOCK_NOT_SET = 91
    INVALID_FILE_FORMAT = 321
    OUT_OF_MEMORY = 31001
    ERROR_SAVING_FILE = 31036
    ERROR_LOADING_FROM_FILE = 31037
End Enum

'List of custom error messages
Enum CustErr
    COLNOTFOUND = 50000
End Enum

'Used when importing 117 to determine the type of report to pull
Enum ReportType
    DS
    BO
    ALL
    INQ
End Enum

'---------------------------------------------------------------------------------------
' Proc : FilterSheet
' Date : 1/29/2013
' Desc : Remove all rows that do not match a specified string
'---------------------------------------------------------------------------------------
Sub FilterSheet(sFilter As String, ColNum As Integer, Match As Boolean)
          Dim Rng As Range
          Dim aRng() As Variant
          Dim aHeaders As Variant
          Dim StartTime As Double
          Dim iCounter As Long
          Dim i As Long
          Dim y As Long

10        StartTime = Timer
20        Set Rng = ActiveSheet.UsedRange
30        aHeaders = Range(Cells(1, 1), Cells(1, ActiveSheet.UsedRange.Columns.Count))
40        iCounter = 1

50        Do While iCounter <= Rng.Rows.Count
60            If Match = True Then
70                If Rng(iCounter, ColNum).Value = sFilter Then
80                    i = i + 1
90                End If
100           Else
110               If Rng(iCounter, ColNum).Value <> sFilter Then
120                   i = i + 1
130               End If
140           End If
150           iCounter = iCounter + 1
160       Loop

170       ReDim aRng(1 To i, 1 To Rng.Columns.Count) As Variant

180       iCounter = 1
190       i = 0
200       Do While iCounter <= Rng.Rows.Count
210           If Match = True Then
220               If Rng(iCounter, ColNum).Value = sFilter Then
230                   i = i + 1
240                   For y = 1 To Rng.Columns.Count
250                       aRng(i, y) = Rng(iCounter, y)
260                   Next
270               End If
280           Else
290               If Rng(iCounter, ColNum).Value <> sFilter Then
300                   i = i + 1
310                   For y = 1 To Rng.Columns.Count
320                       aRng(i, y) = Rng(iCounter, y)
330                   Next
340               End If
350           End If
360           iCounter = iCounter + 1
370       Loop

380       ActiveSheet.Cells.Delete
390       Range(Cells(1, 1), Cells(UBound(aRng, 1), UBound(aRng, 2))) = aRng
400       Rows(1).Insert
410       Range(Cells(1, 1), Cells(1, UBound(aHeaders, 2))) = aHeaders
End Sub

'---------------------------------------------------------------------------------------
' Proc : ExportCode
' Date : 3/19/2013
' Desc : Exports all modules
'---------------------------------------------------------------------------------------
Sub ExportCode()
          Dim comp As Variant
          Dim codeFolder As String
          Dim FileName As String
          Dim File As String

          'References Microsoft Visual Basic for Applications Extensibility 5.3
10        AddReference "{0002E157-0000-0000-C000-000000000046}", 5, 3
20        codeFolder = GetWorkbookPath & "Code\" & Left(ThisWorkbook.Name, Len(ThisWorkbook.Name) - 5) & "\"

30        On Error Resume Next
40        RecMkDir codeFolder
50        On Error GoTo 0

          'Remove all previously exported modules
60        File = Dir(codeFolder)
70        Do While File <> ""
80            DeleteFile codeFolder & File
90            File = Dir
100       Loop

          'Export modules in current project
110       For Each comp In ThisWorkbook.VBProject.VBComponents
120           Select Case comp.Type
                  Case 1
130                   FileName = codeFolder & comp.Name & ".bas"
140                   comp.Export FileName
150               Case 2
160                   FileName = codeFolder & comp.Name & ".cls"
170                   comp.Export FileName
180               Case 3
190                   FileName = codeFolder & comp.Name & ".frm"
200                   comp.Export FileName
210           End Select
220       Next
End Sub

'---------------------------------------------------------------------------------------
' Proc : ImportModule
' Date : 4/4/2013
' Desc : Imports a code module into VBProject
'---------------------------------------------------------------------------------------
Sub ImportModule()
          Dim comp As Variant
          Dim codeFolder As String
          Dim FileName As String
          Dim WkbkPath As String

          'Adds a reference to Microsoft Visual Basic for Applications Extensibility 5.3
10        AddReference "{0002E157-0000-0000-C000-000000000046}", 5, 3

          'Gets the path to this workbook
20        WkbkPath = Left$(ThisWorkbook.fullName, InStr(1, ThisWorkbook.fullName, ThisWorkbook.Name, vbTextCompare) - 1)

          'Gets the path to this workbooks code
30        codeFolder = WkbkPath & "Code\" & Left(ThisWorkbook.Name, Len(ThisWorkbook.Name) - 5) & "\"

40        For Each comp In ThisWorkbook.VBProject.VBComponents
50            If comp.Name <> "All_Helper_Functions" Then
60                Select Case comp.Type
                      Case 1
70                        FileName = codeFolder & comp.Name & ".bas"
80                        ThisWorkbook.VBProject.VBComponents.Remove comp
90                        ThisWorkbook.VBProject.VBComponents.Import FileName
100                   Case 2
110                       FileName = codeFolder & comp.Name & ".cls"
120                       ThisWorkbook.VBProject.VBComponents.Remove comp
130                       ThisWorkbook.VBProject.VBComponents.Import FileName
140                   Case 3
150                       FileName = codeFolder & comp.Name & ".frm"
160                       ThisWorkbook.VBProject.VBComponents.Remove comp
170                       ThisWorkbook.VBProject.VBComponents.Import FileName
180               End Select
190           End If
200       Next
End Sub

'---------------------------------------------------------------------------------------
' Proc : GetWorkbookPath
' Date : 3/19/2013
' Desc : Gets the full path of ThisWorkbook
'---------------------------------------------------------------------------------------
Function GetWorkbookPath() As String
          Dim fullName As String
          Dim wrkbookName As String
          Dim pos As Long

10        wrkbookName = ThisWorkbook.Name
20        fullName = ThisWorkbook.fullName

30        pos = InStr(1, fullName, wrkbookName, vbTextCompare)

40        GetWorkbookPath = Left$(fullName, pos - 1)
End Function

'---------------------------------------------------------------------------------------
' Proc : EndsWith
' Date : 3/19/2013
' Desc : Checks if a string ends in a specified character
'---------------------------------------------------------------------------------------
Function EndsWith(ByVal InString As String, ByVal TestString As String) As Boolean
10        EndsWith = (Right$(InString, Len(TestString)) = TestString)
End Function

'---------------------------------------------------------------------------------------
' Proc : AddReferences
' Date : 3/19/2013
' Desc : Adds a reference to VBProject
'---------------------------------------------------------------------------------------
Sub AddReference(GUID As String, Major As Integer, Minor As Integer)
          Dim ID As Variant
          Dim Ref As Variant
          Dim Result As Boolean


10        For Each Ref In ThisWorkbook.VBProject.References
20            If Ref.GUID = GUID And Ref.Major = Major And Ref.Minor = Minor Then
30                Result = True
40            End If
50        Next

          'References Microsoft Visual Basic for Applications Extensibility 5.3
60        If Result = False Then
70            ThisWorkbook.VBProject.References.AddFromGuid GUID, Major, Minor
80        End If
End Sub

'---------------------------------------------------------------------------------------
' Proc : RemoveReferences
' Date : 3/19/2013
' Desc : Removes a reference from VBProject
'---------------------------------------------------------------------------------------
Sub RemoveReference(GUID As String, Major As Integer, Minor As Integer)
          Dim Ref As Variant

10        For Each Ref In ThisWorkbook.VBProject.References
20            If Ref.GUID = GUID And Ref.Major = Major And Ref.Minor = Minor Then
30                Application.VBE.ActiveVBProject.References.Remove Ref
40            End If
50        Next
End Sub

'---------------------------------------------------------------------------------------
' Proc : ShowReferences
' Date : 4/4/2013
' Desc : Lists all VBProject references
'---------------------------------------------------------------------------------------
Sub ShowReferences()
          Dim i As Variant
          Dim n As Integer

10        ThisWorkbook.Activate
20        On Error GoTo SHEET_EXISTS
30        Sheets("VBA References").Select
40        ActiveSheet.Cells.Delete
50        On Error GoTo 0

60        [A1].Value = "Name"
70        [B1].Value = "Description"
80        [C1].Value = "GUID"
90        [D1].Value = "Major"
100       [E1].Value = "Minor"

110       For i = 1 To ThisWorkbook.VBProject.References.Count
120           n = i + 1
130           With ThisWorkbook.VBProject.References(i)
140               Cells(n, 1).Value = .Name
150               Cells(n, 2).Value = .Description
160               Cells(n, 3).Value = .GUID
170               Cells(n, 4).Value = .Major
180               Cells(n, 5).Value = .Minor
190           End With
200       Next
210       Columns.AutoFit

220       Exit Sub

SHEET_EXISTS:
230       ThisWorkbook.Sheets.Add After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count), Count:=1
240       ActiveSheet.Name = "VBA References"
250       Resume Next
End Sub

'---------------------------------------------------------------------------------------
' Proc : ReportTypeText
' Date : 4/10/2013
' Desc : Returns the report type as a string
'---------------------------------------------------------------------------------------
Function ReportTypeText(RepType As ReportType) As String
10        Select Case RepType
              Case ReportType.BO:
20                ReportTypeText = "BO"
30            Case ReportType.DS:
40                ReportTypeText = "DS"
50            Case ReportType.ALL:
60                ReportTypeText = "ALL"
70            Case ReportType.INQ:
80                ReportTypeText = "INQ"
90        End Select
End Function

'---------------------------------------------------------------------------------------
' Proc : DeleteColumn
' Date : 4/11/2013
' Desc : Removes a column based on text in the column header
'---------------------------------------------------------------------------------------
Sub DeleteColumn(HeaderText As String)
          Dim i As Integer

10        For i = ActiveSheet.UsedRange.Columns.Count To 1 Step -1
20            If Trim(Cells(1, i).Value) = HeaderText Then
30                Columns(i).Delete
40                Exit For
50            End If
60        Next
End Sub

'---------------------------------------------------------------------------------------
' Proc : FindColumn
' Date : 4/11/2013
' Desc : Returns the column number if a match is found
'---------------------------------------------------------------------------------------
Function FindColumn(ByVal HeaderText As String, Optional SearchArea As Range) As Integer
10        Dim i As Integer: i = 0
          Dim ColText As String
          
20        If TypeName(SearchArea) = "Nothing" Or TypeName(SearchArea) = Empty Then
30            Set SearchArea = Range(Cells(1, 1), Cells(1, ActiveSheet.UsedRange.Columns.Count))
40        End If

50        For i = 1 To SearchArea.Columns.Count
60            ColText = Trim(SearchArea.Cells(1, i).Value)

70            Do While InStr(ColText, "  ")
80                ColText = Replace(ColText, "  ", " ")
90            Loop

100           If ColText = HeaderText Then
110               FindColumn = i
120               Exit For
130           End If
140       Next

150       If FindColumn = 0 Then Err.Raise CustErr.COLNOTFOUND, "FindColumn", HeaderText
End Function

