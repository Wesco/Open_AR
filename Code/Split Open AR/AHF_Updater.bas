Attribute VB_Name = "AHF_Updater"
Option Explicit

Private Enum Ver
    Major
    Minor
    Patch
End Enum

'---------------------------------------------------------------------------------------
' Proc : IncrementMajor
' Date : 9/4/2013
' Desc : Increments the macros major version number (major.minor.patch)
'---------------------------------------------------------------------------------------
Sub IncrementMajor()
10        IncrementVer Major
End Sub

'---------------------------------------------------------------------------------------
' Proc : IncrementMinorVersion
' Date : 4/24/2013
' Desc : Increments the macros minor version number (major.minor.patch)
'---------------------------------------------------------------------------------------
Sub IncrementMinor()
10        IncrementVer Minor
End Sub

'---------------------------------------------------------------------------------------
' Proc : IncrementPatch
' Date : 9/4/2013
' Desc : Increments the macros patch number (major.minor.patch)
'---------------------------------------------------------------------------------------
Sub IncrementPatch()
10        IncrementVer Patch
End Sub

'---------------------------------------------------------------------------------------
' Proc : IncrementVer
' Date : 9/4/2013
' Desc :
'---------------------------------------------------------------------------------------
Private Sub IncrementVer(Version As Ver)
          Dim Path As String
          Dim Ver As Variant
          Dim FileNum As Integer
          Dim i As Integer

10        Path = GetWorkbookPath & "Version.txt"
20        FileNum = FreeFile

30        If FileExists(Path) = True Then
40            Open Path For Input As #FileNum
50            Line Input #FileNum, Ver
60            Close FileNum

              'Split version number
70            Ver = Split(Ver, ".")

              'Increment version
80            Select Case Version
                  Case Major
90                    Ver(0) = CInt(Ver(0)) + 1
100               Case Minor
110                   Ver(1) = CInt(Ver(1)) + 1
120               Case Patch
130                   Ver(2) = CInt(Ver(2)) + 1
140           End Select

              'Combine version
150           Ver = Ver(0) & "." & Ver(1) & "." & Ver(2)

160           Open Path For Output As #FileNum
170           Print #FileNum, Ver
180           Close #FileNum
190       Else
200           Open Path For Output As #FileNum
210           Print #FileNum, "1.0.0"
220           Close #FileNum
230       End If
End Sub

'---------------------------------------------------------------------------------------
' Proc : CheckForUpdates
' Date : 4/24/2013
' Desc : Checks to see if the macro is up to date
'---------------------------------------------------------------------------------------
Sub CheckForUpdates(URL As String, Optional RepoName As String = "")
          Dim Ver As Variant
          Dim LocalVer As Variant
          Dim Path As String
          Dim LocalPath As String
          Dim FileNum As Integer
          Dim RegEx As Variant

10        Set RegEx = CreateObject("VBScript.RegExp")
20        Ver = DownloadTextFile(URL)
30        Ver = Ver & vbCrLf
40        Ver = Replace(Ver, vbLf, "")
50        Ver = Replace(Ver, vbCr, "")
60        RegEx.Pattern = "^[0-9]+\.[0-9]+\.[0-9]+$"
70        Path = GetWorkbookPath & "Version.txt"
80        FileNum = FreeFile

90        Open Path For Input As #FileNum
100       Line Input #FileNum, LocalVer
110       Close FileNum

120       If RegEx.test(Ver) Then
130           If Not Ver = LocalVer Then
140               MsgBox Prompt:="An update is available. Please close the macro and get the latest version!", Title:="Update Available"
150               If Not RepoName = "" Then
160                   Shell "C:\Program Files\Internet Explorer\iexplore.exe http://github.com/Wesco/" & RepoName & "/releases/", vbMaximizedFocus
170               End If
180           End If
190       End If
End Sub

'---------------------------------------------------------------------------------------
' Proc : DownloadTextFile
' Date : 4/25/2013
' Desc : Returns the contents of a text file from a website
'---------------------------------------------------------------------------------------
Private Function DownloadTextFile(URL As String) As String
          Dim Success As Boolean
          Dim responseText As String
          Dim oHTTP As Variant

10        Set oHTTP = CreateObject("WinHttp.WinHttpRequest.5.1")

20        oHTTP.Open "GET", URL, False
30        oHTTP.Send
40        Success = oHTTP.WaitForResponse()

50        If Not Success Then
60            DownloadTextFile = ""
70            Exit Function
80        End If

90        responseText = oHTTP.responseText
100       Set oHTTP = Nothing

110       DownloadTextFile = responseText
End Function

