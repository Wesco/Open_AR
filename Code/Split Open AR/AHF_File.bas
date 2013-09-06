Attribute VB_Name = "AHF_File"
Option Explicit

'---------------------------------------------------------------------------------------
' Proc  : Function FileExists
' Date  : 10/10/2012
' Type  : Boolean
' Desc  : Checks if a file exists
' Ex    : FileExists "C:\autoexec.bat"
'---------------------------------------------------------------------------------------
Function FileExists(ByVal sPath As String) As Boolean
          'Remove trailing backslash
10        If InStr(Len(sPath), sPath, "\") > 0 Then sPath = Left(sPath, Len(sPath) - 1)
          'Check to see if the directory exists and return true/false
20        If Dir(sPath, vbDirectory) <> "" Then FileExists = True
End Function

'---------------------------------------------------------------------------------------
' Proc  : Function FolderExists
' Date  : 10/10/2012
' Type  : Boolean
' Desc  : Checks if a folder exists
' Ex    : FolderExists "C:\Program Files\"
'---------------------------------------------------------------------------------------
Function FolderExists(ByVal sPath As String) As Boolean
          'Add trailing backslash
10        If InStr(Len(sPath), sPath, "\") = 0 Then sPath = sPath & "\"
          'If the folder exists return true
20        If Dir(sPath, vbDirectory) <> "" Then FolderExists = True
End Function

'---------------------------------------------------------------------------------------
' Proc  : Sub RecMkDir
' Date  : 10/10/2012
' Desc  : Creates an entire directory tree
' Ex    : RecMkDir "C:\Dir1\Dir2\Dir3\"
'---------------------------------------------------------------------------------------
Sub RecMkDir(ByVal sPath As String)
          Dim sDirArray() As String   'Folder names
          Dim sDrive As String        'Base drive
          Dim sNewPath As String      'Path builder
          Dim i As Long               'Counter

          'Add trailing slash
10        If Right(sPath, 1) <> "\" Then
20            sPath = sPath & "\"
30        End If

          'Split at each \
40        sDirArray = Split(sPath, "\")
50        sDrive = sDirArray(0) & "\"

          'Loop through each directory
60        For i = 1 To UBound(sDirArray) - 1
70            If Len(sNewPath) = 0 Then
80                sNewPath = sDrive & sNewPath & sDirArray(i) & "\"
90            Else
100               sNewPath = sNewPath & sDirArray(i) & "\"
110           End If

120           If Not FolderExists(sNewPath) Then
130               MkDir sNewPath
140           End If
150       Next
End Sub

'---------------------------------------------------------------------------------------
' Proc : DeleteFile
' Date : 3/19/2013
' Desc : Deletes a file
'---------------------------------------------------------------------------------------
Sub DeleteFile(FileName As String)
10        On Error Resume Next
20        Kill FileName
End Sub
