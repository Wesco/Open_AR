Attribute VB_Name = "AHF_Mail"
Option Explicit

'---------------------------------------------------------------------------------------
' Dep : AHF_File
'---------------------------------------------------------------------------------------

'Pauses for x# of milliseconds
'Used for email function to prevent
'all emails from being sent at once
'Example: "Sleep 1500" will pause for 1.5 seconds
Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'---------------------------------------------------------------------------------------
' Proc  : Sub Email
' Date  : 10/11/2012
' Desc  : Sends an email using Outlook
' Ex    : Email SendTo:=email@example.com, Subject:="example email", Body:="Email Body"
'---------------------------------------------------------------------------------------
Sub Email(SendTo As String, Optional CC As String, Optional BCC As String, Optional Subject As String, Optional Body As String, Optional Attachment As Variant)
          Dim s As Variant              'Attachment string if array is passed
          Dim Mail_Object As Variant    'Outlook application object
          Dim Mail_Single As Variant    'Email object

10        Set Mail_Object = CreateObject("Outlook.Application")
20        Set Mail_Single = Mail_Object.CreateItem(0)

30        With Mail_Single
              'Add attachments
40            Select Case TypeName(Attachment)
                  Case "Variant()"
50                    For Each s In Attachment
60                        If s <> Empty Then
70                            If FileExists(s) = True Then
80                                Mail_Single.attachments.Add s
90                            End If
100                       End If
110                   Next
120               Case "String"
130                   If Attachment <> Empty Then
140                       If FileExists(Attachment) = True Then
150                           Mail_Single.attachments.Add Attachment
160                       End If
170                   End If
180           End Select

              'Setup email
190           .Subject = Subject
200           .To = SendTo
210           .CC = CC
220           .BCC = BCC
230           .HTMLbody = Body
240           On Error GoTo SEND_FAILED
250           .Send
260           On Error GoTo 0
270       End With

          'Give the email time to send
280       Sleep 1500
290       Exit Sub

SEND_FAILED:
300       With Mail_Single
310           MsgBox "Mail to '" & .To & "' could not be sent."
320           .Delete
330       End With
340       Resume Next
End Sub
