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

    Set Mail_Object = CreateObject("Outlook.Application")
    Set Mail_Single = Mail_Object.CreateItem(0)

    With Mail_Single
        'Add attachments
        Select Case TypeName(Attachment)
            Case "Variant()"
                For Each s In Attachment
                    If s <> Empty Then
                        If FileExists(s) = True Then
                            Mail_Single.attachments.Add s
                        End If
                    End If
                Next
            Case "String"
                If Attachment <> Empty Then
                    If FileExists(Attachment) = True Then
                        Mail_Single.attachments.Add Attachment
                    End If
                End If
        End Select

        'Setup email
        .Subject = Subject
        .To = SendTo
        .CC = CC
        .BCC = BCC
        .HTMLbody = Body
        On Error GoTo SEND_FAILED
        .Send
        On Error GoTo 0
    End With

    'Give the email time to send
    Sleep 1500
    Exit Sub

SEND_FAILED:
    With Mail_Single
        MsgBox "Mail to '" & .To & "' could not be sent."
        .Delete
    End With
    Resume Next
End Sub
