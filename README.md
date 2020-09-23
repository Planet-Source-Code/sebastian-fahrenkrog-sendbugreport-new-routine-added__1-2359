<div align="center">

## SendBugReport  NEW ROUTINE ADDED


</div>

### Description

Do you ever want to have a easy possibility to get in contact with your users? Here it is! You just have to add the form to your projekt and config it before you compile your projekt! Your users just have to write their comment or bug report in a textbox and hit the send button. You will love this!

I ADDED A NEW ROUTINE TO PREVENT TIMEOUTS!!
 
### More Info
 
You must config it (before you compile it) with your personal data, like:

E-Mail Adress

E-Mail Server

Subjekt Line

...etc.

See the code section for more info's

Just copy the code below and paste it in the notepad! Save it as SendBug.frm and and add it to your projekt...

It send an E-Mail after you hit the Send Button!

Mail me if you find any!


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Sebastian Fahrenkrog](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/sebastian-fahrenkrog.md)
**Level**          |Unknown
**User Rating**    |5.0 (5 globes from 1 user)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Internet/ HTML](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/internet-html__1-34.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/sebastian-fahrenkrog-sendbugreport-new-routine-added__1-2359/archive/master.zip)





### Source Code

```
'Save it as SendBug.frm and compile it!
'-------------------8< Cut here ---------------------------------------
VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1
  BorderStyle   =  0 'Kein
  Caption     =  "Send Bug Report"
  ClientHeight  =  3195
  ClientLeft   =  0
  ClientTop    =  0
  ClientWidth   =  4680
  LinkTopic    =  "Form1"
  MaxButton    =  0  'False
  MinButton    =  0  'False
  ScaleHeight   =  3195
  ScaleWidth   =  4680
  StartUpPosition =  2 'Bildschirmmitte
  Begin MSWinsockLib.Winsock Winsock1
   Left      =  120
   Top       =  120
   _ExtentX    =  741
   _ExtentY    =  741
   _Version    =  393216
  End
  Begin VB.CommandButton Exit
   Caption     =  "Exit"
   Height     =  255
   Left      =  2280
   TabIndex    =  2
   Top       =  2880
   Width      =  2295
  End
  Begin VB.CommandButton Connect
   Caption     =  "Send Bug Report"
   Height     =  255
   Left      =  120
   TabIndex    =  1
   Top       =  2880
   Width      =  2055
  End
  Begin VB.TextBox Bugreporttxt
   Height     =  2655
   Left      =  120
   MultiLine    =  -1 'True
   TabIndex    =  0
   Top       =  120
   Width      =  4455
  End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private bTrans As Boolean
Private m_iStage As Integer
Private strData As String
'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'CHANGE THIS SETTING LIKE YOU NEED IT
'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Private Const mailserver As String = "your-mail-server.com"
Private Const Tobox As String = "youre-mail@adress.com"
Private Const Frombox As String = "theuser@ofthisprogram.com"
Private Const Subject As String = "Heading of the E-Mail send to you!"
'***************************************************************
'Routine for connecting to the server
'***************************************************************
Private Sub Connect_Click()
If Winsock1.State <> sckClosed Then Winsock1.Close
Winsock1.LocalPort = 0
Winsock1.Protocol = sckTCPProtocol
Winsock1.Connect mailserver, "25"
bTrans = True
m_iStage = 0
strData = ""
Call WaitForResponse
End Sub
'***************************************************************
'Transmit the E-Mail
'***************************************************************
Private Sub Transmit(iStage As Integer)
Dim Helo As String, temp As String
Dim pos As Integer
Select Case m_iStage
Case 1:
Helo = Frombox
pos = Len(Helo) - InStr(Helo, "@")
Helo = Right$(Helo, pos)
Winsock1.SendData "HELO " & Helo & vbCrLf
strData = ""
Call WaitForResponse
Case 2:
Winsock1.SendData "MAIL FROM: <" & Trim(Frombox) & ">" & vbCrLf
Call WaitForResponse
Case 3:
Winsock1.SendData "RCPT TO: <" & Trim(Tobox) & ">" & vbCrLf
Call WaitForResponse
Case 4:
Winsock1.SendData "DATA" & vbCrLf
Call WaitForResponse
Case 5:
'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'If you want additional Headers like Date,Message-Id,...etc. !
'simply add them below                   !
'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
temp = temp & "From: " & Frombox & vbNewLine
temp = temp & "To: " & Tobox & vbNewLine
temp = temp & "Subject: " & Subject & vbNewLine
'Header + Message
temp = temp & vbCrLf & Bugreporttxt.Text
'Send the Message & close connection
Winsock1.SendData temp
Winsock1.SendData vbCrLf & "." & vbCrLf
m_iStage = 0
bTrans = False
Call WaitForResponse
End Select
End Sub
'***************************************************************
'Routine for Winsock Errors
'***************************************************************
Private Sub Winsock1_Error(ByVal number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
MsgBox "Error:" & Description, vbOKOnly, "Winsock Error!" ' Show error message
If Winsock1.State <> sckClosed Then
Winsock1.Close
End If
End Sub
'***************************************************************
'Routine for arraving Data
'***************************************************************
Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
Dim messagesent As String
On Error Resume Next
Winsock1.GetData strData, vbString
'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'!If you have problems with sending the E-Mail, you should   !
'!activate the line below and add a Textbox txtStatus, to   !
'!see the Server's response                  !
'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'txtStatus.Text = txtStatus.Text & strData
If bTrans Then
m_iStage = m_iStage + 1
Transmit m_iStage
Else
  If Winsock1.State <> sckClosed Then Winsock1.Close
  messagesent = MsgBox("Bug report sent! Hit exit to end program.", vbOKOnly, "Bug Report")
End If
End Sub
'**************************************************************
'NEW! Waits until time out, while waiting for response
'**************************************************************
Sub WaitForResponse()
Dim Start As Long
Dim Tmr As Long
Start = Timer
While Len(strData) = 0
  Tmr = Timer - Start
  DoEvents ' Let System keep checking for incoming response
  'Wait 50 seconds for response
  If Tmr > 50 Then
    MsgBox "SMTP service error, timed out while waiting for response", 64, "Error!"
    strData = ""
    End
  End If
Wend
End Sub
Private Sub Exit_Click()
On Error Resume Next
If Winsock1.State <> sckClosed Then Winsock1.Close
End
End Sub
```

