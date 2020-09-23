VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form ChatMain 
   Caption         =   "Harvey Chat"
   ClientHeight    =   6015
   ClientLeft      =   1620
   ClientTop       =   1020
   ClientWidth     =   3015
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "FrmClient.frx":0000
   ScaleHeight     =   6015
   ScaleWidth      =   3015
   Begin VB.CommandButton Command5 
      Caption         =   "Font"
      Height          =   375
      Left            =   3480
      TabIndex        =   8
      Top             =   320
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1920
      Top             =   3720
   End
   Begin VB.Timer Timer1 
      Left            =   1440
      Top             =   3720
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Fore Color"
      Height          =   375
      Left            =   5760
      TabIndex        =   4
      Top             =   320
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Back Color"
      Height          =   375
      Left            =   4680
      TabIndex        =   3
      Top             =   320
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ListBox RoomList 
      Height          =   2040
      Left            =   320
      TabIndex        =   2
      Top             =   780
      Width           =   2320
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Connect"
      Height          =   372
      Left            =   360
      TabIndex        =   1
      Top             =   5280
      Width           =   2295
   End
   Begin MSWinsockLib.Winsock SockClient 
      Left            =   360
      Top             =   3720
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.PictureBox UserControl11 
      Height          =   375
      Left            =   1320
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   6
      Top             =   1080
      Width           =   375
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   840
      Top             =   3720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox txtSendData 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3360
      MousePointer    =   3  'I-Beam
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   5240
      Visible         =   0   'False
      Width           =   5260
   End
   Begin VB.TextBox TxtOutPut 
      Height          =   4215
      Left            =   3370
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   9
      Top             =   790
      Visible         =   0   'False
      Width           =   5240
   End
   Begin VB.Label Label3 
      Height          =   6015
      Left            =   3000
      TabIndex        =   7
      Top             =   0
      Width           =   6015
   End
   Begin VB.Line Line2 
      BorderColor     =   &H0000FFFF&
      BorderWidth     =   4
      Visible         =   0   'False
      X1              =   360
      X2              =   360
      Y1              =   5520
      Y2              =   5535
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      Visible         =   0   'False
      X1              =   360
      X2              =   2520
      Y1              =   5520
      Y2              =   5520
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Room List"
      BeginProperty Font 
         Name            =   "News Gothic MT"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   360
      TabIndex        =   5
      Top             =   240
      Width           =   2175
   End
   Begin VB.Menu mnufile 
      Caption         =   "&File"
      Begin VB.Menu mnuquit 
         Caption         =   "Quit"
      End
   End
   Begin VB.Menu mnuconnexion 
      Caption         =   "&Conn"
      Begin VB.Menu mnuconnect 
         Caption         =   "Connect"
      End
      Begin VB.Menu mnuoption 
         Caption         =   "Option"
      End
   End
   Begin VB.Menu mnuchat 
      Caption         =   "Ch&at"
      Begin VB.Menu mnuopenchatroom 
         Caption         =   "Open Chat Room"
      End
      Begin VB.Menu mnuclosechatroom 
         Caption         =   "Close Chat Room"
      End
   End
End
Attribute VB_Name = "ChatMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**************************************************************
'*                                                            *
'*  Chat Client Application Working whit the Specific         *
'*  Server Application comming with the package .zip.         *
'*                                                            *
'*  Creator : CARL HARVEY elterrorista@videotron.ca           *
'*  Creation : 12-22-97                                       *
'*  Last Modif :08-14-99                                      *
'*                                                            *
'*  Please Give credit to my Client engine (HARVEY ENGINE)    *
'*  Thanks and enjoy this code.                               *
'*                                                            *
'**************************************************************


Option Explicit
Dim oksend As Boolean
Dim backcol As Boolean
Dim tryconnexion As Integer
Dim indice_avance
Dim chatroom As Boolean
Dim CloseConn As Boolean
Private Sub chatopengraphic(param As Boolean)
Command2.Visible = param
Command3.Visible = param
Command5.Visible = param
TxtOutPut.Visible = param
txtSendData.Visible = param
Label3.Visible = Not param
End Sub
Private Sub Command1_Click()
Dim rep
If Command1.Caption = "Connect" Then
  connectuser
Else
  closeconnexion
End If
End Sub
Private Sub connectuser()
On Error Resume Next
If SockClient.State = 0 Then
  SockClient.RemoteHost = HostIP
  SockClient.RemotePort = HostPort
  SockClient.Connect
  Start_Stop_Animation True, True, True, False
  Me.MousePointer = vbHourglass
  mnuconnect.Caption = "Cancel"
  Do
   DoEvents
  Loop Until SockClient = 7
  Timer2.Interval = 2
  Me.MousePointer = vbNormal
  txtSendData.SetFocus
  CloseConn = False
End If
End Sub
Private Sub Start_Stop_Animation(ByVal p1 As Boolean, ByVal p2 As Boolean, ByVal p3 As Boolean, ByVal p4 As Boolean)
Timer2.Enabled = p1
Line1.Visible = p2
Line2.Visible = p3
Command1.Visible = p4
End Sub
Private Sub Command2_Click()
 TxtOutPut.BackColor = GetColor()
End Sub
Private Function GetColor() As Long
CommonDialog1.CancelError = True
 On Error GoTo errHandler
 CommonDialog1.Flags = &H1
 CommonDialog1.ShowColor
 GetColor = CommonDialog1.Color
 Exit Function
errHandler:
End Function
Private Sub Command3_Click()
TxtOutPut.ForeColor = GetColor()
End Sub
Private Sub Command5_Click()
 CommonDialog1.CancelError = True
 On Error GoTo errHandler
  CommonDialog1.Flags = &H3  ' Flags property must be set
  CommonDialog1.ShowFont  ' Display Font common dialog box.
  TxtOutPut.FontBold = CommonDialog1.FontBold
  TxtOutPut.FontItalic = CommonDialog1.FontItalic
  TxtOutPut.FontName = CommonDialog1.FontName
  TxtOutPut.FontSize = CommonDialog1.FontSize
 Exit Sub
errHandler:
End Sub
Private Sub RestoreDefault()
Command1.Caption = "Connect"
Command1.Visible = True
Connected = False
RoomList.Clear
TxtOutPut.Text = ""
mnuconnect.Caption = "Connect"
Timer2.Enabled = False
End Sub
Private Sub closeconnexion()
'On Error Resume Next
If SockClient.State = 7 Then
  SockClient.SendData "1003 LEAVE" & Username & Chr(13) & Chr(10)
End If
CloseConn = True
RestoreDefault
End Sub

Private Sub Form_Terminate()
If SockClient.State <> 0 Then closeconnexion
End Sub

Private Sub Form_Unload(Cancel As Integer)
If SockClient.State <> 0 Then closeconnexion
End Sub
Private Sub mnuclosechatroom_Click()
 Me.Width = 3150
 chatroom = False
 chatopengraphic False
End Sub
Private Sub mnuconnect_Click()
Dim rep
If mnuconnect.Caption = "Connect" Then
  connectuser
ElseIf mnuconnect.Caption = "Connect" Then
  closeconnexion
Else
 SockClient.Close
 closeconnexion
End If
End Sub

Private Sub mnuopenchatroom_Click()
  TxtOutPut = ""
  Me.Width = 9120
  chatroom = True
  chatopengraphic True
End Sub
Private Sub mnuoption_Click()
  frmLog.Show
End Sub
Private Sub mnuquit_Click()
  If SockClient.State <> 0 Then closeconnexion
  Unload Me
End Sub
Private Sub SockClient_DataArrival(ByVal bytesTotal As Long)
On Error Resume Next
Dim ChatData As String
  SockClient.GetData ChatData, vbString
  ExecuteData ChatData
  ChatData = ""
End Sub

Private Sub SockClient_SendComplete()
  If CloseConn Then SockClient.Close
End Sub

'Create the animation of connection
Private Sub Timer2_Timer()
  If Line2.X1 >= 2500 Then
    indice_avance = -24
  ElseIf Line2.X1 = 360 Then
   indice_avance = 24
  End If
  Line2.X1 = Line2.X1 + indice_avance
  Line2.X2 = Line2.X2 + indice_avance
End Sub

Private Sub txtSendData_KeyPress(KeyAscii As Integer)
On Error Resume Next
  If KeyAscii = 13 Then
    If SockClient.State = 0 Then MsgBox "You are not Connected !", vbInformation + vbOKOnly, "Expression": Exit Sub
    If SockClient.State = 7 Then
      Dim temp$
      SockClient.SendData "1005 UTALK" & Username & "@" & txtSendData.Text & Chr(13) & Chr(10)
      TxtOutPut.Text = TxtOutPut.Text & Username & ">> " & txtSendData.Text & Chr(13) & Chr(10)
      TxtOutPut.SelStart = Len(TxtOutPut)
      KeyAscii = 0
    End If
    txtSendData.Text = ""
  End If
End Sub

'***************************************************
'*If you want to check if the user is still connected
'
'Private Sub Timer1_Timer()
'  If SockClient.State <> 7 And Connected Then
'    MsgBox "You have lost your connection !"
    'Close Default thing
' End If
'End Sub
'**************************************************
Private Sub ExecuteData(ByVal batdat)
  Dim start As Integer
  Dim strcheck As String
  Dim strcheck2 As String
  Dim pos1, x
  start = 1
  Do
    pos1 = InStr(start, batdat, Chr(10), vbBinaryCompare)
    If pos1 <> 0 Then
      strcheck = Mid(batdat, start, pos1 - start + 1)
      x = pos1
      start = pos1 + 1
      strcheck2 = Left(strcheck, 10)
      Select Case strcheck2
        Case "5001 GOODU": CodeLoginOk strcheck
        Case "1001 LUSER": CodeUser strcheck       'User already in the Room
        Case "1002 UJOIN": CodeJoin strcheck       'User Join the Room
        Case "1003 LEAVE": CodeLeave strcheck      'User leave the Room
'       Case "1004 WHISP": CodeWhisper strcheck    'User whisper you
        Case "1005 UTALK": CodeTalk strcheck       'User Talk (chat text)
'       Case "1007 CHANN": CodeChannel strcheck    'Current Channel
        Case "1018 UINFO": CodeInfo strcheck
        Case "1019":                         'Error
      End Select
      DoEvents
    Else
      Exit Do
    End If
  Loop Until x = Len(batdat) - 1
End Sub

'=============================================================
'*************  WINSOCK CODE FUNCTION ************************
'=============================================================
Private Sub CodeLoginOk(ByVal strcheck As String)
SockClient.SendData "1021 INFOR " & Username & Chr(13) & Chr(10)
Start_Stop_Animation False, False, False, True
Connected = True
Command1.Caption = "DisConnect"
mnuconnect.Caption = "DisConnect"
End Sub
Private Sub CodeTalk(strdata)
Dim mypos, strname, strchat
  mypos = InStr(11, strdata, "@", vbTextCompare)
  strname = Mid(strdata, 11, mypos - 11)
  strchat = Mid(strdata, mypos + 1, Len(strdata) - (mypos + 1) - 1)
  TxtOutPut.Text = TxtOutPut.Text & strname & ">> " & strchat & Chr(13) & Chr(10)
  TxtOutPut.SelStart = Len(TxtOutPut.Text)
End Sub
Private Sub CodeInfo(strdata)
  On Error Resume Next
  Dim strchat
  strchat = Mid(strdata, 11, Len(strdata) - 12)
  TxtOutPut.Text = TxtOutPut.Text & strchat & Chr(13) & Chr(10)
  TxtOutPut.SelStart = Len(TxtOutPut.Text)
End Sub
Private Sub CodeJoin(strdata)
Dim strname
  strname = Trim(Mid(strdata, 11, Len(strdata) - 12))
  RoomList.AddItem strname
End Sub
Private Sub CodeLeave(strdata)
'On Error Resume Next
  Dim founded As Boolean, strname, i
  founded = False
  strname = Trim(Mid(strdata, 11, Len(strdata) - 12))
  For i = 0 To RoomList.ListCount
    If strname = Trim(RoomList.List(i)) Then
      founded = True
      Exit For
    End If
  Next i
  If founded Then RoomList.RemoveItem i
End Sub
Private Sub CodeUser(strdata)
Dim strname
  strname = Mid(strdata, 12, Len(strdata) - 13) '& Mid(strdata, mypos + 4, Len(strdata) - (mypos + 6) + 2)
  RoomList.AddItem strname
End Sub
'Private Sub CodeChannel(strdata)
  'Specifics routine for Channel Change
'End Sub
'Private Sub CodeWhisper(strdata)
  'Specifics routine for User Whisper
'End Sub

