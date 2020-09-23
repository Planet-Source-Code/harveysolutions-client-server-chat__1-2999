VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmserver 
   Caption         =   "Chat Server"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5250
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   -1  'True
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   3135
   ScaleWidth      =   5250
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   3135
      Left            =   0
      ScaleHeight     =   3075
      ScaleWidth      =   5175
      TabIndex        =   2
      Top             =   0
      Width           =   5230
      Begin VB.ListBox List1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1650
         Left            =   0
         TabIndex        =   5
         Top             =   960
         Width           =   5175
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Open Server"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2640
         TabIndex        =   4
         Top             =   0
         Width           =   2535
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3960
         TabIndex        =   3
         Text            =   "1005"
         Top             =   2640
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Welcome Message"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   0
         TabIndex        =   6
         Top             =   0
         Width           =   2655
      End
      Begin VB.Label Label1 
         Caption         =   "User Online"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Port To Listen:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2400
         TabIndex        =   7
         Top             =   2760
         Width           =   1575
      End
   End
   Begin VB.PictureBox UserControl11 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   720
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   1
      Top             =   8760
      Width           =   375
   End
   Begin VB.TextBox Txtdata 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   4560
      Visible         =   0   'False
      Width           =   6135
   End
   Begin MSWinsockLib.Winsock sckServer 
      Index           =   0
      Left            =   360
      Top             =   4080
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "frmserver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**************************************************************
'*                                                            *
'*  Chat Server Application Working whit the Specific         *
'*  Client Application comming with the package .zip.         *
'*                                                            *
'*  Creator : CARL HARVEY elterrorista@videotron.ca           *
'*  Creation : 12-22-97                                       *
'*  Last Modif :08-14-99                                      *
'*                                                            *
'*  Please Give credit to my server engine (HARVEY ENGINE)    *
'*  Thanks and enjoy this code.                                *
'*                                                            *
'**************************************************************
'Option Explicit
Private TotalSocket As Long
Dim useronline() As Boolean
Dim PortToListen As String

Private Sub Command1_Click()
SendToAllOthers 0, "1018 UINFOWelcome To My Server" & Chr(13) & Chr(10)
End Sub

Private Sub Command2_Click()
On Error Resume Next
Dim i
If Command2.Caption = "Open Server" Then
 If PortToListen = "" Then PortToListen = InputBox("Enter a Port To listen on !" & Chr(13) & "From :[ 1000 to 65000 ]", "Server Info")
 sckServer(0).LocalPort = PortToListen
 sckServer(0).Listen
 Command2.Caption = "Close Server"
Else
 sckServer(0).Close
 For i = 1 To TotalSocket
   sckServer(i).Close
   Unload sckServer(i)
 Next
 List1.Clear
 Command2.Caption = "Open Server"
End If
End Sub
Private Sub Form_Load()
PortToListen = Text1.Text
TotalSocket = 0
ReDim useronline(0)
End Sub

Private Sub sckServer_ConnectionRequest(index As Integer, ByVal RequestID As Long)
If index = 0 Then GetConnection RequestID
End Sub
Private Sub GetConnection(ByVal RequestID)
On Error Resume Next
  Dim pointeur
  pointeur = 0
  For i = 1 To TotalSocket 'Search for a new Socket
    If Not useronline(i) Then pointeur = i: Exit For
  Next i
  If pointeur = 0 Then  'If no socket available
    TotalSocket = TotalSocket + 1 'add a socket for the TotalSocket variable
    pointeur = TotalSocket         'Point to the last socket
    ReDim Preserve useronline(TotalSocket)
  End If
  useronline(pointeur) = True
  Load sckServer(pointeur)    'create a new cocket
  sckServer(pointeur).LocalPort = 0
  sckServer(pointeur).Accept RequestID
  sckServer(pointeur).SendData "5001 GOODU " & Chr(13) & Chr(10)
  GiveInfo pointeur, "1018 UINFOWelcome To My Server" & Chr(13) & Chr(10) & "1018 UINFOThere are currently :" & List1.ListCount + 1 & " Users online" & Chr(13) & Chr(10)
End Sub

Private Sub GiveUserList(ByVal pointer)
 Dim Stringtemp As String, i
 For i = 0 To List1.ListCount - 1
   Stringtemp = Stringtemp & "1001 LUSER " & List1.List(i) & Chr(13) & Chr(10)
 Next i
 sckServer(pointer).SendData Stringtemp
End Sub
Private Sub SendToAllOthers(ByVal index As Integer, ByVal InfoStr)
 Dim i
 For i = 1 To TotalSocket
  If i <> index And useronline(i) Then
   sckServer(i).SendData InfoStr & Chr(13) & Chr(10)
   DoEvents
  End If
 Next i
End Sub
Private Sub GiveInfo(ByVal pointer, ByVal InfoStr)
  sckServer(pointer).SendData InfoStr
End Sub
Private Sub ExecuteData(ByVal index, ByVal data As String)
  Dim pos1, strcheck As String, strcheck2 As String
  pos1 = InStr(1, data, Chr(13), vbBinaryCompare)
  If pos1 <> 0 Then
      strcheck = Mid(data, 1, pos1 - start + 1)
      X = pos1
      strcheck2 = Left(data, 10)
      Select Case strcheck2
        Case "1003 LEAVE": CodeLeave index, strcheck        'User leave the Room
        Case "1005 UTALK": CodeTalk index, strcheck         'User Talk (chat text)
        Case "1021 INFOR": CodeReturnInfo index, strcheck   'User Return his info
       'Case "1004 WHISP": GiveInfo index, strcheck         'User whisper you
       'Case "1007 CHANN": CodeChannel strcheck             'Current Channel
     End Select
  End If
End Sub
Private Sub CodeTalk(ByVal index, ByVal data)
'Treat User Chat Text Here If you Want
SendToAllOthers index, data
End Sub
Private Sub CodeReturnInfo(ByVal index, ByVal data)
mypos = InStr(11, data, Chr(13), vbBinaryCompare)
strname = Trim(Mid(data, 11, mypos - 11))
List1.AddItem strname
SendToAllOthers index, "1002 UJOIN " & strname
GiveUserList index
End Sub
Private Sub sckServer_DataArrival(index As Integer, ByVal bytesTotal As Long)
On Error Resume Next
Dim ChatData As String
sckServer(index).GetData ChatData, vbString
ExecuteData index, ChatData
ChatData = ""
End Sub
Private Sub CodeLeave(ByVal index, ByVal data)
Dim mypos, strname
mypos = InStr(11, data, Chr(13), vbBinaryCompare)
strname = Trim(Mid(data, 11, mypos - 11))

For i = 0 To List1.ListCount - 1
  If strname = List1.List(i) Then
    List1.RemoveItem i
    Unload sckServer(index)
    useronline(index) = False
    If index = TotalSocket Then TotalSocket = TotalSocket - 1
    Exit For
  End If
Next
SendToAllOthers 0, "1003 LEAVE " & strname
End Sub
