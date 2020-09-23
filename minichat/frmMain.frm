VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MiniChat"
   ClientHeight    =   2370
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5055
   Icon            =   "frmMain.frx":0000
   MaxButton       =   0   'False
   ScaleHeight     =   2370
   ScaleWidth      =   5055
   StartUpPosition =   2  'CenterScreen
   Begin MSWinsockLib.Winsock WSC 
      Left            =   480
      Top             =   2400
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock WS 
      Index           =   0
      Left            =   1080
      Top             =   2400
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Height          =   2415
      Left            =   0
      TabIndex        =   0
      Top             =   -60
      Width           =   5055
      Begin VB.PictureBox Picture2 
         Height          =   430
         Left            =   3720
         ScaleHeight     =   375
         ScaleWidth      =   1095
         TabIndex        =   5
         Top             =   1800
         Width           =   1160
         Begin VB.CommandButton cmdEnd 
            Caption         =   "&End Session"
            Height          =   375
            Left            =   0
            TabIndex        =   6
            Top             =   0
            Width           =   1095
         End
      End
      Begin VB.PictureBox Picture1 
         Height          =   430
         Left            =   2280
         ScaleHeight     =   375
         ScaleWidth      =   1095
         TabIndex        =   3
         Top             =   1800
         Width           =   1150
         Begin VB.CommandButton cmdChangeNick 
            Caption         =   "Change &Nick"
            Height          =   375
            Left            =   0
            TabIndex        =   4
            Top             =   0
            Width           =   1095
         End
      End
      Begin VB.TextBox txtSend 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   285
         Left            =   60
         MaxLength       =   500
         TabIndex        =   2
         Top             =   1440
         Width           =   4935
      End
      Begin VB.TextBox txtChat 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   1215
         Left            =   60
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   1
         Top             =   150
         Width           =   4935
      End
      Begin VB.Label lblNick 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   2040
         Width           =   2055
      End
      Begin VB.Label lblRemoteIP 
         Caption         =   "You are now chatting as:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   1800
         Width           =   2085
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim strRecivedData As String
    Dim SocketCount As Integer
Private Sub cmdChangeNick_Click()
    frmNick.Show
End Sub

Private Sub cmdEnd_Click()
    frmMain.WSC.SendData "<INFO>  Connection Terminated!!"
    frmMain.WS(0).Close
    frmMain.WSC.Close
    frmMain.Hide
    
    End
End Sub

Private Sub Timer1_Timer()
    lblNick.Caption = frmNick.txtNick.Text
End Sub

Private Sub Form_Load()
    txtChat.Text = txtChat.Text & "<INFO>  Welcome to MiniChat!" & vbCrLf
End Sub

Private Sub Form_Terminate()
    End
End Sub

Private Sub txtChat_Change()
    txtChat.SelStart = Len(txtChat.Text)
End Sub

Private Sub txtSend_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If txtSend.Text = "" Then
            Exit Sub
        Else
            frmMain.WSC.SendData "<" & lblNick.Caption & ">  " & txtSend.Text
            DoEvents
            txtSend.Text = ""
        End If
    End If
End Sub

Private Sub WS_ConnectionRequest(Index As Integer, ByVal requestID As Long)
    SocketCount = SocketCount + 1
    Load WS(SocketCount)
    WS(SocketCount).Accept requestID
End Sub

Private Sub WS_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    Dim strRecivedData As String
    Dim SocketCheck As Integer

    WS(Index).GetData strRecivedData
    
    For SocketCheck = 0 To SocketCount Step 1
    If WS(SocketCheck).State = sckConnected Then
        WS(SocketCheck).SendData strRecivedData
        DoEvents
    End If
    Next SocketCheck
End Sub

Private Sub WSc_DataArrival(ByVal bytesTotal As Long)
    Dim strDataRecived As String

    frmMain.WSC.GetData strDataRecived
    DoEvents
    txtChat.Text = txtChat.Text & strDataRecived & vbCrLf
End Sub
