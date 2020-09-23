VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmStart 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MiniChat Startup Center"
   ClientHeight    =   2670
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4935
   Icon            =   "frmStart.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2670
   ScaleWidth      =   4935
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   2700
      Left            =   0
      TabIndex        =   0
      Top             =   -50
      Width           =   4935
      Begin VB.PictureBox Picture3 
         Height          =   420
         Left            =   3600
         ScaleHeight     =   360
         ScaleWidth      =   960
         TabIndex        =   11
         Top             =   2160
         Width           =   1020
         Begin VB.CommandButton cmdCancel 
            Caption         =   "&Cancel"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   0
            TabIndex        =   12
            Top             =   0
            Width           =   975
         End
      End
      Begin TabDlg.SSTab SSTab1 
         Height          =   1815
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   3201
         _Version        =   393216
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         BackColor       =   8421504
         ForeColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "Server Options"
         TabPicture(0)   =   "frmStart.frx":0442
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "cmdHelp"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Image1"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "Label1"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "Picture2"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).ControlCount=   4
         TabCaption(1)   =   "Client Options"
         TabPicture(1)   =   "frmStart.frx":045E
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "cmdHelp2"
         Tab(1).Control(1)=   "Image2"
         Tab(1).Control(2)=   "Label2"
         Tab(1).Control(3)=   "Picture1"
         Tab(1).Control(4)=   "txtIPAddress"
         Tab(1).ControlCount=   5
         Begin VB.TextBox txtIPAddress 
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   -74040
            MaxLength       =   20
            TabIndex        =   6
            Top             =   1230
            Width           =   2415
         End
         Begin VB.PictureBox Picture1 
            Height          =   420
            Left            =   -71520
            ScaleHeight     =   360
            ScaleWidth      =   960
            TabIndex        =   4
            Top             =   1150
            Width           =   1020
            Begin VB.CommandButton cmdConnect 
               Caption         =   "&Connect"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   0
               TabIndex        =   5
               Top             =   0
               Width           =   975
            End
         End
         Begin VB.PictureBox Picture2 
            Height          =   420
            Left            =   1560
            ScaleHeight     =   360
            ScaleWidth      =   1560
            TabIndex        =   2
            Top             =   1200
            Width           =   1620
            Begin VB.CommandButton cmdStartServer 
               Caption         =   "Start a &Server"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   0
               TabIndex        =   3
               Top             =   0
               Width           =   1575
            End
         End
         Begin VB.Label Label1 
            BackColor       =   &H00404040&
            Caption         =   "To start a server with MiniChat, click the 'Start a Server' button below and wait for your incoming clients to connect."
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   615
            Left            =   840
            TabIndex        =   10
            Top             =   480
            Width           =   3615
         End
         Begin VB.Image Image1 
            Height          =   480
            Left            =   240
            Picture         =   "frmStart.frx":047A
            Top             =   480
            Width           =   480
         End
         Begin VB.Label cmdHelp 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H00404040&
            Caption         =   "Help"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF8080&
            Height          =   195
            Left            =   350
            TabIndex        =   9
            Top             =   1320
            Width           =   405
         End
         Begin VB.Label Label2 
            BackColor       =   &H00404040&
            Caption         =   "To connect to a server as a client, enter the IP address of the server you wish to connect to, and press connect."
            DataMember      =   "&H00FFFFFF&"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   795
            Left            =   -74160
            TabIndex        =   8
            Top             =   480
            Width           =   3615
         End
         Begin VB.Image Image2 
            Height          =   480
            Left            =   -74760
            Picture         =   "frmStart.frx":08BC
            Top             =   480
            Width           =   480
         End
         Begin VB.Label cmdHelp2 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H00404040&
            Caption         =   "Help"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF8080&
            Height          =   195
            Left            =   -74640
            TabIndex        =   7
            Top             =   1320
            Width           =   390
         End
      End
      Begin VB.Label Label3 
         Caption         =   "Welcome to MiniChat!  Click on one of the tab above to start chatting."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   13
         Top             =   2160
         Width           =   3375
      End
   End
End
Attribute VB_Name = "frmStart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    frmMain.WS(0).Close
    frmMain.WSC.Close
    
    End
End Sub

Private Sub cmdConnect_Click()
    If txtIPAddress.Text = "" Then
        MsgBox "MiniChat User Error 1010: Invalid IP Address!" & vbCrLf & vbCrLf & "Could not Connect to Server, no IP address was given.", vbCritical, "MiniChat User Error!!"
    Else
        If frmMain.WSC.LocalIP = Null Then
            MsgBox "MiniChat Client Error 1100:  No internet connection established!" & vbCrLf & vbCrLf & "Could not Connect to Server, Client not connected to Internet.", vbCritical, "MyChat Client Error!!"
        End If
        frmMain.WSC.RemotePort = 789
        frmMain.WSC.Connect txtIPAddress.Text
        frmStart.Hide
        frmNick.Show
    End If
End Sub

Private Sub cmdStartServer_Click()
   If frmMain.WS(0).LocalIP = Null Then
        MsgBox "MiniChat Server Error 1001:  No internet connection established!" & vbCrLf & vbCrLf & "Could not Start Server, not Connected to Internet.", vbCritical, "MiniChat Server Error!!"
    Else
        frmMain.WS(0).LocalPort = 789
        frmMain.WS(0).Listen
        frmMain.WSC.RemotePort = 789
        frmMain.WSC.Connect frmMain.WSC.LocalIP
        frmStart.Hide
        frmNick.Show
    End If
End Sub

Private Sub SSTab1_DblClick(Index As Integer)

End Sub
