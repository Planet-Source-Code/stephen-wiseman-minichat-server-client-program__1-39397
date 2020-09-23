VERSION 5.00
Begin VB.Form frmNick 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MiniChat - Choose your Nick"
   ClientHeight    =   1665
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4470
   Icon            =   "frmNick.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1665
   ScaleWidth      =   4470
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   1695
      Left            =   0
      TabIndex        =   0
      Top             =   -50
      Width           =   4455
      Begin VB.PictureBox Picture1 
         Height          =   420
         Left            =   3240
         ScaleHeight     =   360
         ScaleWidth      =   960
         TabIndex        =   4
         Top             =   1200
         Width           =   1020
         Begin VB.CommandButton cmdAccept 
            Caption         =   "&Accept"
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
      Begin VB.TextBox txtNick 
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
         Height          =   285
         Left            =   1200
         MaxLength       =   15
         TabIndex        =   3
         Top             =   1260
         Width           =   1935
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Nickname:"
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
         Height          =   195
         Left            =   240
         TabIndex        =   6
         Top             =   1275
         Width           =   915
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   120
         Picture         =   "frmNick.frx":0442
         Top             =   360
         Width           =   480
      End
      Begin VB.Label Label2 
         Caption         =   "It must be between 3 and 15 characters in length."
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
         Left            =   720
         TabIndex        =   2
         Top             =   720
         Width           =   2895
      End
      Begin VB.Label Label1 
         Caption         =   "Choose a nickname by entering it in the feild below and click Accept."
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
         Left            =   720
         TabIndex        =   1
         Top             =   240
         Width           =   3135
      End
   End
End
Attribute VB_Name = "frmNick"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAccept_Click()
    frmMain.lblNick.Caption = frmNick.txtNick.Text
    frmNick.Hide
    frmMain.Show
End Sub
