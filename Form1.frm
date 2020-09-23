VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2715
   ClientLeft      =   2820
   ClientTop       =   1515
   ClientWidth     =   4530
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   2715
   ScaleWidth      =   4530
   Begin VB.CommandButton Command1 
      Caption         =   "Close"
      Height          =   405
      Left            =   2955
      TabIndex        =   3
      Top             =   735
      Width           =   900
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1455
      Left            =   360
      ScaleHeight     =   1425
      ScaleWidth      =   1995
      TabIndex        =   0
      Top             =   255
      Width           =   2025
      Begin SHDocVwCtl.WebBrowser WebBrowser1 
         Height          =   1770
         Left            =   -255
         TabIndex        =   1
         Top             =   -270
         Width           =   2625
         ExtentX         =   4630
         ExtentY         =   3122
         ViewMode        =   0
         Offline         =   0
         Silent          =   0
         RegisterAsBrowser=   0
         RegisterAsDropTarget=   1
         AutoArrange     =   0   'False
         NoClientEdge    =   0   'False
         AlignLeft       =   0   'False
         ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
         Location        =   ""
      End
   End
   Begin VB.Label Label1 
      Caption         =   "A work around way to use an animated gif on a form"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   375
      TabIndex        =   2
      Top             =   1860
      Width           =   3555
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Public PicPath As String

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Load()
 
PicPath = App.Path & "\ani.html"
WebBrowser1.Navigate PicPath

End Sub

Private Sub Form_Resize()
WebBrowser1.Height = Me.Height

End Sub

