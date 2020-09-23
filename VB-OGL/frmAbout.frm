VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About VB-OpenGL"
   ClientHeight    =   2700
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   5790
   ClipControls    =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1863.588
   ScaleMode       =   0  'User
   ScaleWidth      =   5437.109
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "&Vote For Me"
      Default         =   -1  'True
      Height          =   345
      Left            =   4440
      TabIndex        =   0
      Top             =   1800
      Width           =   1260
   End
   Begin VB.PictureBox picIcon 
      AutoSize        =   -1  'True
      ClipControls    =   0   'False
      Height          =   540
      Left            =   240
      Picture         =   "frmAbout.frx":0000
      ScaleHeight     =   337.12
      ScaleMode       =   0  'User
      ScaleWidth      =   337.12
      TabIndex        =   2
      Top             =   240
      Width           =   540
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Height          =   345
      Left            =   4440
      TabIndex        =   1
      Top             =   2280
      Width           =   1260
   End
   Begin VB.Label Label2 
      Caption         =   "Clicking Here"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   1755
      MouseIcon       =   "frmAbout.frx":0CCA
      MousePointer    =   99  'Custom
      TabIndex        =   6
      Top             =   1140
      Width           =   960
   End
   Begin VB.Label Label3 
      Caption         =   "If you like this code, Please Vote for me by"
      Height          =   495
      Left            =   1080
      TabIndex        =   7
      Top             =   960
      Width           =   2415
   End
   Begin VB.Label Label1 
      Caption         =   $"frmAbout.frx":0E1C
      Height          =   855
      Left            =   240
      TabIndex        =   5
      Top             =   1680
      Width           =   3975
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   84.515
      X2              =   5309.398
      Y1              =   1107.8
      Y2              =   1107.8
   End
   Begin VB.Label lblDescription 
      Caption         =   "Saadat Ali Shah , shahji_2000@yahoo.com"
      ForeColor       =   &H00000000&
      Height          =   210
      Left            =   1050
      TabIndex        =   3
      Top             =   600
      Width           =   3885
   End
   Begin VB.Label lblTitle 
      Caption         =   "VB-OpenGL 1.0"
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   1050
      TabIndex        =   4
      Top             =   240
      Width           =   3885
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   98.6
      X2              =   5309.398
      Y1              =   1118.153
      Y2              =   1118.153
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const SW_NORMAL = 1

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long


Private Sub cmdOK_Click()
  Unload Me
End Sub

Private Sub Command1_Click()
Label2_Click
End Sub

Private Sub Form_Load()
  Me.Icon = OGLWin.Icon
End Sub

Private Sub Label2_Click()
ShellExecute Me.hDC, "Open", "http://planet-source-code.com/vb/scripts/ShowCode.asp?lngWId=1&txtCodeId=47711", "", "", SW_SHOWNORMAL
End Sub
