VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form OGLWin 
   Caption         =   "OpenGL with VB: Solar System"
   ClientHeight    =   6690
   ClientLeft      =   1215
   ClientTop       =   1455
   ClientWidth     =   9480
   Icon            =   "OGLWin.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   446
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   632
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Frame 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   3855
      Left            =   60
      ScaleHeight     =   3825
      ScaleWidth      =   9465
      TabIndex        =   1
      Top             =   2880
      Width           =   9495
      Begin VB.PictureBox glView2 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1935
         Left            =   100
         ScaleHeight     =   129
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   129
         TabIndex        =   13
         Top             =   100
         Width           =   1935
      End
      Begin VB.Timer Timer1 
         Interval        =   300
         Left            =   8760
         Top             =   0
      End
      Begin VB.PictureBox CtrlFrame 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1695
         Left            =   0
         ScaleHeight     =   1695
         ScaleWidth      =   9135
         TabIndex        =   2
         Top             =   2400
         Width           =   9135
         Begin MSComctlLib.ImageList ImageList1 
            Left            =   240
            Top             =   480
            _ExtentX        =   1005
            _ExtentY        =   1005
            BackColor       =   -2147483643
            ImageWidth      =   32
            ImageHeight     =   16
            MaskColor       =   12632256
            UseMaskColor    =   0   'False
            _Version        =   393216
            BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
               NumListImages   =   10
               BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "OGLWin.frx":0CCA
                  Key             =   ""
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "OGLWin.frx":0F64
                  Key             =   ""
               EndProperty
               BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "OGLWin.frx":11FD
                  Key             =   ""
               EndProperty
               BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "OGLWin.frx":1496
                  Key             =   ""
               EndProperty
               BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "OGLWin.frx":1730
                  Key             =   ""
               EndProperty
               BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "OGLWin.frx":19CA
                  Key             =   ""
               EndProperty
               BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "OGLWin.frx":1C63
                  Key             =   ""
               EndProperty
               BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "OGLWin.frx":1EFC
                  Key             =   ""
               EndProperty
               BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "OGLWin.frx":2195
                  Key             =   ""
               EndProperty
               BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "OGLWin.frx":242C
                  Key             =   ""
               EndProperty
            EndProperty
         End
         Begin MSComctlLib.ImageCombo ClrCombo 
            Height          =   330
            Left            =   2160
            TabIndex        =   38
            Top             =   480
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   582
            _Version        =   393216
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            Locked          =   -1  'True
            ImageList       =   "ImageList1"
         End
         Begin VB.CheckBox ChkPath 
            Caption         =   "&Show Planet Path"
            Height          =   255
            Left            =   6360
            TabIndex        =   7
            Top             =   525
            Value           =   1  'Checked
            Width           =   1935
         End
         Begin VB.CheckBox ChkCam 
            Caption         =   "&Floating Camera"
            Height          =   255
            Left            =   3840
            TabIndex        =   6
            Top             =   525
            Width           =   1575
         End
         Begin VB.CommandButton Command2 
            Cancel          =   -1  'True
            Caption         =   "E&xit"
            Height          =   375
            Left            =   6480
            TabIndex        =   5
            Top             =   1080
            Width           =   1335
         End
         Begin VB.CommandButton Command1 
            Caption         =   "&About"
            Default         =   -1  'True
            Height          =   375
            Left            =   3960
            TabIndex        =   4
            Top             =   1080
            Width           =   1335
         End
         Begin VB.ComboBox Combo 
            Height          =   315
            ItemData        =   "OGLWin.frx":26C1
            Left            =   960
            List            =   "OGLWin.frx":26DA
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   1080
            Width           =   1575
         End
         Begin ComctlLib.Slider Slider2 
            Height          =   420
            Left            =   6600
            TabIndex        =   8
            Top             =   -45
            Width           =   2295
            _ExtentX        =   4048
            _ExtentY        =   741
            _Version        =   327682
            Min             =   1
            Max             =   30
            SelStart        =   10
            TickStyle       =   1
            Value           =   10
         End
         Begin ComctlLib.Slider Slider1 
            Height          =   420
            Left            =   2160
            TabIndex        =   9
            Top             =   -45
            Width           =   2295
            _ExtentX        =   4048
            _ExtentY        =   741
            _Version        =   327682
            Min             =   10
            Max             =   30
            SelStart        =   20
            TickStyle       =   1
            Value           =   20
         End
         Begin VB.Label ClrLabel 
            Caption         =   "Sun Flare Color :"
            Height          =   255
            Left            =   720
            TabIndex        =   12
            Top             =   480
            Width           =   1455
         End
         Begin VB.Label RotLabel 
            AutoSize        =   -1  'True
            Caption         =   "Earth Rotations : "
            Height          =   195
            Left            =   5280
            TabIndex        =   11
            Top             =   0
            Width           =   1230
         End
         Begin VB.Label RevLabel 
            AutoSize        =   -1  'True
            Caption         =   "Earth Revolution Time(Secs) : "
            Height          =   195
            Left            =   0
            TabIndex        =   10
            Top             =   0
            Width           =   2160
         End
      End
      Begin VB.Label Captions 
         AutoSize        =   -1  'True
         Caption         =   "Name : "
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
         Index           =   0
         Left            =   2280
         TabIndex        =   37
         Top             =   120
         Width           =   735
      End
      Begin VB.Label Desc 
         AutoSize        =   -1  'True
         Height          =   255
         Index           =   0
         Left            =   3120
         TabIndex        =   36
         Top             =   120
         Width           =   135
      End
      Begin VB.Label Captions 
         AutoSize        =   -1  'True
         Caption         =   "Type : "
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
         Index           =   1
         Left            =   5280
         TabIndex        =   35
         Top             =   120
         Width           =   615
      End
      Begin VB.Label Desc 
         AutoSize        =   -1  'True
         Height          =   255
         Index           =   1
         Left            =   6240
         TabIndex        =   34
         Top             =   120
         Width           =   135
      End
      Begin VB.Label Captions 
         AutoSize        =   -1  'True
         Caption         =   "Mass : "
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
         Index           =   2
         Left            =   2280
         TabIndex        =   33
         Top             =   400
         Width           =   735
      End
      Begin VB.Label Captions 
         AutoSize        =   -1  'True
         Caption         =   "Diameter : "
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
         Index           =   3
         Left            =   5280
         TabIndex        =   32
         Top             =   400
         Width           =   975
      End
      Begin VB.Label Captions 
         AutoSize        =   -1  'True
         Caption         =   "Surface Pressure : "
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
         Index           =   4
         Left            =   2280
         TabIndex        =   31
         Top             =   680
         Width           =   1695
      End
      Begin VB.Label Captions 
         AutoSize        =   -1  'True
         Caption         =   "Density : "
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
         Index           =   5
         Left            =   5280
         TabIndex        =   30
         Top             =   680
         Width           =   855
      End
      Begin VB.Label Captions 
         AutoSize        =   -1  'True
         Caption         =   "Gravity : "
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
         Index           =   6
         Left            =   2280
         TabIndex        =   29
         Top             =   940
         Width           =   855
      End
      Begin VB.Label Captions 
         AutoSize        =   -1  'True
         Caption         =   "Escape Velocity : "
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
         Index           =   7
         Left            =   5280
         TabIndex        =   28
         Top             =   940
         Width           =   1575
      End
      Begin VB.Label Captions 
         AutoSize        =   -1  'True
         Caption         =   "MeanTemperature : "
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
         Index           =   8
         Left            =   2280
         TabIndex        =   27
         Top             =   1220
         Width           =   1815
      End
      Begin VB.Label Captions 
         AutoSize        =   -1  'True
         Caption         =   "No. of Moons : "
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
         Index           =   9
         Left            =   5280
         TabIndex        =   26
         Top             =   1220
         Width           =   1335
      End
      Begin VB.Label Desc 
         AutoSize        =   -1  'True
         Height          =   255
         Index           =   2
         Left            =   3120
         TabIndex        =   25
         Top             =   400
         Width           =   135
      End
      Begin VB.Label Desc 
         AutoSize        =   -1  'True
         Height          =   255
         Index           =   3
         Left            =   6240
         TabIndex        =   24
         Top             =   400
         Width           =   135
      End
      Begin VB.Label Desc 
         AutoSize        =   -1  'True
         Height          =   255
         Index           =   4
         Left            =   3120
         TabIndex        =   23
         Top             =   680
         Width           =   135
      End
      Begin VB.Label Desc 
         AutoSize        =   -1  'True
         Height          =   255
         Index           =   5
         Left            =   6240
         TabIndex        =   22
         Top             =   680
         Width           =   135
      End
      Begin VB.Label Desc 
         AutoSize        =   -1  'True
         Height          =   255
         Index           =   6
         Left            =   3120
         TabIndex        =   21
         Top             =   940
         Width           =   135
      End
      Begin VB.Label Desc 
         AutoSize        =   -1  'True
         Height          =   255
         Index           =   7
         Left            =   6240
         TabIndex        =   20
         Top             =   940
         Width           =   135
      End
      Begin VB.Label Desc 
         AutoSize        =   -1  'True
         Height          =   255
         Index           =   8
         Left            =   3120
         TabIndex        =   19
         Top             =   1220
         Width           =   135
      End
      Begin VB.Label Captions 
         AutoSize        =   -1  'True
         Caption         =   "Length of Day : "
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
         Index           =   10
         Left            =   2280
         TabIndex        =   18
         Top             =   1480
         Width           =   1455
      End
      Begin VB.Label Captions 
         AutoSize        =   -1  'True
         Caption         =   "Length of Year : "
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
         Index           =   11
         Left            =   5280
         TabIndex        =   17
         Top             =   1480
         Width           =   1455
      End
      Begin VB.Label Desc 
         AutoSize        =   -1  'True
         Height          =   255
         Index           =   10
         Left            =   3120
         TabIndex        =   16
         Top             =   1480
         Width           =   135
      End
      Begin VB.Label Desc 
         AutoSize        =   -1  'True
         Height          =   255
         Index           =   11
         Left            =   6240
         TabIndex        =   15
         Top             =   1480
         Width           =   135
      End
      Begin VB.Label Desc 
         AutoSize        =   -1  'True
         Height          =   255
         Index           =   9
         Left            =   6240
         TabIndex        =   14
         Top             =   1220
         Width           =   135
      End
      Begin VB.Line Line1 
         BorderColor     =   &H8000000A&
         X1              =   600
         X2              =   8520
         Y1              =   2160
         Y2              =   2160
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFFFF&
         X1              =   600
         X2              =   8520
         Y1              =   2150
         Y2              =   2150
      End
   End
   Begin VB.PictureBox glView1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2775
      Left            =   0
      ScaleHeight     =   185
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   641
      TabIndex        =   0
      Top             =   0
      Width           =   9615
   End
End
Attribute VB_Name = "OGLWin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'     This File contains code for Rendering, Animating and picking OpenGL Window      '
'                                       + VB Stuff                                    '
'         (You have)CopyRight © 2003 Saadat Ali Shah, shahji_2000@yahoo.com           '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Public Done As Boolean      'Done=True => Closing App.
Dim Rev!(2 To 5), Rot!(2 To 4), Rt!, Rt2!, Rt3! 'Rev():Revolution angle around Sun (or Earth), Rot(): Rotation angle around your own Centre
Dim Mode As Integer         'Rendering Mode or Mouse Picking Mode
Public Selected As Byte     'Identifies object thats under mouse pointer
Dim Tmr As New VbTimer      'Our very own HiRes Timer
Dim Year(2 To 4) As Double  'Time Scale: 1 Year in secs.
Dim Days_per_year(2 To 4) As Double  'implies how may Rots a planet does in a year
Dim Description(Sun To Moon, 0 To 11) As String, Tip(0 To Moon) As String
Dim TexReady As Boolean     'Waoh! this provides a way around a little Bug,see glView2_paint()
Const Rdn_to_Degree = 3.1415926535 / 180 'Radians to Degree
Dim Clr(0 To 9, 0 To 3) As Single, SunClr(0 To 3) As Single

Sub GetClr(i As Integer)
 If i < 10 Then '<Max color
  SunClr(0) = Clr(i, 0)
  SunClr(1) = Clr(i, 1)
  SunClr(2) = Clr(i, 2)
  SunClr(3) = Clr(i, 3)
Debug.Print vbCrLf & vbCrLf
Debug.Print SunClr(0) & " - " & SunClr(1) & " - " & SunClr(2) & " - " & SunClr(3)
  
 End If
End Sub

Private Sub ClrCombo_Click()
'I wanted to do this with OwnerDrawn Combobox but as u can c ImageCombo does the trick pretty well!
GetClr CInt(ClrCombo.SelectedItem.Image - 1)
End Sub

Private Sub Combo_Click()
Dim i As Byte
Selected = Combo.ItemData(Combo.ListIndex)
If Combo.ListIndex = 1 Then Combo.ListIndex = 0
If Selected = 0 Then glView2_Paint
'Show Description...
If Selected > 0 Then For i = 0 To 11: Desc(i) = Description(Selected, i): Next i
If Selected = Sun Or Selected = Moon Then
  Captions(10).Visible = False: Captions(11).Visible = False
ElseIf Selected > 0 Then
  Captions(10).Visible = True: Captions(11).Visible = True
End If
End Sub

Private Sub Command1_Click()
Tmr.Paused = True
frmAbout.Show 1
Tmr.Paused = False
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load() 'Prog actually start loading in Sub Main this time in 'XPModule.bas'
 Dim i As Integer
 'Initialize...
 Mode = GL_RENDER: Selected = 0
 Done = False: TexReady = False
 hDC1 = glView1.hDC 'Save Hdc of OpenGL Drawing Windows
 hDC2 = glView2.hDC
 Year(Earth) = 20  '1 earth year equals ... secs
 Year(Mercury) = Year(Earth) * (88 / 365) 'Extract Mercury's year length from Earth's
 Year(Venus) = Year(Earth) * (225 / 365)
 Days_per_year(Earth) = 10  'Unfortunately putting real value '365' will make ur head spin (..with our year lenght)
 Days_per_year(Mercury) = 88 / 58: Days_per_year(Venus) = 225 / 243
 
 'Some guys have suggested that i compact following code with Sub procedures etc but i thinks its more clear to user this way
 'Init string to hold data
 Description(Sun, 0) = "Sun"
 Description(Sun, 1) = "Star"
 Description(Sun, 2) = "19,891 x 10^26 Kg"
 Description(Sun, 3) = "1,392,000 Km"
 Description(Sun, 4) = "0.868 mb"
 Description(Sun, 5) = "1408 Kg/m^3"
 Description(Sun, 6) = "274 m/s^2"
 Description(Sun, 7) = "617.7 Km/s"
 Description(Sun, 8) = "5778 K"
 Description(Sun, 9) = "9"
 Description(Sun, 10) = ""
 Description(Sun, 11) = ""
 
 Description(Mercury, 0) = "Mercury"
 Description(Mercury, 1) = "Planet"
 Description(Mercury, 2) = "0.33 x 10^24 Kg"
 Description(Mercury, 3) = "4,879 Km"
 Description(Mercury, 4) = "0 bars"
 Description(Mercury, 5) = "5427 Kg/m^3"
 Description(Mercury, 6) = "3.7 m/s^2"
 Description(Mercury, 7) = "4.3 Km/s"
 Description(Mercury, 8) = "167 C"
 Description(Mercury, 9) = "0"
 Description(Mercury, 10) = "58.65 Earth days"
 Description(Mercury, 11) = "87.97 Earth days"
 
 Description(Venus, 0) = "Venus"
 Description(Venus, 1) = "Planet"
 Description(Venus, 2) = "4.87 x 10^24 Kg"
 Description(Venus, 3) = "12,104 Km"
 Description(Venus, 4) = "92 bars"
 Description(Venus, 5) = "5243 Kg/m^3"
 Description(Venus, 6) = "8.9 m/s^2"
 Description(Venus, 7) = "10.4 Km/s"
 Description(Venus, 8) = "464 C"
 Description(Venus, 9) = "0"
 Description(Venus, 10) = "243 Earth days"
 Description(Venus, 11) = "224.7 Earth days"
  
 Description(Earth, 0) = "Earth"
 Description(Earth, 1) = "Planet"
 Description(Earth, 2) = "5.97 x 10^24 Kg"
 Description(Earth, 3) = "12,756 Km"
 Description(Earth, 4) = "1 bars"
 Description(Earth, 5) = "5515 Kg/m^3"
 Description(Earth, 6) = "9.8 m/s^2"
 Description(Earth, 7) = "11.2 Km/s"
 Description(Earth, 8) = "15 C"
 Description(Earth, 9) = "1"
 Description(Earth, 10) = "23.93 Earth Hours"
 Description(Earth, 11) = "365.26 Earth days"
 
 Description(Moon, 0) = "Moon"
 Description(Moon, 1) = "Moon"
 Description(Moon, 2) = "0.073 x 10^24 Kg"
 Description(Moon, 3) = "3475 Km"
 Description(Moon, 4) = "0 bars"
 Description(Moon, 5) = "3340 Kg/m^3"
 Description(Moon, 6) = "1.6 m/s^2"
 Description(Moon, 7) = "2.4 Km/s"
 Description(Moon, 8) = "-20 C"
 Description(Moon, 9) = "0"
 Description(Moon, 10) = ""
 Description(Moon, 11) = ""
 
 Tip(0) = ""
 Tip(Sun) = "Sun"
 Tip(Mercury) = "Mercury"
 Tip(Venus) = "Venus"
 Tip(Earth) = "Earth"
 Tip(Moon) = "Moon"
 
 'Sun Flare Clrs, In L8r version we can pick any color for Sun Flare
 'But as u can c most of colors r not good choice,even now so its just for fun!
 Clr(0, 0) = 1!:    Clr(0, 1) = 0.6!:  Clr(0, 2) = 0.2!:  Clr(0, 3) = 0.55!
 Clr(1, 0) = 1!:    Clr(1, 1) = 1!:    Clr(1, 2) = 0.65!: Clr(1, 3) = 0.55!
 Clr(2, 0) = 1!:    Clr(2, 1) = 0.1!:  Clr(2, 2) = 0.1!:  Clr(2, 3) = 0.55!
 Clr(3, 0) = 0.5!:  Clr(3, 1) = 1!:    Clr(3, 2) = 0.5!:  Clr(3, 3) = 0.55!
 Clr(4, 0) = 0.25!: Clr(4, 1) = 1!:    Clr(4, 2) = 1!:    Clr(4, 3) = 0.55!
 Clr(5, 0) = 0.8!:  Clr(5, 1) = 0.8!:  Clr(5, 2) = 1!:    Clr(5, 3) = 0.55!
 Clr(6, 0) = 0.4!:  Clr(6, 1) = 0.4!:  Clr(6, 2) = 1!:    Clr(6, 3) = 0.55!
 Clr(9, 0) = 0.7!:  Clr(9, 1) = 0.7!:  Clr(9, 2) = 0.7!:  Clr(9, 3) = 0.55!
 Clr(8, 0) = 1!:    Clr(8, 1) = 1!:    Clr(8, 2) = 1!:    Clr(8, 3) = 0.55!
 Clr(7, 0) = 0.8!:  Clr(7, 1) = 0.55!: Clr(7, 2) = 0.6!:  Clr(7, 3) = 0.55!
  
 ClrCombo.ComboItems.Add 1, , "0", 1
 ClrCombo.ComboItems.Add 2, , "1", 2
 ClrCombo.ComboItems.Add 3, , "2", 3
 ClrCombo.ComboItems.Add 4, , "3", 4
 ClrCombo.ComboItems.Add 5, , "4", 5
 ClrCombo.ComboItems.Add 6, , "5", 6
 ClrCombo.ComboItems.Add 7, , "6", 7
 ClrCombo.ComboItems.Add 8, , "7", 8
 ClrCombo.ComboItems.Add 9, , "8", 9
 ClrCombo.ComboItems.Add 10, , "9", 10
 ClrCombo.Refresh
 ClrCombo.SelectedItem = ClrCombo.ComboItems(1): GetClr 0
 
 Combo.ListIndex = 0
 
 Slider1.ToolTipText = "Current Value = " & Slider1.Value: Slider2.ToolTipText = "Current Value = " & Slider2.Value
 
 Rt3 = 90 ' Initial Floating Camera Angle
 
 
 gHW = Me.hwnd      'Save handle to the form.
 'If u want to debug Code u better comment out Hook and Unhook(in Form_Unload)
 Hook               'Begin subclassing.
 If InitGL = False Then Unload Me: Exit Sub
 
 Me.Show
 Tmr.Paused = False ' start Timer, can Also use: Tmr.Start
 glView2_Paint
 
 'Take Control...
 Do
  DoEvents 'Let VB do its stuff
  If Not Done Then
    'Update Animating Vars...
    Tmr.UpdateTimer 'Get Elapsed Time
    If Tmr.ElapsedSeconds > 0 Then  'If Timer not Paused
      For i = Mercury To Earth
        If Not (Selected = i Or i = Earth) Then ' Only inc Rev if the Object is not selected, Exclude earth for now
          Rev(i) = Rev(i) + 360! / (Year(i) / Tmr.ElapsedSeconds)
          If (Rev(i) >= 360!) Then Rev(i) = 0!
        End If
        Rot(i) = Rot(i) + 360! / ((Year(i) / Days_per_year(i)) / Tmr.ElapsedSeconds)
        If (Rot(i) >= 360!) Then Rot(i) = 0!
      Next i
      If Not (Selected = Moon) Then 'Don't Revolve Earth when Moon is selected
        If Not (Selected = Earth) Then Rev(Earth) = Rev(Earth) + 360! / (Year(Earth) / Tmr.ElapsedSeconds):  If (Rev(Earth) >= 360!) Then Rev(Earth) = 0!
        Rev(Moon) = Rev(Moon) + 360! / ((Year(Earth) / 12) / Tmr.ElapsedSeconds): If (Rev(Moon) >= 360!) Then Rev(Moon) = 0!
      End If
      'Other Rotation Vars...
      Rt = Rt + 360! / (20 / Tmr.ElapsedSeconds): If (Rt >= 360!) Then Rt = 0!
      Rt2 = Rt2 + 360! / (40 / Tmr.ElapsedSeconds): If (Rt2 >= 360!) Then Rt2 = 0!
    End If
    
    'Render Views...
    glView1_Paint
    If Selected > 0 Then glView2_Paint  'Don't Draw glView2 if nothing is selected
  End If
 Loop Until Done = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
 Unhook  'Stop subclassing.
 Done = True
 
 If hglRC1 <> 0 Then
  wglMakeCurrent 0, 0
  wglDeleteContext hglRC1
 End If
 
 If hglRC2 <> 0 Then
  wglMakeCurrent 0, 0
  wglDeleteContext hglRC2
 End If
End Sub

Private Sub Form_Resize()
Dim V2w As Single, V2h As Single, h As Single, w As Single, i As Byte
glView1.Width = Me.ScaleX(Me.Width, vbTwips, vbPixels)
Frame.Width = glView1.Width - 15

glView1.Height = (Me.ScaleY(Me.Height, vbTwips, vbPixels) / 2) - 15 'minus height of window bar n border
If glView1.Height > 2 Then Frame.Height = glView1.Height - 8 ' if to avoid some error cazed when window is minimized

Frame.Top = glView1.Height + 2
V2w = glView2.ScaleX(Frame.Width, vbPixels, vbTwips) / 3
V2h = glView2.ScaleY(Frame.Height, vbPixels, vbTwips) / 2
glView2.Width = IIf(V2w < V2h, V2w, V2h)
glView2.Height = glView2.Width

' Following is quite extraneous but i have done it any way, controls r repositioned on resize
For i = 0 To 11 Step 2
  Captions(i).Left = glView2.Width + 300
  Desc(i).Left = Captions(i).Left + Captions(i).Width
  Captions(i + 1).Left = Captions(i).Left + 3000
  Desc(i + 1).Left = Captions(i + 1).Left + Captions(i + 1).Width + 50
Next i

Line1.Y1 = glView2.Height + 200: Line1.Y2 = Line1.Y1
Line2.Y1 = glView2.Height + 215: Line2.Y2 = Line2.Y1
Line1.X2 = Me.Width - 1000: Line2.X2 = Me.Width - 1000

CtrlFrame.Top = Line1.Y1 + 100
CtrlFrame.Height = Me.ScaleY(Frame.Height, vbPixels, vbTwips) - CtrlFrame.Top - 100
CtrlFrame.Left = (Me.ScaleX(Frame.Width, vbPixels, vbTwips) - CtrlFrame.Width) / 2
ClrLabel.Top = CtrlFrame.Height / 2 - 150
ClrCombo.Top = ClrLabel.Top - 80: ChkCam.Top = ClrLabel.Top: ChkPath.Top = ClrLabel.Top
Combo.Top = CtrlFrame.Height - Combo.Height - 50
Command1.Top = Combo.Top: Command2.Top = Combo.Top

wglMakeCurrent hDC1, hglRC1
glViewport 0, 0, glView1.Width, glView1.Height 'glview1 height, width r simply in pixels
glMatrixMode GL_PROJECTION
glLoadIdentity
gluPerspective 35!, glView1.Width / glView1.Height, 1!, 100!  'calculate aspect ratio of window
glMatrixMode GL_MODELVIEW
glLoadIdentity
            
wglMakeCurrent hDC2, hglRC2
'even though i made ScaleMode = pixels these guies r in twips??!
w = glView2.ScaleX(glView2.Width, vbTwips, vbPixels): h = glView2.ScaleX(glView2.Height, vbTwips, vbPixels)
glViewport 0, 0, w, h
glMatrixMode GL_PROJECTION
glLoadIdentity
gluPerspective 60!, w / h, 1!, 50! 'calculate aspect ratio of window
glMatrixMode GL_MODELVIEW
glLoadIdentity
            
End Sub
Sub Render_glView1()
Dim i As Integer, j As Double, Radius As Single

  glClear clrDepthBufferBit Or clrColorBufferBit
  glLoadIdentity 'reset modelview matrix
  
  'This Func. Sets Camera Position, U can Play with 1st 3 parameters n interchange them for different effects
  gluLookAt -Cos(Rt3 / 180 * 3.14) * 20, Sin(Rt3 / 180 * 3.14) * 5, 0, 0, 2.5, -40, 0, 1, 0
  If ChkCam.Value = vbChecked Then 'Only Update angle if ChkCam is Checked
    If Tmr.ElapsedSeconds > 0 Then Rt3 = Rt3 + 360! / (40 / Tmr.ElapsedSeconds): If (Rt3 >= 360!) Then Rt3 = 0!
  End If
  
  glTranslatef 0!, 5!, -40!: glRotatef 25, 1, 0, 0

  'Path Lines...
  If ChkPath.Value = vbChecked Then
  glColor3f 1, 0.6, 0
  For i = Mercury To Earth
    If i = Mercury Then
      Radius = 10.1
    ElseIf i = Venus Then
      Radius = 15.1
    Else
      Radius = 20.1
    End If
    glBegin GL_LINE_LOOP
    glVertex3f Sin(0) * Radius, 0, Cos(0) * Radius
    For j = 3.14 / 180 To 2 * 3.14 Step 3.14 / 60
      glVertex3f Sin(j) * Radius, 0, Cos(j) * Radius
    Next j
    glEnd
  Next i
  End If

 'Sun...
  glPushMatrix
    'glTranslatef 0!, -2.5!, 0!  'Draw Sun a bit below the Centre
    If Selected = Sun Then      'Look extra cool only if selected
      'Let there be SunShine...
      glDisable GL_DEPTH_TEST
      glEnable GL_BLEND
      glColor4fv SunClr(0) ' 1!, 0.6!, 0.2!, 0.55!
      glPushMatrix
        glBindTexture GL_TEXTURE_2D, TArray(0)
        glRotatef -25, 1, 0, 0 ' undo tilt
        glRotatef Rt, 0, 0, 1 'Rotate CCW
        glBegin GL_QUADS
          glTexCoord2f 0, 1: glVertex3f -6.5!, 6.5!, 0!
          glTexCoord2f 1, 1:  glVertex3f -6.5!, -6.5!, 0!
          glTexCoord2f 1, 0:  glVertex3f 6.5!, -6.5!, 0!
          glTexCoord2f 0, 0:  glVertex3f 6.5!, 6.5!, 0!
        glEnd
        glRotatef -Rt * 2, 0, 0, 1 'Rotate CW
        glBegin GL_QUADS
          glTexCoord2f 0, 1: glVertex3f -6.5!, 6.5!, 0!
          glTexCoord2f 1, 1:  glVertex3f -6.5!, -6.5!, 0!
          glTexCoord2f 1, 0:  glVertex3f 6.5!, -6.5!, 0!
          glTexCoord2f 0, 0:  glVertex3f 6.5!, 6.5!, 0!
        glEnd
      glPopMatrix
      glDisable GL_BLEND
      glEnable GL_DEPTH_TEST
    End If
    glPushName Sun  'Push name for picking purpose
    glEnable GL_LIGHT1: glDisable GL_LIGHT0 'Sun Uses Light1, every one else Light0
    glColor3f 1, 0.6, 0
    glBindTexture GL_TEXTURE_2D, TArray(Sun)
    glRotatef Rt, 0, 1, 0: gluSphere QObj, 3, 20, 20
    glDisable GL_LIGHT1: glEnable GL_LIGHT0
  glPopMatrix
  
  'Mercury...
  glPushMatrix
    glColor3f 0.7!, 0.7!, 0.7!
    glRotatef Rev(Mercury), 0, 1, 0: glTranslatef 10!, 0!, 0! 'Position around Sun
    glRotatef Rot(Mercury), 0!, 1!, 0!                        'Rotation around centre
    glRotatef -90, 1!, 0!, 0!                                 'upright the Texture
    glLoadName Mercury
    glBindTexture GL_TEXTURE_2D, TArray(Mercury): gluSphere QObj, 1, 30, 20
  glPopMatrix
  
  'Venus...
  glPushMatrix
    glColor3f 0.7!, 0.7!, 0.7!
    glRotatef Rev(Venus), 0, 1, 0: glTranslatef 15!, 0!, 0!
    glRotatef Rot(Venus), 0!, 1!, 0!
    glRotatef -90, 1!, 0!, 0!
    glLoadName Venus
    glBindTexture GL_TEXTURE_2D, TArray(Venus): gluSphere QObj, 1.1, 40, 30
  glPopMatrix
  
  'Earth...
  glPushMatrix
    glColor3f 0.7!, 0.7!, 0.7!
    glRotatef Rev(Earth), 0, 1, 0: glTranslatef 20!, 0!, 0!
    glPushMatrix
      glRotatef Rot(Earth), 0!, 1!, 0!
      glRotatef -90, 1!, 0!, 0!
      glLoadName Earth
      glBindTexture GL_TEXTURE_2D, TArray(Earth): gluSphere QObj, 1.5, 50, 30
      glPopName
    glPopMatrix
    If Selected = Earth Then 'The Clouds if selected
      glPushMatrix
        glEnable GL_BLEND
        glColor4f 1!, 1!, 1!, 0.3!
        glBindTexture GL_TEXTURE_2D, TArray(6)
        glRotatef -Rt, 0, 1, 0
        glRotatef -90, 1!, 0!, 0!: gluSphere QObj, 1.7, 30, 30
        glDisable GL_BLEND
      glPopMatrix
    End If
    'Moon...
    glPushMatrix
      glRotatef 45, 1, 0, 0 'Tilt Plane like: /
      glTranslatef 3! * Sin(Rdn_to_Degree * Rev(Moon)), 3! * Cos(Rdn_to_Degree * Rev(Moon)), 0!
      glColor3f 0.5!, 0.5!, 0.5!
      glRotatef Rt2, 0, 1, 0
      glRotatef -90, 1!, 0!, 0!
      glPushName Moon
      glBindTexture GL_TEXTURE_2D, TArray(Moon): gluSphere QObj, 0.6, 20, 20
    glPopMatrix
  glPopMatrix
  
 glPopName 'No more rendering Pop last Name
 
 If Mode = GL_RENDER Then SwapBuffers hDC1
End Sub

Private Sub glView1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim Hits As Long, i As Integer, Idx As Integer
Dim SelectBuf(0 To 511) As Long
Dim NameNos As Integer, MinZ As Double
Dim Viewport(0 To 3) As Long

 Mode = GL_SELECT
  
 'Setup n start Picking...
 wglMakeCurrent hDC1, hglRC1
 glSelectBuffer 512, SelectBuf(0)
 glGetIntegerv GL_VIEWPORT, Viewport(0)
 glRenderMode GL_SELECT
 glInitNames
 glMatrixMode GL_PROJECTION
 glPushMatrix   'save Original Projection Matrix
 glLoadIdentity
 gluPickMatrix X, Viewport(3) - Y, 1, 1, Viewport(0)  'Get Area around Mouse pointer
 gluPerspective 35!, Viewport(2) / Viewport(3), 1!, 100!
 glMatrixMode GL_MODELVIEW
 
 Render_glView1 'Render area around mouse to capture Hits
 
 glMatrixMode GL_PROJECTION
 glPopMatrix
 glMatrixMode GL_MODELVIEW
 glFlush
 Mode = GL_RENDER
 Hits = glRenderMode(GL_RENDER) 'Get no. of Hits
 If Not (Hits = 0) Then
  MinZ = 2147483647 'init minZ to a big value
  Idx = 0
  Selected = 0      'Nothing is selected yet
  
  'To undersatnd Follwing For Loop Remember Selection Buffer's Record Format:
  ' Rec1: |    SelectBuf(0)     | SelectBuf(1)  | SelectBuf(2)  |      SelectBuf( 3... 3+NameNos)     |
  '  | No. of Names for the Hit | Minimum depth | Maximum depth | Names for the Hit (can be 0 to ...) |
  '  ...Next Record and So on!
  ' Rec2: |SelectBuf(0 + 3 + NameNos)| So on...
    For i = 1 To Hits
    NameNos = SelectBuf(Idx)
    If (SelectBuf(Idx + 1) < MinZ) And (NameNos > 0) Then 'If a named object is closer to screen then...
      MinZ = SelectBuf(Idx + 1)
      Selected = SelectBuf(Idx + 3) 'there is only one Name/Hit in the way we render
    End If
    Idx = Idx + 3 + NameNos
  Next i
  If Selected = 0 Then
    glView2_Paint 'if hits r no good clear view
    For i = 0 To 11: Desc(i).Caption = "": Next i ' Clear Text
    Combo.ListIndex = 0
  End If
  If Selected > 0 Then Combo.ListIndex = Selected + 1
  'Show Description...
  If Selected > 0 Then For i = 0 To 11: Desc(i) = Description(Selected, i): Next i
  If Selected = Sun Or Selected = Moon Then
    Captions(10).Visible = False: Captions(11).Visible = False
  ElseIf Selected > 0 Then
    Captions(10).Visible = True: Captions(11).Visible = True
  End If
  
 Else 'if Not Hits =0
  If Selected > 0 Then 'if last time around there was a hit then
    Selected = 0
    glView2_Paint 'clear view
    For i = 0 To 11: Desc(i).Caption = "": Next i
    Combo.ListIndex = 0
  Else
    Selected = 0
  End If
 End If
 glView1.ToolTipText = Tip(Selected) 'Update ToolTip
End Sub

Private Sub glView1_Paint()
  wglMakeCurrent hDC1, hglRC1
  Render_glView1
End Sub

Private Sub glView2_Paint()
  wglMakeCurrent hDC2, hglRC2
  glClear clrColorBufferBit Or clrDepthBufferBit
  glLoadIdentity
  glTranslatef 0!, 0!, -10! 'setup camera
  'Draw Which ever Object was Selected
  If Selected = Sun Then
      glDisable GL_DEPTH_TEST
      glEnable GL_BLEND
      glColor4fv SunClr(0)
      glPushMatrix
        'Attention: OpenGL Gurus out there! can u tell whats wrong here? Mail ME!
        'Problem: if u Comment all 'If Not TexReady...' Lines and ur Computer is not Super
        'Fast (over 1 Ghz). If u Select Sun b4 Selecting any thing else, u will notice
        'Sun Flare FLASH as its not Colored with my specifies color its simply white.
        'But after a Quadric object ( gluSphere() ) is drawn 1st time its ok thereafter.
        'Worst part is there was no problem until a day or two b4 finshing this project.
        
        If Not TexReady Then glScalef 0.1, 0.1, 0.1: TexReady = True
        glBindTexture GL_TEXTURE_2D, TArray2(0)
        'gluSphere QObj, 0.1!, 2, 2
        glRotatef Rt, 0, 0, 1
        glBegin GL_QUADS
          glTexCoord2f 0, 1: glVertex3f -8!, 8!, 0!
          glTexCoord2f 1, 1:  glVertex3f -8!, -8!, 0!
          glTexCoord2f 1, 0:  glVertex3f 8!, -8!, 0!
          glTexCoord2f 0, 0:  glVertex3f 8!, 8!, 0!
        glEnd
        glRotatef -Rt * 2, 0, 0, 1
        glBegin GL_QUADS
          glTexCoord2f 0, 1: glVertex3f -8!, 8!, 0!
          glTexCoord2f 1, 1:  glVertex3f -8!, -8!, 0!
          glTexCoord2f 1, 0:  glVertex3f 8!, -8!, 0!
          glTexCoord2f 0, 0:  glVertex3f 8!, 8!, 0!
        glEnd
      glPopMatrix
      glDisable GL_BLEND
      glEnable GL_DEPTH_TEST
    glBindTexture GL_TEXTURE_2D, TArray2(Sun)
    glRotatef Rt2, 0, 1, 0
    glRotatef -90, 1!, 0!, 0!: gluSphere QObj, 3.5, 30, 30
    'glColor4f 1!, 1!, 1!, 1!
    
  ElseIf Selected = Mercury Then
    glColor4f 1!, 1!, 1!, 1!
    glBindTexture GL_TEXTURE_2D, TArray2(Mercury)
    glRotatef Rot(Mercury), 0, 1, 0
    glRotatef -90, 1!, 0!, 0!: gluSphere QObj, 3.5, 30, 30
    If Not TexReady Then TexReady = True
    
  ElseIf Selected = Venus Then
    glColor4f 1!, 1!, 1!, 1!
    glBindTexture GL_TEXTURE_2D, TArray2(Venus)
    glRotatef Rot(Venus), 0, 1, 0
    glRotatef -90, 1!, 0!, 0!: gluSphere QObj, 3.5, 30, 30
    If Not TexReady Then TexReady = True
    
  ElseIf Selected = Moon Then
    'Give Moon some Halo
    gluSphere QObj, 0.1!, 2, 2
    glDisable GL_DEPTH_TEST
    glEnable GL_BLEND
    glColor4f 0.9!, 0.9!, 0.9!, 0.6!
    glPushMatrix
      If Not TexReady Then TexReady = True
      glBindTexture GL_TEXTURE_2D, TArray2(7)
      glRotatef Rt, 0, 0, 1
      glBegin GL_QUADS
        glTexCoord2f 0, 1: glVertex3f -4.2!, 4.2!, 0!
        glTexCoord2f 1, 1:  glVertex3f -4.2!, -4.2!, 0!
        glTexCoord2f 1, 0:  glVertex3f 4.2!, -4.2!, 0!
        glTexCoord2f 0, 0:  glVertex3f 4.2!, 4.2!, 0!
      glEnd
      glColor4f 1!, 1!, 1!, 0.6!
      glBindTexture GL_TEXTURE_2D, TArray2(0)
      glRotatef -Rt * 2, 0, 0, 1
      glBegin GL_QUADS
        glTexCoord2f 0, 1: glVertex3f -5!, 5!, 0!
        glTexCoord2f 1, 1:  glVertex3f -5!, -5!, 0!
        glTexCoord2f 1, 0:  glVertex3f 5!, -5!, 0!
        glTexCoord2f 0, 0:  glVertex3f 5!, 5!, 0!
      glEnd
    glPopMatrix
    glDisable GL_BLEND
    glEnable GL_DEPTH_TEST
    'Moon itself
    glRotatef Rt2, 0, 1, 0
    glColor4f 1!, 1!, 1!, 1!
    glBindTexture GL_TEXTURE_2D, TArray2(Moon)
    glRotatef -90, 1!, 0!, 0!: gluSphere QObj, 3, 30, 30
    
  ElseIf Selected = Earth Then
    glColor4f 1!, 1!, 1!, 1!
    glBindTexture GL_TEXTURE_2D, TArray2(Earth)
    glPushMatrix
    glRotatef Rot(Earth), 0, 1, 0
    glRotatef -90, 1!, 0!, 0!: gluSphere QObj, 3.5, 30, 30
    glBindTexture GL_TEXTURE_2D, TArray2(6)
    glPopMatrix
    glEnable GL_BLEND
    glColor4f 1!, 1!, 1!, 0.3!
    glRotatef -Rot(Earth), 0, 1, 0
    glRotatef -90, 1!, 0!, 0!: gluSphere QObj, 3.7, 30, 30
    glDisable GL_BLEND
  End If
  
  SwapBuffers hDC2
End Sub

Private Sub Slider1_Change()
Slider1.ToolTipText = "Current Value = " & Slider1.Value
Year(Earth) = Slider1.Value
Year(Mercury) = Year(Earth) * (88 / 365)
 Year(Venus) = Year(Earth) * (225 / 365)
End Sub

Private Sub Slider2_Change()
Slider2.ToolTipText = "Current Value = " & Slider2.Value
Days_per_year(Earth) = Slider2.Value
End Sub

Private Sub Timer1_Timer()
  If Not Done Then Me.Caption = "OpenGL with VB: Solar System  [FPS : " & Format(Tmr.FPS, "###.##") & "]"
End Sub
