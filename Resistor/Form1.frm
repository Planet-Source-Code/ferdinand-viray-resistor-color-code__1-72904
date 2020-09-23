VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H80000012&
   Caption         =   "Resisitor Color Code "
   ClientHeight    =   8640
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   9195
   FillColor       =   &H008080FF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   8640
   ScaleWidth      =   9195
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H80000012&
      Caption         =   "Resistor Type"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1935
      Left            =   840
      TabIndex        =   32
      Top             =   240
      Width           =   7095
      Begin VB.OptionButton Option2 
         BackColor       =   &H80000012&
         Caption         =   "5 - Band"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   4200
         TabIndex        =   34
         Top             =   600
         Width           =   1935
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H80000012&
         Caption         =   "4 -Band"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   1200
         TabIndex        =   33
         Top             =   600
         Width           =   2175
      End
      Begin VB.Image Image7 
         Height          =   525
         Left            =   3120
         Picture         =   "Form1.frx":0000
         Top             =   1080
         Visible         =   0   'False
         Width           =   3900
      End
      Begin VB.Image Image5 
         Height          =   480
         Left            =   120
         Picture         =   "Form1.frx":6AE6
         Top             =   1080
         Visible         =   0   'False
         Width           =   3960
      End
      Begin VB.Image Image3 
         Height          =   375
         Left            =   3720
         Picture         =   "Form1.frx":CE28
         Top             =   1200
         Width           =   2985
      End
      Begin VB.Image Image2 
         Height          =   375
         Left            =   360
         Picture         =   "Form1.frx":10902
         Top             =   1200
         Width           =   3015
      End
      Begin VB.Image Image1 
         Height          =   255
         Left            =   1320
         Top             =   2640
         Width           =   975
      End
   End
   Begin VB.Timer Timer2 
      Left            =   10320
      Top             =   2880
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   9960
      Top             =   3120
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   4
      Left            =   7440
      TabIndex        =   17
      Top             =   2640
      Width           =   1575
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   3
      Left            =   5760
      TabIndex        =   16
      Top             =   2640
      Width           =   1575
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   2
      Left            =   3960
      TabIndex        =   15
      Top             =   2640
      Width           =   1695
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   1
      Left            =   2160
      TabIndex        =   14
      Top             =   2640
      Width           =   1695
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   0
      Left            =   240
      TabIndex        =   9
      Top             =   2640
      Width           =   1815
   End
   Begin VB.Image Image6 
      Height          =   390
      Left            =   5160
      Picture         =   "Form1.frx":14440
      Top             =   7800
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FF8080&
      Height          =   1215
      Left            =   1800
      Top             =   7200
      Width           =   5055
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   6960
      TabIndex        =   38
      Top             =   6360
      Width           =   585
   End
   Begin VB.Image Image12 
      Height          =   765
      Left            =   6240
      Picture         =   "Form1.frx":145FA
      Top             =   6240
      Width           =   2175
   End
   Begin VB.Image Image11 
      Height          =   765
      Left            =   5880
      Picture         =   "Form1.frx":14F2C
      Top             =   6240
      Width           =   3105
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Reset"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   4080
      TabIndex        =   37
      Top             =   6360
      Width           =   810
   End
   Begin VB.Image Image10 
      Height          =   765
      Left            =   3480
      Picture         =   "Form1.frx":15A6A
      Top             =   6240
      Width           =   2175
   End
   Begin VB.Image Image9 
      Height          =   765
      Left            =   3000
      Picture         =   "Form1.frx":1639C
      Top             =   6240
      Visible         =   0   'False
      Width           =   3105
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Get Value"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   960
      TabIndex        =   36
      Top             =   6360
      Width           =   1410
   End
   Begin VB.Image Image8 
      Height          =   765
      Left            =   600
      Picture         =   "Form1.frx":16EDA
      Top             =   6240
      Width           =   2175
   End
   Begin VB.Line Line28 
      BorderWidth     =   20
      Index           =   2
      X1              =   6360
      X2              =   6600
      Y1              =   6000
      Y2              =   6240
   End
   Begin VB.Line Line24 
      BorderColor     =   &H0000C000&
      BorderStyle     =   3  'Dot
      BorderWidth     =   2
      X1              =   8040
      X2              =   6000
      Y1              =   3960
      Y2              =   3960
   End
   Begin VB.Line Line23 
      BorderColor     =   &H0000C000&
      BorderStyle     =   3  'Dot
      BorderWidth     =   2
      X1              =   6000
      X2              =   6000
      Y1              =   3960
      Y2              =   4680
   End
   Begin VB.Line Line16 
      BorderColor     =   &H0000C000&
      BorderStyle     =   3  'Dot
      BorderWidth     =   2
      X1              =   8040
      X2              =   8040
      Y1              =   3000
      Y2              =   3960
   End
   Begin VB.Line Line28 
      BorderWidth     =   17
      Index           =   3
      X1              =   2880
      X2              =   3120
      Y1              =   4320
      Y2              =   4560
   End
   Begin VB.Line Line28 
      BorderWidth     =   20
      Index           =   1
      X1              =   1440
      X2              =   1680
      Y1              =   6000
      Y2              =   6240
   End
   Begin VB.Line Line26 
      BorderWidth     =   11
      Index           =   2
      X1              =   3000
      X2              =   2760
      Y1              =   6000
      Y2              =   6240
   End
   Begin VB.Line Line26 
      BorderWidth     =   11
      Index           =   1
      X1              =   1800
      X2              =   1560
      Y1              =   4320
      Y2              =   4560
   End
   Begin VB.Label lblT 
      Height          =   375
      Left            =   9960
      TabIndex        =   26
      Top             =   4080
      Width           =   615
   End
   Begin VB.Label lblc 
      Height          =   375
      Left            =   9960
      TabIndex        =   25
      Top             =   3600
      Width           =   615
   End
   Begin VB.Line Line29 
      BorderWidth     =   20
      X1              =   8160
      X2              =   7800
      Y1              =   6000
      Y2              =   6360
   End
   Begin VB.Line Line28 
      BorderWidth     =   17
      Index           =   0
      X1              =   7800
      X2              =   8160
      Y1              =   4200
      Y2              =   4560
   End
   Begin VB.Line Line27 
      BorderWidth     =   11
      X1              =   7320
      X2              =   7440
      Y1              =   2520
      Y2              =   2640
   End
   Begin VB.Line Line26 
      BorderWidth     =   11
      Index           =   0
      X1              =   6720
      X2              =   6480
      Y1              =   4320
      Y2              =   4560
   End
   Begin VB.Line Line25 
      BorderWidth     =   18
      X1              =   9000
      X2              =   8760
      Y1              =   2520
      Y2              =   2760
   End
   Begin VB.Label lblTol 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   5520
      TabIndex        =   13
      Top             =   7800
      Width           =   1335
   End
   Begin VB.Label lblVal 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   690
      Left            =   1800
      TabIndex        =   12
      Top             =   7680
      Width           =   3015
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000012&
      Caption         =   "Tolerance"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   405
      Index           =   6
      Left            =   4800
      TabIndex        =   11
      Top             =   7200
      Width           =   1350
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000012&
      Caption         =   "Value"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   405
      Index           =   5
      Left            =   2640
      TabIndex        =   10
      Top             =   7200
      Width           =   750
   End
   Begin VB.Line Line22 
      BorderColor     =   &H8000000E&
      BorderStyle     =   3  'Dot
      BorderWidth     =   2
      Index           =   2
      X1              =   8040
      X2              =   9600
      Y1              =   5520
      Y2              =   5520
   End
   Begin VB.Line Line21 
      BorderColor     =   &H0000C000&
      BorderStyle     =   3  'Dot
      BorderWidth     =   2
      Index           =   2
      X1              =   4200
      X2              =   4200
      Y1              =   3720
      Y2              =   4680
   End
   Begin VB.Line Line20 
      BorderColor     =   &H0000C000&
      BorderStyle     =   3  'Dot
      BorderWidth     =   2
      Index           =   2
      X1              =   4200
      X2              =   4680
      Y1              =   3720
      Y2              =   3720
   End
   Begin VB.Line Line22 
      BorderColor     =   &H8000000E&
      BorderStyle     =   3  'Dot
      BorderWidth     =   2
      Index           =   1
      X1              =   6720
      X2              =   7800
      Y1              =   6120
      Y2              =   6120
   End
   Begin VB.Line Line21 
      BorderColor     =   &H0000C000&
      BorderStyle     =   3  'Dot
      BorderWidth     =   2
      Index           =   1
      X1              =   3000
      X2              =   3000
      Y1              =   3000
      Y2              =   3720
   End
   Begin VB.Line Line20 
      BorderColor     =   &H0000C000&
      BorderStyle     =   3  'Dot
      BorderWidth     =   2
      Index           =   1
      X1              =   4920
      X2              =   4920
      Y1              =   3720
      Y2              =   4680
   End
   Begin VB.Line Line22 
      BorderColor     =   &H8000000E&
      BorderStyle     =   3  'Dot
      BorderWidth     =   2
      Index           =   0
      X1              =   8040
      X2              =   9600
      Y1              =   5040
      Y2              =   5040
   End
   Begin VB.Line Line21 
      BorderColor     =   &H0000C000&
      BorderStyle     =   3  'Dot
      BorderWidth     =   2
      Index           =   0
      X1              =   1080
      X2              =   1080
      Y1              =   3000
      Y2              =   3720
   End
   Begin VB.Line Line20 
      BorderColor     =   &H0000C000&
      BorderStyle     =   3  'Dot
      BorderWidth     =   2
      Index           =   0
      X1              =   1080
      X2              =   2520
      Y1              =   3720
      Y2              =   3720
   End
   Begin VB.Line Line19 
      BorderColor     =   &H0000C000&
      BorderStyle     =   3  'Dot
      BorderWidth     =   2
      X1              =   6480
      X2              =   6480
      Y1              =   3000
      Y2              =   3720
   End
   Begin VB.Line Line18 
      BorderColor     =   &H0000C000&
      BorderStyle     =   3  'Dot
      BorderWidth     =   2
      X1              =   2520
      X2              =   2520
      Y1              =   3720
      Y2              =   4440
   End
   Begin VB.Line Line17 
      BorderColor     =   &H0000C000&
      BorderStyle     =   3  'Dot
      BorderWidth     =   2
      X1              =   4920
      X2              =   6480
      Y1              =   3720
      Y2              =   3720
   End
   Begin VB.Line Line15 
      BorderColor     =   &H0000C000&
      BorderStyle     =   3  'Dot
      BorderWidth     =   2
      X1              =   3480
      X2              =   3480
      Y1              =   3720
      Y2              =   4680
   End
   Begin VB.Line Line14 
      BorderColor     =   &H0000C000&
      BorderStyle     =   3  'Dot
      BorderWidth     =   2
      X1              =   3000
      X2              =   3480
      Y1              =   3720
      Y2              =   3720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000012&
      Caption         =   "Multiplier"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   4
      Left            =   6000
      TabIndex        =   8
      Top             =   2400
      Width           =   885
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000012&
      Caption         =   "Tolerance"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   3
      Left            =   7680
      TabIndex        =   7
      Top             =   2400
      Width           =   945
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000012&
      Caption         =   "3rd  Band"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   2
      Left            =   4200
      TabIndex        =   6
      Top             =   2400
      Width           =   930
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000012&
      Caption         =   "2nd Band"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   1
      Left            =   2520
      TabIndex        =   5
      Top             =   2400
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000012&
      Caption         =   "1st Band"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   240
      Index           =   0
      Left            =   600
      TabIndex        =   4
      Top             =   2400
      Width           =   855
   End
   Begin VB.Line Line13 
      BorderColor     =   &H8000000E&
      BorderWidth     =   2
      X1              =   8040
      X2              =   8040
      Y1              =   4680
      Y2              =   5880
   End
   Begin VB.Line Line12 
      BorderColor     =   &H8000000E&
      BorderWidth     =   2
      X1              =   0
      X2              =   1560
      Y1              =   5040
      Y2              =   5040
   End
   Begin VB.Line Line11 
      BorderColor     =   &H0000C000&
      BorderWidth     =   2
      Index           =   0
      X1              =   4680
      X2              =   4680
      Y1              =   3000
      Y2              =   3720
   End
   Begin VB.Line Line10 
      BorderColor     =   &H8000000E&
      BorderWidth     =   2
      Index           =   0
      X1              =   0
      X2              =   1560
      Y1              =   5520
      Y2              =   5520
   End
   Begin VB.Line Line9 
      BorderColor     =   &H8000000E&
      BorderWidth     =   2
      X1              =   7800
      X2              =   8040
      Y1              =   4440
      Y2              =   4680
   End
   Begin VB.Line Line8 
      BorderColor     =   &H8000000E&
      BorderWidth     =   2
      Index           =   1
      X1              =   6720
      X2              =   7800
      Y1              =   4440
      Y2              =   4440
   End
   Begin VB.Line Line6 
      BorderColor     =   &H8000000E&
      BorderWidth     =   2
      Index           =   1
      X1              =   8040
      X2              =   7800
      Y1              =   5880
      Y2              =   6120
   End
   Begin VB.Line Line3 
      BorderColor     =   &H8000000E&
      BorderWidth     =   2
      Index           =   1
      X1              =   6720
      X2              =   6480
      Y1              =   4440
      Y2              =   4680
   End
   Begin VB.Line Line2 
      BorderColor     =   &H8000000E&
      BorderWidth     =   2
      Index           =   1
      X1              =   6480
      X2              =   6720
      Y1              =   5880
      Y2              =   6120
   End
   Begin VB.Line Line8 
      BorderColor     =   &H8000000E&
      BorderWidth     =   2
      Index           =   0
      X1              =   1560
      X2              =   1560
      Y1              =   4680
      Y2              =   5880
   End
   Begin VB.Line Line7 
      BorderColor     =   &H8000000E&
      BorderWidth     =   2
      Index           =   0
      X1              =   1560
      X2              =   1800
      Y1              =   5880
      Y2              =   6120
   End
   Begin VB.Line Line6 
      BorderColor     =   &H8000000E&
      BorderWidth     =   2
      Index           =   0
      X1              =   1800
      X2              =   1560
      Y1              =   4440
      Y2              =   4680
   End
   Begin VB.Line Line5 
      BorderColor     =   &H8000000E&
      BorderWidth     =   2
      Index           =   0
      X1              =   2760
      X2              =   1800
      Y1              =   6120
      Y2              =   6120
   End
   Begin VB.Line Line4 
      BorderColor     =   &H8000000E&
      BorderWidth     =   2
      Index           =   0
      X1              =   2760
      X2              =   1800
      Y1              =   4440
      Y2              =   4440
   End
   Begin VB.Line Line3 
      BorderColor     =   &H8000000E&
      BorderWidth     =   2
      Index           =   0
      X1              =   3000
      X2              =   2760
      Y1              =   5880
      Y2              =   6120
   End
   Begin VB.Line Line2 
      BorderColor     =   &H8000000E&
      BorderWidth     =   2
      Index           =   0
      X1              =   2760
      X2              =   3000
      Y1              =   4440
      Y2              =   4680
   End
   Begin VB.Line Line1 
      BorderColor     =   &H8000000E&
      BorderWidth     =   2
      Index           =   1
      X1              =   6480
      X2              =   3000
      Y1              =   5880
      Y2              =   5880
   End
   Begin VB.Line Line1 
      BorderColor     =   &H8000000E&
      BorderWidth     =   2
      Index           =   0
      X1              =   6480
      X2              =   3000
      Y1              =   4680
      Y2              =   4680
   End
   Begin VB.Label Label10 
      BackColor       =   &H00008080&
      Height          =   1215
      Index           =   1
      Left            =   3600
      TabIndex        =   19
      Top             =   4680
      Width           =   255
   End
   Begin VB.Label Label10 
      BackColor       =   &H00008080&
      Height          =   1215
      Index           =   2
      Left            =   4320
      TabIndex        =   20
      Top             =   4680
      Width           =   255
   End
   Begin VB.Label Label10 
      BackColor       =   &H00008080&
      Height          =   1215
      Index           =   3
      Left            =   5040
      TabIndex        =   21
      Top             =   4680
      Width           =   735
   End
   Begin VB.Label Label9 
      BackColor       =   &H00404040&
      Height          =   375
      Index           =   0
      Left            =   0
      TabIndex        =   18
      Top             =   5040
      Width           =   1575
   End
   Begin VB.Label Label12 
      BackColor       =   &H00C0C0C0&
      Height          =   255
      Index           =   0
      Left            =   0
      TabIndex        =   24
      Top             =   5280
      Width           =   1575
   End
   Begin VB.Label Label10 
      BackColor       =   &H00008080&
      Height          =   1695
      Index           =   6
      Left            =   1560
      TabIndex        =   27
      Top             =   4440
      Width           =   735
   End
   Begin VB.Label Label9 
      BackColor       =   &H00404040&
      Height          =   375
      Index           =   1
      Left            =   8040
      TabIndex        =   29
      Top             =   5040
      Width           =   1575
   End
   Begin VB.Label Label12 
      BackColor       =   &H00C0C0C0&
      Height          =   255
      Index           =   1
      Left            =   8040
      TabIndex        =   30
      Top             =   5280
      Width           =   1575
   End
   Begin VB.Label Label10 
      BackColor       =   &H00008080&
      Height          =   1695
      Index           =   4
      Left            =   6480
      TabIndex        =   22
      Top             =   4440
      Width           =   1575
   End
   Begin VB.Label Label2 
      ForeColor       =   &H000000FF&
      Height          =   1695
      Left            =   2280
      TabIndex        =   0
      Top             =   4440
      Width           =   495
   End
   Begin VB.Label Label10 
      BackColor       =   &H00008080&
      Height          =   1695
      Index           =   7
      Left            =   2640
      TabIndex        =   28
      Top             =   4440
      Width           =   375
   End
   Begin VB.Label Label10 
      BackColor       =   &H00008080&
      Height          =   1215
      Index           =   0
      Left            =   3000
      TabIndex        =   31
      Top             =   4680
      Width           =   135
   End
   Begin VB.Label Label3 
      Height          =   1215
      Left            =   3120
      TabIndex        =   23
      Top             =   4680
      Width           =   615
   End
   Begin VB.Label Label4 
      Height          =   1215
      Left            =   3840
      TabIndex        =   1
      Top             =   4680
      Width           =   495
   End
   Begin VB.Label Label5 
      Height          =   1215
      Left            =   4560
      TabIndex        =   2
      Top             =   4680
      Width           =   495
   End
   Begin VB.Label Label6 
      Height          =   1215
      Left            =   5760
      TabIndex        =   3
      Top             =   4680
      Width           =   495
   End
   Begin VB.Label Label10 
      BackColor       =   &H00008080&
      Height          =   1215
      Index           =   5
      Left            =   6240
      TabIndex        =   35
      Top             =   4680
      Width           =   375
   End
   Begin VB.Image Image4 
      Height          =   765
      Left            =   120
      Picture         =   "Form1.frx":1780C
      Top             =   6240
      Visible         =   0   'False
      Width           =   3105
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim viray As Double
Private Sub Combo1_Click(Index As Integer)

viray = Val(lblc.Caption)
If Option2.Value = True Then
    If Combo1(0).Text = "Black" Then
    Label2.BackColor = vbBlack
    lblc.Caption = "0"
    ElseIf Combo1(0).Text = "Brown" Then
    Label2.BackColor = &H4080&
    lblc.Caption = "1"
    ElseIf Combo1(0).Text = "Red" Then
    Label2.BackColor = &HFF&
    lblc.Caption = "2"
    ElseIf Combo1(0).Text = "Orange" Then
    Label2.BackColor = &H80FF&
    lblc.Caption = "3"
    ElseIf Combo1(0).Text = "Yellow" Then
    Label2.BackColor = &HFFFF&
    lblc.Caption = "4"
    ElseIf Combo1(0).Text = "Green" Then
    Label2.BackColor = &HFF00&
    lblc.Caption = "5"
    ElseIf Combo1(0).Text = "Blue" Then
    Label2.BackColor = &HC00000
    lblc.Caption = "6"
    ElseIf Combo1(0).Text = "Violet" Then
    Label2.BackColor = &HC000C0
    lblc.Caption = "7"
    ElseIf Combo1(0).Text = "Gray" Then
    Label2.BackColor = &H808080
    lblc.Caption = "8"
    ElseIf Combo1(0).Text = "White" Then
    Label2.BackColor = vbWhite
    lblc.Caption = "9"
    End If

    If Combo1(1).Text = "Black" Then
    Label3.BackColor = vbBlack
    lblc.Caption = lblc.Caption & "0"
    ElseIf Combo1(1).Text = "Brown" Then
    Label3.BackColor = &H4080&
    lblc.Caption = lblc.Caption & "1"
    ElseIf Combo1(1).Text = "Red" Then
    Label3.BackColor = &HFF&
    lblc.Caption = lblc.Caption & "2"
    ElseIf Combo1(1).Text = "Orange" Then
    Label3.BackColor = &H80FF&
    lblc.Caption = lblc.Caption & "3"
    ElseIf Combo1(1).Text = "Yellow" Then
    Label3.BackColor = &HFFFF&
    lblc.Caption = lblc.Caption & "4"
    ElseIf Combo1(1).Text = "Green" Then
    Label3.BackColor = &HFF00&
    lblc.Caption = lblc.Caption & "5"
    ElseIf Combo1(1).Text = "Blue" Then
    Label3.BackColor = &HC00000
    lblc.Caption = lblc.Caption & "6"
    ElseIf Combo1(1).Text = "Violet" Then
    Label3.BackColor = &HC000C0
    lblc.Caption = lblc.Caption & "7"
    ElseIf Combo1(1).Text = "Gray" Then
    Label3.BackColor = &H808080
    lblc.Caption = lblc.Caption & "8"
    ElseIf Combo1(1).Text = "White" Then
    Label3.BackColor = vbWhite
    lblc.Caption = lblc.Caption & "9"
    End If

    If Combo1(2).Text = "Black" Then
    Label4.BackColor = vbBlack
    lblc.Caption = lblc.Caption & "0"
    ElseIf Combo1(2).Text = "Brown" Then
    Label4.BackColor = &H4080&
    lblc.Caption = lblc.Caption & "1"
    ElseIf Combo1(2).Text = "Red" Then
    Label4.BackColor = &HFF&
    lblc.Caption = lblc.Caption & "2"
    ElseIf Combo1(2).Text = "Orange" Then
    Label4.BackColor = &H80FF&
    lblc.Caption = lblc.Caption & "3"
    ElseIf Combo1(2).Text = "Yellow" Then
    Label4.BackColor = &HFFFF&
    lblc.Caption = lblc.Caption & "4"
    ElseIf Combo1(2).Text = "Green" Then
    Label4.BackColor = &HFF00&
    lblc.Caption = lblc.Caption & "5"
    ElseIf Combo1(2).Text = "Blue" Then
    Label4.BackColor = &HC00000
    lblc.Caption = lblc.Caption & "6"
    ElseIf Combo1(2).Text = "Violet" Then
    Label4.BackColor = &HC000C0
    lblc.Caption = lblc.Caption & "7"
    ElseIf Combo1(2).Text = "Gray" Then
    Label4.BackColor = &H808080
    lblc.Caption = lblc.Caption & "8"
    ElseIf Combo1(2).Text = "White" Then
    Label4.BackColor = vbWhite
    lblc.Caption = lblc.Caption & "9"
    End If

    If Combo1(3).Text = "Black" Then
    Label5.BackColor = vbBlack
    lblc.Caption = viray * 1
    ElseIf Combo1(3).Text = "Brown" Then
    Label5.BackColor = &H4080&
    lblc.Caption = viray * 10
    ElseIf Combo1(3).Text = "Red" Then
    Label5.BackColor = &HFF&
    lblc.Caption = viray * 100
    ElseIf Combo1(3).Text = "Orange" Then
    Label5.BackColor = &H80FF&
    lblc.Caption = viray * 1000
    ElseIf Combo1(3).Text = "Yellow" Then
    Label5.BackColor = &HFFFF&
    lblc.Caption = viray * 10000
    ElseIf Combo1(3).Text = "Green" Then
    Label5.BackColor = &HFF00&
    lblc.Caption = viray * 100000
    ElseIf Combo1(3).Text = "Blue" Then
    Label5.BackColor = &HC00000
    lblc.Caption = viray * 1000000
    ElseIf Combo1(3).Text = "Violet" Then
    Label5.BackColor = &HC000C0
    lblc.Caption = viray * 10000000
    ElseIf Combo1(3).Text = "Gray" Then
    Label5.BackColor = &H808080
    lblc.Caption = viray * 100000000
    ElseIf Combo1(3).Text = "White" Then
    Label5.BackColor = vbWhite
    lblc.Caption = viray * 100000000
    ElseIf Combo1(3).Text = "Gold" Then
    Label5.BackColor = &H80C0FF
    lblc.Caption = viray * 0.1
    ElseIf Combo1(3).Text = "Silver" Then
    Label5.BackColor = &HC0C0C0
    lblc.Caption = viray * 0.01
    End If
    
    If Combo1(4).Text = "Black" Then
    Label6.BackColor = vbBlack
    ElseIf Combo1(4).Text = "Brown" Then
    Label6.BackColor = &H4080&
    lblT.Caption = "1 %"
    ElseIf Combo1(4).Text = "Red" Then
    Label6.BackColor = &HFF&
    lblT.Caption = "2 %"
    ElseIf Combo1(4).Text = "Orange" Then
    Label6.BackColor = &H80FF&
    ElseIf Combo1(4).Text = "Yellow" Then
    Label6.BackColor = &HFFFF&
    ElseIf Combo1(4).Text = "Green" Then
    Label6.BackColor = &HFF00&
    lblT.Caption = "0.5 %"
    ElseIf Combo1(4).Text = "Blue" Then
    Label6.BackColor = &HC00000
    lblT.Caption = "0.25 %"
    ElseIf Combo1(4).Text = "Violet" Then
    Label6.BackColor = &HC000C0
    lblT.Caption = "0.10 %"
    ElseIf Combo1(4).Text = "Gray" Then
    Label6.BackColor = &H808080
    lblT.Caption = "0.05 %"
    ElseIf Combo1(4).Text = "White" Then
    Label6.BackColor = vbWhite
    ElseIf Combo1(4).Text = "Gold" Then
    Label6.BackColor = &H80C0FF
    lblT.Caption = "5 %"
    ElseIf Combo1(4).Text = "Silver" Then
    Label6.BackColor = &HC0C0C0
    lblT.Caption = "10 %"
    ElseIf Combo1(4).Text = "No Color" Then
    Label6.BackColor = &HC0C0C0
    lblT.Caption = "20 %"
    End If
End If
If Option1.Value = True Then
    If Combo1(0).Text = "Black" Then
    Label2.BackColor = vbBlack
    lblc.Caption = "0"
    ElseIf Combo1(0).Text = "Brown" Then
    Label2.BackColor = &H4080&
    lblc.Caption = "1"
    ElseIf Combo1(0).Text = "Red" Then
    Label2.BackColor = &HFF&
    lblc.Caption = "2"
    ElseIf Combo1(0).Text = "Orange" Then
    Label2.BackColor = &H80FF&
    lblc.Caption = "3"
    ElseIf Combo1(0).Text = "Yellow" Then
    Label2.BackColor = &HFFFF&
    lblc.Caption = "4"
    ElseIf Combo1(0).Text = "Green" Then
    Label2.BackColor = &HFF00&
    lblc.Caption = "5"
    ElseIf Combo1(0).Text = "Blue" Then
    Label2.BackColor = &HC00000
    lblc.Caption = "6"
    ElseIf Combo1(0).Text = "Violet" Then
    Label2.BackColor = &HC000C0
    lblc.Caption = "7"
    ElseIf Combo1(0).Text = "Gray" Then
    Label2.BackColor = &H808080
    lblc.Caption = "8"
    ElseIf Combo1(0).Text = "White" Then
    Label2.BackColor = vbWhite
    lblc.Caption = "9"
    End If

    If Combo1(1).Text = "Black" Then
    Label3.BackColor = vbBlack
    lblc.Caption = lblc.Caption & "0"
    ElseIf Combo1(1).Text = "Brown" Then
    Label3.BackColor = &H4080&
    lblc.Caption = lblc.Caption & "1"
    ElseIf Combo1(1).Text = "Red" Then
    Label3.BackColor = &HFF&
    lblc.Caption = lblc.Caption & "2"
    ElseIf Combo1(1).Text = "Orange" Then
    Label3.BackColor = &H80FF&
    lblc.Caption = lblc.Caption & "3"
    ElseIf Combo1(1).Text = "Yellow" Then
    Label3.BackColor = &HFFFF&
    lblc.Caption = lblc.Caption & "4"
    ElseIf Combo1(1).Text = "Green" Then
    Label3.BackColor = &HFF00&
    lblc.Caption = lblc.Caption & "5"
    ElseIf Combo1(1).Text = "Blue" Then
    Label3.BackColor = &HC00000
    lblc.Caption = lblc.Caption & "6"
    ElseIf Combo1(1).Text = "Violet" Then
    Label3.BackColor = &HC000C0
    lblc.Caption = lblc.Caption & "7"
    ElseIf Combo1(1).Text = "Gray" Then
    Label3.BackColor = &H808080
    lblc.Caption = lblc.Caption & "8"
    ElseIf Combo1(1).Text = "White" Then
    Label3.BackColor = vbWhite
    lblc.Caption = lblc.Caption & "9"
    End If
    
    If Combo1(2).Text = "Black" Then
    Label4.BackColor = vbBlack
    lblc.Caption = viray * 1
    ElseIf Combo1(2).Text = "Brown" Then
    Label4.BackColor = &H4080&
    lblc.Caption = viray * 10
    ElseIf Combo1(2).Text = "Red" Then
    Label4.BackColor = &HFF&
    lblc.Caption = viray * 100
    ElseIf Combo1(2).Text = "Orange" Then
    Label4.BackColor = &H80FF&
    lblc.Caption = viray * 1000
    ElseIf Combo1(2).Text = "Yellow" Then
    Label4.BackColor = &HFFFF&
    lblc.Caption = viray * 10000
    ElseIf Combo1(2).Text = "Green" Then
    Label4.BackColor = &HFF00&
    lblc.Caption = viray * 100000
    ElseIf Combo1(2).Text = "Blue" Then
    Label4.BackColor = &HC00000
    lblc.Caption = viray * 1000000
    ElseIf Combo1(2).Text = "Violet" Then
    Label4.BackColor = &HC000C0
    lblc.Caption = viray * 10000000
    ElseIf Combo1(2).Text = "Gray" Then
    Label4.BackColor = &H808080
    lblc.Caption = viray * 100000000
    ElseIf Combo1(2).Text = "White" Then
    Label4.BackColor = vbWhite
    lblc.Caption = viray * 100000000
    ElseIf Combo1(2).Text = "Gold" Then
    Label4.BackColor = &H80C0FF
    lblc.Caption = viray * 0.1
    ElseIf Combo1(2).Text = "Silver" Then
    Label4.BackColor = &HC0C0C0
    lblc.Caption = viray * 0.01
    End If
    
    If Combo1(4).Text = "Black" Then
    Label6.BackColor = vbBlack
    ElseIf Combo1(4).Text = "Brown" Then
    Label6.BackColor = &H4080&
    lblT.Caption = "1 %"
    ElseIf Combo1(4).Text = "Red" Then
    Label6.BackColor = &HFF&
    lblT.Caption = "2 %"
    ElseIf Combo1(4).Text = "Orange" Then
    Label6.BackColor = &H80FF&
    ElseIf Combo1(4).Text = "Yellow" Then
    Label6.BackColor = &HFFFF&
    ElseIf Combo1(4).Text = "Green" Then
    Label6.BackColor = &HFF00&
    lblT.Caption = "0.5 %"
    ElseIf Combo1(4).Text = "Blue" Then
    Label6.BackColor = &HC00000
    lblT.Caption = "0.25 %"
    ElseIf Combo1(4).Text = "Violet" Then
    Label6.BackColor = &HC000C0
    lblT.Caption = "0.10 %"
    ElseIf Combo1(4).Text = "Gray" Then
    Label6.BackColor = &H808080
    lblT.Caption = "0.05 %"
    ElseIf Combo1(4).Text = "White" Then
    Label6.BackColor = vbWhite
    ElseIf Combo1(4).Text = "Gold" Then
    Label6.BackColor = &H80C0FF
    lblT.Caption = "5 %"
    ElseIf Combo1(4).Text = "Silver" Then
    Label6.BackColor = &HC0C0C0
    lblT.Caption = "10 %"
    ElseIf Combo1(4).Text = "No Color" Then
    Label6.BackColor = &HC0C0C0
    lblT.Caption = "20 %"
    End If
End If
End Sub



Private Sub Command1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command1.BackColor = vbRed
Command2.BackColor = vbWhite
End Sub

Private Sub Command2_Click()


End Sub

Private Sub Command2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command2.BackColor = vbRed
Command1.BackColor = vbWhite
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image10.Visible = True
Image9.Visible = False
Image8.Visible = True
Image4.Visible = False
Label7.FontSize = 15
Label7.ForeColor = vbBlack
Label8.ForeColor = vbBlack
Label8.FontSize = 15
Image12.Visible = True
Image11.Visible = False
Label11.FontSize = 15
Label11.ForeColor = vbBlack
End Sub

Private Sub Image10_Click()
Combo1(0).Text = ""
Combo1(1).Text = ""
Combo1(2).Text = ""
Combo1(3).Text = ""
Combo1(4).Text = ""
lblVal.Caption = ""
lblTol.Caption = ""
lblc.Caption = ""
lblT.Caption = ""
End Sub

Private Sub Image11_Click()
Unload Me
End Sub

Private Sub Image4_Click()
lblVal.Caption = lblc.Caption & "  ohms"
 lblTol.Caption = lblT.Caption
Image6.Visible = True
End Sub

Private Sub Image8_Click()
lblVal.Caption = lblc.Caption & "  ohms"
 lblTol.Caption = lblT.Caption

End Sub

Private Sub Image9_Click()
Combo1(0).Text = ""
Combo1(1).Text = ""
Combo1(2).Text = ""
Combo1(3).Text = ""
Combo1(4).Text = ""
lblVal.Caption = ""
lblTol.Caption = ""
lblc.Caption = ""
lblT.Caption = ""
Label2.BackColor = vbWhite
Label3.BackColor = vbWhite
Label4.BackColor = vbWhite
Label5.BackColor = vbWhite
Label6.BackColor = vbWhite
Image6.Visible = False
End Sub

Private Sub Label11_Click()
Unload Me
End Sub

Private Sub Label11_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image10.Visible = True
Image9.Visible = False
Image8.Visible = True
Image4.Visible = False
Label7.FontSize = 15
Label7.ForeColor = vbBlack
Label8.ForeColor = vbBlack
Label8.FontSize = 15
Image11.Visible = True
Image12.Visible = False
Label11.FontSize = 18
Label11.ForeColor = vbRed
End Sub

Private Sub Label7_Click()
If Option2.Value = True Then
    If Combo1(0).Text = "" Then
    MsgBox "1st Band Uncolored!!...", vbCritical, "Insufficient Data!"
    ElseIf Combo1(1).Text = "" Then
    MsgBox "2nd Band Uncolored!!...", vbCritical, "Insufficient Data!"
    ElseIf Combo1(2).Text = "" Then
    MsgBox "3rd Band Uncolored!!...", vbCritical, "Insufficient Data!"
    ElseIf Combo1(3).Text = "" Then
    MsgBox "Multiplier,  Uncolored!!...", vbCritical, "Insufficient Data!"
    ElseIf Combo1(4).Text = "" Then
    MsgBox "Tolerance, Uncolored!!...", vbCritical, "Insufficient Data!"
    Else
    lblVal.Caption = lblc.Caption & " ohms"
    lblTol.Caption = lblT.Caption
    Image6.Visible = True
    End If
End If

If Option1.Value = True Then
    If Combo1(0).Text = "" Then
    MsgBox "1st Band, Uncolored!!...", vbCritical, "Insufficient Data!"
    ElseIf Combo1(1).Text = "" Then
    MsgBox "2nd Band, Uncolored!!...", vbCritical, "Insufficient Data!"
    ElseIf Combo1(2).Text = "" Then
    MsgBox "Multiplier, Uncolored!!...", vbCritical, "Insufficient Data!"
    ElseIf Combo1(4).Text = "" Then
    MsgBox "Tolerance, Uncolored!!...", vbCritical, "Insufficient Data!"
    Else
    lblVal.Caption = lblc.Caption & " ohms"
    lblTol.Caption = lblT.Caption
    Image6.Visible = True
    End If
End If


End Sub

Private Sub Label7_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image4.Visible = True
Image8.Visible = False
Image10.Visible = True
Image9.Visible = False
Image11.Visible = False
Image12.Visible = True
Label7.FontSize = 18
Label7.ForeColor = vbRed
Label8.ForeColor = vbBlack
Label8.FontSize = 15
Label11.FontSize = 15
Label11.ForeColor = vbBlack
End Sub

Private Sub Label8_Click()
Combo1(0).Text = ""
Combo1(1).Text = ""
Combo1(2).Text = ""
Combo1(3).Text = ""
Combo1(4).Text = ""
lblVal.Caption = ""
lblTol.Caption = ""
lblc.Caption = ""
lblT.Caption = ""
Label2.BackColor = vbWhite
Label3.BackColor = vbWhite
Label4.BackColor = vbWhite
Label5.BackColor = vbWhite
Label6.BackColor = vbWhite
Image6.Visible = False
End Sub

Private Sub Label8_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image10.Visible = False
Image9.Visible = True
Image8.Visible = True
Image4.Visible = False
Label7.FontSize = 15
Label7.ForeColor = vbBlack
Label8.ForeColor = vbRed
Label8.FontSize = 18
Image11.Visible = False
Image12.Visible = True
Label11.FontSize = 15
Label11.ForeColor = vbBlack
End Sub

Private Sub Option2_Click()
If Option2.Value = True Then
Line17.Visible = True
Line19.Visible = True
Line20(1).Visible = True
Label5.BackColor = &H8000000F
Combo1(3).Visible = True
Label1(2).Caption = "3rd Band"
Label1(4).Caption = "Multiplier"
Label1(3).Caption = "Tolerance"
Image3.Visible = False
Image7.Visible = True
Image2.Visible = True
Image5.Visible = False

Combo1(0).AddItem ("Black")
Combo1(0).AddItem ("Brown")
Combo1(0).AddItem ("Red")
Combo1(0).AddItem ("Orange")
Combo1(0).AddItem ("Yellow")
Combo1(0).AddItem ("Green")
Combo1(0).AddItem ("Blue")
Combo1(0).AddItem ("Violet")
Combo1(0).AddItem ("Gray")
Combo1(0).AddItem ("White")


Combo1(1).AddItem ("Black")
Combo1(1).AddItem ("Brown")
Combo1(1).AddItem ("Red")
Combo1(1).AddItem ("Orange")
Combo1(1).AddItem ("Yellow")
Combo1(1).AddItem ("Green")
Combo1(1).AddItem ("Blue")
Combo1(1).AddItem ("Violet")
Combo1(1).AddItem ("Gray")
Combo1(1).AddItem ("White")

Combo1(2).AddItem ("Black")
Combo1(2).AddItem ("Brown")
Combo1(2).AddItem ("Red")
Combo1(2).AddItem ("Orange")
Combo1(2).AddItem ("Yellow")
Combo1(2).AddItem ("Green")
Combo1(2).AddItem ("Blue")
Combo1(2).AddItem ("Violet")
Combo1(2).AddItem ("Gray")
Combo1(2).AddItem ("White")


Combo1(3).AddItem ("Black")
Combo1(3).AddItem ("Brown")
Combo1(3).AddItem ("Red")
Combo1(3).AddItem ("Orange")
Combo1(3).AddItem ("Yellow")
Combo1(3).AddItem ("Green")
Combo1(3).AddItem ("Blue")
Combo1(3).AddItem ("Violet")
Combo1(3).AddItem ("Gray")
Combo1(3).AddItem ("White")
Combo1(3).AddItem ("Gold")
Combo1(3).AddItem ("Silver")

Combo1(4).AddItem ("Black")
Combo1(4).AddItem ("Brown")
Combo1(4).AddItem ("Red")
Combo1(4).AddItem ("Orange")
Combo1(4).AddItem ("Yellow")
Combo1(4).AddItem ("Green")
Combo1(4).AddItem ("Blue")
Combo1(4).AddItem ("Violet")
Combo1(4).AddItem ("Gray")
Combo1(4).AddItem ("White")
Combo1(4).AddItem ("Gold")
Combo1(4).AddItem ("Silver")
Combo1(4).AddItem ("No Color")
End If
End Sub

Private Sub Option1_Click()
If Option1.Value = True Then
Line19.Visible = False
Line17.Visible = False
Line20(1).Visible = False
Label5.BackColor = &H8080&
Combo1(3).Visible = False
Label1(2).Caption = "Multiplier"
Label1(4).Caption = ""
Image2.Visible = False
Image5.Visible = True
Image3.Visible = True
Image7.Visible = False

Combo1(0).AddItem ("Black")
Combo1(0).AddItem ("Brown")
Combo1(0).AddItem ("Red")
Combo1(0).AddItem ("Orange")
Combo1(0).AddItem ("Yellow")
Combo1(0).AddItem ("Green")
Combo1(0).AddItem ("Blue")
Combo1(0).AddItem ("Violet")
Combo1(0).AddItem ("Gray")
Combo1(0).AddItem ("White")


Combo1(1).AddItem ("Black")
Combo1(1).AddItem ("Brown")
Combo1(1).AddItem ("Red")
Combo1(1).AddItem ("Orange")
Combo1(1).AddItem ("Yellow")
Combo1(1).AddItem ("Green")
Combo1(1).AddItem ("Blue")
Combo1(1).AddItem ("Violet")
Combo1(1).AddItem ("Gray")
Combo1(1).AddItem ("White")

Combo1(2).AddItem ("Black")
Combo1(2).AddItem ("Brown")
Combo1(2).AddItem ("Red")
Combo1(2).AddItem ("Orange")
Combo1(2).AddItem ("Yellow")
Combo1(2).AddItem ("Green")
Combo1(2).AddItem ("Blue")
Combo1(2).AddItem ("Violet")
Combo1(2).AddItem ("Gray")
Combo1(2).AddItem ("White")
Combo1(2).AddItem ("Gold")
Combo1(2).AddItem ("Silver")

Combo1(4).AddItem ("Black")
Combo1(4).AddItem ("Brown")
Combo1(4).AddItem ("Red")
Combo1(4).AddItem ("Orange")
Combo1(4).AddItem ("Yellow")
Combo1(4).AddItem ("Green")
Combo1(4).AddItem ("Blue")
Combo1(4).AddItem ("Violet")
Combo1(4).AddItem ("Gray")
Combo1(4).AddItem ("White")
Combo1(4).AddItem ("Gold")
Combo1(4).AddItem ("Silver")
Combo1(4).AddItem ("No Color")
End If
End Sub



