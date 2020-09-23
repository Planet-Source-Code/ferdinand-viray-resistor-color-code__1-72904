VERSION 5.00
Begin VB.Form frmResistor 
   BackColor       =   &H80000012&
   Caption         =   "Resisitor Color Code Converter"
   ClientHeight    =   9030
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   9195
   ControlBox      =   0   'False
   FillColor       =   &H008080FF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   9030
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
      TabIndex        =   31
      Top             =   360
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
         TabIndex        =   33
         ToolTipText     =   "5-Band Resistor"
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
         TabIndex        =   32
         ToolTipText     =   "4-Band Resistor"
         Top             =   600
         Width           =   2175
      End
      Begin VB.Image Image7 
         Height          =   525
         Left            =   3000
         Picture         =   "frmResistor.frx":0000
         Top             =   1080
         Visible         =   0   'False
         Width           =   3900
      End
      Begin VB.Image Image5 
         Height          =   480
         Left            =   120
         Picture         =   "frmResistor.frx":6AE6
         Top             =   1080
         Visible         =   0   'False
         Width           =   3960
      End
      Begin VB.Image Image3 
         Height          =   375
         Left            =   3720
         Picture         =   "frmResistor.frx":CE28
         Top             =   1200
         Width           =   2985
      End
      Begin VB.Image Image2 
         Height          =   375
         Left            =   360
         Picture         =   "frmResistor.frx":10902
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
      Left            =   9960
      Top             =   3120
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
      ToolTipText     =   "Tolerance"
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
      ToolTipText     =   "2nd Band"
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
      ToolTipText     =   "1st Band"
      Top             =   2640
      Width           =   1815
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
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
      Height          =   450
      Left            =   15000
      TabIndex        =   42
      Top             =   10680
      Width           =   1125
   End
   Begin VB.Label lblc5 
      Height          =   375
      Left            =   15480
      TabIndex        =   41
      Top             =   10680
      Width           =   615
   End
   Begin VB.Label lblc4 
      Height          =   495
      Left            =   15360
      TabIndex        =   40
      Top             =   10680
      Width           =   495
   End
   Begin VB.Label lblc3 
      Height          =   375
      Left            =   15480
      TabIndex        =   39
      Top             =   10680
      Width           =   615
   End
   Begin VB.Label lblc2 
      Height          =   375
      Left            =   15360
      TabIndex        =   38
      Top             =   10800
      Width           =   615
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FF8080&
      Height          =   1335
      Left            =   240
      Top             =   7320
      Width           =   8535
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
      Left            =   7080
      TabIndex        =   37
      ToolTipText     =   "End Program"
      Top             =   6360
      Width           =   585
   End
   Begin VB.Image Image12 
      Height          =   765
      Left            =   6360
      Picture         =   "frmResistor.frx":14440
      Top             =   6240
      Width           =   2175
   End
   Begin VB.Image Image11 
      Height          =   765
      Left            =   5880
      Picture         =   "frmResistor.frx":14D72
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
      TabIndex        =   36
      ToolTipText     =   "Clear all items"
      Top             =   6360
      Width           =   810
   End
   Begin VB.Image Image10 
      Height          =   765
      Left            =   3480
      Picture         =   "frmResistor.frx":158B0
      Top             =   6240
      Width           =   2175
   End
   Begin VB.Image Image9 
      Height          =   765
      Left            =   3000
      Picture         =   "frmResistor.frx":161E2
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
      Left            =   1080
      TabIndex        =   35
      ToolTipText     =   "Get thr Resistor Value"
      Top             =   6360
      Width           =   1410
   End
   Begin VB.Image Image8 
      Height          =   765
      Left            =   720
      Picture         =   "frmResistor.frx":16D20
      Top             =   6240
      Width           =   2175
   End
   Begin VB.Line Line28 
      BorderWidth     =   20
      Index           =   2
      X1              =   6360
      X2              =   6600
      Y1              =   5880
      Y2              =   6120
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
      Y1              =   5880
      Y2              =   6120
   End
   Begin VB.Line Line26 
      BorderWidth     =   18
      Index           =   2
      X1              =   3120
      X2              =   2880
      Y1              =   5880
      Y2              =   6120
   End
   Begin VB.Line Line26 
      BorderWidth     =   11
      Index           =   1
      X1              =   1800
      X2              =   1560
      Y1              =   4320
      Y2              =   4560
   End
   Begin VB.Label lblc1 
      Height          =   375
      Left            =   15480
      TabIndex        =   25
      Top             =   10680
      Width           =   615
   End
   Begin VB.Line Line29 
      BorderWidth     =   20
      X1              =   8040
      X2              =   7680
      Y1              =   5880
      Y2              =   6240
   End
   Begin VB.Line Line28 
      BorderWidth     =   17
      Index           =   0
      X1              =   7800
      X2              =   8160
      Y1              =   4320
      Y2              =   4680
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
   Begin VB.Label lblTol 
      BackColor       =   &H80000007&
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
      Left            =   2400
      TabIndex        =   13
      Top             =   8040
      Width           =   2055
   End
   Begin VB.Label lblVal 
      AutoSize        =   -1  'True
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
      Height          =   450
      Left            =   1800
      TabIndex        =   12
      Top             =   7440
      Width           =   165
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000012&
      Caption         =   "Tolerance:"
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
      Left            =   600
      TabIndex        =   11
      Top             =   8040
      Width           =   1470
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000012&
      Caption         =   "Value:"
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
      Left            =   600
      TabIndex        =   10
      Top             =   7440
      Width           =   870
   End
   Begin VB.Line Line22 
      BorderColor     =   &H8000000E&
      BorderStyle     =   3  'Dot
      BorderWidth     =   2
      Index           =   2
      X1              =   7920
      X2              =   9120
      Y1              =   5400
      Y2              =   5400
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
      X1              =   6600
      X2              =   7680
      Y1              =   6000
      Y2              =   6000
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
      X1              =   7920
      X2              =   9120
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
      X1              =   7920
      X2              =   7920
      Y1              =   4680
      Y2              =   5760
   End
   Begin VB.Line Line12 
      BorderColor     =   &H8000000E&
      BorderWidth     =   2
      X1              =   120
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
      X1              =   120
      X2              =   1560
      Y1              =   5400
      Y2              =   5400
   End
   Begin VB.Line Line9 
      BorderColor     =   &H8000000E&
      BorderWidth     =   2
      X1              =   7680
      X2              =   7920
      Y1              =   4440
      Y2              =   4680
   End
   Begin VB.Line Line8 
      BorderColor     =   &H8000000E&
      BorderWidth     =   2
      Index           =   1
      X1              =   6720
      X2              =   7680
      Y1              =   4440
      Y2              =   4440
   End
   Begin VB.Line Line6 
      BorderColor     =   &H8000000E&
      BorderWidth     =   2
      Index           =   1
      X1              =   7920
      X2              =   7680
      Y1              =   5760
      Y2              =   6000
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
      Y1              =   5760
      Y2              =   6000
   End
   Begin VB.Line Line8 
      BorderColor     =   &H8000000E&
      BorderWidth     =   2
      Index           =   0
      X1              =   1560
      X2              =   1560
      Y1              =   4680
      Y2              =   5760
   End
   Begin VB.Line Line7 
      BorderColor     =   &H8000000E&
      BorderWidth     =   2
      Index           =   0
      X1              =   1560
      X2              =   1800
      Y1              =   5760
      Y2              =   6000
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
      Y1              =   6000
      Y2              =   6000
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
      Y1              =   5760
      Y2              =   6000
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
      Y1              =   5760
      Y2              =   5760
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
      BackColor       =   &H008080FF&
      Height          =   1095
      Index           =   1
      Left            =   3600
      TabIndex        =   19
      Top             =   4680
      Width           =   255
   End
   Begin VB.Label Label10 
      BackColor       =   &H008080FF&
      Height          =   1095
      Index           =   2
      Left            =   4320
      TabIndex        =   20
      Top             =   4680
      Width           =   255
   End
   Begin VB.Label Label10 
      BackColor       =   &H008080FF&
      Height          =   1095
      Index           =   3
      Left            =   5040
      TabIndex        =   21
      Top             =   4680
      Width           =   735
   End
   Begin VB.Label Label9 
      BackColor       =   &H00404040&
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   18
      Top             =   5040
      Width           =   1455
   End
   Begin VB.Label Label12 
      BackColor       =   &H00C0C0C0&
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   24
      Top             =   5160
      Width           =   1455
   End
   Begin VB.Label Label10 
      BackColor       =   &H008080FF&
      Height          =   1575
      Index           =   6
      Left            =   1560
      TabIndex        =   26
      Top             =   4440
      Width           =   735
   End
   Begin VB.Label Label9 
      BackColor       =   &H00404040&
      Height          =   255
      Index           =   1
      Left            =   7920
      TabIndex        =   28
      Top             =   5040
      Width           =   1215
   End
   Begin VB.Label Label12 
      BackColor       =   &H00C0C0C0&
      Height          =   255
      Index           =   1
      Left            =   7920
      TabIndex        =   29
      Top             =   5160
      Width           =   1215
   End
   Begin VB.Label Label10 
      BackColor       =   &H008080FF&
      Height          =   1575
      Index           =   4
      Left            =   6480
      TabIndex        =   22
      Top             =   4440
      Width           =   1455
   End
   Begin VB.Label Label2 
      ForeColor       =   &H000000FF&
      Height          =   1575
      Left            =   2280
      TabIndex        =   0
      Top             =   4440
      Width           =   495
   End
   Begin VB.Label Label10 
      BackColor       =   &H008080FF&
      Height          =   1575
      Index           =   7
      Left            =   2640
      TabIndex        =   27
      Top             =   4440
      Width           =   375
   End
   Begin VB.Label Label10 
      BackColor       =   &H008080FF&
      Height          =   1095
      Index           =   0
      Left            =   3000
      TabIndex        =   30
      Top             =   4680
      Width           =   135
   End
   Begin VB.Label Label3 
      Height          =   1095
      Left            =   3120
      TabIndex        =   23
      Top             =   4680
      Width           =   615
   End
   Begin VB.Label Label4 
      Height          =   1095
      Left            =   3840
      TabIndex        =   1
      Top             =   4680
      Width           =   495
   End
   Begin VB.Label Label5 
      Height          =   1095
      Left            =   4560
      TabIndex        =   2
      Top             =   4680
      Width           =   495
   End
   Begin VB.Label Label6 
      Height          =   1095
      Left            =   5760
      TabIndex        =   3
      Top             =   4680
      Width           =   495
   End
   Begin VB.Label Label10 
      BackColor       =   &H008080FF&
      Height          =   1095
      Index           =   5
      Left            =   6240
      TabIndex        =   34
      Top             =   4680
      Width           =   375
   End
   Begin VB.Image Image4 
      Height          =   765
      Left            =   240
      Picture         =   "frmResistor.frx":17652
      Top             =   6240
      Visible         =   0   'False
      Width           =   3105
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H0000FFFF&
      Height          =   8655
      Left            =   120
      Top             =   240
      Width           =   9015
   End
   Begin VB.Label lblc6 
      Height          =   255
      Left            =   15120
      TabIndex        =   43
      Top             =   10800
      Width           =   1215
   End
End
Attribute VB_Name = "frmResistor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim viray
Private Sub Combo1_Click(Index As Integer)
If Option1.Value = True Then
    Rest
    Rest3
    Short
    ferds
Else
    Rest
    Rest2
    Short
    ferds
End If
End Sub
Private Sub Form_Load()
ferds
Combo1(0).Text = ""
    Combo1(1).Text = ""
    Combo1(2).Text = ""
    Combo1(3).Text = ""
    Combo1(4).Text = ""
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
Private Sub Image4_Click()
    Label7_Click
    Image6.Visible = True
End Sub
Private Sub Image8_Click()
    Label7_Click
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
On Error Resume Next
If Option2.Value = True Then
    If Combo1(0).Text = "" Or Combo1(1).Text = "" Or Combo1(2).Text = "" Or Combo1(3).Text = "" Or Combo1(4).Text = "" Then
    MsgBox "One of the bands were uncolored!!..." & vbCrLf & "" & vbCrLf & "Please fill up the Bands properly.", vbCritical, "Insufficient Data!"
    Else
        If lblc6.Caption <> "" Then
        lblVal.Caption = lblc4.Caption & "  ohms" & "  Or  " & lblc6.Caption & " ohms"
        Label13.Caption = lblc6.Caption
        ferds
        Else
        lblVal.Caption = lblc4.Caption & "  ohms"
        Label13.Caption = lblc6.Caption
        ferds
        End If
    End If
ElseIf Option1.Value = True Then
    If Combo1(0).Text = "" Or Combo1(1).Text = "" Or Combo1(2).Text = "" Or Combo1(4).Text = "" Then
    MsgBox "One of the bands were uncolored!!..." & vbCrLf & "" & vbCrLf & "Please fill up the Bands properly.", vbCritical, "Insufficient Data!"
    Else
        If Label13.Caption <> "" Then
        lblVal.Caption = lblc4.Caption & "  ohms" & "  Or  " & Label13.Caption & " ohms"
        Label13.Caption = lblc6.Caption
        ferds
        Else
        lblVal.Caption = lblc4.Caption & "  ohms"
        Label13.Caption = lblc6.Caption
        ferds
    End If
    End If
End If

Short
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
    lblc1.Caption = ""
    lblc2.Caption = ""
    lblc3.Caption = ""
    lblc4.Caption = ""
    lblc5.Caption = ""
    lblc6.Caption = ""
    Label13.Caption = ""
    Label2.BackColor = vbWhite
    Label3.BackColor = vbWhite
    Label4.BackColor = vbWhite
    Label5.BackColor = vbWhite
    Label6.BackColor = vbWhite
    
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
Call Label8_Click
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

    List
    With Combo1(2)
        .AddItem ("Black")
        .AddItem ("Brown")
        .AddItem ("Red")
        .AddItem ("Orange")
        .AddItem ("Yellow")
        .AddItem ("Green")
        .AddItem ("Blue")
        .AddItem ("Violet")
        .AddItem ("Gray")
        .AddItem ("White")
    End With
    With Combo1(3)
        .AddItem ("Black")
        .AddItem ("Brown")
        .AddItem ("Red")
        .AddItem ("Orange")
        .AddItem ("Yellow")
        .AddItem ("Green")
        .AddItem ("Blue")
        .AddItem ("Violet")
        .AddItem ("Gray")
        .AddItem ("White")
        .AddItem ("Gold")
        .AddItem ("Silver")
    End With
End If
End Sub
Private Sub Option1_Click()
Call Label8_Click
If Option1.Value = True Then
        Line19.Visible = False
        Line17.Visible = False
        Line20(1).Visible = False
        Label5.BackColor = &H8080FF
        Combo1(3).Visible = False
        Label1(2).Caption = "Multiplier"
        Label1(4).Caption = ""
        Image2.Visible = False
        Image5.Visible = True
        Image3.Visible = True
        Image7.Visible = False
     List
    With Combo1(2)
        .AddItem ("Black")
        .AddItem ("Brown")
        .AddItem ("Red")
        .AddItem ("Orange")
        .AddItem ("Yellow")
        .AddItem ("Green")
        .AddItem ("Blue")
        .AddItem ("Violet")
        .AddItem ("Gray")
        .AddItem ("White")
        .AddItem ("Gold")
        .AddItem ("Silver")
    End With
End If
Short
ferds
End Sub
Function Rest()
Select Case Combo1(0).Text
    Case "Black"
        Label2.BackColor = vbBlack
        lblc1.Caption = "0"
    Case "Brown"
        Label2.BackColor = &H4080&
        lblc1.Caption = "1"
    Case "Red"
        Label2.BackColor = &HFF&
        lblc1.Caption = "2"
    Case "Orange"
        Label2.BackColor = &H80FF&
        lblc1.Caption = "3"
    Case "Yellow"
        Label2.BackColor = &HFFFF&
        lblc1.Caption = "4"
    Case "Green"
        Label2.BackColor = &HFF00&
        lblc1.Caption = "5"
    Case "Blue"
        Label2.BackColor = &HC00000
        lblc1.Caption = "6"
    Case "Violet"
        Label2.BackColor = &HC000C0
        lblc1.Caption = "7"
    Case "Gray"
        Label2.BackColor = &H808080
        lblc1.Caption = "8"
    Case "White"
        Label2.BackColor = vbWhite
        lblc1.Caption = "9"
    End Select

Select Case Combo1(1).Text
    Case "Black"
        Label3.BackColor = vbBlack
        lblc2.Caption = "0"
    Case "Brown"
        Label3.BackColor = &H4080&
        lblc2.Caption = "1"
    Case "Red"
        Label3.BackColor = &HFF&
        lblc2.Caption = "2"
    Case "Orange"
        Label3.BackColor = &H80FF&
        lblc2.Caption = "3"
    Case "Yellow"
        Label3.BackColor = &HFFFF&
        lblc2.Caption = "4"
    Case "Green"
        Label3.BackColor = &HFF00&
        lblc2.Caption = "5"
    Case "Blue"
        Label3.BackColor = &HC00000
        lblc2.Caption = "6"
    Case "Violet"
        Label3.BackColor = &HC000C0
        lblc2.Caption = "7"
    Case "Gray"
        Label3.BackColor = &H808080
        lblc2.Caption = "8"
    Case "White"
        Label3.BackColor = vbWhite
        lblc2.Caption = "9"
    End Select

Select Case Combo1(4).Text
    Case "Black"
        Label6.BackColor = vbBlack
        lblTol.Caption = "No Value"
    Case "Brown"
        Label6.BackColor = &H4080&
        lblTol.Caption = "± 1 %"
    Case "Red"
        Label6.BackColor = &HFF&
        lblTol.Caption = "± 2 %"
    Case "Orange"
        Label6.BackColor = &H80FF&
        lblTol.Caption = "No Value"
    Case "Yellow"
        Label6.BackColor = &HFFFF&
        lblTol.Caption = "No Value"
    Case "Green"
        Label6.BackColor = &HFF00&
        lblTol.Caption = "± 0.5 %"
    Case "Blue"
        Label6.BackColor = &HC00000
        lblTol.Caption = "± 0.25 %"
    Case "Violet"
        Label6.BackColor = &HC000C0
        lblTol.Caption = "± 0.10 %"
    Case "Gray"
        Label6.BackColor = &H808080
        lblTol.Caption = "± 0.05 %"
    Case "White"
        Label6.BackColor = vbWhite
        lblTol.Caption = "No Value"
    Case "Gold"
        Label6.BackColor = &H80C0FF
        lblTol.Caption = "± 5 %"
    Case "Silver"
        Label6.BackColor = &HC0C0C0
        lblTol.Caption = "± 10 %"
    Case "No Color"
        Label6.BackColor = &H8080FF
        lblTol.Caption = "± 20 %"
    End Select
ferds
Short
End Function

Function Rest2()
Short
ferds
viray = Val(lblc3.Caption)
Select Case Combo1(2).Text
    Case "Black"
        Label4.BackColor = vbBlack
        lblc5.Caption = "0"
    Case "Brown"
        Label4.BackColor = &H4080&
        lblc5.Caption = "1"
    Case "Red"
        Label4.BackColor = &HFF&
        lblc5.Caption = "2"
    Case "Orange"
        Label4.BackColor = &H80FF&
        lblc5.Caption = "3"
    Case "Yellow"
        Label4.BackColor = &HFFFF&
        lblc5.Caption = "4"
    Case "Green"
        Label4.BackColor = &HFF00&
        lblc5.Caption = "5"
    Case "Blue"
        Label4.BackColor = &HC00000
        lblc5.Caption = "6"
    Case "Violet"
        Label4.BackColor = &HC000C0
        lblc5.Caption = "7"
    Case "Gray"
        Label4.BackColor = &H808080
        lblc5.Caption = "8"
    Case "White"
        Label4.BackColor = vbWhite
        lblc5.Caption = "9"
    End Select


Select Case Combo1(3).Text
    Case "Black"
        Label5.BackColor = vbBlack
        lblc4.Caption = viray * 1
    Case "Brown"
        Label5.BackColor = &H4080&
        lblc4.Caption = viray * 10
    Case "Red"
        Label5.BackColor = &HFF&
        lblc4.Caption = viray * 100
    Case "Orange"
        Label5.BackColor = &H80FF&
        lblc4.Caption = viray * 1000
    Case "Yellow"
        Label5.BackColor = &HFFFF&
        lblc4.Caption = viray * 10000
    Case "Green"
        Label5.BackColor = &HFF00&
        lblc4.Caption = viray * 100000
    Case "Blue"
        Label5.BackColor = &HC00000
        lblc4.Caption = viray * 1000000
    Case "Violet"
        Label5.BackColor = &HC000C0
        lblc4.Caption = viray * 10000000
    Case "Gray"
        Label5.BackColor = &H808080
        lblc4.Caption = viray * 100000000
    Case "White"
        Label5.BackColor = vbWhite
        lblc4.Caption = viray * 1000000000
    Case "Gold"
        Label5.BackColor = &H80C0FF
        lblc4.Caption = viray * 0.1
    Case "Silver"
        Label5.BackColor = &HC0C0C0
        lblc4.Caption = viray * 0.01
    End Select

End Function
Function Rest3()
Dim viray
viray = Val(lblc3.Caption)
Select Case Combo1(2).Text
Case "Black"
    Label4.BackColor = vbBlack
    lblc4.Caption = viray * 1
Case "Brown"
    Label4.BackColor = &H4080&
    lblc4.Caption = viray * 10
Case "Red"
    Label4.BackColor = &HFF&
    lblc4.Caption = viray * 100
Case "Orange"
    Label4.BackColor = &H80FF&
    lblc4.Caption = viray * 1000
Case "Yellow"
    Label4.BackColor = &HFFFF&
    lblc4.Caption = viray * 10000
Case "Green"
    Label4.BackColor = &HFF00&
    lblc4.Caption = viray * 100000
Case "Blue"
    Label4.BackColor = &HC00000
    lblc4.Caption = viray * 1000000
Case "Violet"
    Label4.BackColor = &HC000C0
    lblc4.Caption = viray * 10000000
Case "Gray"
    Label4.BackColor = &H808080
    lblc4.Caption = viray * 100000000
Case "White"
    Label4.BackColor = vbWhite
    lblc4.Caption = viray * 1000000000
Case "Gold"
    Label4.BackColor = &H80C0FF
    lblc4.Caption = viray * 0.1
Case "Silver"
    Label4.BackColor = &HC0C0C0
    lblc4.Caption = viray * 0.01
End Select
viray = Val(lblc3.Caption)
ferds
Short
End Function
Function List()
With Combo1(0)
    .AddItem ("Black")
    .AddItem ("Brown")
    .AddItem ("Red")
    .AddItem ("Orange")
    .AddItem ("Yellow")
    .AddItem ("Green")
    .AddItem ("Blue")
    .AddItem ("Violet")
    .AddItem ("Gray")
    .AddItem ("White")
End With

With Combo1(1)
    .AddItem ("Black")
    .AddItem ("Brown")
    .AddItem ("Red")
    .AddItem ("Orange")
    .AddItem ("Yellow")
    .AddItem ("Green")
    .AddItem ("Blue")
    .AddItem ("Violet")
    .AddItem ("Gray")
    .AddItem ("White")
End With

With Combo1(4)
    .AddItem ("Black")
    .AddItem ("Brown")
    .AddItem ("Red")
    .AddItem ("Orange")
    .AddItem ("Yellow")
    .AddItem ("Green")
    .AddItem ("Blue")
    .AddItem ("Violet")
    .AddItem ("Gray")
    .AddItem ("White")
    .AddItem ("Gold")
    .AddItem ("Silver")
    .AddItem ("No Color")
End With
Short
End Function
Function ferds()
If Option2.Value = True Then
    lblc3.Caption = Val(lblc1.Caption) & Val(lblc2.Caption) & Val(lblc5.Caption)
End If
If Option1.Value = True Then
    lblc3.Caption = Val(lblc1.Caption) & Val(lblc2.Caption)
End If
End Function
Function Short()
viray = Val(lblc3.Caption)
If Option2.Value = True Then
Select Case Combo1(3).Text
    Case "Black"
        Label5.BackColor = vbBlack
        lblc6.Caption = ""
    Case "Brown"
        Label5.BackColor = &H4080&
        lblc6.Caption = ""
    Case "Red"
        Label5.BackColor = &HFF&
        lblc6.Caption = ""
    Case "Orange"
        Label5.BackColor = &H80FF&
        lblc6.Caption = viray * 1 & " Kilo"
    Case "Yellow"
        Label5.BackColor = &HFFFF&
        lblc6.Caption = viray * 10 & " Kilo"
    Case "Green"
        Label5.BackColor = &HFF00&
        lblc6.Caption = viray * 100 & " Kilo"
    Case "Blue"
        Label5.BackColor = &HC00000
        lblc6.Caption = viray * 1 & " Mega"
    Case "Violet"
        Label5.BackColor = &HC000C0
        lblc6.Caption = viray * 10 & " Mega"
    Case "Gray"
        Label5.BackColor = &H808080
        lblc6.Caption = viray * 100 & " Mega"
    Case "White"
        Label5.BackColor = vbWhite
        lblc6.Caption = viray * 1 & " Giga"
    Case "Gold"
        Label5.BackColor = &H80C0FF
        lblc6.Caption = ""
    Case "Silver"
        Label5.BackColor = &HC0C0C0
        lblc6.Caption = ""
    End Select
ElseIf Option1.Value = True Then
Select Case Combo1(2).Text
    Case "Black"
        Label5.BackColor = vbBlack
        lblc6.Caption = ""
    Case "Brown"
        Label5.BackColor = &H4080&
        lblc6.Caption = ""
    Case "Red"
        Label5.BackColor = &HFF&
        lblc6.Caption = ""
    Case "Orange"
        Label5.BackColor = &H80FF&
        lblc6.Caption = viray * 1 & " Kilo"
    Case "Yellow"
        Label5.BackColor = &HFFFF&
        lblc6.Caption = viray * 10 & " Kilo"
    Case "Green"
        Label5.BackColor = &HFF00&
        lblc6.Caption = viray * 100 & " Kilo"
    Case "Blue"
        Label5.BackColor = &HC00000
        lblc6.Caption = viray * 1 & " Mega"
    Case "Violet"
        Label5.BackColor = &HC000C0
        lblc6.Caption = viray * 10 & " Mega"
    Case "Gray"
        Label5.BackColor = &H808080
        lblc6.Caption = viray * 100 & " Mega"
    Case "White"
        Label5.BackColor = vbWhite
        lblc6.Caption = viray * 1 & " Giga"
    Case "Gold"
        Label5.BackColor = &H80C0FF
        lblc6.Caption = ""
    Case "Silver"
        Label5.BackColor = &HC0C0C0
        lblc6.Caption = ""
    End Select
Label13.Caption = lblc6.Caption
End If
End Function
