VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{2724431D-30C7-11D4-B93E-0050DA73070D}#1.0#0"; "Label3Dcontrol.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{DBD1A1A1-6274-43F5-9588-B58E194B0FE9}#3.0#0"; "lavolpeButton.ocx"
Begin VB.Form frmLogin 
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Login"
   ClientHeight    =   3525
   ClientLeft      =   45
   ClientTop       =   225
   ClientWidth     =   11640
   ControlBox      =   0   'False
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3525
   ScaleWidth      =   11640
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   0
      ScaleHeight     =   1095
      ScaleWidth      =   11640
      TabIndex        =   16
      Top             =   0
      Width           =   11640
      Begin Label3DControl.Label3D Label3D1 
         Height          =   480
         Left            =   1800
         TabIndex        =   17
         Top             =   0
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   847
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Segoe UI"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Segoe UI"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackStyle       =   1
         Caption         =   "dimo"
         Caption         =   "dimo"
         Color1          =   16744576
         ShadowLeft      =   35
         AutoSize        =   0
      End
      Begin Label3DControl.Label3D Label3D2 
         Height          =   480
         Left            =   8760
         TabIndex        =   37
         Top             =   600
         Width           =   2610
         _ExtentX        =   4604
         _ExtentY        =   847
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Segoe UI"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Segoe UI"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackStyle       =   1
         Caption         =   "Softnet Solutions"
         Caption         =   "Softnet Solutions"
         Color1          =   16744576
         ShadowLeft      =   35
         AutoSize        =   0
      End
      Begin VB.Image Image3 
         Height          =   735
         Left            =   150
         Picture         =   "frmLogin.frx":08CA
         Top             =   165
         Width           =   870
      End
      Begin VB.Shape Shape3 
         BackColor       =   &H00FFFFFF&
         BorderColor     =   &H00FF8080&
         BorderWidth     =   2
         Height          =   810
         Index           =   0
         Left            =   120
         Shape           =   4  'Rounded Rectangle
         Top             =   120
         Width           =   945
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Ver 0.01  Project Started on 20 October 2016"
         BeginProperty Font 
            Name            =   "Segoe UI Semibold"
            Size            =   9
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   6840
         TabIndex        =   38
         Top             =   120
         Visible         =   0   'False
         Width           =   3495
      End
      Begin VB.Label lblServer 
         BackStyle       =   0  'Transparent
         Caption         =   "Server"
         BeginProperty Font 
            Name            =   "Segoe UI Semibold"
            Size            =   9
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   6840
         TabIndex        =   24
         Top             =   405
         Visible         =   0   'False
         Width           =   3735
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "ystem"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   315
         Index           =   6
         Left            =   4275
         TabIndex        =   23
         Top             =   525
         Width           =   900
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "S"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   4080
         TabIndex        =   22
         Top             =   480
         Width           =   255
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "I"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   1800
         TabIndex        =   21
         Top             =   480
         Width           =   375
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "ncentive"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   315
         Index           =   3
         Left            =   1920
         TabIndex        =   20
         Top             =   525
         Width           =   735
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "C"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   2880
         TabIndex        =   19
         Top             =   480
         Width           =   495
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "ontroll"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   315
         Index           =   4
         Left            =   3075
         TabIndex        =   18
         Top             =   525
         Width           =   855
      End
   End
   Begin VB.PictureBox Picture2 
      Align           =   2  'Align Bottom
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   0
      ScaleHeight     =   375
      ScaleWidth      =   11640
      TabIndex        =   13
      Top             =   3150
      Width           =   11640
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Copyright © 2012. All Right Reserved. Softnet Solutions™   Contact: 0755 527475 Email : Chamindadp@gmail.com"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   15
         Top             =   120
         Width           =   9015
      End
      Begin VB.Image Image4 
         Height          =   375
         Left            =   0
         Picture         =   "frmLogin.frx":2ABC
         Stretch         =   -1  'True
         Top             =   0
         Width           =   12495
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Copyright © 2012. All Right Reserved. Softnet Solution Contact: 0755 527475, 0770 685595"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   0
         TabIndex        =   14
         Top             =   120
         Width           =   8655
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   5160
      Top             =   2760
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   4000
      Left            =   6840
      TabIndex        =   25
      Top             =   960
      Width           =   5775
      Begin VB.CheckBox ChkEOD 
         BackColor       =   &H00FFFFFF&
         Caption         =   "EOD"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1200
         TabIndex        =   10
         Top             =   3600
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.OptionButton optSignIn 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Sign in"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   50
         TabIndex        =   3
         Top             =   2160
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox txtFloat 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   2400
         MaxLength       =   20
         TabIndex        =   8
         Top             =   2160
         Visible         =   0   'False
         Width           =   1395
      End
      Begin VB.OptionButton optLogin 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Log In"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   50
         TabIndex        =   4
         Top             =   2520
         Value           =   -1  'True
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.OptionButton optLogOff 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Log Off"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   50
         TabIndex        =   5
         Top             =   2880
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.OptionButton optSignOff 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Sign Off"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   50
         TabIndex        =   6
         Top             =   3240
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox txtUserName 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   330
         Left            =   1000
         MaxLength       =   20
         TabIndex        =   0
         Top             =   720
         Width           =   3600
      End
      Begin VB.TextBox txtPassword 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   330
         IMEMode         =   3  'DISABLE
         Left            =   1000
         MaxLength       =   20
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   1065
         Width           =   3600
      End
      Begin lavolpeButton.lvButtons_H cmdCancel 
         Height          =   495
         Left            =   3300
         TabIndex        =   12
         Tag             =   "04"
         ToolTipText     =   "Double Click to open"
         Top             =   1530
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   873
         Caption         =   "Cancel"
         CapAlign        =   2
         BackStyle       =   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFore           =   0
         cFHover         =   0
         cBhover         =   12648384
         LockHover       =   3
         cGradient       =   16744576
         Gradient        =   3
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin lavolpeButton.lvButtons_H cmdOK 
         Height          =   495
         Left            =   1920
         TabIndex        =   11
         Tag             =   "04"
         ToolTipText     =   "Double Click to open"
         Top             =   1530
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   873
         Caption         =   "OK"
         CapAlign        =   2
         BackStyle       =   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFore           =   0
         cFHover         =   0
         cBhover         =   12648384
         LockHover       =   3
         cGradient       =   16744576
         Gradient        =   3
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin MSDataListLib.DataCombo DacCompany 
         Height          =   315
         Left            =   1000
         TabIndex        =   2
         Top             =   1425
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         ForeColor       =   -2147483641
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DacLocations 
         Height          =   315
         Left            =   1000
         TabIndex        =   27
         Top             =   1755
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Appearance      =   0
         ForeColor       =   -2147483641
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.TextBox Text1 
         Height          =   195
         Left            =   2160
         TabIndex        =   32
         Top             =   1680
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   2760
         TabIndex        =   33
         Top             =   1680
         Visible         =   0   'False
         Width           =   255
      End
      Begin MSComCtl2.DTPicker SignOnDate 
         Height          =   300
         Left            =   2400
         TabIndex        =   7
         ToolTipText     =   "Enter Insurance Date Range"
         Top             =   2505
         Visible         =   0   'False
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   529
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Segoe UI"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarBackColor=   16744576
         Format          =   149487617
         CurrentDate     =   41609
      End
      Begin MSComCtl2.DTPicker SignOffDate 
         Height          =   300
         Left            =   2400
         TabIndex        =   9
         ToolTipText     =   "Enter Insurance Date Range"
         Top             =   3195
         Visible         =   0   'False
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   529
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Segoe UI"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarBackColor=   16744576
         Format          =   149487617
         CurrentDate     =   41609
      End
      Begin VB.Label lblFloat 
         BackStyle       =   0  'Transparent
         Caption         =   "Float"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   255
         Left            =   1200
         TabIndex        =   36
         Top             =   2160
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label lblSignOn 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Sign on Date"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1200
         TabIndex        =   35
         Top             =   2520
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label lblSignOff 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Sign Off Date"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1200
         TabIndex        =   34
         Top             =   3240
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Company"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   255
         Index           =   5
         Left            =   50
         TabIndex        =   31
         Top             =   1425
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Password"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   255
         Index           =   1
         Left            =   50
         TabIndex        =   30
         Top             =   1065
         Width           =   1215
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Username"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   255
         Left            =   50
         TabIndex        =   29
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Location"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   255
         Index           =   0
         Left            =   50
         TabIndex        =   28
         Top             =   1755
         Width           =   1215
      End
      Begin VB.Label lblTitle 
         BackStyle       =   0  'Transparent
         Caption         =   "User Login"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   50
         TabIndex        =   26
         Top             =   240
         Width           =   3615
      End
      Begin VB.Image Image2 
         Height          =   495
         Left            =   0
         Picture         =   "frmLogin.frx":7FEE8
         Stretch         =   -1  'True
         Top             =   120
         Width           =   4600
      End
   End
   Begin VB.Label lblServer1 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Segoe UI Semilight"
         Size            =   12
         Charset         =   0
         Weight          =   350
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   40
      Top             =   1680
      Width           =   5535
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Ver 0.01  Last Updated 28 Junel 2017"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   39
      Top             =   1320
      Width           =   5535
   End
   Begin VB.Image Image5 
      Height          =   105
      Left            =   0
      Picture         =   "frmLogin.frx":FD314
      Stretch         =   -1  'True
      Top             =   960
      Width           =   11100
   End
   Begin VB.Image Image1 
      Height          =   4215
      Left            =   -6120
      Picture         =   "frmLogin.frx":FFFF8
      Top             =   1080
      Width           =   15000
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'************************************************
'Project Started 20/10/2016
'************************************************
Option Explicit
Public LoginSucceeded As Boolean
Dim retry As Integer
Dim Empl As ClsEmployee
Dim Sec As ClsSecurity
Dim Item As ClsItem
Dim rsCompany As ADODB.recordset
Dim rsUser As ADODB.recordset
Dim ComSetUp As New ADODB.Command
Dim Gen As ClsGeneral
Dim mSerialNo As String
Dim mResponse As String
Dim rsSetup As New ADODB.recordset
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

Private Sub cmdCancel_Click()
On Error Resume Next
Err.Clear
    LoginSucceeded = False
    Unload Me
'    Unload frmSplash
End Sub
Private Sub cmdCancel_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
Err.Clear
'    StatusBar1.Panels(1).Text = "Press Cancel to Quit"
End Sub
Private Sub cmdOK_Click()
If txtUserName = "" Then
    MsgBox "Check User Name", vbInformation, MsgTitle
    Exit Sub
End If
If txtPassword = "" Then
    MsgBox "Check Password", vbInformation, MsgTitle
    Exit Sub
End If
If DacCompany = "" Then
    MsgBox "Check Company", vbInformation, MsgTitle
    Exit Sub
End If

If DacLocations = "" Then
    MsgBox "Check Locations", vbInformation, MsgTitle
    Exit Sub
End If

On Error GoTo errHandler


Set Empl = New ClsEmployee
Set rsUser = Empl.GetUser_Access(txtUserName, txtPassword, DacCompany) 'IUsers file

If rsUser.RecordCount <> 0 Then 'User Ok
    LoginSucceeded = True
    UserId = Trim(rsUser!username)
    username = Trim(rsUser!Name)
    QuotRead = rsUser!QuotRead
    QuotRevision = rsUser!QuotRevision
    QuotEdit = rsUser!QuotEdit
    
    mCompanyCode = Trim(rsUser!companycode)
    mCompany = Trim(rsUser!Company)
    MuserType = rsUser!usertype


''==========================
'
''SignOn
'    If PSystemDate <> SignOnDate Then
'            mResponse = MsgBox("Last EOD Date is  " & rsUser!eoddate & " and Do You want to Starting New Date  ", vbYesNo, MsgTitle)
'            If mResponse = vbYes Then
'                optSignIn.Value = True
'                SignOnDate = rsSetup!SystemDate
'                PSystemDate = SignOnDate
'                SignOnDate.SetFocus
'            ElseIf mResponse = vbNo Then
'                txtUserName.SetFocus
'                Exit Sub
'            End If
'
'
'
'
'
'     ElseIf PSystemDate = SignOnDate Then
'        If optSignIn.Value = True And rsUser!issignon = "0" Then
'            UpdateSignOn
'
'
'
'        ElseIf optSignIn.Value = True And rsUser!issignon = "1" Then
'            MsgBox "You Already Sign On for the Date " & rsUser!SignOnDate & "", vbCritical, MsgTitle
'            optLogin.Value = True
'            Exit Sub
'
'
'        ElseIf optLogin.Value = True Then
'
'
'
'
'
'
'        ElseIf optSignOff.Value = True And rsUser!issignon = "1" Then
'            UpdateSignOFF
'
'        ElseIf optSignOff.Value = True And rsUser!issignoff = "1" Then
'                    MsgBox "You Already Sign Off for the Date " & rsUser!SignOnDate & "", vbCritical, MsgTitle
'                    optLogin.Value = True
'                    Exit Sub
'
'
'        ElseIf optLogOff.Value = True Then
'
'
'
'        End If
'                    '-------------------------------------------- n e w
'                    If optSignIn.Value = True Then
'                        Set Item = New ClsItem
'                        Item.UserLog "SignIn Success"
'                        Unload frmLogin
'                        frmMenu.Show
'                    ElseIf optLogin.Value = True Then
'                        Set Item = New ClsItem
'                        Item.UserLog "Login Success"
'                        Unload frmLogin
'                        frmMenu.Show
'
'                    ElseIf optLogOff.Value = True Then
'                        Set Item = New ClsItem
'                        Item.UserLog "Log Off Success"
'                        Unload frmLogin
'
'
'                    ElseIf optSignOff.Value = True Then
'
'                        Set Item = New ClsItem
'                        Item.UserLog "Sign Off Success"
'                        Unload frmLogin
'                    End If
'                    '--------------------------------------------
'
'
'
'
'
'
'    End If
'
''==========================

      Set Gen = New ClsGeneral
      Set rsSetup = Gen.CheckSetup
      If rsSetup.RecordCount <> 0 Then
            PcurrMonth = rsSetup!currmonth
            PCurrYear = rsSetup!curryear
      End If

'----------------------------------------------------
        Set Item = New ClsItem
        Item.UserLog "Success"


    Unload frmLogin
'    Unload frmSplash
    frmMenu.Show
'----------------------------------------------------
Else

    If retry < 3 Then
        MsgBox "Invalid Entries, try again!", vbCritical, "Login Failure -" + MsgTitle
        txtUserName.SetFocus
'        SendKeys "{Home}+{End}"
        retry = retry + 1
                Set Item = New ClsItem
                Item.UserLog "Invalid Entries " + str(retry)

        If retry = 3 Then
            MsgBox " Access violated, Contact Administrator ", vbCritical, "Login - " + MsgTitle
                Set Item = New ClsItem
                Item.UserLog "Access violated " + str(retry)
            
            Unload Me
'            Unload frmSplash
    '''        mDateTime = Date + Time()
    '''        Call AddLogFile("PassErr", "", "", "", 0, 0, 0, UserId, mDateTime, mCompanyCode)
        End If
    End If

End If
'''On Error GoTo errHandler
'''
'''DacCompany = UCase(DacCompany.Text)
'''Set rsUser = New ADODB.recordset
'''rsUser.Open "SELECT * FROM users WHERE username = '" & txtUserName & "' and password = '" & txtPassword & "' and companycode = '" & DacCompany & "'", cnStock, adOpenKeyset, adLockReadOnly, 8
'''rsUser.Requery
'''    If rsUser.RecordCount <> 0 Then
'''        LoginSucceeded = True
'''        mUserName = txtUserName
'''        UserId = txtUserName
'''        mCompanyCode = DacCompany.Text
'''        Me.Hide
'''        Set rsCompany = New recordset
'''        rsCompany.Open "SELECT * FROM company WHERE companycode='" & DacCompany & "'", cnStock, adOpenKeyset, adLockReadOnly, 8
'''        rsCompany.Requery
'''        mCompany = rsCompany!Company
'''
'''    Unload frmLogin
'''    Load frmMenu
'''    frmMenu.Show
'''    Unload frmSplash
'''
'''    Dim mDateTime As Date
'''    mDateTime = Date + Time()
'''    Call AddLogFile("Login", "", "", "", 0, 0, 0, UserId, mDateTime, mCompanyCode)
'''
'''
'''    Else
'''        If retry = 3 Then
'''            MsgBox " Access violated, Contact Administrator ", vbCritical, "Login"
'''            Unload Me
'''            Unload frmSplash
'''            mDateTime = Date + Time()
'''            Call AddLogFile("PassErr", "", "", "", 0, 0, 0, UserId, mDateTime, mCompanyCode)
'''        Else
'''            MsgBox "Invalid Entries, try again!"
'''            txtUserName.SetFocus
'''            SendKeys "{Home}+{End}"
'''            retry = retry + 1
'''        End If
'''    End If
'''    Exit Sub
       
       
       Exit Sub

errHandler:
    MsgBox Err.Description
    MsgBox "Server not Found", vbCritical, MsgTitle
    Resume
'    Unload frmSplash
    Unload frmLogin
    End
End Sub

Private Sub cmdOK_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
Err.Clear
'    StatusBar1.Panels(1).Text = "Press OK to Login"
End Sub

Private Sub DacCompany_Change()
        mCompanyCode = DacCompany
        Set Gen = New ClsGeneral
        Set rsSetup = Gen.CheckSetup           'Chk From QuotHeader
        If rsSetup.RecordCount <> 0 Then       'QuotNo Existing - QuotHeader
            DacLocations = rsSetup!LocCode
            PSignOnLocaton = rsSetup!LocCode
            '-------------------
            PSystemDate = rsSetup!SystemDate
            SignOnDate = rsSetup!SignOnDate
            PSignOnDate = rsSetup!SignOnDate
            SignOffDate = rsSetup!SignOffDate
            PSignOffDate = rsSetup!SignOffDate
            '-------------------
        End If
End Sub

Private Sub DacCompany_Click(Area As Integer)
On Error Resume Next
Err.Clear
    
'    Set Gen = New ClsGeneral
'    Set rsCompany = Gen.GetListCompanies
'    Set DacCompany.RowSource = rsCompany
'    DacCompany.ListField = "companycode"

    
   
End Sub
Private Sub DacCompany_GotFocus()
On Error Resume Next
Err.Clear
'    StatusBar1.Panels(1).Text = "Select Your Company Code"
End Sub
Private Sub DacCompany_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
Err.Clear
    If KeyCode = vbKeyReturn Then cmdOK.SetFocus

'    If KeyCode = vbKeyReturn Then cmdOK.SetFocus
End Sub

Private Sub DacCompany_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
Err.Clear
'    StatusBar1.Panels(1).Text = "Select Your Company Code"
End Sub

Private Sub DacLocations_Change()
    Set Gen = New ClsGeneral
    Set rsCompany = Gen.CheckLocation(DacLocations)
    If rsCompany.RecordCount <> 0 Then
        PLocationCode = rsCompany!Code
        PLocation = rsCompany!Name
    End If
End Sub

Private Sub DacLocations_Click(Area As Integer)
On Error Resume Next
Err.Clear
    
    Set Gen = New ClsGeneral
    Set rsCompany = Gen.GetListLocation
    Set DacLocations.RowSource = rsCompany
    DacLocations.ListField = "Code"
End Sub

Private Sub Form_Load()
'Set Skinner1.Forms = Forms
'On Error Resume Next
'Err.Clear

''If Server = "" Then
''        Text1.Text = "C:\Stock\Reports\initialCash.exe"
''        If FileExists%(Text1.Text) = True Then
''            Kill "C:\Stock\Reports\initialCash.exe"
''        End If

'    Call Valid
    Open "C:\DimoIncentive\DimoIncentive.txt" For Input As #1
'    Open "D:\DimoIncentive\DimoIncentive.txt" For Input As #1

    Input #1, Server, cDatabase, Reportpath, CompHeader, MsgTitle, CompID
    Close #1
    
    lblServer.caption = "Server " + Server
    lblServer1.caption = "Server " + Server
    With frmLogin
        .Left = ((Screen.Width - .Width) / 2)
        .Top = ((Screen.Height - .Height) / 2) + .Height
    End With
    Call CheckCompanySerial
    Call valid1
    Label3D1.caption = CompHeader

    txtUserName = "000"
    txtPassword = "123"
    DacCompany = "DMO"

'---------------------------------------
    mCompanyCode = ""
    mCompanyCodeDot = "DM."
'-------------------------------------------
        Set Gen = New ClsGeneral
        Set rsSetup = Gen.CheckSetup           'Chk From QuotHeader
        If rsSetup.RecordCount <> 0 Then       'QuotNo Existing - QuotHeader
            DacLocations = rsSetup!LocCode
            PSignOnLocaton = rsSetup!LocCode
            '-------------------
            PSystemDate = rsSetup!SystemDate
            SignOnDate = rsSetup!SignOnDate
            PSignOnDate = rsSetup!SignOnDate
            SignOffDate = rsSetup!SignOffDate
            PSignOffDate = rsSetup!SignOffDate
            '-------------------
        End If
'-------------------------------------------------------------
'SignOnDate = Date


'-------------------------------------------------------------
'    Call ValidNew
'    Call Valid
'    MsgBox "Current Server - " + Server + " / Database - " + cDatabase
'frmSplash.Show
'frmLogin.cmdOK = True
End Sub


Private Sub optLogin_Click()
    Call optSignIn_Click
    Call optSignOff_Click
End Sub

Private Sub optLogOff_Click()
    Call optSignIn_Click
    Call optSignOff_Click
End Sub

Private Sub optSignIn_Click()
If optSignIn.Value = True Then
    lblFloat.Visible = True
    txtFloat.Visible = True
    
    lblSignOn.Visible = True
    SignOnDate.Visible = True
    lblSignOff.Visible = False
    SignOffDate.Visible = False
    ChkEOD.Visible = False
Else
    lblFloat.Visible = False
    txtFloat.Visible = False
    lblSignOn.Visible = False
    SignOnDate.Visible = False

End If
    

End Sub

Private Sub optSignOff_Click()
If optSignOff.Value = True Then
    lblFloat.Visible = False
    txtFloat.Visible = False
    
    lblSignOff.Visible = True
    SignOffDate.Visible = True
    ChkEOD.Visible = True
Else
    lblFloat.Visible = False
    txtFloat.Visible = False
    lblSignOff.Visible = False
    SignOffDate.Visible = False
    ChkEOD.Visible = False
End If
End Sub

Private Sub SignOnDate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then txtFloat.SetFocus
End Sub

Private Sub txtFloat_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call cmdOK_Click
End Sub

Private Sub txtFloat_LostFocus()
    If Val(txtFloat) = 0 Then txtFloat.SetFocus
End Sub

'Private Sub Timer1_Timer()
'    If GetAsyncKeyState(vbKeyF1) Then
'        Dim Exp As SHDocVw.InternetExplorer
'        Set Exp = New SHDocVw.InternetExplorer
'        Exp.Visible = True
'        Exp.Navigate Helppath
'    End If
'End Sub

'Private Sub Timer1_Timer()
'    If GetAsyncKeyState(vbKeyF1) Then
'        Dim Exp As SHDocVw.InternetExplorer
'        Set Exp = New SHDocVw.InternetExplorer
'        Exp.Visible = True
'        Exp.Navigate HelpPath
'    End If
'End Sub

Private Sub txtPassword_GotFocus()
On Error Resume Next
Err.Clear
'    StatusBar1.Panels(1).Text = "Enter Your Password"
End Sub
Private Sub txtPassword_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
Err.Clear
    If KeyCode = vbKeyReturn Then DacCompany.SetFocus
End Sub
Private Sub txtPassword_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
Err.Clear
'    StatusBar1.Panels(1).Text = "Enter Your Password"
End Sub
Private Sub txtUserName_GotFocus()
'    StatusBar1.Panels(1).Text = "Enter Your User Name"
End Sub
Private Sub txtUserName_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
Err.Clear
    If KeyCode = vbKeyReturn Then txtPassword.SetFocus
End Sub

Private Sub txtUserName_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
Err.Clear
'    StatusBar1.Panels(1).Text = "Enter Your User Name"
End Sub
Private Sub txtUserName_Validate(Cancel As Boolean)
On Error Resume Next
Err.Clear
    txtUserName = UCase(txtUserName)
End Sub

Public Function Valid()
    Text1.Text = GetStringValue("HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion", "RegisteredOwner")
    Text2.Text = GetStringValue("HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion", "RegisteredOrganization")

   If Text1 <> "Chaminda Dhammajith" Or Text2 <> "Softnet" Then
        MsgBox "Illegal Copy, Contact System's Administrator", vbCritical, "Systems Error"
        End
   End If
    
    Text1.Text = GetStringValue("HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion", "SystemSerialNo")
    Text2.Text = GetStringValue("HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion", "SystemExpiredDate")

   If Text1 <> "19700530" Then
       MsgBox "Illegal Copy, Contact System's Administrator", vbCritical, "Systems Error"
        End
   End If
   If Date > CDate(Text2) Then
        MsgBox "Time Expired, Contact System's Administrator", vbCritical, "Systems Error"
        SetStringValue "HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion", "SystemSerialNo", ""
        End
   End If
End Function
Public Sub CheckCompanySerial()
On Error Resume Next
Err.Clear
mCompanyCode = "DMO"
mCompany = "Diesel and Motor Engineering PLC"
'mSerialNo = "2787-3376-2849-62215-2787"
mSerialNo = "2016-4425-3645-62215-2017"

Set Empl = New ClsEmployee
Set rsUser = Empl.GetSec_Auth(mCompanyCode, mCompany, mSerialNo)

If rsUser.RecordCount = 0 Then 'User Ok
    MsgBox "Authontication Prohibited", vbCritical, MsgTitle
    End
End If

End Sub
Public Function ValidNew()
'On Error GoTo errHandler

Set Sec = New ClsSecurity
Set rsUser = Sec.GetSec_Access()

If rsUser.RecordCount <> 0 Then
Text1 = rsUser!DatabaseVersion1
Text2 = rsUser!DatabaseVersion2
        If ((rsUser!DatabaseVersion2 = "AVS50-81SG00S-G61002U") Or (rsUser!DatabaseVersion2 = "AVS50-81SG00S-G61002U1970") Or (rsUser!DatabaseVersion2 = "AVS50-81SG00S-G61002U1971") Or (rsUser!DatabaseVersion2 = "AVS50-81SG00S-G61002U1972") Or (rsUser!DatabaseVersion2 = "AVS50-81SG00S-G61002U1973") Or (rsUser!DatabaseVersion2 = "AVS50-81SG00S-G61002U1974") Or (rsUser!DatabaseVersion2 = "AVS50-81SG00S-G61002U1975")) Then
                If Date > CDate(Text1) Then
                    MsgBox "Time Expired, Contact System's Administrator", vbCritical, "Systems Error"
                    End
                End If
        Else
            MsgBox "Illegal Copy, Contact System's Administrator", vbCritical, "Systems Error"
            End
        End If
Else
        MsgBox "Invalid Licence, Contact System's Administrator", vbCritical, "Systems Error"
        End
End If
End Function
Function valid1()
Text2 = "2018/ 01/ 30"
    If Date > CDate(Text2) Then
        '----------------------------------
        mQuotNo = ""
'        mQuotNoShort = ""
        Dim Gen As New ClsGeneral
        Set Gen = New ClsGeneral
        Gen.AddSetUp "SNo", mCompanyCode, mCompany, mQuotNo, ""
        '----------------------------------
        MsgBox "Licence Expired and Renew for This/Next Year", vbCritical, MsgTitle
        End
    End If
End Function

Function FileExists%(fname$)
 On Local Error Resume Next
 
 Dim ff%
 ff% = FreeFile
 Open fname$ For Input As ff%
 If Err Then
  FileExists% = False
 Else
  FileExists% = True
 End If
 
 Close ff%
End Function

Public Function UpdateSignOn()
  Dim Dll As ClsConnection
  Set Dll = New ClsConnection

  If Dll.IsConnectionDeclared <> 1 Then
     'not declared
     If Dll.DeclareConnection <> 1 Then
        MsgBox "Couldn't Connect", vbCritical
     End If
  End If
  Dll.Conn.Open

  'open recordset and issue command

        Set Gen = New ClsGeneral
        Set rsSetup = Gen.CheckSetup          'Chk From QuotHeader
        If rsSetup.RecordCount <> 0 Then       'QuotNo Existing - QuotHeader
                'Update isetup
                Set ComSetUp = New Command
                ComSetUp.ActiveConnection = Dll.Conn
                ComSetUp.CommandText = "Update ISetUp set SignOnDate='" & Format(PSystemDate, "mm/dd/yyyy") & "' , SignOffDate='" & Format(PSystemDate, "mm/dd/yyyy") & "' , EodDate='" & Format(PSystemDate, "mm/dd/yyyy") & "' , SystemDate='" & Format(PSystemDate, "mm/dd/yyyy") & "' WHERE LocCode='" & PLocationCode & "' and CompanyCode='" & mCompanyCode & "'"
                ComSetUp.Execute
         End If
                
                'Update Iusers

        Set Gen = New ClsGeneral
        Set rsSetup = Gen.CheckSignOn          'Chk From QuotHeader
        If rsSetup.RecordCount <> 0 Then       'QuotNo Existing - QuotHeader
                'Update isetup
                Set ComSetUp = New Command
                ComSetUp.ActiveConnection = Dll.Conn
                ComSetUp.CommandText = "Update IUsers set SignOnDate='" & Format(PSystemDate, "mm/dd/yyyy") & "' , SignOffDate='" & Format(PSystemDate, "mm/dd/yyyy") & "' , EodDate='" & Format(PSystemDate, "mm/dd/yyyy") & "' , Float='" & Val(frmLogin.txtFloat) & "' , CompID='" & CompID & "' , LocCode='" & PLocationCode & "' , IsSignOn='1', IsSignOff='0', EOD='0'  Where UserName='" & UserId & "' and CompanyCode='" & mCompanyCode & "'"
                ComSetUp.Execute
        End If

End Function


Public Function UpdateSignOFF()
  Dim Dll As ClsConnection
  Set Dll = New ClsConnection

  If Dll.IsConnectionDeclared <> 1 Then
     'not declared
     If Dll.DeclareConnection <> 1 Then
        MsgBox "Couldn't Connect", vbCritical
     End If
  End If
  Dll.Conn.Open

  'open recordset and issue command

        Set Gen = New ClsGeneral
        Set rsSetup = Gen.CheckSetup          'Chk From QuotHeader
        If rsSetup.RecordCount <> 0 Then       'QuotNo Existing - QuotHeader
                'Update isetup
                Set ComSetUp = New Command
                ComSetUp.ActiveConnection = Dll.Conn
                ComSetUp.CommandText = "Update ISetUp set SignOnDate='" & Format(PSystemDate, "mm/dd/yyyy") & "' , SignOffDate='" & Format(PSystemDate, "mm/dd/yyyy") & "' , EodDate='" & Format(PSystemDate, "mm/dd/yyyy") & "' , SystemDate='" & Format(PSystemDate + 1, "mm/dd/yyyy") & "' WHERE LocCode='" & PLocationCode & "' and CompanyCode='" & mCompanyCode & "'"
                ComSetUp.Execute
         End If
                
                'Update Iusers

        Set Gen = New ClsGeneral
        Set rsSetup = Gen.CheckSignOn          'Chk From QuotHeader
        If rsSetup.RecordCount <> 0 Then       'QuotNo Existing - QuotHeader
                'Update isetup
                Set ComSetUp = New Command
                ComSetUp.ActiveConnection = Dll.Conn
                ComSetUp.CommandText = "Update IUsers set SignOnDate='" & Format(PSystemDate, "mm/dd/yyyy") & "' , SignOffDate='" & Format(PSystemDate, "mm/dd/yyyy") & "' , EodDate='" & Format(PSystemDate, "mm/dd/yyyy") & "' , Float='" & Val(frmLogin.txtFloat) & "' , IsSignOn='0', IsSignOff='1', EOD='1'  Where UserName='" & UserId & "' and CompanyCode='" & mCompanyCode & "'"
                ComSetUp.Execute
        End If

End Function

