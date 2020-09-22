VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm mdimain 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   Caption         =   "Point of Sales and Inventory System for Aeulala "
   ClientHeight    =   11010
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8370
   LinkTopic       =   "MDIForm1"
   Moveable        =   0   'False
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture2 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1695
      Left            =   0
      ScaleHeight     =   1695
      ScaleWidth      =   8370
      TabIndex        =   14
      Top             =   9315
      Width           =   8370
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   360
         Left            =   14640
         Picture         =   "MDIForm1.frx":0000
         Stretch         =   -1  'True
         Top             =   1080
         Width           =   480
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Log Out"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   375
         Left            =   12480
         TabIndex        =   21
         Top             =   1200
         Width           =   1935
      End
      Begin VB.Label Label7 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Current User"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   0
         TabIndex        =   18
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label lblname 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Current User"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   0
         TabIndex        =   17
         Top             =   600
         Width           =   3015
      End
      Begin VB.Label Label8 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "User Level"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   0
         TabIndex        =   16
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label lbldep 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Current User"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   0
         TabIndex        =   15
         Top             =   1200
         Width           =   3015
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   3  'Align Left
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   8460
      Left            =   0
      ScaleHeight     =   8460
      ScaleWidth      =   3360
      TabIndex        =   1
      Top             =   855
      Width           =   3360
      Begin VB.PictureBox picabout 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   855
         Left            =   960
         ScaleHeight     =   855
         ScaleWidth      =   2415
         TabIndex        =   19
         Top             =   6720
         Visible         =   0   'False
         Width           =   2415
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "About Us"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Left            =   0
            TabIndex        =   20
            Top             =   120
            Width           =   1935
         End
         Begin VB.Line Line34 
            BorderWidth     =   3
            X1              =   0
            X2              =   1200
            Y1              =   840
            Y2              =   840
         End
         Begin VB.Line Line33 
            BorderWidth     =   3
            X1              =   1200
            X2              =   1200
            Y1              =   840
            Y2              =   480
         End
         Begin VB.Line Line32 
            BorderWidth     =   3
            X1              =   120
            X2              =   2040
            Y1              =   480
            Y2              =   480
         End
         Begin VB.Line Line31 
            BorderWidth     =   3
            X1              =   120
            X2              =   120
            Y1              =   480
            Y2              =   240
         End
         Begin VB.Line Line30 
            BorderWidth     =   3
            X1              =   2040
            X2              =   2040
            Y1              =   480
            Y2              =   240
         End
      End
      Begin VB.PictureBox picdate 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   855
         Left            =   960
         ScaleHeight     =   855
         ScaleWidth      =   2415
         TabIndex        =   12
         Top             =   5640
         Visible         =   0   'False
         Width           =   2415
         Begin VB.Line Line29 
            BorderWidth     =   3
            X1              =   2040
            X2              =   2040
            Y1              =   480
            Y2              =   240
         End
         Begin VB.Line Line28 
            BorderWidth     =   3
            X1              =   120
            X2              =   120
            Y1              =   480
            Y2              =   240
         End
         Begin VB.Line Line27 
            BorderWidth     =   3
            X1              =   120
            X2              =   2040
            Y1              =   480
            Y2              =   480
         End
         Begin VB.Line Line26 
            BorderWidth     =   3
            X1              =   1200
            X2              =   1200
            Y1              =   840
            Y2              =   480
         End
         Begin VB.Line Line25 
            BorderWidth     =   3
            X1              =   0
            X2              =   1200
            Y1              =   840
            Y2              =   840
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Date/Time Setting"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   615
            Left            =   120
            TabIndex        =   13
            Top             =   0
            Width           =   1935
         End
      End
      Begin VB.PictureBox picuser 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   855
         Left            =   960
         ScaleHeight     =   855
         ScaleWidth      =   2415
         TabIndex        =   10
         Top             =   4680
         Visible         =   0   'False
         Width           =   2415
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "User Account"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   255
            Left            =   360
            TabIndex        =   11
            Top             =   0
            Width           =   1695
         End
         Begin VB.Line Line24 
            BorderWidth     =   3
            X1              =   0
            X2              =   1200
            Y1              =   720
            Y2              =   720
         End
         Begin VB.Line Line23 
            BorderWidth     =   3
            X1              =   1200
            X2              =   1200
            Y1              =   720
            Y2              =   360
         End
         Begin VB.Line Line22 
            BorderWidth     =   3
            X1              =   120
            X2              =   2040
            Y1              =   360
            Y2              =   360
         End
         Begin VB.Line Line21 
            BorderWidth     =   3
            X1              =   120
            X2              =   120
            Y1              =   360
            Y2              =   120
         End
         Begin VB.Line Line20 
            BorderWidth     =   3
            X1              =   2040
            X2              =   2040
            Y1              =   360
            Y2              =   120
         End
      End
      Begin VB.PictureBox picprod 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1335
         Left            =   960
         ScaleHeight     =   1335
         ScaleWidth      =   2415
         TabIndex        =   8
         Top             =   -120
         Visible         =   0   'False
         Width           =   2415
         Begin VB.Label lblprod 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Product Maintenance"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   615
            Left            =   0
            TabIndex        =   9
            Top             =   240
            Width           =   1935
         End
         Begin VB.Line l2 
            BorderWidth     =   3
            X1              =   0
            X2              =   1200
            Y1              =   1200
            Y2              =   1200
         End
         Begin VB.Line Line2 
            BorderWidth     =   3
            X1              =   1200
            X2              =   1200
            Y1              =   1200
            Y2              =   840
         End
         Begin VB.Line Line3 
            BorderWidth     =   3
            X1              =   120
            X2              =   2040
            Y1              =   840
            Y2              =   840
         End
         Begin VB.Line Line5 
            BorderWidth     =   3
            X1              =   120
            X2              =   120
            Y1              =   840
            Y2              =   600
         End
         Begin VB.Line Line6 
            BorderWidth     =   3
            X1              =   2040
            X2              =   2040
            Y1              =   840
            Y2              =   600
         End
      End
      Begin VB.PictureBox piccat 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   855
         Left            =   960
         ScaleHeight     =   855
         ScaleWidth      =   2415
         TabIndex        =   6
         Top             =   1440
         Visible         =   0   'False
         Width           =   2415
         Begin VB.Line Line1 
            BorderWidth     =   3
            X1              =   2040
            X2              =   2040
            Y1              =   360
            Y2              =   120
         End
         Begin VB.Line Line4 
            BorderWidth     =   3
            X1              =   120
            X2              =   120
            Y1              =   360
            Y2              =   120
         End
         Begin VB.Line Line7 
            BorderWidth     =   3
            X1              =   120
            X2              =   2040
            Y1              =   360
            Y2              =   360
         End
         Begin VB.Line Line8 
            BorderWidth     =   3
            X1              =   1200
            X2              =   1200
            Y1              =   720
            Y2              =   360
         End
         Begin VB.Line Line9 
            BorderWidth     =   3
            X1              =   0
            X2              =   1200
            Y1              =   720
            Y2              =   720
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Category"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   255
            Left            =   240
            TabIndex        =   7
            Top             =   0
            Width           =   1695
         End
      End
      Begin VB.PictureBox picsales 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1095
         Left            =   960
         ScaleHeight     =   1095
         ScaleWidth      =   2415
         TabIndex        =   4
         Top             =   2280
         Visible         =   0   'False
         Width           =   2415
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Sales Transaction"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   495
            Left            =   480
            TabIndex        =   5
            Top             =   0
            Width           =   1455
         End
         Begin VB.Line Line10 
            BorderWidth     =   3
            X1              =   0
            X2              =   1200
            Y1              =   960
            Y2              =   960
         End
         Begin VB.Line Line11 
            BorderWidth     =   3
            X1              =   1200
            X2              =   1200
            Y1              =   960
            Y2              =   600
         End
         Begin VB.Line Line12 
            BorderWidth     =   3
            X1              =   120
            X2              =   2040
            Y1              =   600
            Y2              =   600
         End
         Begin VB.Line Line13 
            BorderWidth     =   3
            X1              =   120
            X2              =   120
            Y1              =   600
            Y2              =   360
         End
         Begin VB.Line Line14 
            BorderWidth     =   3
            X1              =   2040
            X2              =   2040
            Y1              =   600
            Y2              =   360
         End
      End
      Begin VB.PictureBox picreport 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   855
         Left            =   960
         ScaleHeight     =   855
         ScaleWidth      =   2415
         TabIndex        =   2
         Top             =   3600
         Visible         =   0   'False
         Width           =   2415
         Begin VB.Line Line15 
            BorderWidth     =   3
            X1              =   2040
            X2              =   2040
            Y1              =   360
            Y2              =   120
         End
         Begin VB.Line Line16 
            BorderWidth     =   3
            X1              =   120
            X2              =   120
            Y1              =   360
            Y2              =   120
         End
         Begin VB.Line Line17 
            BorderWidth     =   3
            X1              =   120
            X2              =   2040
            Y1              =   360
            Y2              =   360
         End
         Begin VB.Line Line18 
            BorderWidth     =   3
            X1              =   1200
            X2              =   1200
            Y1              =   720
            Y2              =   360
         End
         Begin VB.Line Line19 
            BorderWidth     =   3
            X1              =   0
            X2              =   1200
            Y1              =   720
            Y2              =   720
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Sales Report"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   255
            Left            =   360
            TabIndex        =   3
            Top             =   0
            Width           =   1695
         End
      End
      Begin VB.Image imgabout 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   600
         Left            =   240
         Picture         =   "MDIForm1.frx":0E42
         Stretch         =   -1  'True
         Top             =   7080
         Width           =   720
      End
      Begin VB.Image imgdate 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   600
         Left            =   240
         Picture         =   "MDIForm1.frx":161D
         Stretch         =   -1  'True
         Top             =   6000
         Width           =   720
      End
      Begin VB.Image imguser 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   600
         Left            =   240
         Picture         =   "MDIForm1.frx":1D9C
         Stretch         =   -1  'True
         Top             =   4920
         Width           =   720
      End
      Begin VB.Image imgprod 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   600
         Left            =   240
         Picture         =   "MDIForm1.frx":21DE
         Stretch         =   -1  'True
         Top             =   600
         Width           =   720
      End
      Begin VB.Image imgcat 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   600
         Left            =   240
         Picture         =   "MDIForm1.frx":2B56
         Stretch         =   -1  'True
         Top             =   1680
         Width           =   720
      End
      Begin VB.Image imgsales 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   600
         Left            =   240
         Picture         =   "MDIForm1.frx":3473
         Stretch         =   -1  'True
         Top             =   2760
         Width           =   720
      End
      Begin VB.Image imgreport 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   600
         Left            =   240
         Picture         =   "MDIForm1.frx":94B4
         Stretch         =   -1  'True
         Top             =   3840
         Width           =   720
      End
   End
   Begin MSComctlLib.ImageList img2 
      Left            =   6240
      Top             =   4080
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   21
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":9D7E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":9DDC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":9E3A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":9E98
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":9EF6
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":9F54
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":9FB2
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":A010
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":A06E
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":A0CC
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":A12A
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":A188
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":A1E6
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":A244
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":A2A2
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":A300
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":A35E
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":A3BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":A41A
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":A478
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":A4D6
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox Picture3 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   0
      ScaleHeight     =   855
      ScaleWidth      =   8370
      TabIndex        =   0
      Top             =   0
      Width           =   8370
      Begin Project1.Label Label1 
         Height          =   855
         Left            =   0
         Top             =   0
         Width           =   15735
         _extentx        =   27755
         _extenty        =   1508
         alignment       =   1
         backcolor1      =   16711680
         backcolor2      =   16777215
         backstyle       =   3
         themecolor      =   5
         bordercolor1    =   6019061
         bordercolor2    =   484732
         broderstyle     =   3
         caption         =   "Charles Cell Zone Sales and Inventory System"
         effects         =   4
         effectcolor     =   4210752
         font            =   "MDIForm1.frx":A534
         forecolor1      =   8454143
         forecolor2      =   6019061
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   14
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":A560
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":A5BE
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":A61C
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":A67A
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":A6D8
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":A736
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":A794
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":A7F2
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":A850
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":A8AE
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":A90C
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":A96A
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":A9C8
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":AA26
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   4320
      Top             =   3000
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":AA84
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":AAE2
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":AB40
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":AB9E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":ABFC
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":AC5A
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":ACB8
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":AD16
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":AD74
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":ADD2
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":AE30
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList3 
      Left            =   4920
      Top             =   3000
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":AE8E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":AEEC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":AF4A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":AFA8
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":B006
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":B064
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":B0C2
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":B120
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":B17E
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":B1DC
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":B23A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList itb32x32 
      Left            =   5520
      Top             =   3000
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   17
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":B298
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":B2F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":B354
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":B3B2
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":B410
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":B46E
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":B4CC
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":B52A
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":B588
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":B5E6
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":B644
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":B6A2
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":B700
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":B75E
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":B7BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":B81A
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":B878
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "mdimain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Image1_Click()
Label10_Click
End Sub

Private Sub imgabout_DblClick()
frmaboutus.Show
End Sub

Private Sub imgabout_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgabout.Appearance = 1
imgabout.BorderStyle = 1
picabout.Visible = True
End Sub

Private Sub imgcat_DblClick()
frmcat.Show
End Sub

Private Sub imgcat_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgcat.Appearance = 1
imgcat.BorderStyle = 1
piccat.Visible = True
End Sub

Private Sub imgdate_DblClick()
 Shell ("C:\WINDOWS\system32\control.exe date/time")
End Sub

Private Sub imgdate_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgdate.Appearance = 1
imgdate.BorderStyle = 1
picdate.Visible = True
End Sub

Private Sub imgprod_DblClick()
frmprod.Show

End Sub

Private Sub imgprod_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgprod.Appearance = 1
imgprod.BorderStyle = 1
picprod.Visible = True

End Sub

Private Sub imgreport_Click()
 frmreport.txtch = "Sales"
frmreport.lblreports = "-== Sales Report ==-"
End Sub

Private Sub imgreport_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgreport.Appearance = 1
imgreport.BorderStyle = 1
picreport.Visible = True
End Sub

Private Sub imgsales_DblClick()
frmsales.Show
End Sub

Private Sub imgsales_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgsales.Appearance = 1
imgsales.BorderStyle = 1
picsales.Visible = True
End Sub

Private Sub imguser_DblClick()
frmuser.Show
End Sub

Private Sub imguser_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
imguser.Appearance = 1
imguser.BorderStyle = 1
picuser.Visible = True
End Sub

Private Sub Label10_Click()
frmlogin.Show
Unload Me
End Sub

Private Sub MDIForm_Load()
connection.connect
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgprod.Appearance = 0
imgprod.BorderStyle = 0
picprod.Visible = False

imgcat.Appearance = 0
imgcat.BorderStyle = 0
piccat.Visible = False

imgsales.Appearance = 0
imgsales.BorderStyle = 0
picsales.Visible = False

imgreport.Appearance = 0
imgreport.BorderStyle = 0
picreport.Visible = False


imguser.Appearance = 0
imguser.BorderStyle = 0
picuser.Visible = False

imgdate.Appearance = 0
imgdate.BorderStyle = 0
picdate.Visible = False

imgabout.Appearance = 0
imgabout.BorderStyle = 0
picabout.Visible = False
End Sub
