VERSION 5.00
Begin VB.Form frmaboutus 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "About us"
   ClientHeight    =   5550
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6645
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5550
   ScaleWidth      =   6645
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4215
      Left            =   0
      ScaleHeight     =   4215
      ScaleWidth      =   6615
      TabIndex        =   3
      Top             =   1320
      Width           =   6615
      Begin VB.Line Line1 
         X1              =   720
         X2              =   6120
         Y1              =   480
         Y2              =   480
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Point Of Sales And Inventory System"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   0
         Width           =   7335
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00808080&
         Caption         =   " Copyrights (c) 2007 Developed By gieoww"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Index           =   1
         Left            =   2880
         TabIndex        =   7
         Top             =   3960
         Width           =   3585
      End
      Begin VB.Label lblWarning 
         BackStyle       =   0  'Transparent
         Caption         =   "Unauthorized Used of This Software Is Strictly Prohibited"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   240
         TabIndex        =   6
         Top             =   1440
         Width           =   6735
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "gieoww_23@yahoo.com"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   240
         TabIndex        =   5
         Top             =   2160
         Width           =   2025
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "For Comments Or Suggestions Please Email Us At "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   2
         Left            =   240
         TabIndex        =   4
         Top             =   1800
         Width           =   4095
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1335
      Left            =   0
      ScaleHeight     =   1335
      ScaleWidth      =   6615
      TabIndex        =   0
      Top             =   0
      Width           =   6615
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Charles Cell Zone"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   495
         Left            =   600
         TabIndex        =   2
         Top             =   120
         Width           =   5655
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Mega Center the Mall"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   495
         Left            =   720
         TabIndex        =   1
         Top             =   720
         Width           =   5415
      End
   End
End
Attribute VB_Name = "frmaboutus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Unload Me
End Sub

Private Sub imgclose_Click()
Unload Me
End Sub
