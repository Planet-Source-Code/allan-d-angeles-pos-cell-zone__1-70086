VERSION 5.00
Begin VB.Form frmreport 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Reports"
   ClientHeight    =   3165
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3300
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3165
   ScaleWidth      =   3300
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txttot 
      Height          =   285
      Left            =   240
      TabIndex        =   12
      Text            =   "Text1"
      Top             =   3480
      Width           =   1575
   End
   Begin VB.TextBox txtch 
      Height          =   285
      Left            =   1320
      TabIndex        =   11
      Top             =   3600
      Width           =   1695
   End
   Begin VB.PictureBox picmonth 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   120
      ScaleHeight     =   1095
      ScaleWidth      =   3015
      TabIndex        =   5
      Top             =   1560
      Width           =   3015
      Begin VB.TextBox yyyy 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   1080
         MaxLength       =   4
         TabIndex        =   8
         Text            =   "yyyy"
         Top             =   360
         Width           =   855
      End
      Begin VB.ComboBox mm 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         ItemData        =   "frmreport.frx":0000
         Left            =   240
         List            =   "frmreport.frx":0061
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   360
         Width           =   855
      End
      Begin VB.OptionButton optm 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Monthly Reports"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   0
         Width           =   2655
      End
      Begin VB.Shape Shape2 
         Height          =   855
         Left            =   0
         Top             =   120
         Width           =   3015
      End
   End
   Begin VB.PictureBox picday 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   120
      ScaleHeight     =   1095
      ScaleWidth      =   3135
      TabIndex        =   0
      Top             =   480
      Width           =   3135
      Begin VB.TextBox txtyyyy 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   1920
         MaxLength       =   4
         TabIndex        =   4
         Text            =   "yyyy"
         Top             =   360
         Width           =   855
      End
      Begin VB.ComboBox cbomm 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         ItemData        =   "frmreport.frx":00E1
         Left            =   240
         List            =   "frmreport.frx":0109
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   360
         Width           =   855
      End
      Begin VB.ComboBox cbodd 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         ItemData        =   "frmreport.frx":013D
         Left            =   1080
         List            =   "frmreport.frx":019E
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   360
         Width           =   855
      End
      Begin VB.OptionButton optd 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Daily Reports"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   0
         Width           =   1935
      End
      Begin VB.Shape Shape1 
         Height          =   855
         Left            =   0
         Top             =   120
         Width           =   3015
      End
   End
   Begin Project1.chameleonButton cmdprint 
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   2640
      Width           =   1335
      _extentx        =   2355
      _extenty        =   661
      btype           =   3
      tx              =   "Print"
      enab            =   0   'False
      font            =   "frmreport.frx":021E
      coltype         =   1
      focusr          =   -1  'True
      bcol            =   14215660
      bcolo           =   14215660
      fcol            =   0
      fcolo           =   0
      mcol            =   12632256
      mptr            =   1
      micon           =   "frmreport.frx":024A
      umcol           =   -1  'True
      soft            =   0   'False
      picpos          =   0
      ngrey           =   0   'False
      fx              =   0
      hand            =   0   'False
      check           =   0   'False
      value           =   0   'False
   End
   Begin Project1.chameleonButton cmdclose 
      Height          =   375
      Left            =   1800
      TabIndex        =   10
      Top             =   2640
      Width           =   1335
      _extentx        =   2355
      _extenty        =   661
      btype           =   3
      tx              =   "&Close"
      enab            =   -1  'True
      font            =   "frmreport.frx":0268
      coltype         =   1
      focusr          =   -1  'True
      bcol            =   14215660
      bcolo           =   14215660
      fcol            =   0
      fcolo           =   0
      mcol            =   12632256
      mptr            =   1
      micon           =   "frmreport.frx":0294
      umcol           =   -1  'True
      soft            =   0   'False
      picpos          =   0
      ngrey           =   0   'False
      fx              =   0
      hand            =   0   'False
      check           =   0   'False
      value           =   0   'False
   End
   Begin VB.Label lblreports 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "asa"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   3255
   End
End
Attribute VB_Name = "frmreport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdclose_Click()
Unload Me
End Sub

Private Sub cmdprint_Click()
If txtch = "Sales" Then
If optd.Value = True Then

rst.Open "Select * from tblsales where mm='" & cbomm & "' and dd='" & cbodd & "' and yyyy='" & txtyyyy & "'", con, adOpenDynamic, adLockOptimistic
While rst.EOF = False

Set dtpsreport.DataSource = rst
dtpsreport.Sections("Section2").Controls.Item("lbldaily").Caption = "Daily"
'dtpsreport.Sections("Section5").Controls.Item("lbluser").Caption = mdimain.lblname
dtpsreport.Sections("Section5").Controls.Item("lbldate").Caption = Date
dtpsreport.Sections("Section5").Controls.Item("lbltot").Caption = txttot
dtpsreport.Show
cmdprint.Enabled = False
rst.MoveNext
Wend

Else

rst.Open "Select * from tblsales where mm='" & mm & "' and yyyy='" & yyyy & "'", con, adOpenDynamic, adLockOptimistic
While rst.EOF = False

Set dtpsreport.DataSource = rst
dtpsreport.Sections("Section2").Controls.Item("lbldaily").Caption = "Monthly"
'dtpsreport.Sections("Section5").Controls.Item("lbluser").Caption = mdimain.lblname
dtpsreport.Sections("Section5").Controls.Item("lbldate").Caption = Date
dtpsreport.Sections("Section5").Controls.Item("lbltot").Caption = txttot
dtpsreport.Show
cmdprint.Enabled = False
rst.MoveNext
Wend

End If


End If

End Sub

Private Sub optd_Click()
cbomm.Enabled = True
cbodd.Enabled = True
txtyyyy.Enabled = True
mm.Enabled = False
yyyy.Enabled = False
optd.Value = True
optm.Value = False
End Sub

Private Sub optm_Click()
cbomm.Enabled = False
cbodd.Enabled = False
txtyyyy.Enabled = False
mm.Enabled = True
yyyy.Enabled = True
optd.Value = False
optm.Value = True
End Sub

Private Sub txttot_Change()
txttot = Format(txttot, "00.00")
End Sub

Private Sub txtyyyy_Change()
If Len(txtyyyy) = 4 Then
If txtch = "Sales" Then
rst.Open "Select * from tblsales", con, adOpenDynamic, adLockOptimistic
txttot = ""
 While rst.EOF = False
 If cbomm = rst!mm And cbodd = rst!dd And txtyyyy = rst!yyyy Then
     txttot = Val(txttot) + Val(rst!tot)
 End If
 rst.MoveNext
 Wend
 rst.Close

End If

cmdprint.Enabled = True
cmdprint.SetFocus
End If
End Sub

Private Sub txtyyyy_Click()
txtyyyy = ""
End Sub

Private Sub yyyy_Change()
If Len(yyyy) = 4 Then
If txtch = "Sales" Then
rst.Open "Select * from tblsales", con, adOpenDynamic, adLockOptimistic
txttot = ""
 While rst.EOF = False
 If mm = rst!mm And yyyy = rst!yyyy Then
     txttot = Val(txttot) + Val(rst!tot)
 End If
 rst.MoveNext
 Wend
 rst.Close

End If

cmdprint.Enabled = True
cmdprint.SetFocus
End If

End Sub

Private Sub yyyy_Click()
yyyy = ""
End Sub
