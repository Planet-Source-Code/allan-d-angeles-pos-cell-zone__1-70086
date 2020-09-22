VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmcat 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Category Maintenance"
   ClientHeight    =   3720
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   7380
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3720
   ScaleWidth      =   7380
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ListView lst 
      Height          =   1695
      Left            =   120
      TabIndex        =   9
      ToolTipText     =   "Double click to update record"
      Top             =   1800
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   2990
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      HideColumnHeaders=   -1  'True
      FlatScrollBar   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483635
      BackColor       =   -2147483643
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Category ID"
         Object.Width           =   4048
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Description"
         Object.Width           =   5953
      EndProperty
   End
   Begin Project1.chameleonButton cmdnew 
      Height          =   375
      Left            =   5880
      TabIndex        =   0
      Top             =   1440
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      BTYPE           =   14
      TX              =   "&NEW"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   65535
      BCOLO           =   65535
      FCOL            =   16711680
      FCOLO           =   16711680
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmcat.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.TextBox txtdesc 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1800
      MaxLength       =   25
      TabIndex        =   4
      Top             =   600
      Width           =   3735
   End
   Begin Project1.chameleonButton cmdsave 
      Height          =   375
      Left            =   5880
      TabIndex        =   1
      Top             =   2040
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      BTYPE           =   14
      TX              =   "&SAVE"
      ENAB            =   0   'False
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   65535
      BCOLO           =   65535
      FCOL            =   16711680
      FCOLO           =   16711680
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmcat.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Project1.chameleonButton cmdupdate 
      Height          =   375
      Left            =   5880
      TabIndex        =   3
      Top             =   3240
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      BTYPE           =   14
      TX              =   "&UPDATE"
      ENAB            =   0   'False
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   65535
      BCOLO           =   65535
      FCOL            =   16711680
      FCOLO           =   16711680
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmcat.frx":0038
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Project1.chameleonButton cmdedit 
      Height          =   375
      Left            =   5880
      TabIndex        =   2
      Top             =   2640
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      BTYPE           =   14
      TX              =   "&EDIT"
      ENAB            =   0   'False
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   65535
      BCOLO           =   65535
      FCOL            =   16711680
      FCOLO           =   16711680
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmcat.frx":0054
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000006&
      Caption         =   "Description"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   255
      Left            =   2280
      TabIndex        =   11
      Top             =   1440
      Width           =   3375
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000006&
      Caption         =   "Category ID"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   1440
      Width           =   2295
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   1440
      TabIndex        =   8
      Top             =   720
      Width           =   135
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
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
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   720
      Width           =   2295
   End
   Begin VB.Label lblcat 
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   1800
      TabIndex        =   6
      Top             =   120
      Width           =   3735
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Category ID "
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
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   240
      Width           =   2295
   End
End
Attribute VB_Name = "frmcat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdedit_Click()
txtdesc.SetFocus
cmdedit.Enabled = False
cmdupdate.Enabled = True
txtdesc.Locked = False
cmdnew.Enabled = False

End Sub

Private Sub cmdnew_Click()
txtdesc.Locked = False
lblcat = ""
txtdesc = ""

rst.Open "Select * from tblcat", con, adOpenDynamic, adLockOptimistic
rst.MoveLast
lblcat = "CAT" + Format(Val(Right(rst!catid, 4)) + 1, "000#")
txtdesc.SetFocus
rst.Close
cmdnew.Enabled = False
cmdsave.Enabled = True
cmdedit.Enabled = False

End Sub

Private Sub cmdsave_Click()
If txtdesc = "" Then
    MsgBox "Description is empty!", vbCritical + vbInformation, "System Message"
    txtdesc.SetFocus
Else
Dim q As VbMsgBoxResult
q = MsgBox("Save record " + lblcat, vbQuestion + vbYesNo, "System Message")
If q = vbYes Then
rst.Open "Select * from tblcat", con, adOpenDynamic, adLockOptimistic
rst.AddNew

rst!catid = lblcat
rst!Desc = txtdesc

rst.Update
rst.Close

Call reload

MsgBox "Saved!", vbInformation, "System Message"

cmdnew.Enabled = True
cmdsave.Enabled = False

lblcat = ""
txtdesc = ""

End If
End If

End Sub

Private Sub cmdupdate_Click()
If txtdesc = "" Then
    MsgBox "Description is empty!", vbCritical + vbInformation, "System Message"
    txtdesc.SetFocus
Else
Dim q As VbMsgBoxResult
q = MsgBox("Update record " + lblcat, vbQuestion + vbYesNo, "System Message")
If q = vbYes Then
rst.Open "Select * from tblcat", con, adOpenDynamic, adLockOptimistic
While rst.EOF = False
If rst!catid = lblcat Then
rst!Desc = txtdesc

rst.Update
End If
rst.MoveNext
Wend
rst.Close

Call reload

MsgBox "Updated!", vbInformation, "System Message"

cmdnew.Enabled = True
cmdsave.Enabled = False
cmdupdate.Enabled = False

lblcat = ""
txtdesc = ""

End If
End If

End Sub

Private Sub Form_Load()
Call reload
End Sub

Function reload()
rst.Open "Select * from tblcat", con, adOpenDynamic, adLockOptimistic
lst.ListItems.clear
While rst.EOF = False
lst.ListItems.Add , , rst!catid
lst.ListItems(lst.ListItems.Count).ListSubItems.Add , , rst!Desc
rst.MoveNext
Wend
rst.Close

End Function

Private Sub lst_Click()
lblcat = lst.SelectedItem
txtdesc = lst.SelectedItem.SubItems(1)
txtdesc.Locked = True
cmdsave.Enabled = False

End Sub

Private Sub lst_DblClick()
cmdedit.Enabled = True
End Sub
