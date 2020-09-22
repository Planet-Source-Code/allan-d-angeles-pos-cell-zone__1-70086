VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmprod 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Product Maintenance"
   ClientHeight    =   7770
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   11460
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7770
   ScaleWidth      =   11460
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtprice 
      Alignment       =   2  'Center
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
      Left            =   7680
      Locked          =   -1  'True
      MaxLength       =   12
      TabIndex        =   24
      Top             =   2400
      Width           =   1575
   End
   Begin VB.TextBox txtstocks 
      Alignment       =   2  'Center
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
      Left            =   7680
      Locked          =   -1  'True
      MaxLength       =   12
      TabIndex        =   21
      Top             =   1800
      Width           =   1575
   End
   Begin VB.TextBox category 
      Height          =   285
      Left            =   5640
      TabIndex        =   20
      Text            =   "Text1"
      Top             =   2520
      Width           =   150
   End
   Begin VB.ComboBox cbocat 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1920
      Locked          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   2520
      Width           =   3735
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
      Height          =   1215
      Left            =   1920
      Locked          =   -1  'True
      MaxLength       =   100
      MultiLine       =   -1  'True
      TabIndex        =   10
      Top             =   1200
      Width           =   3735
   End
   Begin VB.TextBox txtpname 
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
      Left            =   1920
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   4
      Top             =   720
      Width           =   3735
   End
   Begin Project1.chameleonButton cmdnew 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   7200
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
      MICON           =   "frmprod.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Project1.chameleonButton cmdsave 
      Height          =   375
      Left            =   1800
      TabIndex        =   1
      Top             =   7200
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
      MICON           =   "frmprod.frx":001C
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
      Left            =   5160
      TabIndex        =   2
      Top             =   7200
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
      MICON           =   "frmprod.frx":0038
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
      Left            =   3480
      TabIndex        =   3
      Top             =   7200
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
      MICON           =   "frmprod.frx":0054
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSComctlLib.ListView lst 
      Height          =   3255
      Left            =   240
      TabIndex        =   15
      ToolTipText     =   "Double click to update record"
      Top             =   3480
      Width           =   11055
      _ExtentX        =   19500
      _ExtentY        =   5741
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
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "product id"
         Object.Width           =   4048
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "product name"
         Object.Width           =   5953
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "description"
         Object.Width           =   5953
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "category id"
         Object.Width           =   5953
      EndProperty
   End
   Begin Project1.chameleonButton cmdstocks 
      Height          =   375
      Left            =   9360
      TabIndex        =   22
      Top             =   1800
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      BTYPE           =   14
      TX              =   "&UPDATE STOCKS"
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
      MICON           =   "frmprod.frx":0070
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Project1.chameleonButton cmdprice 
      Height          =   375
      Left            =   9360
      TabIndex        =   25
      Top             =   2400
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      BTYPE           =   14
      TX              =   "&UPDATE PRICE"
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
      MICON           =   "frmprod.frx":008C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label15 
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
      Left            =   7560
      TabIndex        =   28
      Top             =   1920
      Width           =   135
   End
   Begin VB.Label Label14 
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
      Left            =   7560
      TabIndex        =   27
      Top             =   2400
      Width           =   135
   End
   Begin VB.Label Label13 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Current Price"
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
      Left            =   5880
      TabIndex        =   26
      Top             =   2520
      Width           =   2295
   End
   Begin VB.Label Label12 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Current Stocks"
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
      Left            =   5880
      TabIndex        =   23
      Top             =   1920
      Width           =   2295
   End
   Begin VB.Label Label11 
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
      Left            =   9120
      TabIndex        =   19
      Top             =   3240
      Width           =   2295
   End
   Begin VB.Label Label10 
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
      Left            =   5760
      TabIndex        =   18
      Top             =   3240
      Width           =   3375
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000006&
      Caption         =   "Product ID"
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
      Left            =   240
      TabIndex        =   17
      Top             =   3240
      Width           =   2295
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000006&
      Caption         =   "Product Name"
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
      Left            =   2400
      TabIndex        =   16
      Top             =   3240
      Width           =   3375
   End
   Begin VB.Label Label7 
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
      Left            =   1800
      TabIndex        =   14
      Top             =   2520
      Width           =   135
   End
   Begin VB.Label Label6 
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
      Left            =   1800
      TabIndex        =   13
      Top             =   1200
      Width           =   135
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Category"
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
      Left            =   240
      TabIndex        =   11
      Top             =   2520
      Width           =   2295
   End
   Begin VB.Label Label4 
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
      Left            =   240
      TabIndex        =   9
      Top             =   1200
      Width           =   2295
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Product Name"
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
      Left            =   240
      TabIndex        =   8
      Top             =   840
      Width           =   2295
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Product  ID "
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
      Left            =   240
      TabIndex        =   7
      Top             =   360
      Width           =   2295
   End
   Begin VB.Label lblpcode 
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
      Left            =   1920
      TabIndex        =   6
      Top             =   240
      Width           =   3735
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
      Left            =   1800
      TabIndex        =   5
      Top             =   720
      Width           =   135
   End
End
Attribute VB_Name = "frmprod"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cat As String



Private Sub cbocat_Click()
rst1.Open "Select * from tblcat", con, adOpenDynamic, adLockOptimistic
While rst1.EOF = False
If cbocat = rst1!Desc Then
 category = rst1!catid
End If
rst1.MoveNext
Wend
rst1.Close
End Sub

Private Sub cmdedit_Click()
Call bukas
txtpname.SetFocus
cmdedit.Enabled = False
cmdupdate.Enabled = True
End Sub
Function bukas()
txtpname.Locked = False
txtdesc.Locked = False
cbocat.Locked = False
End Function
Private Sub cmdnew_Click()
Call bukas
Call clear

rst.Open "Select * from tblprod", con, adOpenDynamic, adLockOptimistic
rst.MoveLast
lblpcode = "P" + Format(Date, "mmyy") + Format(Time, "hhmmss")
txtpname.SetFocus
rst.Close
cmdnew.Enabled = False
cmdsave.Enabled = True
cmdedit.Enabled = False
cmdupdate.Enabled = False


End Sub
Function clear()
lblpcode = ""
txtpname = ""
txtdesc = ""

End Function

Private Sub cmdprice_Click()
txtprice.Locked = False

If cmdprice.Caption = "SAVE" Then
rst.Open "Select * from tblstocks", con, adOpenDynamic, adLockOptimistic
While rst.EOF = False
If rst!pcode = lblpcode Then

rst!price = txtprice
rst.Update

MsgBox "Updated!", vbInformation, "System Message"

cmdprice.Caption = "&UPDATE PRICE"
txtprice.Locked = True
End If
rst.MoveNext
Wend
rst.Close
Else
cmdprice.Caption = "SAVE"
txtprice.SetFocus
End If
End Sub

Private Sub cmdsave_Click()
If lblpcode = "" Then
     MsgBox "Please click new button to add new record!", vbInformation, "System Message"
     cmdnew.SetFocus
ElseIf txtpname = "" Then
    MsgBox "Pname is empty!", vbCritical, "System Message"
    txtpname.SetFocus
ElseIf txtdesc = "" Then
     MsgBox "Product Description is empty!", vbCritical, "System Message"
     txtdesc.SetFocus
ElseIf cbocat = "" Then
    MsgBox "Please choose in category!", vbInformation, "System Message"
    cbocat.SetFocus
Else
  Dim q As VbMsgBoxResult
  q = MsgBox("Save record " + lblpcode + " ?", vbQuestion + vbYesNo, "System Message")
  If q = vbYes Then
  rst.Open "Select * from tblprod", con, adOpenDynamic, adLockPessimistic
  rst.AddNew
  
  rst!pcode = lblpcode
  rst!pname = txtpname
  rst!pdesc = txtdesc
  rst!pcat = category
  
  rst.Update
  rst.Close
  
  rst.Open "Select * from tblstocks", con, adOpenDynamic, adLockOptimistic
  rst.AddNew
    rst!pcode = lblpcode
  rst.Update
  rst.Close
  
  MsgBox "Saved!", vbInformation, "System Message"
  
  lblpcode = ""
  txtpname = ""
  txtdesc = ""

  
  cmdnew.Enabled = True
  cmdsave.Enabled = False
  
  Call sarado
  Call reload
  
  End If
End If

End Sub
Function sarado()
txtpname.Locked = True
txtdesc.Locked = True
cbocat.Locked = True
End Function


Private Sub cmdstocks_Click()
txtstocks.Locked = False

If cmdstocks.Caption = "SAVE" Then
rst.Open "Select * from tblstocks", con, adOpenDynamic, adLockOptimistic
While rst.EOF = False
If rst!pcode = lblpcode Then

rst!stocks = txtstocks
rst.Update

MsgBox "Updated!", vbInformation, "System Message"

cmdstocks.Caption = "&UPDATE STOCKS"
txtstocks.Locked = True
End If
rst.MoveNext
Wend
rst.Close
Else
cmdstocks.Caption = "SAVE"
txtstocks.SetFocus
End If

End Sub

Private Sub cmdupdate_Click()
If lblpcode = "" Then
     MsgBox "Please click new button to add new record!", vbInformation, "System Message"
     cmdnew.SetFocus
ElseIf txtpname = "" Then
    MsgBox "Pname is empty!", vbCritical, "System Message"
    txtpname.SetFocus
ElseIf txtdesc = "" Then
     MsgBox "Product Description is empty!", vbCritical, "System Message"
     txtdesc.SetFocus
ElseIf cbocat = "" Then
    MsgBox "Please choose in category!", vbInformation, "System Message"
    cbocat.SetFocus
Else
  Dim q As VbMsgBoxResult
  q = MsgBox("Update record " + lblpcode + " ?", vbQuestion + vbYesNo, "System Message")
  If q = vbYes Then
  rst.Open "Select * from tblprod", con, adOpenDynamic, adLockPessimistic
  While rst.EOF = False
  If rst!pcode = lblpcode Then
  rst!pname = txtpname
  rst!pdesc = txtdesc
  rst!pcat = category
  
  rst.Update
  End If
  rst.MoveNext
  Wend
  rst.Close
  
  MsgBox "Updated!", vbInformation, "System Message"
  
  lblpcode = ""
  txtpname = ""
  txtdesc = ""

  
  cmdnew.Enabled = True
  cmdsave.Enabled = False
  cmdupdate.Enabled = False
  
  Call sarado
  Call reload
  
  End If
End If
End Sub

Private Sub Form_Load()
cbocat.clear
rst.Open "Select * from tblcat", con, adOpenDynamic, adLockOptimistic
While rst.EOF = False
cbocat.AddItem rst!Desc
rst.MoveNext
Wend
rst.Close

Call reload
End Sub

Function reload()
rst.Open "select * from tblprod", con, adOpenDynamic, adLockOptimistic

lst.ListItems.clear
While rst.EOF = False
lst.ListItems.Add , , rst!pcode
lst.ListItems(lst.ListItems.Count).ListSubItems.Add , , rst!pname
lst.ListItems(lst.ListItems.Count).ListSubItems.Add , , rst!pdesc
lst.ListItems(lst.ListItems.Count).ListSubItems.Add , , rst!pcat
rst.MoveNext
Wend
rst.Close

End Function

Private Sub lst_Click()
cmdnew.Enabled = True
cmdsave.Enabled = False
lblpcode = lst.SelectedItem
txtpname = lst.SelectedItem.SubItems(1)
txtdesc = lst.SelectedItem.SubItems(2)
cat = lst.SelectedItem.SubItems(3)

rst.Open "select * from tblcat", con, adOpenDynamic, adLockOptimistic
While rst.EOF = False
If cat = rst!catid Then
cbocat = rst!Desc
End If
rst.MoveNext
Wend
rst.Close

rst.Open "Select * from tblstocks", con, adOpenDynamic, adLockOptimistic
While rst.EOF = False
If lst.SelectedItem = rst!pcode Then
    cmdstocks.Enabled = True
    cmdprice.Enabled = True

    If IsNull(rst!stocks) Then
    txtstocks = ""
    Else
    txtstocks = rst!stocks
    End If
    
    If IsNull(rst!price) Then
    txtprice = ""
    Else
    txtprice = rst!price
    End If
End If
rst.MoveNext
Wend
rst.Close

End Sub

Private Sub lst_DblClick()
cmdedit.Enabled = True
End Sub

Private Sub txtprice_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
Case Is < 32
Case 48 To 57
Case 46
Case Else
    KeyAscii = 0
End Select
If KeyAscii = 13 Then cmdprice_Click

End Sub

Private Sub txtstocks_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
Case Is < 32
Case 48 To 57
Case 46
Case Else
    KeyAscii = 0
End Select
If KeyAscii = 13 Then cmdstocks_Click

End Sub
