VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmsales 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Sales Invoice"
   ClientHeight    =   7905
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   9525
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7905
   ScaleWidth      =   9525
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picpayment 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0FF&
      ForeColor       =   &H80000008&
      Height          =   2175
      Left            =   2520
      ScaleHeight     =   2145
      ScaleWidth      =   4185
      TabIndex        =   23
      Top             =   3360
      Visible         =   0   'False
      Width           =   4215
      Begin VB.TextBox txtcash 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Book Antiqua"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   1200
         TabIndex        =   26
         Top             =   360
         Width           =   2775
      End
      Begin VB.CommandButton cmdprint 
         Caption         =   "Print"
         Height          =   255
         Left            =   240
         TabIndex        =   25
         Top             =   1800
         Width           =   855
      End
      Begin VB.CommandButton cmdcancel 
         Caption         =   "Cancel"
         Height          =   255
         Left            =   1320
         TabIndex        =   24
         Top             =   1800
         Width           =   855
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Cash"
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
         TabIndex        =   29
         Top             =   600
         Width           =   2295
      End
      Begin VB.Label Label12 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Change"
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
         TabIndex        =   28
         Top             =   1200
         Width           =   2295
      End
      Begin VB.Label lblchange 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   1200
         TabIndex        =   27
         Top             =   1200
         Width           =   2655
      End
   End
   Begin VB.PictureBox picinfo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   6855
      Left            =   0
      ScaleHeight     =   6855
      ScaleWidth      =   9615
      TabIndex        =   2
      Top             =   0
      Width           =   9615
      Begin VB.TextBox pcode 
         Height          =   285
         Left            =   9720
         TabIndex        =   20
         Text            =   "Text1"
         Top             =   360
         Width           =   150
      End
      Begin VB.ComboBox cboprod 
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
         TabIndex        =   5
         Top             =   720
         Width           =   3735
      End
      Begin VB.TextBox txtqty 
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
         Top             =   2040
         Width           =   3735
      End
      Begin Project1.chameleonButton cmdpurchase 
         Height          =   375
         Left            =   6720
         TabIndex        =   3
         Top             =   2160
         Width           =   1095
         _extentx        =   1931
         _extenty        =   661
         btype           =   5
         tx              =   "PURCHASE"
         enab            =   -1  'True
         font            =   "frmsales.frx":0000
         coltype         =   2
         focusr          =   -1  'True
         bcol            =   14215660
         bcolo           =   14215660
         fcol            =   0
         fcolo           =   0
         mcol            =   12632256
         mptr            =   1
         micon           =   "frmsales.frx":002C
         umcol           =   -1  'True
         soft            =   0   'False
         picpos          =   0
         ngrey           =   0   'False
         fx              =   0
         hand            =   0   'False
         check           =   0   'False
         value           =   0   'False
      End
      Begin Project1.chameleonButton cmddrop 
         Height          =   375
         Left            =   8160
         TabIndex        =   6
         Top             =   2160
         Width           =   1095
         _extentx        =   1931
         _extenty        =   661
         btype           =   5
         tx              =   "DROP ITEM"
         enab            =   0   'False
         font            =   "frmsales.frx":004A
         coltype         =   2
         focusr          =   -1  'True
         bcol            =   14215660
         bcolo           =   14215660
         fcol            =   0
         fcolo           =   0
         mcol            =   12632256
         mptr            =   1
         micon           =   "frmsales.frx":0076
         umcol           =   -1  'True
         soft            =   0   'False
         picpos          =   0
         ngrey           =   0   'False
         fx              =   0
         hand            =   0   'False
         check           =   0   'False
         value           =   0   'False
      End
      Begin MSComctlLib.ListView lst 
         Height          =   3255
         Left            =   240
         TabIndex        =   7
         ToolTipText     =   "Double click to update record"
         Top             =   2880
         Width           =   9015
         _ExtentX        =   15901
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
            Text            =   "Qty"
            Object.Width           =   1931
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Product Name"
            Object.Width           =   5953
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Price"
            Object.Width           =   4048
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Amount"
            Object.Width           =   4048
         EndProperty
      End
      Begin VB.Label Label11 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Total"
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
         Left            =   5520
         TabIndex        =   22
         Top             =   6480
         Width           =   2295
      End
      Begin VB.Label lbltot 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
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
         Left            =   6960
         TabIndex        =   21
         Top             =   6480
         Width           =   2175
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Transaction No."
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
         TabIndex        =   19
         Top             =   360
         Width           =   2295
      End
      Begin VB.Label lbltrans 
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
         TabIndex        =   18
         Top             =   240
         Width           =   3735
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Product"
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
         TabIndex        =   17
         Top             =   840
         Width           =   2295
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Stocks"
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
         TabIndex        =   16
         Top             =   1200
         Width           =   2295
      End
      Begin VB.Label lblstocks 
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
         TabIndex        =   15
         Top             =   1080
         Width           =   3735
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Price"
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
         TabIndex        =   14
         Top             =   1680
         Width           =   2295
      End
      Begin VB.Label lblprice 
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
         TabIndex        =   13
         Top             =   1560
         Width           =   3735
      End
      Begin VB.Label Label7 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Quantity"
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
         TabIndex        =   12
         Top             =   2160
         Width           =   2295
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         Caption         =   "Price"
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
         Left            =   4560
         TabIndex        =   11
         Top             =   2640
         Width           =   2295
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         Caption         =   "Quantity"
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
         Top             =   2640
         Width           =   1095
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
         Left            =   1200
         TabIndex        =   9
         Top             =   2640
         Width           =   3375
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         Caption         =   "Amount"
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
         Left            =   6840
         TabIndex        =   8
         Top             =   2640
         Width           =   2415
      End
   End
   Begin Project1.chameleonButton cmdnew 
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   7320
      Width           =   1335
      _extentx        =   2355
      _extenty        =   873
      btype           =   14
      tx              =   "&NEW"
      enab            =   -1  'True
      font            =   "frmsales.frx":0094
      coltype         =   2
      focusr          =   -1  'True
      bcol            =   65535
      bcolo           =   65535
      fcol            =   16711680
      fcolo           =   16711680
      mcol            =   12632256
      mptr            =   1
      micon           =   "frmsales.frx":00C0
      umcol           =   -1  'True
      soft            =   0   'False
      picpos          =   0
      ngrey           =   0   'False
      fx              =   0
      hand            =   0   'False
      check           =   0   'False
      value           =   0   'False
   End
   Begin Project1.chameleonButton cmdpayment 
      Height          =   495
      Left            =   1800
      TabIndex        =   1
      Top             =   7320
      Width           =   1335
      _extentx        =   2355
      _extenty        =   873
      btype           =   14
      tx              =   "&PAYMENT"
      enab            =   0   'False
      font            =   "frmsales.frx":00DE
      coltype         =   2
      focusr          =   -1  'True
      bcol            =   65535
      bcolo           =   65535
      fcol            =   16711680
      fcolo           =   16711680
      mcol            =   12632256
      mptr            =   1
      micon           =   "frmsales.frx":010A
      umcol           =   -1  'True
      soft            =   0   'False
      picpos          =   0
      ngrey           =   0   'False
      fx              =   0
      hand            =   0   'False
      check           =   0   'False
      value           =   0   'False
   End
End
Attribute VB_Name = "frmsales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cboprod_Click()
rst.Open "Select * from tblprod", con, adOpenDynamic, adLockOptimistic
While rst.EOF = False
If cboprod = rst!pname Then
pcode = rst!pcode
End If
rst.MoveNext
Wend
rst.Close


rst.Open "Select * from tblstocks", con, adOpenDynamic, adLockOptimistic
While rst.EOF = False
If pcode = rst!pcode Then
    lblstocks = rst!stocks
    lblprice = rst!price
End If
rst.MoveNext
Wend
rst.Close

End Sub

Private Sub cmdcancel_Click()
picpayment.Visible = False

lbltot = ""
txtcash = ""
lblchange = ""
Call clear
End Sub

Private Sub cmddrop_Click()
Dim X As Boolean
rst.Open "Select * from tblsales", con, adOpenDynamic, adLockOptimistic
While rst.EOF = False
If lbltrans = rst!transid Then
X = True
End If
rst.MoveNext
Wend
rst.Close

If X = True Then
    rst.Open "Select * from sales", con, adOpenDynamic, adLockOptimistic
    While rst.EOF = False
    If lbltrans = rst!transid And rst!qty = lst.SelectedItem And rst!pdesc = lst.SelectedItem.SubItems(1) Then
    rst.Delete
    rst.Update
    lbltot = Val(lbltot) - lst.SelectedItem.SubItems(3)
    End If
    rst.MoveNext
    Wend
    rst.Close
    
    rst.Open "Select * from tblstocks", con, adOpenDynamic, adLockOptimistic
    While rst.EOF = False
    If pcode = rst!pcode Then
        rst!stocks = Val(rst!stocks) + lst.SelectedItem
        rst.Update
    End If
    rst.MoveNext
    Wend
    rst.Close
    Call reload
    
    cmddrop.Enabled = False
    
End If

End Sub

Private Sub cmdnew_Click()
picinfo.Enabled = True
lbltrans = "TRANS" + Format(Date, "mmyy") + Format(Time, "hhmmss")
cboprod.SetFocus
Call bukas
cmdnew.Enabled = False
cmdpayment.Enabled = True

End Sub

Private Sub cmdpayment_Click()
picpayment.Visible = True
picinfo.Enabled = False
End Sub

Private Sub cmdprint_Click()
txtcash = Format(txtcash, "00.00")
 If txtcash >= lbltot Then
  lblchange = Val(txtcash) - Val(lbltot)
 Else
    MsgBox "Insufficient amount!", vbCritical, "System Message"
    txtcash.SetFocus
End If
 
rst.Open "Select * from tblsales", con, adOpenDynamic, adLockOptimistic
While rst.EOF = False
If lbltrans = rst!transid Then
rst!tot = lbltot
rst.Update
End If
rst.MoveNext
Wend
rst.Close
 
 
 rst.Open "Select * from sales where transid ='" & lbltrans & "' ", con, adOpenDynamic, adLockOptimistic
    Set dtpresibo.DataSource = rst
    dtpresibo.Sections("Section5").Controls.Item("lbldate").Caption = Now
    dtpresibo.Sections("Section5").Controls.Item("lbltot").Caption = lbltot
    dtpresibo.Sections("Section5").Controls.Item("lblcash").Caption = txtcash
    dtpresibo.Sections("Section5").Controls.Item("lblchange").Caption = lblchange
    dtpresibo.Show
cmdnew.Enabled = True
cmdpayment.Enabled = False
picpayment.Visible = False

lbltot = ""
txtcash = ""
lblchange = ""
Call clear
End Sub

Private Sub cmdpurchase_Click()
If cboprod = "" Then
    MsgBox "Product is empty!", vbCritical, "System Message"
    cboprod.SetFocus
ElseIf Val(lblstocks) <= 0 Then
    MsgBox "Out of stocks!", vbCritical, "System Message"
ElseIf Val(txtqty) > Val(lblstocks) Then
    MsgBox "Insuficient stocks!", vbCritical, "System Message"
    txtqty.SetFocus
ElseIf txtqty = "" Then
    MsgBox "Quantity is empty!", vbCritical, "System Message"
Else
 Dim q As VbMsgBoxResult, X As Boolean
 q = MsgBox("Purchase " + cboprod + " ?", vbQuestion + vbYesNo, "System Message")
 If q = vbYes Then
 rst.Open "Select * from tblsales", con, adOpenDynamic, adLockOptimistic
 While rst.EOF = False
 If lbltrans = rst!transid Then
    X = True
 End If
 rst.MoveNext
 Wend
rst.Close
 
If X = False Then
rst.Open "Select * from tblsales", con, adOpenDynamic, adLockOptimistic
rst.AddNew
rst!transid = lbltrans
rst!mm = Format(Date, "mm")
rst!dd = Format(Date, "dd")
rst!yyyy = Format(Date, "yyyy")
rst.Update
rst.Close

rst.Open "Select * from sales", con, adOpenDynamic, adLockOptimistic
rst.AddNew
rst!transid = lbltrans
rst!pdesc = cboprod
rst!qty = txtqty
rst!price = lblprice
rst!amount = Val(txtqty) * Val(lblprice)

rst.Update
rst.Close

Call reload
Call alis

lbltot = Val(lbltot) + (Val(txtqty) * Val(lblprice))

lblstocks = ""
lblprice = ""
txtqty = ""

ElseIf X = True Then
rst.Open "Select * from sales", con, adOpenDynamic, adLockOptimistic
rst.AddNew
rst!transid = lbltrans
rst!pdesc = cboprod
rst!qty = txtqty
rst!price = lblprice
rst!amount = Val(txtqty) * Val(lblprice)

rst.Update
rst.Close

Call reload
Call alis

lbltot = Val(lbltot) + (Val(txtqty) * Val(lblprice))

lblstocks = ""
lblprice = ""
txtqty = ""

End If
 
 End If


End If
End Sub

Function alis()
rst.Open "Select * from tblstocks", con, adOpenDynamic, adLockOptimistic
While rst.EOF = False
If pcode = rst!pcode Then
 rst!stocks = Val(rst!stocks) - Val(txtqty)
rst.Update
End If
rst.MoveNext
Wend
rst.Close

End Function
Function reload()
rst.Open "Select * from sales", con, adOpenDynamic, adLockOptimistic
lst.ListItems.clear
While rst.EOF = False
If lbltrans = rst!transid Then
lst.ListItems.Add , , rst!qty
lst.ListItems(lst.ListItems.Count).ListSubItems.Add , , rst!pdesc
lst.ListItems(lst.ListItems.Count).ListSubItems.Add , , rst!price
lst.ListItems(lst.ListItems.Count).ListSubItems.Add , , rst!amount
End If
rst.MoveNext
Wend
rst.Close
End Function

Private Sub Form_Load()
rst.Open "Select * from tblprod", con, adOpenDynamic, adLockOptimistic
cboprod.clear
While rst.EOF = False
cboprod.AddItem rst!pname
rst.MoveNext
Wend
rst.Close
End Sub
Function clear()
lbltrans = ""
lblprice = ""
lblstocks = ""
txtqty = ""
lst.ListItems.clear
lbltot = ""

End Function
Function sarado()
txtqty.Locked = True
cboprod.Locked = True
End Function

Function bukas()
txtqty.Locked = False
cboprod.Locked = False

End Function

Private Sub lblchange_Change()
lblchange = Format(lblchange, "00.00")
End Sub

Private Sub lbltot_Change()
lbltot = Format(lbltot, "00.00")
End Sub

Private Sub lst_DblClick()
cmddrop.Enabled = True
cmddrop.SetFocus

rst.Open "Select * from tblprod", con, adOpenDynamic, adLockOptimistic
While rst.EOF = False
If lst.SelectedItem.SubItems(1) = rst!pname Then
    pcode = rst!pcode
End If
rst.MoveNext
Wend
rst.Close

End Sub

Private Sub txtcash_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
Case Is < 32
Case 48 To 57
Case 46
Case Else
    KeyAscii = 0
End Select

If KeyAscii = 13 Then
 cmdprint_Click
End If

End Sub


Private Sub txtqty_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
Case Is < 32
Case 48 To 57
Case Else
    KeyAscii = 0
End Select
End Sub
