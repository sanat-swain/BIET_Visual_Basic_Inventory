VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmstock 
   BackColor       =   &H00008000&
   Caption         =   "Form4"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form4"
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   Begin VB.CommandButton command1 
      BackColor       =   &H00FF0000&
      Caption         =   "show stock"
      Height          =   495
      Left            =   7920
      TabIndex        =   29
      Top             =   6240
      Width           =   2055
   End
   Begin VB.TextBox txtquantity 
      Height          =   375
      Left            =   3360
      TabIndex        =   28
      Text            =   " "
      Top             =   4440
      Width           =   1215
   End
   Begin VB.ComboBox cmbsname 
      Height          =   315
      Left            =   2880
      TabIndex        =   4
      Top             =   2400
      Width           =   3735
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00404080&
      Caption         =   "Frame2"
      Height          =   1335
      Left            =   3240
      TabIndex        =   27
      Top             =   5760
      Width           =   4095
      Begin VB.CommandButton CMDEXIT 
         Caption         =   "EXIT"
         Height          =   375
         Left            =   2880
         TabIndex        =   15
         Top             =   720
         Width           =   975
      End
      Begin VB.CommandButton CMDFIND 
         Caption         =   "FIND"
         Height          =   375
         Left            =   240
         TabIndex        =   13
         Top             =   720
         Width           =   975
      End
      Begin VB.CommandButton CMDDELETE 
         Caption         =   "DELETE"
         Height          =   375
         Left            =   1560
         TabIndex        =   14
         Top             =   720
         Width           =   975
      End
      Begin VB.CommandButton CMDUPDATE 
         Caption         =   "UPDATE"
         Height          =   375
         Left            =   2880
         TabIndex        =   12
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton CMDSAVE 
         Caption         =   "SAVE"
         Height          =   375
         Left            =   1560
         TabIndex        =   11
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton CMDADD 
         Caption         =   "ADD"
         Height          =   375
         Left            =   240
         TabIndex        =   10
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.TextBox txtsadd 
      Height          =   375
      Left            =   2880
      TabIndex        =   5
      Top             =   3000
      Width           =   2895
   End
   Begin VB.TextBox txtpcode 
      Height          =   375
      Left            =   2880
      TabIndex        =   3
      Top             =   1800
      Width           =   855
   End
   Begin VB.ComboBox CMBPNAME 
      Height          =   315
      Left            =   2880
      TabIndex        =   2
      Top             =   1320
      Width           =   2535
   End
   Begin VB.TextBox TXTSLNO 
      Height          =   375
      Left            =   2880
      TabIndex        =   1
      Top             =   840
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000040&
      Caption         =   "Frame1"
      Height          =   1815
      Left            =   600
      TabIndex        =   20
      Top             =   3600
      Width           =   10455
      Begin MSComCtl2.DTPicker cmbdate 
         Height          =   375
         Left            =   2760
         TabIndex        =   6
         Top             =   360
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         Format          =   24444929
         CurrentDate     =   36527
      End
      Begin VB.TextBox txtstock 
         Height          =   495
         Left            =   8160
         TabIndex        =   9
         Top             =   960
         Width           =   1695
      End
      Begin VB.TextBox txtpurprice 
         Height          =   375
         Left            =   8160
         TabIndex        =   8
         Top             =   360
         Width           =   1695
      End
      Begin VB.TextBox txtunit 
         Height          =   375
         Left            =   2760
         TabIndex        =   7
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label LBLSTOCK 
         AutoSize        =   -1  'True
         Caption         =   "STOCK"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   5400
         TabIndex        =   25
         Top             =   1200
         Width           =   1050
      End
      Begin VB.Label LBLPPRICE 
         AutoSize        =   -1  'True
         Caption         =   "PURCHASE PRICE"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   5280
         TabIndex        =   24
         Top             =   360
         Width           =   2370
      End
      Begin VB.Label LBLUNIT 
         AutoSize        =   -1  'True
         Caption         =   "UNIT"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   360
         TabIndex        =   23
         Top             =   1320
         Width           =   930
      End
      Begin VB.Label LBLQUANTITY 
         AutoSize        =   -1  'True
         Caption         =   "QUANTITY"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   360
         TabIndex        =   22
         Top             =   840
         Width           =   1410
      End
      Begin VB.Label LBLPDATE 
         AutoSize        =   -1  'True
         Caption         =   "PURCHASE DATE"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   360
         TabIndex        =   21
         Top             =   360
         Width           =   2370
      End
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "STOCK-OF OUR SHOP"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   555
      Left            =   3960
      TabIndex        =   26
      Top             =   0
      Width           =   5415
   End
   Begin VB.Label LBLSADDRESS 
      AutoSize        =   -1  'True
      Caption         =   "SUPPLIER ADDRESS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   480
      TabIndex        =   19
      Top             =   3120
      Width           =   2280
   End
   Begin VB.Label LBLSNAME 
      AutoSize        =   -1  'True
      Caption         =   "SUPPLIER NAME"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   480
      TabIndex        =   18
      Top             =   2520
      Width           =   1830
   End
   Begin VB.Label LBLPCODE 
      AutoSize        =   -1  'True
      Caption         =   "PARTCODE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   480
      TabIndex        =   17
      Top             =   1920
      Width           =   1260
   End
   Begin VB.Label LBLPNAME 
      AutoSize        =   -1  'True
      Caption         =   "PARTNAME"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   480
      TabIndex        =   16
      Top             =   1440
      Width           =   1275
   End
   Begin VB.Label LBLSLNO 
      AutoSize        =   -1  'True
      Caption         =   "SL NO :-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   480
      TabIndex        =   0
      Top             =   960
      Width           =   870
   End
End
Attribute VB_Name = "frmstock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sql As String
Dim cn As ADODB.Connection
Dim rs As ADODB.Recordset
Dim CMD As ADODB.Command
Private Sub CMBPNAME_click()
Set rs = New ADODB.Recordset
rs.CursorType = adOpenKeyset
rs.LockType = adLockOptimistic
rs.Source = "Select PCODE,uprice,unit from partmaster where PName like '" & Trim(CMBPNAME.Text) & "'; "
rs.ActiveConnection = cn
rs.Open
txtpcode.Text = rs(0)
txtpurprice.Text = rs(1)
txtunit.Text = rs(2)
End Sub
Private Sub cmbsname_click()
Set rs = New ADODB.Recordset
rs.CursorType = adOpenKeyset
rs.LockType = adLockOptimistic
rs.Source = "Select address from dealer  where sname like '" & Trim(cmbsname.Text) & "'; "
rs.ActiveConnection = cn
rs.Open
txtsadd.Text = rs(0)
End Sub
Private Sub CMDADD_Click()
clear
generatesno
CMBPNAME.SetFocus
CMDSAVE.Enabled = True
CMDADD.Enabled = False
End Sub

Private Sub cmddelete_Click()
Dim i As String
i = "Are you sure to delete?"
If MsgBox(i, vbYesNo, "warning?") = vbYes Then
Set rs = New ADODB.Recordset
rs.CursorType = adOpenKeyset
rs.LockType = adLockOptimistic
rs.Source = "delete sname,address from stock  where slno like'" & Trim(TXTSLNO.Text) & "';"
rs.ActiveConnection = cn
rs.Open
'Set CMD = New ADODB.Command
'CMD.CommandText = "update dealer set sname=" ",address=""  where sno like'" & Trim(Text1.Text) & "';"
'CMD.CommandText = "update dealer set sname="",address=""  where sno like'" & Trim(Text1.Text) & "';"
'CMD.CommandType = adCmdText
'Set CMD.ActiveConnection = CN
'CMD.Execute
MsgBox "Successfully deleted", vbExclamation, "Deletion"
clear2
ElseIf MsgBox(i, vbYesNo, "Warning?") = vbNo Then
MsgBox "Do you want to exit", vbOKOnly, "Stop"
End If
'End If
Set rs = Nothing
End Sub

Private Sub CMDEXIT_Click()
Unload Me
End Sub

Private Sub CMDFIND_Click()
Dim i As String
i = InputBox("Enter The slno U want to find:")
clear
Set rs = New ADODB.Recordset
rs.CursorType = adOpenKeyset
rs.LockType = adLockOptimistic
rs.Source = "Select * from stock where  slno like '" & i & "'"
rs.ActiveConnection = cn
rs.Open
If rs.EOF Then
MsgBox ("The PARTY With This Code is Not Exist")
CMDDELETE.Enabled = False
Else
clear
load
End If
CMDDELETE.Enabled = True
CMDUPDATE.Enabled = True
CMDADD.Enabled = False
rs.Close
End Sub

Private Sub CMDSAVE_Click()
Dim i As String
If TXTSLNO.Text = "" Or CMBPNAME.Text = "" Then
  MsgBox "You Should Fill SLNO And PNAME "
 Else
  i = MsgBox("Do You Want to Save   ", vbYesNo, "Save")
    If i = vbYes Then
       Set rs = New ADODB.Recordset
       rs.CursorType = adOpenKeyset
       rs.LockType = adLockOptimistic
       rs.Source = "select pcode from stockist where pcode like '" & Trim(txtpcode.Text) & "';"
       rs.ActiveConnection = cn
       rs.Open
       If rs.EOF = True Then
       rs.Close
       Set rs = New ADODB.Recordset
       rs.CursorType = adOpenKeyset
       rs.LockType = adLockOptimistic
       rs.Source = "stockist"
       rs.ActiveConnection = cn
       rs.Open
       rs.AddNew
       ASSIGN2
       rs.Update
       rs.Close
       Else
      Set CMD = New ADODB.Command
      x = Val(txtquantity.Text)
      MsgBox (x)
      CMD.CommandText = "update stockist set stock=stock+ " & Val(txtquantity.Text) & " where pcode like '" & Trim(txtpcode.Text) & "';"
      CMD.CommandType = adCmdText
      Set CMD.ActiveConnection = cn
      CMD.Execute
      MsgBox ("The Party is Successfully Modified")
      Set CMD = Nothing
      End If
       Set rs = New ADODB.Recordset
       rs.CursorType = adOpenKeyset
       rs.LockType = adLockOptimistic
       rs.Source = "stock"
       rs.ActiveConnection = cn
       rs.Open
       rs.AddNew
       Call assign
       rs.Update
       rs.Close
       Set rs = Nothing
         
End If
End If
CMBPNAME.SetFocus
CMDADD.Enabled = True
CMDSAVE.Enabled = False
End Sub

Private Sub CMDUPDATE_Click()
Set CMD = New ADODB.Command
Set rs = New ADODB.Recordset
rs.CursorType = adOpenKeyset
rs.LockType = adLockOptimistic
    CMD.CommandText = "update stock set purprice='" & Trim(txtpurprice.Text) & "'where slno like '" & Trim(TXTSLNO.Text) & "';"
    CMD.CommandType = adCmdText
    Set CMD.ActiveConnection = cn
    CMD.Execute
    MsgBox ("The Party is Successfully Modified")
    Set CMD = Nothing
    clear2
    CMDDELETE.Enabled = False
    CMDUPDATE.Enabled = False
    CMDSAVE.Enabled = False
    CMDADD.Enabled = True
End Sub

Private Sub command1_Click()
frmstockshow.Show
End Sub

Private Sub Form_Load()
Set cn = New ADODB.Connection
cn.Provider = "Microsoft.jet.oledb.4.0"
cn.Open ("D:\Sanat\project1\stockist.mdb")
Set rs = New ADODB.Recordset
rs.CursorType = adOpenKeyset
rs.LockType = adLockOptimistic
rs.Source = "SELECT PNAME FROM PARTMASTER"
rs.ActiveConnection = cn
rs.Open
While Not rs.EOF
CMBPNAME.AddItem rs(0)
rs.MoveNext
Wend
rs.Close
Set rs = New ADODB.Recordset
rs.CursorType = adOpenKeyset
rs.LockType = adLockOptimistic
rs.Source = "SELECT sname FROM dealer"
rs.ActiveConnection = cn
rs.Open
While Not rs.EOF
cmbsname.AddItem rs(0)
rs.MoveNext
Wend
frmstock.WindowState = 2
CMDSAVE.Enabled = False
CMDDELETE.Enabled = False
CMDADD.Enabled = True
CMDFIND.Enabled = True
CMDEXIT.Enabled = True
End Sub
Public Sub assign()
rs(0) = TXTSLNO.Text
rs(2) = CMBPNAME.Text
rs(1) = txtpcode.Text
rs(3) = cmbsname.Text
rs(4) = txtsadd.Text
rs(5) = cmbdate.Value
rs(6) = txtquantity.Text
rs(7) = txtunit.Text
rs(8) = txtpurprice.Text
'RS(9) = txtstock.Text
End Sub
Public Sub clear()
TXTSLNO.Text = ""
CMBPNAME.Text = ""
txtpcode.Text = ""
cmbsname.Text = ""
txtsadd.Text = ""
'cmbdate.Value = ""
txtquantity.Text = ""
txtunit.Text = ""
txtpurprice.Text = ""
txtstock.Text = ""
End Sub
Public Sub generatesno()
Set rs = New ADODB.Recordset
rs.CursorType = adOpenKeyset
rs.LockType = adLockOptimistic
rs.Source = "select  Count(*)  from stock ;"
rs.ActiveConnection = cn
rs.Open
TXTSLNO.Text = (rs(0) + 1)
End Sub

Public Sub load()
TXTSLNO.Text = rs(0)
CMBPNAME.Text = rs(2)
txtpcode.Text = rs(1)
cmbsname.Text = rs(3)
txtsadd.Text = rs(4)
cmbdate.Value = rs(5)
txtquantity.Text = rs(6)
txtunit.Text = rs(7)
txtpurprice.Text = rs(8)
'txtstock.Text = RS(9)
End Sub

Public Sub clear2()
CMBPNAME.Text = ""
txtpcode.Text = ""
cmbsname.Text = ""
txtsadd.Text = ""
'cmbdate.Value = ""
txtquantity.Text = ""
txtunit.Text = ""
txtpurprice.Text = ""
txtstock.Text = ""
End Sub
Private Sub txtstock_GotFocus()
'Set RS = New ADODB.Recordset
'RS.CursorType = adOpenKeyset
'RS.LockType = adLockOptimistic
'RS.Source = "select stock from stock"
'RS.ActiveConnection = CN
'RS.Open
End Sub

Public Sub ASSIGN2()
rs(0) = CMBPNAME.Text
rs(1) = txtpcode.Text
rs(2) = txtquantity.Text
End Sub
