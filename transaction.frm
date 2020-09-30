VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmtransaction 
   BackColor       =   &H00800000&
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   Begin VB.CommandButton cmdstock 
      Caption         =   "See The Stock"
      Height          =   495
      Left            =   6000
      TabIndex        =   28
      Top             =   6480
      Width           =   2175
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H000000FF&
      Caption         =   "Frame1"
      Height          =   1215
      Left            =   600
      TabIndex        =   21
      Top             =   6480
      Width           =   4695
      Begin VB.CommandButton cmdexit 
         Caption         =   "EXIT"
         Height          =   375
         Left            =   3120
         TabIndex        =   27
         Top             =   720
         Width           =   1215
      End
      Begin VB.CommandButton cmdfind 
         Caption         =   "FIND"
         Height          =   375
         Left            =   3120
         TabIndex        =   26
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdupdate 
         Caption         =   "UPDATE"
         Height          =   375
         Left            =   1680
         TabIndex        =   25
         Top             =   720
         Width           =   1215
      End
      Begin VB.CommandButton cmddelete 
         Caption         =   "DELETE"
         Height          =   375
         Left            =   1680
         TabIndex        =   24
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdsave 
         Caption         =   "SAVE"
         Height          =   375
         Left            =   240
         TabIndex        =   23
         Top             =   720
         Width           =   1215
      End
      Begin VB.CommandButton cmdadd 
         Caption         =   "ADD"
         Height          =   375
         Left            =   240
         TabIndex        =   22
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.TextBox txtstock 
      Height          =   375
      Left            =   3720
      TabIndex        =   20
      Top             =   4080
      Width           =   1215
   End
   Begin VB.ComboBox cmbpname 
      Height          =   315
      Left            =   3960
      TabIndex        =   19
      Top             =   2160
      Width           =   2535
   End
   Begin VB.TextBox txttcode 
      Height          =   405
      Left            =   5400
      TabIndex        =   18
      Top             =   960
      Width           =   975
   End
   Begin VB.TextBox txtunit 
      Height          =   375
      Left            =   3720
      TabIndex        =   17
      Top             =   3360
      Width           =   1215
   End
   Begin VB.TextBox txtquantity 
      Height          =   375
      Left            =   3720
      TabIndex        =   16
      Top             =   2760
      Width           =   1215
   End
   Begin VB.TextBox txttotal 
      Height          =   375
      Left            =   8880
      TabIndex        =   15
      Top             =   3480
      Width           =   1215
   End
   Begin VB.TextBox txtsprice 
      Height          =   375
      Left            =   8880
      TabIndex        =   14
      Top             =   2880
      Width           =   1215
   End
   Begin VB.TextBox txtpcode 
      Height          =   375
      Left            =   8880
      TabIndex        =   13
      Top             =   2160
      Width           =   1215
   End
   Begin VB.TextBox txtcname 
      Height          =   375
      Left            =   5400
      TabIndex        =   12
      Top             =   1560
      Width           =   4455
   End
   Begin MSComCtl2.DTPicker cmbdate 
      Height          =   375
      Left            =   8160
      TabIndex        =   11
      Top             =   960
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      _Version        =   393216
      Format          =   24576001
      CurrentDate     =   37902
   End
   Begin VB.Label lblstock 
      AutoSize        =   -1  'True
      BackColor       =   &H00808080&
      Caption         =   "STOCK :-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   1800
      TabIndex        =   10
      Top             =   4080
      Width           =   1335
   End
   Begin VB.Label lblsprice 
      AutoSize        =   -1  'True
      BackColor       =   &H00808080&
      Caption         =   "SELL PRICE :-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   6720
      TabIndex        =   9
      Top             =   2880
      Width           =   2025
   End
   Begin VB.Label lbltotal 
      AutoSize        =   -1  'True
      BackColor       =   &H00808080&
      Caption         =   "TOTAL :-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   6720
      TabIndex        =   8
      Top             =   3480
      Width           =   1305
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "TRANSACTION OF THE SHOOP"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   555
      Left            =   2400
      TabIndex        =   7
      Top             =   120
      Width           =   7605
   End
   Begin VB.Label lbltcode 
      AutoSize        =   -1  'True
      BackColor       =   &H00808080&
      Caption         =   "TRANSACTION CODE :-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   1800
      TabIndex        =   6
      Top             =   960
      Width           =   3450
   End
   Begin VB.Label lbldate 
      AutoSize        =   -1  'True
      BackColor       =   &H00808080&
      Caption         =   "DATE :-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   6720
      TabIndex        =   5
      Top             =   960
      Width           =   1125
   End
   Begin VB.Label lblcname 
      AutoSize        =   -1  'True
      BackColor       =   &H00808080&
      Caption         =   "CUSTOMER NAME :-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   1800
      TabIndex        =   4
      Top             =   1560
      Width           =   3015
   End
   Begin VB.Label lblpcode 
      AutoSize        =   -1  'True
      BackColor       =   &H00808080&
      Caption         =   "PART CODE :-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   6720
      TabIndex        =   3
      Top             =   2160
      Width           =   2070
   End
   Begin VB.Label lblpname 
      AutoSize        =   -1  'True
      BackColor       =   &H00808080&
      Caption         =   "PARTNAME :-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   1800
      TabIndex        =   2
      Top             =   2160
      Width           =   2010
   End
   Begin VB.Label lblquantity 
      AutoSize        =   -1  'True
      BackColor       =   &H00808080&
      Caption         =   "QUANTITY :-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   1800
      TabIndex        =   1
      Top             =   2760
      Width           =   1845
   End
   Begin VB.Label lblunit 
      AutoSize        =   -1  'True
      BackColor       =   &H00808080&
      Caption         =   "UNIT :-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   1800
      TabIndex        =   0
      Top             =   3360
      Width           =   1005
   End
End
Attribute VB_Name = "frmtransaction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As ADODB.Recordset
Dim cn As ADODB.Connection
Dim CMD As ADODB.Command
Dim sql As String
Private Sub CMBPNAME_click()
 Set rs = New ADODB.Recordset
 rs.CursorType = adOpenKeyset
 rs.LockType = adLockOptimistic
 rs.Source = "select pcode,unit,sprice,stockist.stock from partmaster,stockist where pname like '" & Trim(CMBPNAME.Text) & "';"
 rs.ActiveConnection = cn
 rs.Open
 txtpcode.Text = rs(0)
 txtunit.Text = rs(1)
 txtsprice = rs(2)
 rs.Close
 txtpcode.Enabled = False
 txtstock.Enabled = False
End Sub

Private Sub CMDADD_Click()
clear
CMDSAVE.Enabled = True
txtcname.SetFocus
generatetcode
txttcode.Enabled = False
CMBPNAME.Refresh
End Sub

Private Sub cmddelete_Click()
Dim i As String
i = MsgBox("Do You Want to Delete   ", vbYesNo, "Save")
    If i = vbYes Then
    Set rs = New ADODB.Recordset
    rs.CursorType = adOpenKeyset
    rs.LockType = adLockOptimistic
    rs.Source = "select Count(tcode) from transaction2 where tcode like '" & Trim(txttcode.Text) & "';"
    rs.ActiveConnection = cn
    rs.Open
    If (rs(0) > 1) Then
    Set CMD = New ADODB.Command
    CMD.CommandText = "delete from transaction2 where tcode like '" & Trim(txttcode.Text) & "';"
    CMD.CommandType = adCmdText
     Set CMD.ActiveConnection = cn
     CMD.Execute
     Set CMD = Nothing
     Else
    Set CMD = New ADODB.Command
    'CMD.CommandText = "Update partmaster set pname='" & Trim(TXTPNAME.Text) & "',pcatagory='" & Trim(TXTPCATAGORY.Text) & "',uprice='" & Trim(TXTUNITPRICE.Text) & "',sprice='" & Trim(TXTSELLPRICE.Text) & "',unit='" & Trim(txtunit.Text) & "';"
    CMD.CommandType = adCmdText
    Set CMD.ActiveConnection = cn
    CMD.Execute
    Set CMD = Nothing
    Set CMD = New ADODB.Command
    CMD.CommandText = "Delete from transaction where tcode like '" & Trim(txttcode.Text) & "';"
    CMD.CommandType = adCmdText
    Set CMD.ActiveConnection = cn
    CMD.Execute
    Set CMD = Nothing
    MsgBox ("The Record is successfully deleted")
    'Else
    
    'End If
    End If
    clear
    'cmbSoName.SetFocus
    CMDDELETE.Enabled = True
    CMDUPDATE.Enabled = False
    'End If
    End If
End Sub

Private Sub CMDEXIT_Click()
End
End Sub

Private Sub CMDFIND_Click()
Dim i As String
i = InputBox("Enter The Party Code U want to find:")
clear
Set rs = New ADODB.Recordset
rs.CursorType = adOpenKeyset
rs.LockType = adLockOptimistic
rs.Source = "Select * from transaction2 where  tcode like '" & i & "';"
rs.ActiveConnection = cn
rs.Open
If rs.EOF Then
MsgBox ("TRANSACTION WITH THIS CODE DOSENOT EXIST")
CMDDELETE.Enabled = False
cmdModify.Enabled = False
Else
clear
load
End If
rs.Close
End Sub

Private Sub CMDSAVE_Click()
Dim i As String
' If TXTSLNO.Text = "" Or cmbpname.Text = "" Then
  'MsgBox "You Should Fill SLNO And PNAME "
 'Else
  i = MsgBox("Do You Want to Save   ", vbYesNo, "Save")
    If i = vbYes Then
       Set rs = New ADODB.Recordset
       rs.CursorType = adOpenKeyset
       rs.LockType = adLockOptimistic
       rs.Source = "TRANSACTION2"
       rs.ActiveConnection = cn
       rs.Open
       rs.AddNew
       assign
       rs.Update
       rs.Close
       MsgBox "SUCCESSFULLY SAVED"
    End If
    Set CMD = New ADODB.Command
    CMD.CommandText = "Update STOCKIST set STOCK=" & Val(txtstock.Text) - Val(txtquantity.Text) & " where pname like '" & Trim(CMBPNAME.Text) & "'; "
    CMD.CommandType = adCmdText
    Set CMD.ActiveConnection = cn
    CMD.Execute
    Set CMD = Nothing
    clear
    Set CMD = New ADODB.Command
    CMD.CommandText = "delete from stockist where stock=0; "
    CMD.CommandType = adCmdText
    Set CMD.ActiveConnection = cn
    CMD.Execute
    Set CMD = Nothing
  
CMBPNAME.SetFocus
CMDADD.Enabled = True
CMDSAVE.Enabled = False
End Sub

Private Sub cmdstock_Click()
frmstockshow.Show
End Sub

Private Sub Form_Load()
frmtransaction.WindowState = 2
Set cn = New ADODB.Connection
cn.Provider = "Microsoft.jet.oledb.4.0"
cn.Open ("D:\sanat\project1\stockist.mdb")
Set rs = New ADODB.Recordset
rs.CursorType = adOpenKeyset
rs.LockType = adLockOptimistic
rs.Source = "SELECT PNAME From stockist"
rs.ActiveConnection = cn
rs.Open
While Not rs.EOF
CMBPNAME.AddItem rs(0)
rs.MoveNext
Wend
rs.Close
CMDSAVE.Enabled = False
CMDDELETE.Enabled = False
CMDADD.Enabled = True
CMDFIND.Enabled = True
CMDEXIT.Enabled = True
End Sub
Public Sub ASSIGN2()
rs(0) = txttcode.Text
rs(1) = cmbdate.Value
rs(2) = txtcname.Text
End Sub


Public Sub generatetcode()
Set rs = New ADODB.Recordset
rs.CursorType = adOpenKeyset
rs.LockType = adLockOptimistic
rs.Source = "select  count(*)  from transaction2 ;"
rs.ActiveConnection = cn
rs.Open
txttcode.Text = (rs(0) + 1)
End Sub

Public Sub clear()
txtcname.Text = ""
CMBPNAME.Text = ""
txtpcode.Text = ""
txtquantity.Text = ""
txtunit.Text = ""
txttotal.Text = ""
txtstock.Text = ""
txtsprice.Text = ""


End Sub

Private Sub txttotal_GotFocus()
txttotal.Text = Val(txtquantity.Text) * Val(txtsprice.Text)
End Sub

Public Sub assign()
rs(0) = Trim(txttcode.Text)
rs(1) = cmbdate.Value
rs(2) = Trim(txtcname.Text)
rs(3) = CMBPNAME.Text
rs(4) = Trim(txtpcode.Text)
rs(5) = Trim(txtquantity.Text)
rs(6) = Trim(txtsprice.Text)
rs(7) = Trim(txtunit.Text)
rs(8) = Trim(txttotal.Text)
End Sub

Public Sub load()
txttcode.Text = rs(0)
cmbdate.Value = rs(1)
txtcname.Text = rs(2)
CMBPNAME.Text = rs(3)
txtpcode.Text = rs(4)
txtquantity.Text = rs(5)
txtsprice.Text = rs(6)
txtunit.Text = rs(7)
txttotal.Text = rs(8)
'txtstock.Text = ""
'txtsprice.Text = ""

End Sub
