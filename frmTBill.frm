VERSION 5.00
Begin VB.Form frmTBill 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   Begin VB.TextBox txttotal 
      Height          =   375
      Left            =   2520
      TabIndex        =   12
      Text            =   " "
      Top             =   3480
      Width           =   1575
   End
   Begin VB.TextBox txtuprice 
      Height          =   375
      Left            =   2520
      TabIndex        =   11
      Text            =   " "
      Top             =   3000
      Width           =   1575
   End
   Begin VB.TextBox txtquantity 
      Height          =   375
      Left            =   2520
      TabIndex        =   10
      Text            =   " "
      Top             =   2400
      Width           =   1575
   End
   Begin VB.TextBox txtpname 
      Height          =   375
      Left            =   2520
      TabIndex        =   9
      Text            =   " "
      Top             =   1680
      Width           =   1575
   End
   Begin VB.CommandButton cmdview 
      Caption         =   "VIEW"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4440
      TabIndex        =   4
      Top             =   1920
      Width           =   1215
   End
   Begin VB.ComboBox cmbtcode 
      Height          =   315
      Left            =   2520
      TabIndex        =   3
      Text            =   " "
      Top             =   480
      Width           =   1575
   End
   Begin VB.TextBox txtcname 
      Height          =   375
      Left            =   2520
      TabIndex        =   2
      Text            =   " "
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
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
      Height          =   240
      Left            =   1440
      TabIndex        =   8
      Top             =   3480
      Width           =   555
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Unit  Price "
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
      Left            =   840
      TabIndex        =   7
      Top             =   3000
      Width           =   1140
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
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
      Height          =   240
      Left            =   1080
      TabIndex        =   6
      Top             =   2400
      Width           =   855
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Part Name"
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
      Left            =   840
      TabIndex        =   5
      Top             =   1800
      Width           =   1110
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Customer Name "
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
      Left            =   360
      TabIndex        =   1
      Top             =   1200
      Width           =   1725
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Transaction Code "
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
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   1920
   End
End
Attribute VB_Name = "frmTBill"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cn As ADODB.Connection
Dim rs As ADODB.Recordset

Private Sub cmbtcode_click()
Set rs = New ADODB.Recordset
rs.CursorType = adOpenKeyset
rs.LockType = adLockOptimistic
rs.Source = "Select  cname,pname,quantity,sprice,total from transaction2 where tcode like '" & Trim(cmbtcode.Text) & "'; "
rs.ActiveConnection = cn
rs.Open
txtcname.Text = rs(0)
txtpname.Text = rs(1)
txtquantity.Text = rs(2)
txtuprice.Text = rs(3)
txttotal.Text = rs(4)

rs.Close
End Sub

Private Sub cmdview_Click()
DataReport2.Sections("Section2").Controls("label2").Caption = cmbtcode.Text
DataReport2.Sections("Section2").Controls("label4").Caption = txtcname.Text
DataReport2.Sections("Section2").Controls("label6").Caption = txtpname.Text
DataReport2.Sections("Section2").Controls("label8").Caption = txtuprice.Text
DataReport2.Sections("Section2").Controls("label10").Caption = txtquantity.Text
DataReport2.Sections("Section2").Controls("label2").Caption = txttotal.Text

'.Sections("Section2").Controls("lblConName").Caption = txtConName.Text
'RptCashMemo.Sections("Section4").Controls("label17").Caption = Date$
'Set rs = New ADODB.Recordset
'rs.Source = "SELECT "
  'rs.Source = "Select Messer_No,Transaction_Date,Item_Name from Trans where Messer_No ='" & Trim(cmbMNo.Text) & "';"
  'If (Frame4.Enabled = True) And (CmbCSociety.Text <> "") Then
  'rs.Source = "Select * from payment_Received_Ledger;"
 ' rs.Source = "Select Transaction_Code,Item_Name,Quality,Quantity,Unit_Price,Total_Price from Transaction_Ledger where Transaction_Code like '" & Trim(cmbtcode.Text) & "'; "
  'ElseIf (Frame3.Enabled = True) And (cmbLSociety.Text <> "") Then
 ' rs.Source = "Select Loan_Ledger.Member_Code,Loan_Ledger.Member_Name,Shg_Ledger.Group_Name,Loan_Ledger.Sanction_Date,Loan_Ledger.Loan_Amount,Loan_Ledger.Loan_Purpose,Loan_Ledger.Return_Date from Loan_Ledger,Shg_Ledger where Shg_Ledger.Society like '" & Trim(cmbLSociety.Text) & "'And Shg_Ledger.Group_Name=Loan_Ledger.Group_Name ;"
  'End If
 rs.ActiveConnection = cn
  rs.Open
  Set DataReport2.DataSource = rs
 ' RptCashMemo.Show
 DataReport2.Show

End Sub

Private Sub Form_Load()
Set cn = New ADODB.Connection
cn.Provider = "Microsoft.jet.oledb.4.0"
cn.Open ("D:\Sanat\project1\stockist.mdb")
Set rs = New ADODB.Recordset
rs.CursorType = adOpenKeyset
rs.LockType = adLockOptimistic
rs.Source = "SELECT tcode from transaction2"
rs.ActiveConnection = cn
rs.Open
While Not rs.EOF
cmbtcode.AddItem rs(0)
rs.MoveNext
Wend
rs.Close
End Sub
