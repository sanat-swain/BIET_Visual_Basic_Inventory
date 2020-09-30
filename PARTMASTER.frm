VERSION 5.00
Begin VB.Form frmpmaster 
   BackColor       =   &H00400040&
   Caption         =   "Form3"
   ClientHeight    =   6795
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9480
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   ScaleHeight     =   6795
   ScaleWidth      =   9480
   Begin VB.VScrollBar VScroll1 
      Height          =   7095
      Left            =   11160
      TabIndex        =   21
      Top             =   1440
      Width           =   255
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   735
      Left            =   360
      TabIndex        =   20
      Top             =   7440
      Width           =   2175
   End
   Begin VB.CommandButton cmdview 
      Caption         =   "VIEW"
      Height          =   495
      Left            =   9720
      TabIndex        =   19
      Top             =   6720
      Width           =   1215
   End
   Begin VB.CommandButton cmdexit 
      Caption         =   "EXIT"
      Height          =   495
      Left            =   8280
      TabIndex        =   18
      Top             =   6720
      Width           =   1215
   End
   Begin VB.CommandButton cmdfind 
      Caption         =   "FIND"
      Height          =   495
      Left            =   6840
      TabIndex        =   17
      Top             =   6720
      Width           =   1215
   End
   Begin VB.CommandButton cmddelete 
      Caption         =   "DELETE"
      Height          =   495
      Left            =   5400
      TabIndex        =   16
      Top             =   6720
      Width           =   1215
   End
   Begin VB.CommandButton cmdupdate 
      Caption         =   "UPDATE"
      Height          =   495
      Left            =   3960
      TabIndex        =   15
      Top             =   6720
      Width           =   1215
   End
   Begin VB.CommandButton cmdsave 
      Caption         =   "SAVE"
      Height          =   495
      Left            =   2520
      TabIndex        =   14
      Top             =   6720
      Width           =   1215
   End
   Begin VB.CommandButton cmdadd 
      Caption         =   "ADD"
      Height          =   495
      Left            =   1080
      TabIndex        =   13
      Top             =   6720
      Width           =   1215
   End
   Begin VB.TextBox TXTUNIT 
      Height          =   495
      Left            =   4320
      TabIndex        =   12
      Top             =   5880
      Width           =   1215
   End
   Begin VB.TextBox TXTSELLPRICE 
      Height          =   495
      Left            =   4320
      TabIndex        =   11
      Top             =   4920
      Width           =   1215
   End
   Begin VB.TextBox TXTUNITPRICE 
      Height          =   495
      Left            =   4320
      TabIndex        =   10
      Top             =   4080
      Width           =   1215
   End
   Begin VB.TextBox TXTPCATAGORY 
      Height          =   495
      Left            =   4320
      TabIndex        =   9
      Top             =   3240
      Width           =   1215
   End
   Begin VB.TextBox TXTPNAME 
      Height          =   495
      Left            =   4320
      TabIndex        =   8
      Top             =   2400
      Width           =   1215
   End
   Begin VB.TextBox TXTPCODE 
      Height          =   495
      Left            =   4320
      TabIndex        =   7
      Top             =   1560
      Width           =   2535
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF00FF&
      Caption         =   "PART MASTER"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   4200
      TabIndex        =   6
      Top             =   360
      Width           =   3060
   End
   Begin VB.Label LBLUNIT 
      AutoSize        =   -1  'True
      Caption         =   "UNIT"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   720
      TabIndex        =   5
      Top             =   5880
      Width           =   840
   End
   Begin VB.Label LBLSELLPRICE 
      AutoSize        =   -1  'True
      Caption         =   "SELL PRICE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   720
      TabIndex        =   4
      Top             =   5040
      Width           =   2130
   End
   Begin VB.Label LBLUNITPRICE 
      AutoSize        =   -1  'True
      Caption         =   "UNIT PRICE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   720
      TabIndex        =   3
      Top             =   4200
      Width           =   2040
   End
   Begin VB.Label LBLPCATAGORY 
      AutoSize        =   -1  'True
      Caption         =   "PART CATEGORY"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   720
      TabIndex        =   2
      Top             =   3360
      Width           =   3090
   End
   Begin VB.Label LBLPNAME 
      AutoSize        =   -1  'True
      Caption         =   "PART NAME"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   720
      TabIndex        =   1
      Top             =   2520
      Width           =   2145
   End
   Begin VB.Label LBLPCODE 
      AutoSize        =   -1  'True
      Caption         =   "PART CODE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   720
      TabIndex        =   0
      Top             =   1680
      Width           =   2100
   End
End
Attribute VB_Name = "frmpmaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As ADODB.Recordset
Dim cn As ADODB.Connection
Dim CMD As ADODB.Command
Dim sql As String

Private Sub CMDADD_Click()
generatepartycode
clear
txtpname.SetFocus
CMDSAVE.Enabled = True
CMDADD.Enabled = False
End Sub

Private Sub cmddelete_Click()
Dim i As String
i = MsgBox("Do You Want to Delete   ", vbYesNo, "Save")
    If i = vbYes Then
    Set rs = New ADODB.Recordset
    rs.CursorType = adOpenKeyset
    rs.LockType = adLockOptimistic
    rs.Source = "select Count(pcode) from partmaster where pcode like '" & Trim(txtpcode.Text) & "';"
    rs.ActiveConnection = cn
    rs.Open
    If (rs(0) > 1) Then
    Set CMD = New ADODB.Command
    CMD.CommandText = "update partmaster set txtpname.text="", txtpcatagory.text="" ,txtunitprice.text=,TXTSELLPRICE.Text="",txtunit.Text="" where pcode like '" & Trim(txtpcode.Text) & "';"
    'CMD.CommandText = "delete from partmaster where pcode like '" & Trim(txtpcode.Text) & "';"
    CMD.CommandType = adCmdText
     Set CMD.ActiveConnection = cn
     CMD.Execute
     Set CMD = Nothing
     Else
    Set CMD = New ADODB.Command
    CMD.CommandText = "Update partmaster set pname='" & Trim(txtpname.Text) & "',pcatagory='" & Trim(TXTPCATAGORY.Text) & "',uprice='" & Trim(TXTUNITPRICE.Text) & "',sprice='" & Trim(TXTSELLPRICE.Text) & "',unit='" & Trim(txtunit.Text) & "';"
    CMD.CommandType = adCmdText
    Set CMD.ActiveConnection = cn
    CMD.Execute
    Set CMD = Nothing
    Set CMD = New ADODB.Command
    CMD.CommandText = "Delete from partmaster where pcode like '" & Trim(txtpcode.Text) & "';"
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
    CMDDELETE.Enabled = False
    CMDUPDATE.Enabled = False
    'End If
    End If
End Sub

Private Sub CMDEXIT_Click()
Unload Me
End Sub

Private Sub CMDFIND_Click()
Dim i As String
i = InputBox("Enter The sno U want to find:")
clear
Set rs = New ADODB.Recordset
rs.CursorType = adOpenKeyset
rs.LockType = adLockOptimistic
rs.Source = "Select * from partmaster where  pcode like '" & i & "'"
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
CMDSAVE.Enabled = False
CMDADD.Enabled = True
rs.Close
End Sub

Private Sub CMDSAVE_Click()
If txtpcode.Text = "" Or txtpname.Text = "" Or TXTPCATAGORY.Text = "" Or TXTUNITPRICE.Text = "" Then
  MsgBox "You Should Fill all the data in fields "
 Else
  i = MsgBox("Do You Want to Save   ", vbYesNo, "Save")
    If i = vbYes Then
       Set rs = New ADODB.Recordset
       rs.CursorType = adOpenKeyset
       rs.LockType = adLockOptimistic
       rs.Source = "partmaster"
       rs.ActiveConnection = cn
       rs.Open
       rs.AddNew
       Call assign
       rs.Update
       rs.Close
       Set rs = Nothing
       Call clear
       'Else
       'MsgBox "sss"
       End If
End If
'End If
txtpname.SetFocus
'Exit Sub
CMDADD.Enabled = True
CMDSAVE.Enabled = False
End Sub

Private Sub CMDUPDATE_Click()
Set rs = New ADODB.Recordset
rs.CursorType = adOpenKeyset
rs.LockType = adLockOptimistic
    Set CMD = New ADODB.Command
    CMD.CommandText = "update partmaster set pname='" & Trim(txtpname.Text) & "', pcatagory='" & Trim(TXTPCATAGORY.Text) & "',uprice='" & Trim(TXTUNITPRICE.Text) & "',sprice='" & Trim(TXTSELLPRICE.Text) & "',unit='" & Trim(txtunit.Text) & "'where pcode like '" & Trim(txtpcode.Text) & "';"
    CMD.CommandType = adCmdText
    Set CMD.ActiveConnection = cn
    CMD.Execute
    MsgBox ("The Party is Successfully Modified")
    Set CMD = Nothing
    clear
    CMDDELETE.Enabled = False
    CMDUPDATE.Enabled = False
    CMDSAVE.Enabled = False
    CMDADD.Enabled = True
End Sub

Private Sub Form_Load()
Set cn = New ADODB.Connection
cn.Provider = "Microsoft.jet.oledb.4.0"
cn.Open "d:\sanat\project1\stockist.mdb"
Set rs = New ADODB.Recordset
rs.CursorType = adOpenKeyset
rs.LockType = adLockOptimistic
'rs.Source = "Select distinct Party_Name from Master_Ledger"
'rs.ActiveConnection = cn
'rs.Open
'While Not rs.EOF
'cmbIName.AddParty rs(0)
'rs.MoveNext
'Wend
'rs.Close
frmpmaster.WindowState = 2
CMDUPDATE.Enabled = False
CMDSAVE.Enabled = False
CMDDELETE.Enabled = False
CMDADD.Enabled = True
End Sub

Public Sub generatepartycode()
Set rs = New ADODB.Recordset
rs.CursorType = adOpenKeyset
rs.LockType = adLockOptimistic
'rs.Source = "Select Society_Code from Society_Ledger"
rs.Source = "select  Count(*)  from PARTMASTER ;"
'Transaction_Ledger where Messer_No having (select distinct(Messer_No) from Transaction_Ledger)  ;"
rs.ActiveConnection = cn
rs.Open
'While Not rs.EOF
'MsgBox (rs(0, 1))
'rs (0) + 1
'Wend
txtpcode.Text = "P" & (rs(0) + 1)
txtpcode.Enabled = False
txtpcode.BackColor = RGB(220, 220, 220)
End Sub

Public Sub clear()
txtpname = ""
TXTPCATAGORY = ""
TXTUNITPRICE = ""
TXTSELLPRICE = ""
txtunit = ""
End Sub

Public Sub assign()
rs(0) = txtpcode.Text
rs(1) = txtpname.Text
rs(2) = TXTPCATAGORY.Text
rs(3) = TXTUNITPRICE.Text
rs(4) = TXTSELLPRICE.Text
rs(5) = txtunit.Text

End Sub

Public Sub load()
txtpcode.Text = rs(0)
txtpname.Text = rs(1)
TXTPCATAGORY.Text = rs(2)
TXTUNITPRICE.Text = rs(3)
TXTSELLPRICE.Text = rs(4)
txtunit.Text = rs(5)
End Sub
