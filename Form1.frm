VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmsupplier 
   BackColor       =   &H00004040&
   Caption         =   "stock"
   ClientHeight    =   5415
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6675
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5415
   ScaleWidth      =   6675
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Adodc1 
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   5160
      ScaleHeight     =   675
      ScaleWidth      =   1515
      TabIndex        =   15
      Top             =   5160
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   3480
      TabIndex        =   14
      Text            =   " "
      Top             =   1920
      Width           =   3975
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Form1.frx":0000
      Height          =   1095
      Left            =   6480
      TabIndex        =   13
      Top             =   240
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   1931
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton CMDEXIT 
      Caption         =   "exit"
      Enabled         =   0   'False
      Height          =   495
      Left            =   3480
      TabIndex        =   11
      Top             =   5280
      Width           =   1215
   End
   Begin VB.CommandButton CMDFIND 
      Caption         =   "find"
      Enabled         =   0   'False
      Height          =   495
      Left            =   3480
      TabIndex        =   10
      Top             =   4680
      Width           =   1215
   End
   Begin VB.CommandButton CMDUPDATE 
      Caption         =   "update"
      Enabled         =   0   'False
      Height          =   495
      Left            =   1800
      TabIndex        =   9
      Top             =   5280
      Width           =   1215
   End
   Begin VB.CommandButton CMDDEL 
      Caption         =   "delete"
      Enabled         =   0   'False
      Height          =   495
      Left            =   120
      TabIndex        =   8
      Top             =   5280
      Width           =   1215
   End
   Begin VB.CommandButton CMDSAVE 
      Caption         =   "save"
      Enabled         =   0   'False
      Height          =   495
      Left            =   1800
      TabIndex        =   7
      Top             =   4680
      Width           =   1215
   End
   Begin VB.CommandButton CMDADD 
      Caption         =   " add"
      Enabled         =   0   'False
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   4680
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      Height          =   1215
      Left            =   3480
      TabIndex        =   5
      Text            =   " "
      Top             =   2640
      Width           =   3975
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   3480
      TabIndex        =   4
      Text            =   " "
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000080&
      Caption         =   "Frame1"
      Enabled         =   0   'False
      Height          =   1695
      Left            =   0
      TabIndex        =   12
      Top             =   4320
      Width           =   4935
   End
   Begin VB.Label Label4 
      Caption         =   " address"
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
      Left            =   1080
      TabIndex        =   3
      Top             =   2640
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "suppname"
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
      Left            =   1080
      TabIndex        =   2
      Top             =   1920
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   " suppplier no"
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
      Left            =   1080
      TabIndex        =   1
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   " stockist details"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   555
      Left            =   2760
      TabIndex        =   0
      Top             =   360
      Width           =   3525
   End
End
Attribute VB_Name = "frmsupplier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cn As ADODB.Connection
Dim rs As ADODB.Recordset
Dim sql As String
Dim CMD As ADODB.Command

Private Sub CMDADD_Click()
clear
generatesno
Text2.SetFocus
CMDSAVE.Enabled = True
CMDADD.Enabled = False
End Sub

Private Sub CMDDEL_Click()
Dim i As String
i = "Are you sure to delete?"
If MsgBox(i, vbYesNo, "warning?") = vbYes Then
Text2.Text = ""
Text3.Text = ""
 Set CMD = New ADODB.Command
    'cmd.CommandText = "delete from Party_Master where Party_Code='" & Trim(txtICode.Text) & "' "
    CMD.CommandText = "Update dealer set sname='" & Trim(Text2.Text) & "',Address='" & Trim(Text3.Text) & "'  where sno like '" & Trim(Text1.Text) & "';"
    CMD.CommandType = adCmdText
    Set CMD.ActiveConnection = cn
    CMD.Execute
    Set CMD = Nothing
MsgBox "Successfully deleted", vbExclamation, "Deletion"
clear
ElseIf MsgBox(i, vbYesNo, "Warning?") = vbNo Then
MsgBox "Do you want to exit", vbOKOnly, "Stop"
End If
Set rs = Nothing
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
rs.Source = "Select * from dealer where  sno like '" & i & "'"
rs.ActiveConnection = cn
rs.Open
If rs.EOF Then
MsgBox ("The PARTY With This Code is Not Exist")
CMDDEL.Enabled = False
Else
clear
load
End If
CMDDEL.Enabled = True
CMDUPDATE.Enabled = True
CMDADD.Enabled = True
rs.Close
End Sub

Private Sub CMDSAVE_Click()
If Text1.Text = "" Or Text2.Text = "" Then
  MsgBox "You Should Fill Code And Item NAME "
 Else
  i = MsgBox("Do You Want to Save   ", vbYesNo, "Save")
    If i = vbYes Then
       Set rs = New ADODB.Recordset
       rs.CursorType = adOpenKeyset
       rs.LockType = adLockOptimistic
       rs.Source = "dealer"
       rs.ActiveConnection = cn
       rs.Open
       rs.AddNew
       Call assign
       rs.Update
       rs.Close
       Set rs = Nothing
       Call clear
       End If
End If
Text2.SetFocus
CMDADD.Enabled = True
CMDSAVE.Enabled = False
End Sub

Private Sub CMDUPDATE_Click()
Set CMD = New ADODB.Command
Set rs = New ADODB.Recordset
rs.CursorType = adOpenKeyset
rs.LockType = adLockOptimistic
    CMD.CommandText = "update dealer set sname='" & Trim(Text2.Text) & "', Address='" & Trim(Text3.Text) & "'where sno ='" & Trim(Text1.Text) & "'"
    CMD.CommandType = adCmdText
    Set CMD.ActiveConnection = cn
    CMD.Execute
    MsgBox ("The Party is Successfully Modified")
    Set CMD = Nothing
    clear
    CMDDEL.Enabled = False
    CMDUPDATE.Enabled = False
    CMDSAVE.Enabled = False
    CMDADD.Enabled = True
    
End Sub

Private Sub DataGrid1_Click()
DisplayData
'RS.Refresh
CMDADD.Enabled = False
CMDSAVE.Enabled = True
CMDDEL.Enabled = True
CMDUPDATE.Enabled = True
End Sub

Private Sub Form_Load()
Set cn = New ADODB.Connection
cn.Provider = "Microsoft.jet.oledb.4.0"
cn.Open "D:\Sanat\project1\stockist.mdb"
Set rs = New ADODB.Recordset
rs.CursorType = adOpenKeyset
rs.LockType = adLockOptimistic
frmsupplier.WindowState = 2
generatesno
CMDSAVE.Enabled = False
CMDDEL.Enabled = False
CMDADD.Enabled = True
CMDFIND.Enabled = True
CMDEXIT.Enabled = True
End Sub

Public Sub generatesno()
Set rs = New ADODB.Recordset
rs.CursorType = adOpenKeyset
rs.LockType = adLockOptimistic
rs.Source = "select  Count(*)  from dealer ;"
rs.ActiveConnection = cn
rs.Open
Text1.Text = "s" & (rs(0) + 1)
Text1.Enabled = False
Text1.BackColor = RGB(220, 220, 220)
End Sub

Public Sub clear()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
End Sub

Public Sub assign()
rs(0) = Trim(Text1.Text)
rs(1) = Trim(Text2.Text)
rs(2) = Trim(Text3.Text)
End Sub

Public Sub load()
Text1.Text = rs(0)
Text2.Text = rs(1)
Text3.Text = rs(2)
End Sub

Public Sub DisplayData()
    Text1.Text = DataGrid1.Columns(0)
    Text2.Text = DataGrid1.Columns(1)
    Text3.Text = DataGrid1.Columns(2)
End Sub


