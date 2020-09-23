VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmpurchasereporting 
   BackColor       =   &H00EDDDD1&
   Caption         =   "  *Puchase Order Reporting * "
   ClientHeight    =   9165
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10515
   Icon            =   "frmpogrnrepots.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9165
   ScaleWidth      =   10515
   WindowState     =   2  'Maximized
   Begin VB.ComboBox cbodept 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   8760
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   240
      Width           =   3015
   End
   Begin VB.ComboBox cbopono 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1320
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   240
      Width           =   1575
   End
   Begin VB.ComboBox cbosup 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   4800
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   240
      Width           =   3855
   End
   Begin VB.TextBox txttotamt 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   12000
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   9840
      Width           =   3015
   End
   Begin VB.TextBox txttotpo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   9840
      Width           =   1695
   End
   Begin VB.CommandButton cmdlist1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "List All "
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   13560
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   " Show All PO 's "
      Top             =   120
      Width           =   1575
   End
   Begin MSFlexGridLib.MSFlexGrid ReportGrid 
      Height          =   8655
      Left            =   240
      TabIndex        =   5
      ToolTipText     =   "  Note : # Indicate Place Click The First Row of Grid To  Open the Filter Options "
      Top             =   840
      Width           =   14775
      _ExtentX        =   26061
      _ExtentY        =   15266
      _Version        =   393216
      Rows            =   1
      Cols            =   6
      FixedCols       =   0
      BackColor       =   16777215
      BackColorFixed  =   12542735
      ForeColorFixed  =   16777215
      BackColorSel    =   16777215
      ForeColorSel    =   12647934
      BackColorBkg    =   15588820
      GridColorFixed  =   4194368
      GridLines       =   2
      GridLinesFixed  =   1
      MergeCells      =   1
      AllowUserResizing=   3
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
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H000000FF&
      Height          =   8895
      Left            =   120
      Top             =   720
      Width           =   15015
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Total Purchase Amount (Rs)"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   1
      Left            =   8880
      TabIndex        =   7
      Top             =   9840
      Width           =   3135
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Total Number of PO"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   2
      Left            =   240
      TabIndex        =   6
      Top             =   9840
      Width           =   2175
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FF0000&
      Height          =   615
      Left            =   120
      Top             =   9720
      Width           =   15015
   End
End
Attribute VB_Name = "frmpurchasereporting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim op As Variant
Dim cn As ADODB.Connection
Dim rs As ADODB.Recordset
Dim ww As ADODB.Recordset
Private Sub cbodept_Click()
    Call purchasedeptloads
    Call cals
End Sub
Private Sub cbodept_dropdown()
On Error GoTo X
    cbodept.Clear
        Set rs = cn.Execute("select deptname from postatus_details group by deptname")
        rs.MoveFirst
        Do While Not rs.EOF()
        cbodept.additem (rs(0))
        rs.MoveNext
        Loop
        cbodept.SetFocus
X:
End Sub
Private Sub cbopono_Click()
    Call purchasenoloads
    Call cals
End Sub
Private Sub cbopono_dropdown()
    On Error GoTo X
    cbopono.Clear
        Set rs = cn.Execute("select pono from postatus_details order by pono")
        rs.MoveFirst
        Do While Not rs.EOF()
        cbopono.additem (rs(0))
        rs.MoveNext
        Loop
        cbopono.SetFocus
X:
End Sub
Private Sub cbosup_Click()
    Call purchasesuploads
    Call cals
End Sub

Private Sub cmdlist1_Click()
    Call purchaseloads
    Call visi
    Call cals
End Sub
Private Sub Form_Load()
    Set cn = New ADODB.Connection
    Set ww = New ADODB.Recordset
    cn.Open "Provider=Microsoft.jet.oledb.4.0; Data Source=" & App.path & "\Database\Data.mdb"
    ww.Open "Select * from postatus_details", cn, adOpenKeyset, adLockOptimistic
    
    Call frmpurchaserepots
    Call purchaseloads
    Call cals
    Call visi
End Sub
Private Function purchaseloads()
    Dim i As Integer
    Dim rs As New ADODB.Recordset
    rs.Open "Select * from postatus_details", cn, adOpenKeyset, adLockOptimistic
    i = 1
    If rs.BOF = False Then rs.MoveFirst
    While rs.EOF = False
    ReportGrid.Rows = rs.RecordCount + 1
    ReportGrid.TextMatrix(i, 0) = i
    ReportGrid.TextMatrix(i, 1) = rs![pono]
    ReportGrid.TextMatrix(i, 2) = rs![podate]
    ReportGrid.TextMatrix(i, 3) = rs![supname]
    ReportGrid.TextMatrix(i, 4) = rs![deptname]
    ReportGrid.TextMatrix(i, 5) = Format(rs![netamt], "0.00")
    rs.MoveNext
    i = i + 1
    Wend
End Function
Private Function cals()
    Dim i As Integer
    txttotamt = 0
    For i = 1 To ReportGrid.Rows - 1
    txttotamt.Text = Format(Val(txttotamt.Text) + Val(ReportGrid.TextMatrix(i, 5)), "0.00")
    Next i
    txttotpo.Text = ReportGrid.Rows - 1
End Function
Private Function purchasenoloads()
    Dim i As Integer
    Dim rs As New ADODB.Recordset
    rs.Open "Select * from postatus_details where pono='" & cbopono.Text & "'", cn, adOpenKeyset, adLockOptimistic
    i = 1
    If rs.BOF = False Then rs.MoveFirst
    While rs.EOF = False
    ReportGrid.Rows = ReportGrid.Rows + 1
    ReportGrid.TextMatrix(i, 0) = i
    ReportGrid.TextMatrix(i, 1) = rs![pono]
    ReportGrid.TextMatrix(i, 2) = rs![podate]
    ReportGrid.TextMatrix(i, 3) = rs![supname]
    ReportGrid.TextMatrix(i, 4) = rs![deptname]
    ReportGrid.TextMatrix(i, 5) = Format(rs![netamt], "0.00")
    rs.MoveNext
    i = i + 1
    Wend
    ReportGrid.Rows = rs.RecordCount + 1
End Function
Sub visi()
    cbosup.Visible = False
    cbopono.Visible = False
    cbodept.Visible = False
    cbosup.Clear
    cbopono.Clear
    cbodept.Clear
End Sub
Private Sub ReportGrid_Click()
    If ReportGrid.Col = 1 And ReportGrid.Row = 1 Then
        cbopono.Visible = True
    Else
        cbopono.Visible = False
    End If
    If ReportGrid.Col = 3 And ReportGrid.Row = 1 Then
        cbosup.Visible = True
    Else
        cbosup.Visible = False
    End If
    If ReportGrid.Col = 4 And ReportGrid.Row = 1 Then
        cbodept.Visible = True
    Else
        cbodept.Visible = False
    End If
End Sub
Private Sub cbosup_dropdown()
    On Error GoTo X
    cbosup.Clear
        Set rs = cn.Execute("select supname from postatus_details Group by supname")
        rs.MoveFirst
        Do While Not rs.EOF()
        cbosup.additem (rs(0))
        rs.MoveNext
        Loop
        cbosup.SetFocus
X:
End Sub
Private Function purchasesuploads()
    Dim i As Integer
    Dim rs As New ADODB.Recordset
    rs.Open "Select * from postatus_details where supname='" & cbosup.Text & "'", cn, adOpenKeyset, adLockOptimistic
    i = 1
    If rs.BOF = False Then rs.MoveFirst
    While rs.EOF = False
    ReportGrid.Rows = ReportGrid.Rows + 1
    ReportGrid.TextMatrix(i, 0) = i
    ReportGrid.TextMatrix(i, 1) = rs![pono]
    ReportGrid.TextMatrix(i, 2) = rs![podate]
    ReportGrid.TextMatrix(i, 3) = rs![supname]
    ReportGrid.TextMatrix(i, 4) = rs![deptname]
    ReportGrid.TextMatrix(i, 5) = Format(rs![netamt], "0.00")
    rs.MoveNext
    i = i + 1
    Wend
    ReportGrid.Rows = rs.RecordCount + 1
End Function
Private Function purchasedeptloads()
    Dim i As Integer
    Dim rs As New ADODB.Recordset
    rs.Open "Select * from postatus_details where deptname='" & cbodept.Text & "'", cn, adOpenKeyset, adLockOptimistic
    i = 1
    If rs.BOF = False Then rs.MoveFirst
    While rs.EOF = False
    ReportGrid.Rows = ReportGrid.Rows + 1
    ReportGrid.TextMatrix(i, 0) = i
    ReportGrid.TextMatrix(i, 1) = rs![pono]
    ReportGrid.TextMatrix(i, 2) = rs![podate]
    ReportGrid.TextMatrix(i, 3) = rs![supname]
    ReportGrid.TextMatrix(i, 4) = rs![deptname]
    ReportGrid.TextMatrix(i, 5) = Format(rs![netamt], "0.00")
    rs.MoveNext
    i = i + 1
    Wend
    ReportGrid.Rows = rs.RecordCount + 1
End Function
