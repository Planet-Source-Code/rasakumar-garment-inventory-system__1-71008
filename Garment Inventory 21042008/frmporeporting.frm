VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmporeporting 
   BackColor       =   &H00EDDDD1&
   Caption         =   " * Purchase Order Reporting *"
   ClientHeight    =   8505
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10035
   Icon            =   "frmporeporting.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   11115
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
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
      TabIndex        =   9
      ToolTipText     =   " Show All PO 's "
      Top             =   120
      Width           =   1575
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
      TabIndex        =   7
      Top             =   9840
      Width           =   1695
   End
   Begin VB.TextBox txttotgrnqty 
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
      Left            =   12480
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   9840
      Width           =   2535
   End
   Begin VB.TextBox txttotpoqty 
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
      Left            =   7680
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   9840
      Width           =   2415
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
      TabIndex        =   2
      ToolTipText     =   " Filter The Supplier Wise "
      Top             =   240
      Width           =   3855
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
      TabIndex        =   1
      ToolTipText     =   " Filter The PO No Wise "
      Top             =   240
      Width           =   1575
   End
   Begin MSFlexGridLib.MSFlexGrid ReportGrid 
      Height          =   8655
      Left            =   240
      TabIndex        =   0
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
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FF0000&
      Height          =   615
      Left            =   120
      Top             =   9720
      Width           =   15015
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
      TabIndex        =   8
      Top             =   9840
      Width           =   2175
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Total Rec. Qty"
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
      Left            =   10200
      TabIndex        =   6
      Top             =   9840
      Width           =   2295
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Total PO.Qty"
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
      Index           =   0
      Left            =   5400
      TabIndex        =   5
      Top             =   9840
      Width           =   2295
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H000000FF&
      Height          =   8895
      Left            =   120
      Top             =   720
      Width           =   15015
   End
End
Attribute VB_Name = "frmporeporting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim op As Variant
Dim cn As ADODB.Connection
Dim rs As ADODB.Recordset
Dim ww As ADODB.Recordset

Private Sub cbopono_Click()
    Call ponofilter
    Call grnnofilter
    Call cals
End Sub
Private Sub cbopono_dropdown()
    On Error GoTo X
    cbopono.Clear
        Set rs = cn.Execute("select pono from poreport_details order by pono")
        rs.MoveFirst
        Do While Not rs.EOF()
        cbopono.additem (rs(0))
        rs.MoveNext
        Loop
        cbopono.SetFocus
X:
End Sub
Private Sub cbosup_Click()
    Call posupfilter
    Call grnsupfilter
    Call cals
End Sub

Private Sub cbosup_dropdown()
    On Error GoTo X
    cbosup.Clear
        Set rs = cn.Execute("select supname from poreport_details Group by supname")
        rs.MoveFirst
        Do While Not rs.EOF()
        cbosup.additem (rs(0))
        rs.MoveNext
        Loop
        cbosup.SetFocus
X:
End Sub
Private Sub cmdlist1_Click()
    Call visi
    Unload Me
    frmporeporting.Show
End Sub
Private Sub Form_Load()
    Set cn = New ADODB.Connection
    Set ww = New ADODB.Recordset
    cn.Open "Provider=Microsoft.jet.oledb.4.0; Data Source=" & App.path & "\Database\Data.mdb"
    ww.Open "Select * from poreport_details", cn, adOpenKeyset, adLockOptimistic
    
    Call frmporeportingsload
    Call pogridloads
    Call pogridloads1
    Call cals
    Call visi
End Sub
Private Function pogridloads()
    Dim i As Integer
    Dim rs As New ADODB.Recordset
    rs.Open "Select * from poreport_details", cn, adOpenKeyset, adLockOptimistic
    i = 1
    If rs.BOF = False Then rs.MoveFirst
    While rs.EOF = False
    ReportGrid.Rows = ReportGrid.Rows + 1
    ReportGrid.TextMatrix(i, 0) = i
    ReportGrid.TextMatrix(i, 1) = rs![pono]
    ReportGrid.TextMatrix(i, 2) = rs![podate]
    ReportGrid.TextMatrix(i, 3) = rs![supname]
    ReportGrid.TextMatrix(i, 4) = Format(rs![SumOfqty], "0.000")
    rs.MoveNext
    i = i + 1
    Wend
    End Function
Private Function pogridloads1()
    Dim i As Integer
    Dim rs As New ADODB.Recordset
    rs.Open "Select * from grnreport_details", cn, adOpenKeyset, adLockOptimistic
    i = 1
    If rs.BOF = False Then rs.MoveFirst
    While rs.EOF = False
    ReportGrid.TextMatrix(i, 5) = Format(rs![SumOfrecqty], "0.000")
    rs.MoveNext
    i = i + 1
    Wend
End Function
Private Function cals()
    Dim i As Integer
    txttotpoqty = 0
    txttotgrnqty = 0
    For i = 1 To ReportGrid.Rows - 1
    txttotpoqty.Text = Format(Val(txttotpoqty.Text) + Val(ReportGrid.TextMatrix(i, 4)), "0.000")
    txttotgrnqty.Text = Format(Val(txttotgrnqty.Text) + Val(ReportGrid.TextMatrix(i, 5)), "0.000")
    Next i
    txttotpo.Text = ReportGrid.Rows - 1
End Function
Private Function ponofilter()
    Dim rs As New ADODB.Recordset
    rs.Open "select * from poreport_details where pono= " & Trim(cbopono.Text), cn, adOpenKeyset, adLockOptimistic
    i = 1
    If rs.BOF = False Then rs.MoveFirst
    While rs.EOF = False
    ReportGrid.Rows = ReportGrid.Rows + 1
    ReportGrid.TextMatrix(i, 0) = i
    ReportGrid.TextMatrix(i, 1) = rs![pono]
    ReportGrid.TextMatrix(i, 2) = rs![podate]
    ReportGrid.TextMatrix(i, 3) = rs![supname]
    ReportGrid.TextMatrix(i, 4) = Format(rs![SumOfqty], "0.000")
    rs.MoveNext
    i = i + 1
    Wend
    ReportGrid.Rows = rs.RecordCount + 1
End Function
Private Function grnnofilter()
    Dim rs As New ADODB.Recordset
    rs.Open "select * from grnreport_details where pono= '" & Trim(cbopono.Text) & "'", cn, adOpenKeyset, adLockOptimistic
    i = 1
    If rs.BOF = False Then rs.MoveFirst
    While rs.EOF = False
    ReportGrid.Rows = ReportGrid.Rows + 1
    ReportGrid.TextMatrix(i, 5) = Format(rs![SumOfrecqty], "0.000")
    rs.MoveNext
    i = i + 1
    Wend
    ReportGrid.Rows = rs.RecordCount + 1
End Function
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
End Sub
Sub visi()
    cbosup.Visible = False
    cbopono.Visible = False
    cbosup.Clear
    cbopono.Clear
End Sub
Private Function posupfilter()
    Dim rs As New ADODB.Recordset
    rs.Open "select * from poreport_details where supname= '" & Trim(cbosup.Text) & "'", cn, adOpenKeyset, adLockOptimistic
    i = 1
    If rs.BOF = False Then rs.MoveFirst
    While rs.EOF = False
    ReportGrid.Rows = ReportGrid.Rows + 1
    ReportGrid.TextMatrix(i, 0) = i
    ReportGrid.TextMatrix(i, 1) = rs![pono]
    ReportGrid.TextMatrix(i, 2) = rs![podate]
    ReportGrid.TextMatrix(i, 3) = rs![supname]
    ReportGrid.TextMatrix(i, 4) = Format(rs![SumOfqty], "0.000")
    rs.MoveNext
    i = i + 1
    Wend
    ReportGrid.Rows = rs.RecordCount + 1
End Function
Private Function grnsupfilter()
    Dim rs As New ADODB.Recordset
    rs.Open "select * from grnreport_details where supname= '" & Trim(cbosup.Text) & "'", cn, adOpenKeyset, adLockOptimistic
    i = 1
    If rs.BOF = False Then rs.MoveFirst
    While rs.EOF = False
    ReportGrid.Rows = ReportGrid.Rows + 1
    ReportGrid.TextMatrix(i, 5) = Format(rs![SumOfrecqty], "0.000")
    rs.MoveNext
    i = i + 1
    Wend
    ReportGrid.Rows = rs.RecordCount + 1
End Function
