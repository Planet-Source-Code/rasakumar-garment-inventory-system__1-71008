VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmpomain 
   BackColor       =   &H00EDDDD1&
   Caption         =   " * Puchase Order's *"
   ClientHeight    =   9810
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9885
   Icon            =   "frmpomain.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   11115
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdclosure 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Closure PO's"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   13320
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   " To Use Close The PO's "
      Top             =   9600
      Width           =   1695
   End
   Begin VB.CommandButton cmdprint 
      BackColor       =   &H00FFC0C0&
      Caption         =   "&Print PO"
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
      Left            =   7800
      MaskColor       =   &H00C0FFC0&
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "  To use Print PO  "
      Top             =   9600
      UseMaskColor    =   -1  'True
      Width           =   1575
   End
   Begin VB.CommandButton cmddelete 
      BackColor       =   &H00FFC0C0&
      Caption         =   "&Delete PO"
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
      Left            =   6240
      MaskColor       =   &H00C0FFC0&
      Style           =   1  'Graphical
      TabIndex        =   9
      Tag             =   " "
      ToolTipText     =   " To use Delete PO  "
      Top             =   9600
      UseMaskColor    =   -1  'True
      Width           =   1455
   End
   Begin VB.CommandButton cmdedit 
      BackColor       =   &H00FFC0C0&
      Caption         =   "&View PO"
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
      Left            =   4680
      MaskColor       =   &H00C0FFC0&
      Style           =   1  'Graphical
      TabIndex        =   8
      Tag             =   " "
      ToolTipText     =   " To Use View  PO "
      Top             =   9600
      UseMaskColor    =   -1  'True
      Width           =   1455
   End
   Begin VB.CommandButton cmdadd 
      BackColor       =   &H00FFC0C0&
      Caption         =   "&Add PO"
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
      Left            =   3120
      MaskColor       =   &H00C0FFC0&
      Style           =   1  'Graphical
      TabIndex        =   7
      Tag             =   " "
      ToolTipText     =   " To Use Add PO "
      Top             =   9600
      UseMaskColor    =   -1  'True
      Width           =   1455
   End
   Begin VB.CommandButton cmdexit 
      BackColor       =   &H00FFC0C0&
      Caption         =   "&Exit"
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
      Left            =   9480
      MaskColor       =   &H00C0FFC0&
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "  Exit Window  "
      Top             =   9600
      UseMaskColor    =   -1  'True
      Width           =   1695
   End
   Begin VB.TextBox txtponos 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   9600
      Visible         =   0   'False
      Width           =   1455
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
      Left            =   1440
      Style           =   2  'Dropdown List
      TabIndex        =   3
      ToolTipText     =   " Filter The PO No Wise "
      Top             =   1080
      Width           =   2055
   End
   Begin VB.CommandButton cmdlist1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "List All PO's"
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
      Left            =   13440
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   " Show All PO 's "
      Top             =   360
      Width           =   1575
   End
   Begin VB.ComboBox cbosup 
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
      Left            =   5160
      Style           =   2  'Dropdown List
      TabIndex        =   1
      ToolTipText     =   " Filter The Supplier Wise "
      Top             =   1080
      Width           =   3375
   End
   Begin VB.ComboBox cbodepts 
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
      Left            =   11160
      Style           =   2  'Dropdown List
      TabIndex        =   0
      ToolTipText     =   " Filter The Department Wise "
      Top             =   1080
      Width           =   2775
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   360
      Top             =   360
   End
   Begin MSFlexGridLib.MSFlexGrid PoMainGrid 
      Height          =   6615
      Left            =   240
      TabIndex        =   4
      ToolTipText     =   "  Note : # Indicate Place Click The First Row of Grid To  Open the Filter Options "
      Top             =   1560
      Width           =   14775
      _ExtentX        =   26061
      _ExtentY        =   11668
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
   Begin VB.Label Label1 
      BackColor       =   &H00EDDDD1&
      Caption         =   "Note : # Indicate Place Click The First Row of Grid To  Open the Filter Options "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   13
      Top             =   8760
      Width           =   10215
   End
   Begin VB.Label Label1 
      BackColor       =   &H00EDDDD1&
      Caption         =   "Purchase Order's ( PO )"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   495
      Index           =   0
      Left            =   360
      TabIndex        =   11
      Top             =   360
      Width           =   5415
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H000000FF&
      Height          =   735
      Index           =   1
      Left            =   240
      Top             =   240
      Width           =   14775
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H000000FF&
      Height          =   975
      Index           =   0
      Left            =   120
      Top             =   8400
      Width           =   15015
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00008000&
      BorderColor     =   &H000000C0&
      BorderWidth     =   2
      Height          =   735
      Left            =   120
      Top             =   9480
      Width           =   15015
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Height          =   8175
      Left            =   120
      Top             =   120
      Width           =   15015
   End
End
Attribute VB_Name = "frmpomain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim op As Variant
Dim cn As ADODB.Connection
Dim rs As ADODB.Recordset
Dim ww As ADODB.Recordset
Private Sub cbopono_Click()
    Call filterpono
End Sub
Private Sub cmdadd_Click()
    frmpos.Show
    frmpos.cmdsave.Visible = True
    frmpos.cmdupdate.Visible = False
    Unload Me
End Sub

Private Sub cmdclosure_Click()
    Unload Me
    frmpoclosure.Show
End Sub
Private Sub cmddelete_Click()
    Dim rs As New ADODB.Recordset
    Dim rs1 As New ADODB.Recordset
    Dim rs2 As New ADODB.Recordset
    
    If PoMainGrid.Rows <= 1 Then
        MsgBox "No Record Found", vbInformation, "Information"
        PoMainGrid.SetFocus
    ElseIf Trim(txtponos.Text) = "" Then
        MsgBox "Please Select The PO Number", vbCritical, "Selecting Error"
    Else
        rs.Open "Select * from grn_details where pono ='" & txtponos.Text & "'", cn, adOpenKeyset, adLockOptimistic
        rs1.Open "Select * from po_details where pono =" & txtponos.Text, cn, adOpenKeyset, adLockOptimistic
        rs2.Open "Select * from postatus_details where pono ='" & txtponos.Text & "'", cn, adOpenKeyset, adLockOptimistic
        
    If rs.RecordCount >= 1 Then
        MsgBox "Some Quantity Already Received Against This Purchase Order! So Caanot Delete!", vbInformation, "Cannot This Purchase Order Delete "
    Else
    If MsgBox("Are You Sure Delete PO No  " & txtponos.Text & " ? ", vbQuestion + vbYesNo, "Confirm To Delete") = vbYes Then
        i = 1
    If rs1.BOF = False Then rs1.MoveFirst
    While rs1.EOF = False
        rs1.Delete
        rs1.MoveNext
        i = i + 1
    Wend
        j = 1
    If rs2.BOF = False Then rs2.MoveFirst
    While rs2.EOF = False
        rs2.Delete
        rs2.MoveNext
        j = j + 1
    Wend
    MsgBox "One Record Deleted Successfully", vbInformation, "Information"
    Unload Me
    frmpomain.Show
End If

End If
End If
End Sub
Private Sub cmdedit_Click()
    Dim rs1 As New ADODB.Recordset
    rs1.Open "Select * from grn_details where pono= '" & txtponos.Text & "'", cn, 1, 3
     
    If rs1.RecordCount > 1 Then
    MsgBox "Already Some Of Qty Received Against This PO No! So Cann't Update", vbCritical, "Update Error"
    frmpos.cmdupdate.Enabled = False
    frmpos.PoGrid.Enabled = False
    End If
    
    If Trim(txtponos.Text) = "" Then
    MsgBox "Please Select the PO Number", vbCritical, "PO Number Selection Error "
    Else
    frmpos.cmdsave.Visible = False
    frmpos.cmdaddrow.Enabled = False
    frmpos.cmddeleterow.Enabled = False
    frmpos.cmdupdate.Visible = True
    frmpos.Show
    
    Dim i As Integer
    Dim rs As New ADODB.Recordset
    rs.Open "Select * from po_details where pono= " & frmpomain.txtponos.Text, cn, adOpenKeyset, adLockOptimistic
    i = 1
    If rs.BOF = False Then rs.MoveFirst
    While rs.EOF = False
               frmpos.PoGrid.Rows = frmpos.PoGrid.Rows + 1
               frmpos.txtpono.Text = rs![pono]
               frmpos.txtcomid.Text = rs![compname]
               frmpos.dt.Value = rs![podate]
               frmpos.cbosupplier.Text = rs![supname]
               frmpos.cbodept.Text = rs![deptname]
               frmpos.PoGrid.TextMatrix(i, 0) = i
               frmpos.PoGrid.TextMatrix(i, 1) = rs![itemname]
               frmpos.PoGrid.TextMatrix(i, 2) = rs![colour]
               frmpos.PoGrid.TextMatrix(i, 3) = rs![sizes]
               frmpos.PoGrid.TextMatrix(i, 4) = rs![qty]
               frmpos.PoGrid.TextMatrix(i, 5) = rs![uom]
               frmpos.PoGrid.TextMatrix(i, 6) = Format(rs![rates], "0.00")
               frmpos.PoGrid.TextMatrix(i, 7) = Format(rs![totamts], "0.00")
               frmpos.txttot.Text = Format(rs![totamt], "0.00")
               frmpos.txttax.Text = rs![tax]
               frmpos.taxamt.Text = Format(rs![taxamt], "0.00")
               frmpos.txtnet = Format(rs![netamt], "0.00")
               frmpos.txtremarks.Text = rs![remarks]
               frmpos.tw.Text = rs![words]
    rs.MoveNext
    i = i + 1
    Wend
    frmpos.PoGrid.Rows = frmpos.PoGrid.Rows - 1
    Unload Me
    End If
    
    
End Sub
Private Sub cmdexit_Click()
    op = MsgBox("Are You Sure To Close ?", vbYesNo + vbQuestion, "Confirm Close ?")
    If op = vbYes Then
    Unload Me
    Else
    End If
End Sub
Private Sub cmdlist1_Click()
    Call allpos
    Call visibles
End Sub
Private Sub cmdprint_Click()
    
    Dim rscomp As New ADODB.Recordset
    Dim rspo As New ADODB.Recordset
    Dim cn As New ADODB.Connection
    Dim rssup As New ADODB.Recordset
    Dim rsItem As New ADODB.Recordset
    Dim rstc As New ADODB.Recordset
    Dim irc As Integer
    Dim filenum As Integer
    
    cn.Open "Provider=Microsoft.jet.oledb.4.0; Data Source=" & App.path & "\Database\Data.mdb"
    rspo.Open "SELECT * FROM po_details where pono = " & PoMainGrid.TextMatrix(PoMainGrid.Row, 1), cn, 1, 3
    rssup.Open "select * from sup_details where supname ='" & PoMainGrid.TextMatrix(PoMainGrid.Row, 3) & "'", cn, 1, 3
    rscomp.Open "select * from com_details ", cn, 1, 3
    rstc.Open "Select * from condit_details", cn, 1, 3
    
    If Trim(txtponos.Text) = "" Then
    MsgBox "Please Select The PO No ", vbCritical, "PO No Selection Error"
    Else
       
    ChDrive App.path
    ChDir App.path
    
    Kill App.path & "\Reports\PrintPOGeneral.txt"
    Open App.path & "\Reports\PrintPOGeneral.txt" For Output As #1
    
    Print #1,
    Print #1, Tab(42); "PURCHASE ORDER";
    Print #1,
    Print #1, String(100, "-");
    Print #1, Tab(2); "PO Number: "; rspo![pono]; Tab(75); "PO Date :"; rspo![podate];
    Print #1,
    Print #1, String(100, "-");
    Print #1, Tab(2); "From,"; Tab(50); "|"; "To,";
    Print #1, Tab(7); rscomp![compname]; Tab(50); "|"; Tab(56); rspo![supname];
    Print #1, Tab(7); rscomp![address]; Tab(50); "|"; Tab(56); rssup![address];
    Print #1, Tab(7); rscomp![city]; Tab(50); "|"; Tab(56); rssup![city];
    Print #1, Tab(7); "Phone  :"; rscomp![phonenumber]; Tab(50); "|"; Tab(56); "Phone  :"; rssup![phonenumber];
    Print #1, Tab(7); "Mobile :"; rscomp![mobilenumber]; Tab(50); "|"; Tab(56); "Mobile :"; rssup![mobilenumber];
    Print #1, Tab(7); "Email  :"; rscomp![email]; Tab(50); "|"; Tab(56); "Email  :"; rssup![email];
    Print #1, Tab(7); "TIN No :"; rscomp![tinno]; Tab(50); "|"; Tab(56); "TIN No :"; rssup![tinno];
    Print #1, Tab(7); "TIN Date :"; rscomp![tindate]; Tab(50); "|"; Tab(56); "TIN Date :"; rssup![tindate];
    Print #1,
    Print #1, String(100, "-");
    Print #1, Tab(2); "S.No"; Tab(8); "Item Name"; Tab(24); "Colour"; Tab(38); "Size"; Tab(59); "Qty", Tab(71); "UOM"; Tab(81); "Rate"; Tab(90); "Tot.Amt";
    
    Print #1,
    Print #1, String(100, "-");
    rsItem.Open "Select * from po_details where pono =" & txtponos.Text, cn, 1, 3
    i = 1
    If rsItem.BOF = False Then rsItem.MoveFirst
    While rsItem.EOF = False
    Print #1,
    Print #1, rsItem![sno]; Tab(8); rsItem![itemname]; Tab(24); rsItem![colour]; Tab(38); rsItem![sizes]; Tab(58); Format(rsItem![qty], "0.000"); Tab(71); rsItem![uom]; Tab(81); Format(rsItem![rates], "0.00"); Tab(92); Format(rsItem![totamts], "0.00");
    rsItem.MoveNext
    i = i + 1
    Wend
    
    Print #1,
    Print #1, String(100, "-");
    Print #1, Tab(70); "Total Amount:        "; Format(rspo![totamt], "0.00");
    Print #1, Tab(50); String(50, "-");
    Print #1, Tab(45); "Tax (% ):  "; Format(rspo![tax], "0.00"); Tab(72); "Tax Amount:        "; Format(rspo![taxamt], "0.00");
    Print #1, Tab(50); String(50, "-");
    Print #1, Tab(72); "Net Amount:      "; Format(rspo![netamt], "0.00");
    Print #1,
    Print #1, String(100, "-");
    Print #1,
    
    Print #1, Tab(2); "In Words :"; rspo![words];
    Print #1,
    Print #1, String(100, "-");
    Print #1,
    
    Print #1, " * Terms And Conditions *";
    Print #1,
    i = 1
    If rstc.BOF = False Then rstc.MoveFirst
    While rstc.EOF = False
    Print #1,
    Print #1, Tab(5); rstc![con];
    rstc.MoveNext
    i = i + 1
    Wend
    
    Print #1,
    Print #1,
    Print #1,
    Print #1,
    Print #1,
    Print #1,
    Print #1,
    Print #1, String(100, "-");
    
    Print #1,
    Print #1, Tab(2); "Party's Signature"; Tab(25); "Dept.Incharge"; Tab(45); "Merchandiser"; Tab(65); "Approved By";
    Print #1,
    Print #1, String(100, "-");
    
    Close #1
    frmporeport.Show
    End If
End Sub
Private Sub Form_Load()
    Dim i As Integer
    Set cn = New ADODB.Connection
    Set ww = New ADODB.Recordset
    cn.Open "Provider=Microsoft.jet.oledb.4.0; Data Source=" & App.path & "\Database\Data.mdb"
    ww.Open "Select * From postatus_details", cn, adOpenKeyset, adLockOptimistic
    cn.CursorLocation = adUseClient
    
    Call pogridmainalign
    Call gridload
    
    cbopono.Visible = False
    cbosup.Visible = False
    cbodepts.Visible = False
    
    frmpomain.WindowState = 2
    
End Sub
Sub gridload()
    Dim i As Integer
    Dim rs As New ADODB.Recordset
    rs.Open "Select * from postatus_details where postatus = 'Open'", cn, adOpenKeyset, adLockOptimistic
    i = 1
    If rs.RecordCount = 0 Then
    MsgBox "No Record Found", vbInformation, "Information"
    Else
    If rs.BOF = False Then rs.MoveFirst
    While rs.EOF = False
    PoMainGrid.Rows = PoMainGrid.Rows + 1
    PoMainGrid.TextMatrix(i, 0) = i
    PoMainGrid.TextMatrix(i, 1) = rs![pono]
    PoMainGrid.TextMatrix(i, 2) = rs![podate]
    PoMainGrid.TextMatrix(i, 3) = rs![supname]
    PoMainGrid.TextMatrix(i, 4) = Format(rs![netamt], "0.00")
    PoMainGrid.TextMatrix(i, 5) = rs![deptname]
    rs.MoveNext
    i = i + 1
    Wend
    PoMainGrid.Rows = rs.RecordCount + 1
    End If
End Sub
Private Sub PoMainGrid_Click()
        If PoMainGrid.Col = 1 Or PoMainGrid.Col = 2 Or PoMainGrid.Col = 3 Or PoMainGrid.Col = 4 Or PoMainGrid.Col = 5 And PoMainGrid.Row = 1 Then
        txtponos.Text = PoMainGrid.TextMatrix(PoMainGrid.Row, 1)
        End If
        If PoMainGrid.Col = 1 And PoMainGrid.Row = 1 Then
        cbopono.Visible = True
        Else
        cbopono.Visible = False
        End If
        If PoMainGrid.Col = 3 And PoMainGrid.Row = 1 Then
        cbosup.Visible = True
        Else
        cbosup.Visible = False
        End If
        If PoMainGrid.Col = 5 And PoMainGrid.Row = 1 Then
        cbodepts.Visible = True
        Else
        cbodepts.Visible = False
        End If
End Sub
Private Sub cbodepts_Click()
    Call deptfilter
End Sub
Private Sub cbodepts_DropDown()
     On Error GoTo X
        cbodepts.Clear
        Set rs = cn.Execute("select deptname from postatus_details where postatus='Open' Group by deptname")
        rs.MoveFirst
        Do While Not rs.EOF()
        cbodepts.additem (rs(0))
        rs.MoveNext
        Loop
        cbodepts.SetFocus
X:
End Sub
Private Sub cbosup_Click()
    Call supfilter
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
Sub filterpono()
    Dim i As Integer
    Dim rs As New ADODB.Recordset
    rs.Open "Select * from postatus_details where pono='" & cbopono.Text & "'", cn, adOpenKeyset, adLockOptimistic
         i = 1
                PoMainGrid.TextMatrix(i, 1) = rs![pono]
                PoMainGrid.TextMatrix(i, 2) = rs![podate]
                PoMainGrid.TextMatrix(i, 3) = rs![supname]
                PoMainGrid.TextMatrix(i, 4) = Format(rs![netamt], "0.00")
                PoMainGrid.TextMatrix(i, 5) = rs![deptname]
        PoMainGrid.Rows = rs.RecordCount + 1
End Sub
Sub supfilter()
    Dim j As Integer
    Dim rs As New ADODB.Recordset
    rs.Open "select * from postatus_details where supname= '" & Trim(cbosup.Text) & "'", cn, adOpenKeyset, adLockOptimistic
    i = 1
    If rs.BOF = False Then rs.MoveFirst
    While rs.EOF = False
    PoMainGrid.Rows = PoMainGrid.Rows + 1
    PoMainGrid.TextMatrix(i, 0) = i
    PoMainGrid.TextMatrix(i, 1) = rs![pono]
    PoMainGrid.TextMatrix(i, 2) = rs![podate]
    PoMainGrid.TextMatrix(i, 3) = rs![supname]
    PoMainGrid.TextMatrix(i, 4) = Format(rs![netamt], "0.00")
    PoMainGrid.TextMatrix(i, 5) = rs![deptname]
    rs.MoveNext
    i = i + 1
    Wend
    PoMainGrid.Rows = rs.RecordCount + 1
End Sub
Sub deptfilter()
    Dim j As Integer
    Dim rs As New ADODB.Recordset
    rs.Open "select * from postatus_details where deptname= '" & Trim(cbodepts.Text) & "'", cn, 1, 3
    i = 1
    If rs.BOF = False Then rs.MoveFirst
    While rs.EOF = False
    PoMainGrid.Rows = PoMainGrid.Rows + 1
    PoMainGrid.TextMatrix(i, 0) = i
    PoMainGrid.TextMatrix(i, 1) = rs![pono]
    PoMainGrid.TextMatrix(i, 2) = rs![podate]
    PoMainGrid.TextMatrix(i, 3) = rs![supname]
    PoMainGrid.TextMatrix(i, 4) = Format(rs![netamt], "0.00")
    PoMainGrid.TextMatrix(i, 5) = rs![deptname]
    rs.MoveNext
    i = i + 1
    Wend
    PoMainGrid.Rows = rs.RecordCount + 1
End Sub
Sub allpos()
    Dim j As Integer
    Dim rs As New ADODB.Recordset
    rs.Open "select * from postatus_details where postatus='Open'", cn, adOpenKeyset, adLockOptimistic
    If rs.RecordCount = 0 Then
    MsgBox "No Record Found", vbInformation, "Information"
    Else
    i = 1
    If rs.BOF = False Then rs.MoveFirst
    While rs.EOF = False
    PoMainGrid.Rows = PoMainGrid.Rows + 1
    PoMainGrid.TextMatrix(i, 0) = i
    PoMainGrid.TextMatrix(i, 1) = rs![pono]
    PoMainGrid.TextMatrix(i, 2) = rs![podate]
    PoMainGrid.TextMatrix(i, 3) = rs![supname]
    PoMainGrid.TextMatrix(i, 4) = Format(rs![netamt], "0.00")
    PoMainGrid.TextMatrix(i, 5) = rs![deptname]
    rs.MoveNext
    i = i + 1
    Wend
    PoMainGrid.Rows = rs.RecordCount + 1
    Call clears
    End If
End Sub
Sub clears()
    cbopono.Clear
    cbosup.Clear
    cbodepts.Clear
End Sub
Private Sub cbopono_dropdown()
        On Error GoTo X
        cbopono.Clear
        Set rs = cn.Execute("select pono from postatus_details where postatus='Open'")
        rs.MoveFirst
        Do While Not rs.EOF()
        cbopono.additem (rs(0))
        rs.MoveNext
        Loop
        cbopono.SetFocus
X:
End Sub
Sub visibles()
    cbopono.Visible = False
    cbosup.Visible = False
    cbodepts.Visible = False
End Sub
