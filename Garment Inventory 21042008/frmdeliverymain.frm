VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmdeliverymain 
   BackColor       =   &H00EDDDD1&
   Caption         =   " * Delivery Challans *"
   ClientHeight    =   8085
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10020
   Icon            =   "frmdeliverymain.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   11115
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.TextBox txttypeno 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   2040
      TabIndex        =   12
      Top             =   8760
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton cmdprint 
      BackColor       =   &H00FFC0C0&
      Caption         =   "&Print DC"
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
      Left            =   9000
      MaskColor       =   &H00C0FFC0&
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "  To use Print DC "
      Top             =   9600
      UseMaskColor    =   -1  'True
      Width           =   1335
   End
   Begin VB.CommandButton cmddelete 
      BackColor       =   &H00FFC0C0&
      Caption         =   "&Delete DC"
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
      Left            =   7440
      MaskColor       =   &H00C0FFC0&
      Style           =   1  'Graphical
      TabIndex        =   8
      Tag             =   " "
      ToolTipText     =   " To use Delete DC"
      Top             =   9600
      UseMaskColor    =   -1  'True
      Width           =   1455
   End
   Begin VB.CommandButton cmdedit 
      BackColor       =   &H00FFC0C0&
      Caption         =   "&View DC"
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
      Left            =   5880
      MaskColor       =   &H00C0FFC0&
      Style           =   1  'Graphical
      TabIndex        =   7
      Tag             =   " "
      ToolTipText     =   " To Use View DC "
      Top             =   9600
      UseMaskColor    =   -1  'True
      Width           =   1455
   End
   Begin VB.CommandButton cmdadd 
      BackColor       =   &H00FFC0C0&
      Caption         =   "&Add DC"
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
      Left            =   4320
      MaskColor       =   &H00C0FFC0&
      Style           =   1  'Graphical
      TabIndex        =   6
      Tag             =   " "
      ToolTipText     =   " To Use Add DC"
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
      Left            =   10440
      MaskColor       =   &H00C0FFC0&
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "  Exit Window  "
      Top             =   9600
      UseMaskColor    =   -1  'True
      Width           =   1335
   End
   Begin VB.TextBox txtedit 
      Alignment       =   2  'Center
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
      Height          =   375
      Left            =   720
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   8760
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.ComboBox cboDeliveryFilter 
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
      TabIndex        =   3
      ToolTipText     =   " To Use Filter The GRN No Wise "
      Top             =   1080
      Width           =   1935
   End
   Begin VB.ComboBox cbosupfilter 
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
      Left            =   5520
      Style           =   2  'Dropdown List
      TabIndex        =   2
      ToolTipText     =   " To Use Filter The Supplier Name Wise "
      Top             =   1080
      Width           =   4935
   End
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
      Left            =   10440
      Style           =   2  'Dropdown List
      TabIndex        =   1
      ToolTipText     =   " To Use Filter The Department Wise "
      Top             =   1080
      Width           =   3495
   End
   Begin VB.CommandButton cmdlist1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "List All DC's"
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
      Left            =   13320
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   " Show All DC 's "
      Top             =   360
      Width           =   1575
   End
   Begin MSFlexGridLib.MSFlexGrid DeliveryMainGrid 
      Height          =   6615
      Left            =   240
      TabIndex        =   10
      ToolTipText     =   "  Note : # Indicate Place Click The First Row of Grid To  Open the Filter Options "
      Top             =   1560
      Width           =   14775
      _ExtentX        =   26061
      _ExtentY        =   11668
      _Version        =   393216
      Rows            =   1
      Cols            =   5
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
   Begin VB.Shape Shape1 
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
   Begin VB.Shape Shape3 
      BorderColor     =   &H000000FF&
      Height          =   975
      Index           =   0
      Left            =   120
      Top             =   8400
      Width           =   15015
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H000000FF&
      Height          =   735
      Index           =   1
      Left            =   240
      Top             =   240
      Width           =   14775
   End
   Begin VB.Label Label1 
      BackColor       =   &H00EDDDD1&
      Caption         =   "Delivery Challan ( PO GRN Against )"
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
      Width           =   7335
   End
End
Attribute VB_Name = "frmdeliverymain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim op As Variant
Dim cn As ADODB.Connection
Dim rs As ADODB.Recordset
Dim ww As ADODB.Recordset
Private Sub cboDeliveryFilter_Click()
    Call filterdcno
End Sub
Private Sub cboDeliveryFilter_dropdown()
Dim rs As New ADODB.Recordset
    On Error GoTo X
        cboDeliveryFilter.Clear
        Set rs = cn.Execute("select dcno from deliverymain_details order by dcno")
        rs.MoveFirst
        Do While Not rs.EOF()
        cboDeliveryFilter.additem (rs(0))
        rs.MoveNext
        Loop
        cboDeliveryFilter.SetFocus
X:
End Sub
Private Sub cbodept_Click()
    Call filterdept
End Sub
Private Sub cbodept_dropdown()
Dim rs As New ADODB.Recordset
    On Error GoTo X
        cbodept.Clear
        Set rs = cn.Execute("select deptname from deliverymain_details group by deptname")
        rs.MoveFirst
        Do While Not rs.EOF()
        cbodept.additem (rs(0))
        rs.MoveNext
        Loop
        cbodept.SetFocus
X:
End Sub
Private Sub cbosupfilter_Click()
    Call filtersupname
End Sub
Private Sub cbosupfilter_dropdown()
Dim rs As New ADODB.Recordset
    On Error GoTo X
        cbosupfilter.Clear
        Set rs = cn.Execute("select supname from deliverymain_details group by supname")
        rs.MoveFirst
        Do While Not rs.EOF()
        cbosupfilter.additem (rs(0))
        rs.MoveNext
        Loop
        cbosupfilter.SetFocus
X:
End Sub
Private Sub cbotypefilter_Click()
    Call filtertype
End Sub
Private Sub cbotypefilter_dropdown()
    Dim rs As New ADODB.Recordset
    On Error GoTo X
        cbotypefilter.Clear
        Set rs = cn.Execute("select deliverytype from deliverymain_details group by deliverytype")
        rs.MoveFirst
        Do While Not rs.EOF()
        cbotypefilter.additem (rs(0))
        rs.MoveNext
        Loop
        cbotypefilter.SetFocus
X:
End Sub
Private Sub cmdadd_Click()
    frmDeliverys.Show
    frmDeliverys.allgrnopt.Value = True
    Unload Me
End Sub
Private Function DeliveryMainGridload()
    Dim i As Integer
    Dim rs As New ADODB.Recordset
    Dim rs1 As New ADODB.Recordset
    rs.Open "select * from deliverymain_details ", cn, adOpenKeyset, adLockOptimistic
    i = 1
    If rs.BOF = False Then rs.MoveFirst
    While rs.EOF = False
        DeliveryMainGrid.Rows = rs.RecordCount + 1
        DeliveryMainGrid.TextMatrix(i, 0) = i
        DeliveryMainGrid.TextMatrix(i, 1) = rs![dcno]
        DeliveryMainGrid.TextMatrix(i, 2) = rs![dcdate]
        DeliveryMainGrid.TextMatrix(i, 3) = rs![supname]
        DeliveryMainGrid.TextMatrix(i, 4) = rs![deptname]
     
        rs.MoveNext
        i = i + 1
    Wend
        DeliveryMainGrid.Rows = rs.RecordCount + 1
End Function

Private Sub cmddelete_Click()
    Dim rs As New ADODB.Recordset
    If Trim(txtedit.Text) = "" Then
        MsgBox "Please Select The DC Number", vbCritical, "Selecting Error"
    Else
        rs.Open "Select * from delivery_details where dcno= '" & txtedit.Text & "'", cn, adOpenKeyset, adLockOptimistic
        If rs.RecordCount <> 0 Then
        If MsgBox("Are You Sure Delete PO No  " & txtedit.Text & " ? ", vbQuestion + vbYesNo, "Confirm To Delete") = vbYes Then
        i = 1
        If rs.BOF = False Then rs.MoveFirst
        While rs.EOF = False
        rs.Delete
        rs.MoveNext
        i = i + 1
        Wend
        MsgBox "One Record Deleted Successfully", vbInformation, "Information"
        Unload Me
        frmdeliverymain.Show
End If
End If
End If
End Sub
Private Sub cmdedit_Click()
    If Trim(txtedit.Text) = "" Then
    MsgBox "Please Select the DC Number", vbCritical, "PO Number Selection Error "
    Else
    frmDeliverys.Show
    frmDeliverys.cmdsave.Enabled = False
    frmDeliverys.cmdexits.Visible = True
    frmDeliverys.DetailsGrid.Visible = False
    frmDeliverys.cmdadditem.Visible = False
    frmDeliverys.cmddeleteitem.Visible = False
    frmDeliverys.txtdcno.Enabled = False
    frmDeliverys.DeliveryGrid.Enabled = False
    frmDeliverys.DeliveryMainGrid.Visible = False
    frmDeliverys.Shape2.Visible = False
    frmDeliverys.Shape1(0).Height = 9300
    frmDeliverys.cbodept.Enabled = False
    frmDeliverys.cbosupplier.Enabled = False
    frmDeliverys.txtremarks.Enabled = False
    frmDeliverys.allgrnopt.Enabled = False
    frmDeliverys.supopt.Enabled = False
    frmDeliverys.cbosup.Enabled = False
    frmDeliverys.cbogrnno.Enabled = False
    
    frmDeliverys.DeliveryGrid.Height = 8100
    frmDeliverys.DeliveryGrid.Width = 14800
    
    Dim i As Integer
    Dim rs As New ADODB.Recordset
    Dim rs1 As New ADODB.Recordset
    rs.Open "Select * from delivery_details where dcno= '" & frmdeliverymain.txtedit.Text & "'", cn, adOpenKeyset, adLockOptimistic
    i = 1
    If rs.BOF = False Then rs.MoveFirst
    While rs.EOF = False
        frmDeliverys.DeliveryGrid.Rows = rs.RecordCount + 1
        frmDeliverys.txtdcno.Text = rs![dcno]
        frmDeliverys.dt1.Value = rs![dcdate]
        frmDeliverys.cbosupplier.Text = rs![supname]
        frmDeliverys.cbodept.Text = rs![deptname]
        frmDeliverys.DeliveryGrid.TextMatrix(i, 0) = rs![sno]
        frmDeliverys.DeliveryGrid.TextMatrix(i, 1) = rs![grnno]
        frmDeliverys.DeliveryGrid.TextMatrix(i, 2) = rs![grndate]
        frmDeliverys.DeliveryGrid.TextMatrix(i, 3) = rs![itemname]
        frmDeliverys.DeliveryGrid.TextMatrix(i, 4) = rs![colour]
        frmDeliverys.DeliveryGrid.TextMatrix(i, 5) = rs![sizes]
        frmDeliverys.DeliveryGrid.TextMatrix(i, 6) = Format(rs![stockqty], "0.000")
        frmDeliverys.DeliveryGrid.TextMatrix(i, 7) = rs![uom]
        frmDeliverys.DeliveryGrid.TextMatrix(i, 8) = Format(rs![delqty], "0.000")
        frmDeliverys.DeliveryGrid.TextMatrix(i, 9) = Format(rs![grnid], "0.000")
        frmDeliverys.DeliveryGrid.TextMatrix(i, 10) = rs![grnnos]
        frmDeliverys.txtremarks.Text = rs![remarks]
       
       
        rs.MoveNext
        i = i + 1
        Wend
    End If
    Unload Me

End Sub

Private Sub cmdexit_Click()
     op = MsgBox("Are you Sure To Close ?", vbYesNo + vbQuestion, "Confirm Close ?")
    If op = vbYes Then
     Unload Me
    Else
    End If
End Sub
Private Sub cmdlist1_Click()
    Call filteralldc
    
    cboDeliveryFilter.Visible = False
    cbosupfilter.Visible = False
    cbodept.Visible = False
    cboDeliveryFilter.Clear
    cbosupfilter.Clear
    cbodept.Clear
    If DeliveryMainGrid.Rows <= 1 Then
    MsgBox "No Record Found", vbInformation, "Information"
    End If
End Sub
Private Sub cmdprint_Click()
    Dim rscomp As New ADODB.Recordset
    Dim rsgrn As New ADODB.Recordset
    Dim cn As New ADODB.Connection
    Dim rssup As New ADODB.Recordset
    Dim rsItem As New ADODB.Recordset
    Dim irc As Integer
    Dim filenum As Integer
    
    cn.Open "Provider=Microsoft.jet.oledb.4.0; Data Source=" & App.path & "\Database\Data.mdb"
    rsgrn.Open "SELECT * FROM delivery_details where dcno ='" & txtedit.Text & "'", cn, 1, 3
    rssup.Open "select * from sup_details where supname ='" & DeliveryMainGrid.TextMatrix(DeliveryMainGrid.Row, 3) & "'", cn, 1, 3
    rscomp.Open "select * from com_details ", cn, 1, 3
    
    If Trim(txtedit.Text) = "" Then
    MsgBox "Please Select The DC No ", vbCritical, "DC No Selection Error"
    Else
    
    ChDrive App.path
    ChDir App.path
    
    Kill App.path & "\Reports\Printpogrndelivery.txt"
    Open App.path & "\Reports\Printpogrndelivery.txt" For Output As #1
   
    Print #1,
    Print #1, Tab(42); "Delivery Challan ( PO GRN Against )";
    Print #1,
    Print #1, String(100, "-");
    Print #1, Tab(2); "DC No: "; rsgrn![dcno]; Tab(75); "DC Date :"; rsgrn![dcdate];
    Print #1,
    Print #1, String(100, "-");
    Print #1, Tab(2); "From,"; Tab(50); "|"; "To,";
    Print #1, Tab(7); rscomp![compname]; Tab(50); "|"; Tab(56); rsgrn![supname];
    Print #1, Tab(7); rscomp![address]; Tab(50); "|"; Tab(56); rssup![address];
    Print #1, Tab(7); rscomp![city]; Tab(50); "|"; Tab(56); rssup![city];
    Print #1, Tab(7); "Phone  :"; rscomp![phonenumber]; Tab(50); "|"; Tab(56); "Phone  :"; rssup![phonenumber];
    Print #1, Tab(7); "Mobile :"; rscomp![mobilenumber]; Tab(50); "|"; Tab(56); "Mobile :"; rssup![mobilenumber];
    Print #1, Tab(7); "Email  :"; rscomp![email]; Tab(50); "|"; Tab(56); "Email  :"; rssup![email];
    Print #1, Tab(7); "TIN No :"; rscomp![tinno]; Tab(50); "|"; Tab(56); "TIN No :"; rssup![tinno];
    Print #1, Tab(7); "TIN Date :"; rscomp![tindate]; Tab(50); "|"; Tab(56); "TIN Date :"; rssup![tindate];
    Print #1,
    Print #1, String(100, "-");
    Print #1, Tab(2); "S.No"; Tab(8); "Item Name"; Tab(25); "Colour"; Tab(40); "Size"; Tab(59); "Rec.Qty", Tab(71); "UOM"; Tab(85); "GRN No"; 'Tab(92); "Tot.Amt";
    
    Print #1,
    Print #1, String(100, "-");
    rsItem.Open "Select * from delivery_details where dcno ='" & txtedit.Text & "'", cn, 1, 3
    i = 1
    If rsItem.BOF = False Then rsItem.MoveFirst
    While rsItem.EOF = False
    Print #1,
    Print #1, rsItem![sno]; Tab(8); rsItem![itemname]; Tab(25); rsItem![colour]; Tab(40); rsItem![sizes]; Tab(59); Format(rsItem![delqty], "0.000"); Tab(71); rsItem![uom];
    rsItem.MoveNext
    i = i + 1
    Wend
    Print #1,
    Print #1, String(100, "-");
    Print #1,
    Print #1, String(100, "-");
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
    frmpogrndeliveryreport.Show
    End If
End Sub

Private Sub DeliveryMainGrid_Click()
    If DeliveryMainGrid.Col = 0 Or DeliveryMainGrid.Col = 1 Or DeliveryMainGrid.Col = 2 Or DeliveryMainGrid.Col = 3 Or DeliveryMainGrid.Col = 4 Or DeliveryMainGrid.Col = 5 Or DeliveryMainGrid.Col = 6 Then
         txtedit.Text = DeliveryMainGrid.TextMatrix(DeliveryMainGrid.Row, 1)
    End If
        If DeliveryMainGrid.Col = 1 And DeliveryMainGrid.Row = 1 Then
        cboDeliveryFilter.Visible = True
        Else
        cboDeliveryFilter.Visible = False
        End If
        If DeliveryMainGrid.Col = 3 And DeliveryMainGrid.Row = 1 Then
        cbosupfilter.Visible = True
        Else
        cbosupfilter.Visible = False
        End If
        If DeliveryMainGrid.Col = 4 And DeliveryMainGrid.Row = 1 Then
        cbodept.Visible = True
        Else
        cbodept.Visible = False
        End If
End Sub
Private Sub Form_Load()
    Set cn = New ADODB.Connection
    Set ww = New ADODB.Recordset
    Set rs = New ADODB.Recordset
    cn.Open "Provider=Microsoft.jet.oledb.4.0; Data Source=" & App.path & "\Database\Data.mdb"
    ww.Open "Select * From delivery_details", cn, adOpenKeyset, adLockOptimistic
    Call DeliveryMainGridload
    Call frmdeliverymaingridloads
    
    cboDeliveryFilter.Visible = False
    cbosupfilter.Visible = False
    cbodept.Visible = False
   
End Sub
Private Function filtersupname()
    Dim rs As New ADODB.Recordset
    rs.Open "select * from deliverymain_details where supname= '" & Trim(cbosupfilter.Text) & "'", cn, adOpenKeyset, adLockOptimistic
    i = 1
    If rs.BOF = False Then rs.MoveFirst
    While rs.EOF = False
    DeliveryMainGrid.Rows = DeliveryMainGrid.Rows + 1
    DeliveryMainGrid.TextMatrix(i, 0) = i
    DeliveryMainGrid.TextMatrix(i, 1) = rs![dcno]
    DeliveryMainGrid.TextMatrix(i, 2) = rs![dcdate]
    DeliveryMainGrid.TextMatrix(i, 3) = rs![supname]
    DeliveryMainGrid.TextMatrix(i, 4) = rs![deptname]
    rs.MoveNext
    i = i + 1
    Wend
    DeliveryMainGrid.Rows = rs.RecordCount + 1
End Function
Private Function filterdcno()
    Dim rs As New ADODB.Recordset
    rs.Open "select * from deliverymain_details where dcno= '" & Trim(cboDeliveryFilter.Text) & "'", cn, adOpenKeyset, adLockOptimistic
    i = 1
    If rs.BOF = False Then rs.MoveFirst
    While rs.EOF = False
    DeliveryMainGrid.Rows = DeliveryMainGrid.Rows + 1
    DeliveryMainGrid.TextMatrix(i, 0) = i
    DeliveryMainGrid.TextMatrix(i, 1) = rs![dcno]
    DeliveryMainGrid.TextMatrix(i, 2) = rs![dcdate]
    DeliveryMainGrid.TextMatrix(i, 3) = rs![supname]
    DeliveryMainGrid.TextMatrix(i, 4) = rs![deptname]
    rs.MoveNext
    i = i + 1
    Wend
    DeliveryMainGrid.Rows = rs.RecordCount + 1
End Function
Private Function filterdept()
    Dim rs As New ADODB.Recordset
    rs.Open "select * from deliverymain_details where deptname= '" & Trim(cbodept.Text) & "'", cn, adOpenKeyset, adLockOptimistic
    i = 1
    If rs.BOF = False Then rs.MoveFirst
    While rs.EOF = False
    DeliveryMainGrid.Rows = DeliveryMainGrid.Rows + 1
    DeliveryMainGrid.TextMatrix(i, 0) = i
    DeliveryMainGrid.TextMatrix(i, 1) = rs![dcno]
    DeliveryMainGrid.TextMatrix(i, 2) = rs![dcdate]
    DeliveryMainGrid.TextMatrix(i, 3) = rs![supname]
    DeliveryMainGrid.TextMatrix(i, 4) = rs![deptname]
    rs.MoveNext
    i = i + 1
    Wend
    DeliveryMainGrid.Rows = rs.RecordCount + 1
End Function
Private Function filtertype()
    Dim rs As New ADODB.Recordset
    rs.Open "select * from deliverymain_details where deliverytype= '" & Trim(cbotypefilter.Text) & "'", cn, adOpenKeyset, adLockOptimistic
    i = 1
    If rs.BOF = False Then rs.MoveFirst
    While rs.EOF = False
    DeliveryMainGrid.Rows = DeliveryMainGrid.Rows + 1
    DeliveryMainGrid.TextMatrix(i, 0) = i
    DeliveryMainGrid.TextMatrix(i, 1) = rs![dcno]
    DeliveryMainGrid.TextMatrix(i, 2) = rs![dcdate]
    DeliveryMainGrid.TextMatrix(i, 3) = rs![supname]
    DeliveryMainGrid.TextMatrix(i, 4) = rs![deptname]
    rs.MoveNext
    i = i + 1
    Wend
    DeliveryMainGrid.Rows = rs.RecordCount + 1
End Function
Private Function filteralldc()
    Dim rs As New ADODB.Recordset
    rs.Open "select * from deliverymain_details", cn, 1, 3
    i = 1
    If rs.BOF = False Then rs.MoveFirst
    While rs.EOF = False
    DeliveryMainGrid.Rows = DeliveryMainGrid.Rows + 1
    DeliveryMainGrid.TextMatrix(i, 0) = i
    DeliveryMainGrid.TextMatrix(i, 1) = rs![dcno]
    DeliveryMainGrid.TextMatrix(i, 2) = rs![dcdate]
    DeliveryMainGrid.TextMatrix(i, 3) = rs![supname]
    DeliveryMainGrid.TextMatrix(i, 4) = rs![deptname]
    rs.MoveNext
    i = i + 1
    Wend
    DeliveryMainGrid.Rows = rs.RecordCount + 1
End Function

