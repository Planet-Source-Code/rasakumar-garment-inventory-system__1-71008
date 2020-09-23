VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmopendeliverymain 
   BackColor       =   &H00EDDDD1&
   Caption         =   " * Open Delivery Main  ( Against Open GRN )"
   ClientHeight    =   9150
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10410
   Icon            =   "frmopendeliverymain.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9150
   ScaleWidth      =   10410
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdlist1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "List All GRN's"
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
      TabIndex        =   10
      ToolTipText     =   " Show All PO 's "
      Top             =   360
      Width           =   1575
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
      Left            =   9840
      Style           =   2  'Dropdown List
      TabIndex        =   9
      ToolTipText     =   " To Use Filter The Department Wise "
      Top             =   1080
      Width           =   2895
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
      Left            =   5280
      Style           =   2  'Dropdown List
      TabIndex        =   8
      ToolTipText     =   " To Use Filter The Supplier Name Wise "
      Top             =   1080
      Width           =   4455
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
      TabIndex        =   7
      ToolTipText     =   " To Use Filter The GRN No Wise "
      Top             =   1080
      Width           =   1935
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
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   8760
      Visible         =   0   'False
      Width           =   1335
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
      Left            =   10320
      MaskColor       =   &H00C0FFC0&
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "  Exit Window  "
      Top             =   9600
      UseMaskColor    =   -1  'True
      Width           =   1695
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
      Left            =   3600
      MaskColor       =   &H00C0FFC0&
      Style           =   1  'Graphical
      TabIndex        =   4
      Tag             =   " "
      ToolTipText     =   " To Use Add GRN"
      Top             =   9600
      UseMaskColor    =   -1  'True
      Width           =   1575
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
      Left            =   5280
      MaskColor       =   &H00C0FFC0&
      Style           =   1  'Graphical
      TabIndex        =   3
      Tag             =   " "
      ToolTipText     =   " To Use View GRN "
      Top             =   9600
      UseMaskColor    =   -1  'True
      Width           =   1575
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
      Left            =   6960
      MaskColor       =   &H00C0FFC0&
      Style           =   1  'Graphical
      TabIndex        =   2
      Tag             =   " "
      ToolTipText     =   " To use Delete GRN "
      Top             =   9600
      UseMaskColor    =   -1  'True
      Width           =   1575
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
      Left            =   8640
      MaskColor       =   &H00C0FFC0&
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "  To use Print GRN   "
      Top             =   9600
      UseMaskColor    =   -1  'True
      Width           =   1575
   End
   Begin VB.TextBox txttypeno 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   2040
      TabIndex        =   0
      Top             =   8760
      Visible         =   0   'False
      Width           =   1455
   End
   Begin MSFlexGridLib.MSFlexGrid DeliveryMainGrid 
      Height          =   6615
      Left            =   240
      TabIndex        =   11
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
   Begin VB.Label Label1 
      BackColor       =   &H00EDDDD1&
      Caption         =   "Delivery Challan ( Open GRN  Against )"
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
      TabIndex        =   12
      Top             =   360
      Width           =   8295
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
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Height          =   8175
      Left            =   120
      Top             =   120
      Width           =   15015
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H000000C0&
      BorderWidth     =   2
      Height          =   735
      Left            =   120
      Top             =   9480
      Width           =   15015
   End
End
Attribute VB_Name = "frmopendeliverymain"
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
        Set rs = cn.Execute("select dcno from opendeliverymain_details order by dcno")
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
        Set rs = cn.Execute("select deptname from opendeliverymain_details group by deptname")
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
        Set rs = cn.Execute("select supname from opendeliverymain_details group by supname")
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
        Set rs = cn.Execute("select deliverytype from opendeliverymain_details group by deliverytype")
        rs.MoveFirst
        Do While Not rs.EOF()
        cbotypefilter.additem (rs(0))
        rs.MoveNext
        Loop
        cbotypefilter.SetFocus
X:
End Sub
Private Sub cmdadd_Click()
    frmopenDeliverys.Show
    Unload Me
End Sub
Private Function DeliveryMainGridload()
    Dim i As Integer
    Dim rs As New ADODB.Recordset
    Dim rs1 As New ADODB.Recordset
    rs.Open "select * from opendeliverymain_details ", cn, adOpenKeyset, adLockOptimistic
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
    Dim rs1 As New ADODB.Recordset
    Dim rs2 As New ADODB.Recordset
    
    If Trim(txtedit.Text) = "" Then
        MsgBox "Please Select The DC Number", vbCritical, "Selecting Error"
    Else
        rs1.Open "Select * from deliveryopen_details where dcno ='" & txtedit.Text & "'", cn, adOpenKeyset, adLockOptimistic
'        rs2.Open "Select * from opengrnstatus_details where opengrnno ='" & txteditgrn.Text & "'", cn, adOpenKeyset, adLockOptimistic
    If rs1.RecordCount <> 0 Then

    If MsgBox("Are You Sure Delete PO No  " & txtedit.Text & " ? ", vbQuestion + vbYesNo, "Confirm To Delete") = vbYes Then
    i = 1
    If rs1.BOF = False Then rs1.MoveFirst
    While rs1.EOF = False
    rs1.Delete
    rs1.MoveNext
    i = i + 1
    Wend
'         j = 1
'    If rs2.BOF = False Then rs2.MoveFirst
'    While rs2.EOF = False
'        rs2.Delete
'        rs2.MoveNext
'        j = j + 1
'    Wend
    MsgBox "One Record Deleted Successfully", vbInformation, "Information"
    Unload Me
    frmopendeliverymain.Show
End If
End If
End If

End Sub
Private Sub cmdedit_Click()
    If Trim(txtedit.Text) = "" Then
    MsgBox "Please Select the DC Number", vbCritical, "PO Number Selection Error "
    Else
    frmopenDeliverys.Show
    frmopenDeliverys.cmdsave.Enabled = False
    frmopenDeliverys.cmdexits.Visible = True
    frmopenDeliverys.DetailsGrid.Visible = False
    frmopenDeliverys.cmdadditem.Visible = False
    frmopenDeliverys.cmddeleteitem.Visible = False
    frmopenDeliverys.txtdcno.Enabled = False
    frmopenDeliverys.DeliveryGrid.Enabled = False
    frmopenDeliverys.DeliveryMainGrid.Visible = False
    frmopenDeliverys.Shape2.Visible = False
    frmopenDeliverys.Shape1(0).Height = 9300
    frmopenDeliverys.cbodept.Enabled = False
    frmopenDeliverys.cbosupplier.Enabled = False
    frmopenDeliverys.txtremarks.Enabled = False
    frmopenDeliverys.allgrnopt.Enabled = False
    frmopenDeliverys.supopt.Enabled = False
    frmopenDeliverys.cbosup.Visible = False
    frmopenDeliverys.cbogrnno.Enabled = False
    
    frmopenDeliverys.DeliveryGrid.Height = 8100
    frmopenDeliverys.DeliveryGrid.Width = 14800
    
    Dim i As Integer
    Dim rs As New ADODB.Recordset
    Dim rs1 As New ADODB.Recordset
    rs.Open "Select * from deliveryopen_details where dcno= '" & frmopendeliverymain.txtedit.Text & "'", cn, adOpenKeyset, adLockOptimistic
    i = 1
    If rs.BOF = False Then rs.MoveFirst
    While rs.EOF = False
        frmopenDeliverys.DeliveryGrid.Rows = rs.RecordCount + 1
        frmopenDeliverys.txtdcno.Text = rs![dcno]
        frmopenDeliverys.dt1.Value = rs![dcdate]
        frmopenDeliverys.cbosupplier.Text = rs![supname]
        frmopenDeliverys.cbodept.Text = rs![deptname]
        frmopenDeliverys.DeliveryGrid.TextMatrix(i, 0) = rs![sno]
        frmopenDeliverys.DeliveryGrid.TextMatrix(i, 1) = rs![opengrnnos]
        frmopenDeliverys.DeliveryGrid.TextMatrix(i, 2) = rs![opengrndates]
        frmopenDeliverys.DeliveryGrid.TextMatrix(i, 3) = rs![itemname]
        frmopenDeliverys.DeliveryGrid.TextMatrix(i, 4) = rs![colour]
        frmopenDeliverys.DeliveryGrid.TextMatrix(i, 5) = rs![sizes]
        frmopenDeliverys.DeliveryGrid.TextMatrix(i, 6) = Format(rs![stockqty], "0.000")
        frmopenDeliverys.DeliveryGrid.TextMatrix(i, 7) = rs![uom]
        frmopenDeliverys.DeliveryGrid.TextMatrix(i, 8) = Format(rs![delqty], "0.000")
        frmopenDeliverys.DeliveryGrid.TextMatrix(i, 9) = rs![grnid]
        frmopenDeliverys.DeliveryGrid.TextMatrix(i, 11) = rs![grnnos]
        frmopenDeliverys.txtremarks.Text = rs![remarks]
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
    rsgrn.Open "SELECT * FROM deliveryopen_details where dcno ='" & txtedit.Text & "'", cn, 1, 3
    rssup.Open "select * from sup_details where supname ='" & DeliveryMainGrid.TextMatrix(DeliveryMainGrid.Row, 3) & "'", cn, 1, 3
    rscomp.Open "select * from com_details ", cn, 1, 3
    
    If Trim(txtedit.Text) = "" Then
    MsgBox "Please Select The Open GRN No ", vbCritical, "GRN No Selection Error"
    txtedit.SetFocus
    Else
    
    ChDrive App.path
    ChDir App.path
    
    Kill App.path & "\Reports\Printopengrndelivery.txt"
    Open App.path & "\Reports\Printopengrndelivery.txt" For Output As #1
   
    Print #1,
    Print #1, Tab(42); "Delivery Challan ( Open Goods Receipt Against )";
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
    rsItem.Open "Select * from deliveryopen_details where dcno ='" & txtedit.Text & "'", cn, 1, 3
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
    frmopengrndeliveryreport.Show
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
    ww.Open "Select * From opendeliverymain_details", cn, adOpenKeyset, adLockOptimistic
    Call frmdeliverygridloadmains
    Call DeliveryMainGridload
    cboDeliveryFilter.Visible = False
    cbosupfilter.Visible = False
    cbodept.Visible = False
End Sub
Private Function filtersupname()
    Dim rs As New ADODB.Recordset
    rs.Open "select * from opendeliverymain_details where supname= '" & Trim(cbosupfilter.Text) & "'", cn, adOpenKeyset, adLockOptimistic
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
    rs.Open "select * from opendeliverymain_details where dcno= '" & Trim(cboDeliveryFilter.Text) & "'", cn, adOpenKeyset, adLockOptimistic
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
    rs.Open "select * from opendeliverymain_details where deptname= '" & Trim(cbodept.Text) & "'", cn, adOpenKeyset, adLockOptimistic
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
    rs.Open "select * from opendeliverymain_details where deliverytype= '" & Trim(cbotypefilter.Text) & "'", cn, adOpenKeyset, adLockOptimistic
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
    rs.Open "select * from opendeliverymain_details", cn, 1, 3
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


