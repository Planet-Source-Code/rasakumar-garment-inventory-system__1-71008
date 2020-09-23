VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmopeninvoicemain 
   BackColor       =   &H00EDDDD1&
   Caption         =   " * Open Invoice * "
   ClientHeight    =   9015
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9990
   Icon            =   "frmopeninvoicemain.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   11115
   ScaleWidth      =   15240
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
      Left            =   10080
      Style           =   2  'Dropdown List
      TabIndex        =   12
      ToolTipText     =   " Filter The Department "
      Top             =   1080
      Width           =   2295
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
      Left            =   6720
      Style           =   2  'Dropdown List
      TabIndex        =   9
      ToolTipText     =   " Filter The Supplier "
      Top             =   1080
      Width           =   3255
   End
   Begin VB.ComboBox cbosupbill 
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
      Left            =   5040
      Style           =   2  'Dropdown List
      TabIndex        =   8
      ToolTipText     =   " Filter The Supplie Bill No "
      Top             =   1080
      Width           =   1575
   End
   Begin VB.CommandButton cmdlist1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "List All Invoice's"
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
      Left            =   13080
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   " Show All Invoice 's "
      Top             =   360
      Width           =   1935
   End
   Begin VB.ComboBox cboinvoice 
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
      Left            =   1200
      Style           =   2  'Dropdown List
      TabIndex        =   6
      ToolTipText     =   " Filter The Invoice No "
      Top             =   1080
      Width           =   1815
   End
   Begin VB.TextBox txtedit 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   9600
      Visible         =   0   'False
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
      Left            =   10320
      MaskColor       =   &H00C0FFC0&
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "  Exit Window  "
      Top             =   9600
      UseMaskColor    =   -1  'True
      Width           =   1935
   End
   Begin VB.CommandButton cmdadd 
      BackColor       =   &H00FFC0C0&
      Caption         =   "&Add Invoice"
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
      Left            =   3240
      MaskColor       =   &H00C0FFC0&
      Style           =   1  'Graphical
      TabIndex        =   3
      Tag             =   " "
      ToolTipText     =   " To Use Add Invoice "
      Top             =   9600
      UseMaskColor    =   -1  'True
      Width           =   1695
   End
   Begin VB.CommandButton cmdedit 
      BackColor       =   &H00FFC0C0&
      Caption         =   "&View Invoice"
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
      Left            =   5040
      MaskColor       =   &H00C0FFC0&
      Style           =   1  'Graphical
      TabIndex        =   2
      Tag             =   " "
      ToolTipText     =   " To Use View Invoice "
      Top             =   9600
      UseMaskColor    =   -1  'True
      Width           =   1575
   End
   Begin VB.CommandButton cmddelete 
      BackColor       =   &H00FFC0C0&
      Caption         =   "&Delete Invoice"
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
      Left            =   6720
      MaskColor       =   &H00C0FFC0&
      Style           =   1  'Graphical
      TabIndex        =   1
      Tag             =   " "
      ToolTipText     =   " To use Delete Invoice "
      Top             =   9600
      UseMaskColor    =   -1  'True
      Width           =   1695
   End
   Begin VB.CommandButton cmdprint 
      BackColor       =   &H00FFC0C0&
      Caption         =   "&Print Invoice"
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
      Left            =   8520
      MaskColor       =   &H00C0FFC0&
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "  To use Print Invoice "
      Top             =   9600
      UseMaskColor    =   -1  'True
      Width           =   1695
   End
   Begin MSFlexGridLib.MSFlexGrid invoiceMainGrid 
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
      Cols            =   7
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
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Height          =   8175
      Left            =   120
      Top             =   120
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
      Caption         =   "Open Invoice"
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
End
Attribute VB_Name = "frmopeninvoicemain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim op As Variant
Dim cn As ADODB.Connection
Dim rs As ADODB.Recordset
Dim ww As ADODB.Recordset
Private Sub cbodept_Click()
    Call deptfilter
End Sub
Private Sub cbodept_dropdown()
On Error GoTo X
        cbodept.Clear
        Set rs = cn.Execute("select deptname from openinvoicemain_details Group by deptname")
        rs.MoveFirst
        Do While Not rs.EOF()
        cbodept.additem (rs(0))
        rs.MoveNext
        Loop
        cbodept.SetFocus
X:
End Sub
Private Sub cboinvoice_Click()
    Call filterinvoiceno
End Sub

Private Sub cboinvoice_dropdown()
On Error GoTo X
        cboinvoice.Clear
        Set rs = cn.Execute("select invoiceno from openinvoicemain_details order by invoiceno")
        rs.MoveFirst
        Do While Not rs.EOF()
        cboinvoice.additem (rs(0))
        rs.MoveNext
        Loop
        cboinvoice.SetFocus
X:
End Sub
Private Sub cbosup_Click()
    Call supbiilfilter
End Sub

Private Sub cbosup_dropdown()
On Error GoTo X
        cbosup.Clear
        Set rs = cn.Execute("select supname from openinvoicemain_details Group by supname")
        rs.MoveFirst
        Do While Not rs.EOF()
        cbosup.additem (rs(0))
        rs.MoveNext
        Loop
        cbosup.SetFocus
X:
End Sub
Private Sub cbosupbill_Click()
    Call supplierbiilfilter
End Sub

Private Sub cbosupbill_dropdown()
On Error GoTo X
        cbosupbill.Clear
        Set rs = cn.Execute("select supbillno from openinvoicemain_details Group by supbillno")
        rs.MoveFirst
        Do While Not rs.EOF()
        cbosupbill.additem (rs(0))
        rs.MoveNext
        Loop
        cbosupbill.SetFocus
X:
End Sub

Private Sub cmdadd_Click()
    frmopeninvoice.Show
    frmopeninvoice.cmdsave.Visible = True
    frmopeninvoice.cmdupdate.Visible = False
    Unload Me
End Sub

Private Sub cmddelete_Click()
    Dim rs As New ADODB.Recordset
    If Trim(txtedit.Text) = "" Then
        MsgBox "Please Select The Invoice Number", vbCritical, "Selecting Error"
    Else
        rs.Open "Select * from openinvoice_details where invoiceno= '" & txtedit.Text & "'", cn, adOpenKeyset, adLockOptimistic
        If rs.RecordCount <> 0 Then
        If MsgBox("Are You Sure Delete Invoice No  " & txtedit.Text & " ? ", vbQuestion + vbYesNo, "Confirm To Delete") = vbYes Then
        i = 1
        If rs.BOF = False Then rs.MoveFirst
        While rs.EOF = False
        rs.Delete
        rs.MoveNext
        i = i + 1
        Wend
        MsgBox "One Record Deleted Successfully", vbInformation, "Information"
        Unload Me
        frmopeninvoicemain.Show
End If
End If
End If
End Sub
Private Sub cmdedit_Click()
    If Trim(txtedit.Text) = "" Then
        MsgBox "Please Select The Invoice Number", vbCritical, "Invoice NO Error"
        invoiceMainGrid.Col = 1
        invoiceMainGrid.SetFocus
    Else
        frmopeninvoice.Show
        frmopeninvoice.cmdsave.Visible = False
        frmopeninvoice.cmdaddrow.Enabled = False
        frmopeninvoice.cmddeleterow.Enabled = False
        
        
    Dim i As Integer
    Dim rs As New ADODB.Recordset
    rs.Open "Select * from openinvoice_details where invoiceno= '" & txtedit.Text & "'", cn, adOpenKeyset, adLockOptimistic
    i = 1
    If rs.BOF = False Then rs.MoveFirst
    While rs.EOF = False
        frmopeninvoice.InvoiceGrid.Rows = rs.RecordCount + 1
        frmopeninvoice.txtinvoiceno.Text = rs![invoiceno]
        frmopeninvoice.dt1.Value = rs![invoicedate]
        frmopeninvoice.txtsupbillno.Text = rs![supbillno]
        frmopeninvoice.dt2.Value = rs![supbilldate]
        frmopeninvoice.cbosupplier.Text = rs![supname]
        frmopeninvoice.cbodept.Text = rs![deptname]
        frmopeninvoice.InvoiceGrid.TextMatrix(i, 0) = rs![sno]
        frmopeninvoice.InvoiceGrid.TextMatrix(i, 1) = rs![itemname]
        frmopeninvoice.InvoiceGrid.TextMatrix(i, 2) = rs![colour]
        frmopeninvoice.InvoiceGrid.TextMatrix(i, 3) = rs![sizes]
        frmopeninvoice.InvoiceGrid.TextMatrix(i, 4) = Format(rs![invoiceqty], "0.000")
        frmopeninvoice.InvoiceGrid.TextMatrix(i, 5) = rs![uom]
        frmopeninvoice.InvoiceGrid.TextMatrix(i, 6) = Format(rs![invoicerate], "0.000")
        frmopeninvoice.InvoiceGrid.TextMatrix(i, 7) = Format(rs![totalamtgrid], "0.00")
        frmopeninvoice.txttot.Text = rs![totalamt]
        frmopeninvoice.txttax.Text = rs![taxes]
        frmopeninvoice.taxamt.Text = rs![taxamt]
        frmopeninvoice.txtnet.Text = rs![netamt]
        frmopeninvoice.tw.Text = rs![words]
        frmopeninvoice.txtremarks.Text = rs![remarks]
        rs.MoveNext
        i = i + 1
        Wend
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
    Call openinvoicemaingridload
    cboinvoice.Clear
    cbosup.Clear
    cbosupbill.Clear
    cbodept.Clear
    cboinvoice.Visible = False
    cbosup.Visible = False
    cbodept.Visible = False
    cbosupbill.Visible = False
End Sub
Private Sub cmdprint_Click()
    Dim rscomp As New ADODB.Recordset
    Dim rsgrn As New ADODB.Recordset
    Dim cn As New ADODB.Connection
    Dim rssup As New ADODB.Recordset
    Dim rsItem As New ADODB.Recordset
    Dim irc As Integer
    Dim filenum As Integer
    Dim rssups As New ADODB.Recordset
    
    cn.Open "Provider=Microsoft.jet.oledb.4.0; Data Source=" & App.path & "\Database\Data.mdb"
    rsgrn.Open "SELECT * FROM openinvoicemain_details where invoiceno ='" & txtedit.Text & "'", cn, 1, 3
    rssups.Open "SELECT * FROM openinvoice_details where invoiceno ='" & txtedit.Text & "'", cn, 1, 3
    rssup.Open "select * from sup_details where supname ='" & invoiceMainGrid.TextMatrix(invoiceMainGrid.Row, 4) & "'", cn, 1, 3
    rscomp.Open "select * from com_details ", cn, 1, 3
    
    If Trim(txtedit.Text) = "" Then
    MsgBox "Please Select The Invoice No ", vbCritical, "Invoice No Selection Error"
    txtedit.SetFocus
    Else
    
    ChDrive App.path
    ChDir App.path
    
    Kill App.path & "\Reports\Printopeninvoice.txt"
    Open App.path & "\Reports\Printopeninvoice.txt" For Output As #1
   
    Print #1,
    Print #1, Tab(42); "Open Invoice";
    Print #1,
    Print #1, String(100, "-");
    Print #1, Tab(2); "Invoice No: "; rsgrn![invoiceno]; Tab(65); "Invoice Date :"; rsgrn![invoicedate];
    Print #1, Tab(2); "Sup.Bill.No:"; rssups![supbillno]; Tab(65); "Sup.Bill.Date :"; rssups![supbilldate];
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
    Print #1, Tab(2); "S.No"; Tab(8); "Item Name"; Tab(25); "Colour"; Tab(40); "Size"; Tab(55); "Inv.Qty", Tab(72); "UOM"; Tab(82); "Inv.Rate"; Tab(92); "Tot.Amt(Rs)";
    
    Print #1,
    Print #1, String(100, "-");
    rsItem.Open "Select * from openinvoice_details where invoiceno ='" & txtedit.Text & "'", cn, 1, 3
    i = 1
    If rsItem.BOF = False Then rsItem.MoveFirst
    While rsItem.EOF = False
    Print #1,
    Print #1, rsItem![sno]; Tab(8); rsItem![itemname]; Tab(25); rsItem![colour]; Tab(40); rsItem![sizes]; Tab(55); Format(rsItem![invoiceqty], "0.000"); Tab(72); rsItem![uom]; Tab(82); Format(rsItem![invoicerate], "0.00"); ; Tab(92); Format(rsItem![totalamtgrid], "0.00");
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
    frmopeninvoicereport.Show
    End If
End Sub

Private Sub Form_Load()

    Set cn = New ADODB.Connection
    Set ww = New ADODB.Recordset
    Set rs = New ADODB.Recordset
    cn.Open "Provider=Microsoft.jet.oledb.4.0; Data Source=" & App.path & "\Database\Data.mdb"
    ww.Open "Select * From openinvoicemain_details", cn, adOpenKeyset, adLockOptimistic
    
    Call frmopeninvoicemaingridload
    Call openinvoicemaingridload
    Call clears
End Sub
Private Function openinvoicemaingridload()
    Dim i As Integer
    Dim rs As New ADODB.Recordset
    rs.Open "Select * from openinvoicemain_details ", cn, adOpenKeyset, adLockOptimistic
    i = 1
    If rs.RecordCount = 0 Then
    MsgBox "No Record Found", vbInformation, "Information"
    Else
    If rs.BOF = False Then rs.MoveFirst
    While rs.EOF = False
    invoiceMainGrid.Rows = invoiceMainGrid.Rows + 1
    invoiceMainGrid.TextMatrix(i, 0) = i
    invoiceMainGrid.TextMatrix(i, 1) = rs![invoiceno]
    invoiceMainGrid.TextMatrix(i, 2) = rs![invoicedate]
    invoiceMainGrid.TextMatrix(i, 3) = rs![supbillno]
    invoiceMainGrid.TextMatrix(i, 6) = Format(rs![netamt], "0.00")
    invoiceMainGrid.TextMatrix(i, 5) = rs![deptname]
    invoiceMainGrid.TextMatrix(i, 4) = rs![supname]
    rs.MoveNext
    i = i + 1
    Wend
    invoiceMainGrid.Rows = rs.RecordCount + 1
    End If
End Function
Private Function filterinvoiceno()
    Dim i As Integer
    Dim rs As New ADODB.Recordset
    rs.Open "Select * from openinvoicemain_details where invoiceno='" & cboinvoice.Text & "' ", cn, adOpenKeyset, adLockOptimistic
    i = 1
    invoiceMainGrid.TextMatrix(i, 0) = i
    invoiceMainGrid.TextMatrix(i, 1) = rs![invoiceno]
    invoiceMainGrid.TextMatrix(i, 2) = rs![invoicedate]
    invoiceMainGrid.TextMatrix(i, 3) = rs![supbillno]
    invoiceMainGrid.TextMatrix(i, 6) = Format(rs![netamt], "0.00")
    invoiceMainGrid.TextMatrix(i, 5) = rs![deptname]
    invoiceMainGrid.TextMatrix(i, 4) = rs![supname]
    invoiceMainGrid.Rows = rs.RecordCount + 1
End Function
Private Function supplierbiilfilter()
  Dim rs As New ADODB.Recordset
    rs.Open "select * from openinvoicemain_details where supbillno= '" & Trim(cbosupbill.Text) & "'", cn, adOpenKeyset, adLockOptimistic
    i = 1
    If rs.BOF = False Then rs.MoveFirst
    While rs.EOF = False
    invoiceMainGrid.Rows = invoiceMainGrid.Rows + 1
    invoiceMainGrid.TextMatrix(i, 0) = i
    invoiceMainGrid.TextMatrix(i, 1) = rs![invoiceno]
    invoiceMainGrid.TextMatrix(i, 2) = rs![invoicedate]
    invoiceMainGrid.TextMatrix(i, 3) = rs![supbillno]
    invoiceMainGrid.TextMatrix(i, 4) = rs![supname]
    invoiceMainGrid.TextMatrix(i, 5) = rs![deptname]
    invoiceMainGrid.TextMatrix(i, 6) = Format(rs![netamt], "0.00")
    rs.MoveNext
    i = i + 1
    Wend
    invoiceMainGrid.Rows = rs.RecordCount + 1
End Function
Private Function supbiilfilter()
  Dim rs As New ADODB.Recordset
    rs.Open "select * from openinvoicemain_details where supname= '" & Trim(cbosup.Text) & "'", cn, adOpenKeyset, adLockOptimistic
    i = 1
    If rs.BOF = False Then rs.MoveFirst
    While rs.EOF = False
    invoiceMainGrid.Rows = invoiceMainGrid.Rows + 1
    invoiceMainGrid.TextMatrix(i, 0) = i
    invoiceMainGrid.TextMatrix(i, 1) = rs![invoiceno]
    invoiceMainGrid.TextMatrix(i, 2) = rs![invoicedate]
    invoiceMainGrid.TextMatrix(i, 3) = rs![supbillno]
    invoiceMainGrid.TextMatrix(i, 4) = rs![supname]
    invoiceMainGrid.TextMatrix(i, 5) = rs![deptname]
    invoiceMainGrid.TextMatrix(i, 6) = Format(rs![netamt], "0.00")
    rs.MoveNext
    i = i + 1
    Wend
    invoiceMainGrid.Rows = rs.RecordCount + 1
End Function
Private Function deptfilter()
  Dim rs As New ADODB.Recordset
    rs.Open "select * from openinvoicemain_details where deptname= '" & Trim(cbodept.Text) & "'", cn, adOpenKeyset, adLockOptimistic
    i = 1
    If rs.BOF = False Then rs.MoveFirst
    While rs.EOF = False
    invoiceMainGrid.Rows = invoiceMainGrid.Rows + 1
    invoiceMainGrid.TextMatrix(i, 0) = i
    invoiceMainGrid.TextMatrix(i, 1) = rs![invoiceno]
    invoiceMainGrid.TextMatrix(i, 2) = rs![invoicedate]
    invoiceMainGrid.TextMatrix(i, 3) = rs![supbillno]
    invoiceMainGrid.TextMatrix(i, 4) = rs![supname]
    invoiceMainGrid.TextMatrix(i, 5) = rs![deptname]
    invoiceMainGrid.TextMatrix(i, 6) = Format(rs![netamt], "0.00")
    rs.MoveNext
    i = i + 1
    Wend
    invoiceMainGrid.Rows = rs.RecordCount + 1
End Function
Private Sub invoiceMainGrid_Click()
    If invoiceMainGrid.Col = 0 Or invoiceMainGrid.Col = 1 Or invoiceMainGrid.Col = 2 Or invoiceMainGrid.Col = 3 Or invoiceMainGrid.Col = 4 Or invoiceMainGrid.Col = 5 Or invoiceMainGrid.Col = 6 Then
         txtedit.Text = invoiceMainGrid.TextMatrix(invoiceMainGrid.Row, 1)
    End If
        If invoiceMainGrid.Col = 1 And invoiceMainGrid.Row = 1 Then
        cboinvoice.Visible = True
        Else
        cboinvoice.Visible = False
        End If
        If invoiceMainGrid.Col = 4 And invoiceMainGrid.Row = 1 Then
        cbosup.Visible = True
        Else
        cbosup.Visible = False
        End If
        If invoiceMainGrid.Col = 3 And invoiceMainGrid.Row = 1 Then
        cbosupbill.Visible = True
        Else
        cbosupbill.Visible = False
        End If
        If invoiceMainGrid.Col = 5 And invoiceMainGrid.Row = 1 Then
        cbodept.Visible = True
        Else
        cbodept.Visible = False
        End If
End Sub
Private Function clears()
    cboinvoice.Clear
    cbosup.Clear
    cbosupbill.Clear
    cbodept.Clear
    cboinvoice.Visible = False
    cbosup.Visible = False
    cbodept.Visible = False
    cbosupbill.Visible = False
End Function
