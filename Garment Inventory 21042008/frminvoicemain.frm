VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frminvoicemain 
   BackColor       =   &H00EDDDD1&
   Caption         =   " * Invoice Details *"
   ClientHeight    =   9165
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10380
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   11115
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
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
      Left            =   12720
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   " Show All Invoice 's "
      Top             =   360
      Width           =   2175
   End
   Begin VB.ComboBox cbosupbillfilter 
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
      TabIndex        =   9
      ToolTipText     =   " To Use Filter The Department Wise "
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
      Left            =   7320
      Style           =   2  'Dropdown List
      TabIndex        =   8
      ToolTipText     =   " To Use Filter The Supplier Name Wise "
      Top             =   1080
      Width           =   4335
   End
   Begin VB.ComboBox cboinvoiceFilter 
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
      Left            =   10920
      MaskColor       =   &H00C0FFC0&
      Style           =   1  'Graphical
      TabIndex        =   5
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
      Left            =   3600
      MaskColor       =   &H00C0FFC0&
      Style           =   1  'Graphical
      TabIndex        =   4
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
      Left            =   5400
      MaskColor       =   &H00C0FFC0&
      Style           =   1  'Graphical
      TabIndex        =   3
      Tag             =   " "
      ToolTipText     =   " To Use View Invoice "
      Top             =   9600
      UseMaskColor    =   -1  'True
      Width           =   1695
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
      Left            =   7200
      MaskColor       =   &H00C0FFC0&
      Style           =   1  'Graphical
      TabIndex        =   2
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
      Left            =   9000
      MaskColor       =   &H00C0FFC0&
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "  To use Print Invoice "
      Top             =   9600
      UseMaskColor    =   -1  'True
      Width           =   1815
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
   Begin MSFlexGridLib.MSFlexGrid invoiceMainGrid 
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
      Caption         =   " Invoice Details"
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
      Width           =   7335
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
Attribute VB_Name = "frminvoicemain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Dim op As Variant
    Dim cn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim ww As ADODB.Recordset
    Dim i As Integer
Private Sub cmdadd_Click()
    frminvoice.Show
    Unload Me
End Sub
Private Function grnmainload()
    Dim i As Integer
    Dim rs As New ADODB.Recordset
    rs.Open "Select * from invoicemain_details ", cn, adOpenKeyset, adLockOptimistic
    i = 1
    If rs.BOF = False Then rs.MoveFirst
    While rs.EOF = False
    invoiceMainGrid.Rows = invoiceMainGrid.Rows + 1
    invoiceMainGrid.TextMatrix(i, 0) = i
    invoiceMainGrid.TextMatrix(i, 1) = rs![invoiceno]
    invoiceMainGrid.TextMatrix(i, 2) = rs![invoicedate]
    invoiceMainGrid.TextMatrix(i, 4) = rs![supname]
    invoiceMainGrid.TextMatrix(i, 3) = rs![supbillno]
    invoiceMainGrid.TextMatrix(i, 5) = Format(rs![netamt], "0.00")
    rs.MoveNext
    i = i + 1
    Wend
    invoiceMainGrid.Rows = rs.RecordCount + 1
End Function
Private Sub cbodept_Click()
    If InvoiceMiainGrid.Rows = 1 Then
        MsgBox "No Record Found ", vbInformation, "Information"
    End If
End Sub
Private Sub cbosupbillfilter_dropdown()
Dim rs As New ADODB.Recordset
        On Error GoTo X
        cbosupbillfilter.Clear
        Set rs = cn.Execute("select supbillno from invoicemain_details Group by supbillno")
        rs.MoveFirst
        Do While Not rs.EOF()
        cbosupbillfilter.additem (rs(0))
        rs.MoveNext
        Loop
        cbosupbillfilter.SetFocus
X:
End Sub
Private Sub cboinvoicefilter_Click()
    If invoiceMainGrid.Rows = 1 Then
    MsgBox "No Record Found ", vbInformation, "Information"
    End If
    Call filterinvoiceno
End Sub
Private Sub cboinvoicefilter_DropDown()
Dim rs As New ADODB.Recordset
        On Error GoTo X
        cboinvoiceFilter.Clear
        Set rs = cn.Execute("select invoiceno from invoicemain_details order by invoiceno")
        rs.MoveFirst
        Do While Not rs.EOF()
        cboinvoiceFilter.additem (rs(0))
        rs.MoveNext
        Loop
        cboinvoiceFilter.SetFocus
X:
End Sub
Private Sub cbosupfilter_Click()
    If invoiceMainGrid.Rows = 1 Then
    MsgBox "No Record Found ", vbInformation, "Information"
    End If
    Call filtersupname
End Sub
Private Sub cbosupfilter_dropdown()
        Dim rs As New ADODB.Recordset
        On Error GoTo X
        cbosupfilter.Clear
        Set rs = cn.Execute("select supname from grnstatus_details where grnstatus = 'Open' Group by supname")
        rs.MoveFirst
        Do While Not rs.EOF()
        cbosupfilter.additem (rs(0))
        rs.MoveNext
        Loop
        cbosupfilter.SetFocus
X:
End Sub
Private Sub cmddelete_Click()
    Dim rs As New ADODB.Recordset
    Dim rs1 As New ADODB.Recordset
    Dim rs2 As New ADODB.Recordset
    If Trim(txtedit.Text) = "" Then
        MsgBox "Please Select The Invoice  Number", vbCritical, "Selecting Error"
    Else
        rs.Open "Select * from invoice_details where invoiceno= '" & txtedit.Text & "'", cn, adOpenKeyset, adLockOptimistic
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
        frminvoicemain.Show
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
        frminvoice.Show
        frminvoice.txtinvoiceno.Enabled = False
        frminvoice.txtsupbillno.Enabled = False
        frminvoice.cmdsave.Enabled = False
        frminvoice.cmdadditem.Visible = False
        frminvoice.cmddeleteitem.Visible = False
        frminvoice.InvoiceGrid.Enabled = False
        frminvoice.cbogrnno.Enabled = False
        frminvoice.cbosup.Enabled = False
        frminvoice.InvoiceGrid.Height = 3550
        frminvoice.InvoiceDetailsGrid.Visible = False
        frminvoice.invoiceMainGrid.Visible = False
        frminvoice.Shape2.Visible = False
        frminvoice.Shape3.Visible = False
        frminvoice.txttotamt.Top = 5400
        frminvoice.txttotamt.Left = 12600
        frminvoice.Label2(1).Top = 5022
        frminvoice.Label2(1).Left = 12600
        frminvoice.Label2(4).Top = 5022
        frminvoice.txtremarks.Top = 5022
        frminvoice.txtremarks.Width = 11000
        frminvoice.txttaxamt.Locked = True
        
    Dim i As Integer
    Dim rs As New ADODB.Recordset
    rs.Open "Select * from invoice_details where invoiceno= '" & frminvoicemain.txtedit.Text & "'", cn, adOpenKeyset, adLockOptimistic
    i = 1
    If rs.BOF = False Then rs.MoveFirst
    While rs.EOF = False
        frminvoice.InvoiceGrid.Rows = rs.RecordCount + 1
        frminvoice.txtinvoiceno.Text = rs![invoiceno]
        frminvoice.dt1.Value = rs![invoicedate]
        frminvoice.txtsupbillno.Text = rs![supbillno]
        frminvoice.dt2.Value = rs![supbilldate]
        frminvoice.cbosup.Text = rs![supname]
        frminvoice.cbogrnno.Text = rs![grnno]
        frminvoice.InvoiceGrid.TextMatrix(i, 0) = rs![sno]
        frminvoice.InvoiceGrid.TextMatrix(i, 1) = rs![itemname]
        frminvoice.InvoiceGrid.TextMatrix(i, 2) = rs![colour]
        frminvoice.InvoiceGrid.TextMatrix(i, 3) = rs![sizes]
        frminvoice.InvoiceGrid.TextMatrix(i, 4) = Format(rs![grnqty], "0.000")
        frminvoice.InvoiceGrid.TextMatrix(i, 5) = rs![uom]
        frminvoice.InvoiceGrid.TextMatrix(i, 6) = Format(rs![invoiceqty], "0.000")
        frminvoice.InvoiceGrid.TextMatrix(i, 7) = Format(rs![invoicerate], "0.00")
        frminvoice.InvoiceGrid.TextMatrix(i, 8) = Format(rs![totalamtgrid], "0.000")
        frminvoice.txtremarks.Text = rs![remarks]
        frminvoice.txttotamt.Text = rs![totalamt]
        frminvoice.txtword.Text = rs![words]
        frminvoice.txttaxamt.Text = rs![taxamt]
        
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
    Call allgrns
    Call nonvisibles
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
    rsgrn.Open "SELECT * FROM invoicemain_details where invoiceno ='" & txtedit.Text & "'", cn, 1, 3
    rssup.Open "select * from sup_details where supname ='" & invoiceMainGrid.TextMatrix(invoiceMainGrid.Row, 4) & "'", cn, 1, 3
    rscomp.Open "select * from com_details ", cn, 1, 3
    
    If Trim(txtedit.Text) = "" Then
    MsgBox "Please Select The Invoice No ", vbCritical, "Invoice No Selection Error"
    txtedit.SetFocus
    Else
    
    ChDrive App.path
    ChDir App.path
    
    Kill App.path & "\Reports\Printinvoice.txt"
    Open App.path & "\Reports\Printinvoice.txt" For Output As #1
   
    Print #1,
    Print #1, Tab(42); "Invoice";
    Print #1,
    Print #1, String(100, "-");
    Print #1, Tab(2); "Invoice No: "; rsgrn![invoiceno]; Tab(65); "Invoice Date :"; rsgrn![invoicedate];
    Print #1, Tab(2); "Sup.Bill.No:"; rsgrn![supbillno]; Tab(65); "Sup.Bill.Date :"; rsgrn![supbilldate];
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
    rsItem.Open "Select * from invoice_details where invoiceno ='" & txtedit.Text & "'", cn, 1, 3
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
    frminvoicereport.Show
    End If
End Sub
Private Sub Form_Load()

    Set cn = New ADODB.Connection
    Set ww = New ADODB.Recordset
    
    cn.Open "Provider=Microsoft.jet.oledb.4.0; Data Source=" & App.path & "\Database\Data.mdb"
    cn.CursorLocation = adUseClient
    
    Call grnmainload
    Call frminvoicemaingridload
    
    cboinvoiceFilter.Visible = False
    cbosupfilter.Visible = False
    cbosupbillfilter.Visible = False
        
End Sub
Private Sub invoiceMainGrid_Click()
    If invoiceMainGrid.Col = 0 Or invoiceMainGrid.Col = 1 Or invoiceMainGrid.Col = 2 Or invoiceMainGrid.Col = 3 Or invoiceMainGrid.Col = 4 Or invoiceMainGrid.Col = 5 Or invoiceMainGrid.Col = 6 Then
         txtedit.Text = invoiceMainGrid.TextMatrix(invoiceMainGrid.Row, 1)
    End If
        If invoiceMainGrid.Col = 1 And invoiceMainGrid.Row = 1 Then
        cboinvoiceFilter.Visible = True
        Else
        cboinvoiceFilter.Visible = False
        End If
        If invoiceMainGrid.Col = 4 And invoiceMainGrid.Row = 1 Then
        cbosupfilter.Visible = True
        Else
        cbosupfilter.Visible = False
        End If
        If invoiceMainGrid.Col = 3 And invoiceMainGrid.Row = 1 Then
        cbosupbillfilter.Visible = True
        Else
        cbosupbillfilter.Visible = False
        End If
End Sub
Sub filterinvoiceno()
    Dim i As Integer
    Dim rs As New ADODB.Recordset
    rs.Open "Select * from invoicemain_details where invoiceno='" & cboinvoiceFilter.Text & "'", cn, adOpenKeyset, adLockOptimistic
         i = 1
                invoiceMainGrid.TextMatrix(i, 1) = rs![invoiceno]
                invoiceMainGrid.TextMatrix(i, 2) = rs![invoicedate]
                invoiceMainGrid.TextMatrix(i, 3) = rs![supbillno]
                invoiceMainGrid.TextMatrix(i, 4) = rs![supname]
                invoiceMainGrid.TextMatrix(i, 5) = Format(rs![netamt], "0.00")
                invoiceMainGrid.Rows = rs.RecordCount + 1
End Sub
Sub filtersupname()
    Dim rs As New ADODB.Recordset
    rs.Open "select * from invoicemain_details where supname= '" & Trim(cbosupfilter.Text) & "'", cn, adOpenKeyset, adLockOptimistic
    i = 1
    If rs.BOF = False Then rs.MoveFirst
    While rs.EOF = False
    invoiceMainGrid.Rows = invoiceMainGrid.Rows + 1
    invoiceMainGrid.TextMatrix(i, 0) = i
    invoiceMainGrid.TextMatrix(i, 1) = rs![invoiceno]
    invoiceMainGrid.TextMatrix(i, 2) = rs![invoicedate]
    invoiceMainGrid.TextMatrix(i, 3) = rs![supbillno]
    invoiceMainGrid.TextMatrix(i, 4) = rs![supname]
    invoiceMainGrid.TextMatrix(i, 5) = Format(rs![netamt], "0.00")
    rs.MoveNext
    i = i + 1
    Wend
    invoiceMainGrid.Rows = rs.RecordCount + 1
End Sub
Sub filtersupbill()
    Dim rs As New ADODB.Recordset
    rs.Open "select * from invoicemain_details where supbillno= '" & Trim(cbosupbillfilter.Text) & "'", cn, adOpenKeyset, adLockOptimistic
    i = 1
    If rs.BOF = False Then rs.MoveFirst
    While rs.EOF = False
    invoiceMainGrid.Rows = invoiceMainGrid.Rows + 1
    invoiceMainGrid.TextMatrix(i, 0) = i
    invoiceMainGrid.TextMatrix(i, 1) = rs![invoiceno]
    invoiceMainGrid.TextMatrix(i, 2) = rs![invoicedate]
    invoiceMainGrid.TextMatrix(i, 3) = rs![supbillno]
    invoiceMainGrid.TextMatrix(i, 4) = rs![supname]
    invoiceMainGrid.TextMatrix(i, 5) = Format(rs![netamt], "0.00")
    rs.MoveNext
    i = i + 1
    Wend
    invoiceMainGrid.Rows = rs.RecordCount + 1
End Sub
Sub allgrns()
    Dim i As Integer
    Dim rs As New ADODB.Recordset
    rs.Open "Select * from invoicemain_details ", cn, adOpenKeyset, adLockOptimistic
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
        invoiceMainGrid.TextMatrix(i, 4) = rs![supname]
        invoiceMainGrid.TextMatrix(i, 5) = Format(rs![netamt], "0.00")
        rs.MoveNext
        i = i + 1
        Wend
        invoiceMainGrid.Rows = rs.RecordCount + 1
    End If
End Sub
Sub visibles()
    cboinvoiceFilter.Clear
    cbosupfilter.Clear
    cbosupbillfilter.Clear
End Sub
Sub nonvisibles()
    cboinvoiceFilter.Visible = False
    cbosupfilter.Visible = False
    cbosupbillfilter.Visible = False
End Sub

