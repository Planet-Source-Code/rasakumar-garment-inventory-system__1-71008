VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmpaymentmain 
   BackColor       =   &H00EDDDD1&
   Caption         =   " * Payment Details * "
   ClientHeight    =   9105
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10155
   Icon            =   "frmpaymentmain.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   11115
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.ComboBox cboinvoice 
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
      Left            =   4440
      Style           =   2  'Dropdown List
      TabIndex        =   13
      ToolTipText     =   " Filter The Invoice No "
      Top             =   1080
      Width           =   1575
   End
   Begin VB.TextBox txttypeno 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   2040
      TabIndex        =   10
      Top             =   8760
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton cmdprint 
      BackColor       =   &H00FFC0C0&
      Caption         =   "&Print Payment"
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
      ToolTipText     =   "  To use Print Payment  "
      Top             =   9600
      UseMaskColor    =   -1  'True
      Width           =   1815
   End
   Begin VB.CommandButton cmddelete 
      BackColor       =   &H00FFC0C0&
      Caption         =   "&Delete Payment"
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
      TabIndex        =   8
      Tag             =   " "
      ToolTipText     =   " To use Delete Payment  "
      Top             =   9600
      UseMaskColor    =   -1  'True
      Width           =   1695
   End
   Begin VB.CommandButton cmdedit 
      BackColor       =   &H00FFC0C0&
      Caption         =   "&View Payment"
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
      TabIndex        =   7
      Tag             =   " "
      ToolTipText     =   " To Use View Payment "
      Top             =   9600
      UseMaskColor    =   -1  'True
      Width           =   1695
   End
   Begin VB.CommandButton cmdadd 
      BackColor       =   &H00FFC0C0&
      Caption         =   "&Add Payment"
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
      TabIndex        =   6
      Tag             =   " "
      ToolTipText     =   " To Use Add Payment "
      Top             =   9600
      UseMaskColor    =   -1  'True
      Width           =   1695
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
   Begin VB.ComboBox cbopayFilter 
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
      ToolTipText     =   " To Use Filter The Pay No "
      Top             =   1080
      Width           =   1575
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
      Left            =   8160
      Style           =   2  'Dropdown List
      TabIndex        =   2
      ToolTipText     =   " To Use Filter The Supplier Name Wise "
      Top             =   1080
      Width           =   3375
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
      Left            =   6120
      Style           =   2  'Dropdown List
      TabIndex        =   1
      ToolTipText     =   " To Use Filter Supplier Bill No Wise "
      Top             =   1080
      Width           =   1935
   End
   Begin VB.CommandButton cmdlist1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "List All Payment's"
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
      TabIndex        =   0
      ToolTipText     =   " Show All Payment's"
      Top             =   360
      Width           =   2175
   End
   Begin MSFlexGridLib.MSFlexGrid PaymentMainGrid 
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
      TabIndex        =   14
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
      Caption         =   " Payment Details"
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
End
Attribute VB_Name = "frmpaymentmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim op As Variant
Dim cn As ADODB.Connection
Dim rs As ADODB.Recordset
Dim ww As ADODB.Recordset
Private Sub cboinvoice_Click()
    Call invoicefilterload
End Sub

Private Sub cboinvoice_dropdown()
Dim rs As New ADODB.Recordset
        On Error GoTo X
        cboinvoice.Clear
        Set rs = cn.Execute("select invoiceno from paymentmain_details Order by invoiceno")
        rs.MoveFirst
        Do While Not rs.EOF()
        cboinvoice.additem (rs(0))
        rs.MoveNext
        Loop
        cboinvoice.SetFocus
X:
End Sub
Private Sub cbopayFilter_Click()
    Call paynofilterload
End Sub
Private Sub cbopayFilter_dropdown()
Dim rs As New ADODB.Recordset
        On Error GoTo X
        cbopayFilter.Clear
        Set rs = cn.Execute("select payno from paymentmain_details Order by payno")
        rs.MoveFirst
        Do While Not rs.EOF()
        cbopayFilter.additem (rs(0))
        rs.MoveNext
        Loop
        cbopayFilter.SetFocus
X:
End Sub
Private Sub cbosupbillfilter_Click()
    Call supbillfilterload
End Sub

Private Sub cbosupbillfilter_dropdown()
Dim rs As New ADODB.Recordset
        On Error GoTo X
        cbosupbillfilter.Clear
        Set rs = cn.Execute("select supbillno from paymentmain_details Group by supbillno")
        rs.MoveFirst
        Do While Not rs.EOF()
        cbosupbillfilter.additem (rs(0))
        rs.MoveNext
        Loop
        cbosupbillfilter.SetFocus
X:
End Sub
Private Sub cbosupfilter_Click()
    Call supnamefilterload
End Sub
Private Sub cbosupfilter_dropdown()
Dim rs As New ADODB.Recordset
        On Error GoTo X
        cbosupfilter.Clear
        Set rs = cn.Execute("select supname from paymentmain_details Group by supname")
        rs.MoveFirst
        Do While Not rs.EOF()
        cbosupfilter.additem (rs(0))
        rs.MoveNext
        Loop
        cbosupfilter.SetFocus
X:
End Sub
Private Sub cmdadd_Click()
    frmpayment.Show
    Unload Me
End Sub
Private Sub cmddelete_Click()
    Dim rs As New ADODB.Recordset
    Dim rs1 As New ADODB.Recordset
    Dim rs2 As New ADODB.Recordset
    
    If Trim(txtedit.Text) = "" Then
        MsgBox "Please Select The Payment Number", vbCritical, "Selecting Error"
    Else
        rs1.Open "Select * from payment_details where payno ='" & txtedit.Text & "'", cn, adOpenKeyset, adLockOptimistic
        rs2.Open "Select * from debit_details where payno ='" & txtedit.Text & "'", cn, adOpenKeyset, adLockOptimistic
    If rs1.RecordCount <> 0 Then
    If MsgBox("Are You Sure Delete PO No  " & txtedit.Text & " ? ", vbQuestion + vbYesNo, "Confirm To Delete") = vbYes Then
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
    frmpaymentmain.Show
End If
End If
End If

End Sub

Private Sub cmdedit_Click()
    If Trim(txtedit.Text) = "" Then
        MsgBox "Please Select The Payment Number", vbCritical, "Invoice NO Error"
        PaymentMainGrid.Col = 1
        PaymentMainGrid.SetFocus
    Else
        frmpayment.cmdsave.Enabled = False
        frmpayment.cmdadddebt.Enabled = False
        frmpayment.cmddeletedbt.Enabled = False
        frmpayment.cmdsave.Enabled = False
        frmpayment.cboinvoiceno.Enabled = False
        frmpayment.cbosup.Enabled = False
        frmpayment.PaymentGrid.Enabled = False
        frmpayment.DebtGrid.Enabled = False
        frmpayment.txtremarks.Enabled = False
        
    
    Dim i As Integer
    Dim rs As New ADODB.Recordset
    Dim rs1 As New ADODB.Recordset
    rs.Open "Select * from payment_details where payno= '" & txtedit.Text & "'", cn, adOpenKeyset, adLockOptimistic
    rs1.Open "Select * from debit_details where payno= '" & txtedit.Text & "'", cn, adOpenKeyset, adLockOptimistic
    i = 1
    If rs.BOF = False Then rs.MoveFirst
    While rs.EOF = False
        frmpayment.PaymentGrid.Rows = rs.RecordCount + 1
        frmpayment.txtpayno.Text = rs![payno]
        frmpayment.dt1.Value = rs![paydate]
        frmpayment.cbosup.Text = rs![supname]
        frmpayment.cboinvoiceno.Text = rs![invoiceno]
        frmpayment.PaymentGrid.TextMatrix(i, 0) = rs![sno]
        frmpayment.PaymentGrid.TextMatrix(i, 1) = rs![invoiceno]
        frmpayment.PaymentGrid.TextMatrix(i, 2) = rs![invoicedate]
        frmpayment.PaymentGrid.TextMatrix(i, 3) = rs![supbillno]
        frmpayment.PaymentGrid.TextMatrix(i, 4) = Format(rs![invoiceamt], "0.00")
        frmpayment.PaymentGrid.TextMatrix(i, 5) = Format(rs![payamtgrid], "0.00")
        frmpayment.txtremarks.Text = rs![remarks]
        frmpayment.txttot.Text = Format(rs![totpayamt], "0.00")
        frmpayment.txtword.Text = rs![words]
        frmpayment.txtnetpay.Text = Format(rs![paynetamt], "0.00")
        rs.MoveNext
        i = i + 1
        Wend
        j = 1
        If rs1.BOF = False Then rs1.MoveFirst
        While rs1.EOF = False
        frmpayment.DebtGrid.Rows = rs1.RecordCount + 1
        frmpayment.DebtGrid.TextMatrix(j, 0) = rs1![sno]
        frmpayment.DebtGrid.TextMatrix(j, 1) = rs1![debtreason]
        frmpayment.DebtGrid.TextMatrix(j, 2) = Format(rs1![debtamt], "0.00")
        frmpayment.txtdebt.Text = Format(rs1![debttotamt], "0.00")
        rs1.MoveNext
        j = j + 1
        Wend
  Unload Me
  End If
End Sub

Private Sub cmdexit_Click()
    op = MsgBox("Are You Sure To Close ?", vbQuestion + vbYesNo, "Confirm To Close ?")
        If op = vbYes Then
            Unload Me
        Else
        End If
End Sub
Private Sub cmdlist1_Click()
    Call visi
    Call paymaingridlaod
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
    rsgrn.Open "SELECT * FROM paymentmain_details where payno ='" & txtedit.Text & "'", cn, 1, 3
    rssup.Open "select * from sup_details where supname ='" & PaymentMainGrid.TextMatrix(PaymentMainGrid.Row, 5) & "'", cn, 1, 3
    rscomp.Open "select * from com_details ", cn, 1, 3
    
    If Trim(txtedit.Text) = "" Then
    MsgBox "Please Select The Invoice No ", vbCritical, "Invoice No Selection Error"
    txtedit.SetFocus
    Else
    
    ChDrive App.path
    ChDir App.path
    
    Kill App.path & "\Reports\Printpayment.txt"
    Open App.path & "\Reports\Printpayment.txt" For Output As #1
   
    Print #1,
    Print #1, Tab(42); "Payment";
    Print #1,
    Print #1, String(100, "-");
    Print #1, Tab(2); "Payment No: "; rsgrn![payno]; Tab(65); "Payment Date :"; rsgrn![paydate];
    Print #1, Tab(2); "Sup.Bill.No:"; rsgrn![supbillno];
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
    Print #1, Tab(2); "Invoice No"; Tab(25); "Invoice Date"; Tab(40); "Invoice Amt"; Tab(60); "Pay Net Amt"; Tab(80); "Pay Net Amt";
    
    Print #1,
    Print #1, String(100, "-");
    rsItem.Open "Select * from payment_details where invoiceno ='" & txtedit.Text & "'", cn, 1, 3
    i = 1
    If rsItem.BOF = False Then rsItem.MoveFirst
    While rsItem.EOF = False
    Print #1,
    Print #1, Tab(2); rsItem![invoiceno]; Tab(25); rsItem![invoicedate]; Tab(40); Format(rsItem![invoiceamt], "0.00"); Tab(60); Format(rsItem![paynetamt], "0.00");
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
    frmpaymentreport.Show
    End If
End Sub

Private Sub Form_Load()
        
    Set cn = New ADODB.Connection
    Set ww = New ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    cn.Open "Provider=Microsoft.jet.oledb.4.0; Data Source=" & App.path & "\Database\Data.mdb"
    ww.Open "Select * From paymentmain_details", cn, adOpenKeyset, adLockOptimistic
    cn.CursorLocation = adUseClient
    
    Call frmpaymentmaingridload
    Call paymaingridlaod
    Call visi
End Sub
Private Function paymaingridlaod()
    Dim i As Integer
    Dim rs As New ADODB.Recordset
    rs.Open "Select * from paymentmain_details", cn, adOpenKeyset, adLockOptimistic
    If rs.RecordCount = 0 Then
    MsgBox "No Record Found", vbInformation, "Information"
    Else
    i = 1
    If rs.BOF = False Then rs.MoveFirst
    While rs.EOF = False
    PaymentMainGrid.Rows = PaymentMainGrid.Rows + 1
    PaymentMainGrid.TextMatrix(i, 0) = i
    PaymentMainGrid.TextMatrix(i, 1) = rs![payno]
    PaymentMainGrid.TextMatrix(i, 2) = rs![paydate]
    PaymentMainGrid.TextMatrix(i, 3) = rs![invoiceno]
    PaymentMainGrid.TextMatrix(i, 4) = rs![supbillno]
    PaymentMainGrid.TextMatrix(i, 5) = rs![supname]
    PaymentMainGrid.TextMatrix(i, 6) = Format(rs![SumOfpaynetamt], "0.00")
    rs.MoveNext
    i = i + 1
    Wend
    PaymentMainGrid.Rows = rs.RecordCount + 1
    End If
End Function
Private Sub PaymentMainGrid_Click()
    If PaymentMainGrid.Col = 1 Or PaymentMainGrid.Col = 2 Or PaymentMainGrid.Col = 3 Or PaymentMainGrid.Col = 4 Or PaymentMainGrid.Col = 5 Or PaymentMainGrid.Col = 6 Then
         txtedit.Text = PaymentMainGrid.TextMatrix(PaymentMainGrid.Row, 1)
    End If
        If PaymentMainGrid.Col = 1 And PaymentMainGrid.Row = 1 Then
        cbopayFilter.Visible = True
        Else
        cbopayFilter.Visible = False
        End If
        If PaymentMainGrid.Col = 4 And PaymentMainGrid.Row = 1 Then
        cbosupbillfilter.Visible = True
        Else
        cbosupbillfilter.Visible = False
        End If
        If PaymentMainGrid.Col = 3 And PaymentMainGrid.Row = 1 Then
        cboinvoice.Visible = True
        Else
        cboinvoice.Visible = False
        End If
        If PaymentMainGrid.Col = 5 And PaymentMainGrid.Row = 1 Then
        cbosupfilter.Visible = True
        Else
        cbosupfilter.Visible = False
        End If
End Sub
Sub visi()
    cbopayFilter.Visible = False
    cbopayFilter.Clear
    cbosupbillfilter.Visible = False
    cbosupbillfilter.Clear
    cboinvoice.Visible = False
    cboinvoice.Clear
    cbosupfilter.Visible = False
    cbosupfilter.Clear
End Sub
Private Function paynofilterload()
    Dim i As Integer
    Dim rs As New ADODB.Recordset
    rs.Open "Select * from paymentmain_details where payno='" & cbopayFilter.Text & "'", cn, adOpenKeyset, adLockOptimistic
    i = 1
    PaymentMainGrid.TextMatrix(i, 1) = rs![payno]
    PaymentMainGrid.TextMatrix(i, 2) = rs![paydate]
    PaymentMainGrid.TextMatrix(i, 3) = rs![invoiceno]
    PaymentMainGrid.TextMatrix(i, 4) = rs![supbillno]
    PaymentMainGrid.TextMatrix(i, 5) = rs![supname]
    PaymentMainGrid.TextMatrix(i, 6) = Format(rs![SumOfpaynetamt], "0.00")
    PaymentMainGrid.Rows = rs.RecordCount + 1
End Function
 Private Function invoicefilterload()
    Dim i As Integer
    Dim rs As New ADODB.Recordset
    rs.Open "Select * from paymentmain_details where invoiceno='" & cboinvoice.Text & "'", cn, adOpenKeyset, adLockOptimistic
    i = 1
    PaymentMainGrid.TextMatrix(i, 1) = rs![payno]
    PaymentMainGrid.TextMatrix(i, 2) = rs![paydate]
    PaymentMainGrid.TextMatrix(i, 3) = rs![invoiceno]
    PaymentMainGrid.TextMatrix(i, 4) = rs![supbillno]
    PaymentMainGrid.TextMatrix(i, 5) = rs![supname]
    PaymentMainGrid.TextMatrix(i, 6) = Format(rs![SumOfpaynetamt], "0.00")
    PaymentMainGrid.Rows = rs.RecordCount + 1
End Function
Private Function supbillfilterload()
    Dim i As Integer
    Dim rs As New ADODB.Recordset
    rs.Open "Select * from paymentmain_details where supbillno='" & cbosupbillfilter.Text & "'", cn, adOpenKeyset, adLockOptimistic
    i = 1
    PaymentMainGrid.TextMatrix(i, 1) = rs![payno]
    PaymentMainGrid.TextMatrix(i, 2) = rs![paydate]
    PaymentMainGrid.TextMatrix(i, 3) = rs![invoiceno]
    PaymentMainGrid.TextMatrix(i, 4) = rs![supbillno]
    PaymentMainGrid.TextMatrix(i, 5) = rs![supname]
    PaymentMainGrid.TextMatrix(i, 6) = Format(rs![SumOfpaynetamt], "0.00")
    PaymentMainGrid.Rows = rs.RecordCount + 1
End Function
Private Function supnamefilterload()
    Dim i As Integer
    Dim rs As New ADODB.Recordset
    rs.Open "Select * from paymentmain_details where supname='" & cbosupfilter.Text & "'", cn, adOpenKeyset, adLockOptimistic
    i = 1
    PaymentMainGrid.TextMatrix(i, 1) = rs![payno]
    PaymentMainGrid.TextMatrix(i, 2) = rs![paydate]
    PaymentMainGrid.TextMatrix(i, 3) = rs![invoiceno]
    PaymentMainGrid.TextMatrix(i, 4) = rs![supbillno]
    PaymentMainGrid.TextMatrix(i, 5) = rs![supname]
    PaymentMainGrid.TextMatrix(i, 6) = Format(rs![SumOfpaynetamt], "0.00")
    PaymentMainGrid.Rows = rs.RecordCount + 1
End Function

