VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmopengrnmain 
   BackColor       =   &H00EDDDD1&
   Caption         =   " * Open Goods Receipts *"
   ClientHeight    =   9180
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9945
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   11115
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdclosure 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Closure GRN's"
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
      Left            =   13200
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   " To Use Close GRN's "
      Top             =   9600
      Width           =   1815
   End
   Begin VB.TextBox txteditgrn 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   9720
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton cmdprint 
      BackColor       =   &H00FFC0C0&
      Caption         =   "&Print GRN"
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
      Left            =   7920
      MaskColor       =   &H00C0FFC0&
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "  To use Print GRN   "
      Top             =   9600
      UseMaskColor    =   -1  'True
      Width           =   1575
   End
   Begin VB.CommandButton cmddelete 
      BackColor       =   &H00FFC0C0&
      Caption         =   "&Delete GRN"
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
      TabIndex        =   7
      Tag             =   " "
      ToolTipText     =   " To use Delete GRN "
      Top             =   9600
      UseMaskColor    =   -1  'True
      Width           =   1575
   End
   Begin VB.CommandButton cmdedit 
      BackColor       =   &H00FFC0C0&
      Caption         =   "&View GRN"
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
      TabIndex        =   6
      Tag             =   " "
      ToolTipText     =   " To Use View GRN "
      Top             =   9600
      UseMaskColor    =   -1  'True
      Width           =   1455
   End
   Begin VB.CommandButton cmdadd 
      BackColor       =   &H00FFC0C0&
      Caption         =   "&Add GRN"
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
      TabIndex        =   5
      Tag             =   " "
      ToolTipText     =   " To Use Add GRN"
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
      Left            =   9600
      MaskColor       =   &H00C0FFC0&
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "  Exit Window  "
      Top             =   9600
      UseMaskColor    =   -1  'True
      Width           =   1575
   End
   Begin VB.ComboBox cboopengrnnofilter 
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
      Left            =   1080
      Style           =   2  'Dropdown List
      TabIndex        =   3
      ToolTipText     =   " To Use Filter The GRN No Wise "
      Top             =   1080
      Width           =   2055
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
      Left            =   5040
      Style           =   2  'Dropdown List
      TabIndex        =   2
      ToolTipText     =   " To Use Filter The Supplier Name Wise "
      Top             =   1080
      Width           =   3975
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
      Left            =   10320
      Style           =   2  'Dropdown List
      TabIndex        =   1
      ToolTipText     =   " To Use Filter The Department Wise "
      Top             =   1080
      Width           =   2175
   End
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
      Left            =   13200
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   " Show All GRN 's "
      Top             =   360
      Width           =   1695
   End
   Begin MSFlexGridLib.MSFlexGrid GrnMainGrid 
      Height          =   6615
      Left            =   240
      TabIndex        =   9
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
      Caption         =   " Goods Receipts ( Open GRN )"
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
      Width           =   6135
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H000000FF&
      Height          =   735
      Left            =   240
      Top             =   240
      Width           =   14775
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
      Left            =   120
      Top             =   8400
      Width           =   15015
   End
End
Attribute VB_Name = "frmopengrnmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim op As Variant
Dim cn As ADODB.Connection
Dim rs As ADODB.Recordset
Dim ww As ADODB.Recordset
Private Sub cbodept_Click()
     Call opengrndeptfilter
     If GrnMainGrid.Rows = 1 Then
     MsgBox "No Record Found", vbInformation, "Information"
     End If
End Sub
Private Sub cbodept_dropdown()
      On Error GoTo X
       cbodept.Clear
        Set rs = cn.Execute("select deptname from opengrnstatus_details where opengrnstatus='Open' Group by deptname")
        rs.MoveFirst
        Do While Not rs.EOF()
        cbodept.additem (rs(0))
        rs.MoveNext
        Loop
        cbodept.SetFocus
X:
End Sub
Private Sub cboopengrnnofilter_Click()
    Call opengrnnofilter
End Sub
Private Sub cboopengrnnofilter_DropDown()
    On Error GoTo X
       cboopengrnnofilter.Clear
        Set rs = cn.Execute("select opengrnno from opengrnstatus_details where opengrnstatus='Open' order by opengrnno")
        rs.MoveFirst
        Do While Not rs.EOF()
        cboopengrnnofilter.additem (rs(0))
        rs.MoveNext
        Loop
        cboopengrnnofilter.SetFocus
X:
End Sub
Private Sub cbosupfilter_Click()
    Call opengrnsupfilter
End Sub
Private Sub cbosupfilter_dropdown()
 On Error GoTo X
       cbosupfilter.Clear
        Set rs = cn.Execute("select supname from opengrnstatus_details where opengrnstatus = 'Open' Group by supname")
        rs.MoveFirst
        Do While Not rs.EOF()
        cbosupfilter.additem (rs(0))
        rs.MoveNext
        Loop
       cbosupfilter.SetFocus
X:
End Sub
Private Sub cmdadd_Click()
    Unload Me
    frmopenrec.Show
    frmopenrec.cmdsave.Visible = True
    frmopenrec.cmdupdate.Visible = False
End Sub

Private Sub cmdclosure_Click()
    frmopengrnclosure.Show
    Unload Me
End Sub
Private Sub cmddelete_Click()
    Dim rs As New ADODB.Recordset
    Dim rs1 As New ADODB.Recordset
    Dim rs2 As New ADODB.Recordset
    
    If Trim(txteditgrn.Text) = "" Then
        MsgBox "Please Select The GRN Number", vbCritical, "Selecting Error"
    Else
        rs.Open "Select * from deliveryopen_details where opengrnnos ='" & txteditgrn.Text & "'", cn, adOpenKeyset, adLockOptimistic
    If rs.RecordCount >= 1 Then
        MsgBox "Some Quantity Already Delivered Against This Good Receipt! So Caanot Delete!", vbInformation, "Cannot This GRN No Delete "
    Else
        rs1.Open "Select * from opengrn_details where opengrnno =" & txteditgrn.Text, cn, adOpenKeyset, adLockOptimistic
        rs2.Open "Select * from opengrnstatus_details where opengrnno ='" & txteditgrn.Text & "'", cn, adOpenKeyset, adLockOptimistic
    If rs1.RecordCount <> 0 Then

    If MsgBox("Are You Sure Delete PO No  " & txteditgrn.Text & " ? ", vbQuestion + vbYesNo, "Confirm To Delete") = vbYes Then
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
    frmopengrnmain.Show
End If
End If
End If
End If
End Sub
Private Sub cmdedit_Click()
    
    Dim rs1 As New ADODB.Recordset
    rs1.Open "Select * from deliveryopen_details where opengrnnos= '" & txteditgrn.Text & "'", cn, 1, 3
     
    If rs1.RecordCount >= 1 Then
    MsgBox "Already Some Of Qty Delivered Against This GRN No! So Cann't Update", vbCritical, "Update Error"
    frmopenrec.cmdupdate.Enabled = False
    frmopenrec.GrnGrid.Enabled = False
    frmopenrec.cmdsave.Visible = False
    End If
    
    If Trim(txteditgrn.Text) = "" Then
    MsgBox "Please Select Select The GRN Number", vbCritical, "GRN No Selection Error"
    GrnMainGrid.Col = 1
    GrnMainGrid.SetFocus
    Else
    
    frmopenrec.cmdaddrow.Enabled = False
    frmopenrec.cmddeleterow.Enabled = False
'    frmopenrec.cmdupdate.Visible = True
    frmopenrec.cmdsave.Visible = False
    frmopenrec.txtdc.Enabled = False
    frmopenrec.cmdupdate.Enabled = True
    
    Dim i As Integer
    Dim rs As New ADODB.Recordset
    rs.Open "select *from opengrn_details where opengrnno = " & frmopengrnmain.txteditgrn.Text, cn, 1, 3
    i = 1
    If rs.BOF = False Then rs.MoveFirst
    While rs.EOF = False
        frmopenrec.GrnGrid.Rows = rs.RecordCount + 1
        frmopenrec.txtgrnno.Text = rs![opengrnno]
        frmopenrec.dt1.Value = rs![opengrndate]
        frmopenrec.txtdc.Text = rs![dcno]
        frmopenrec.dt2.Value = rs![dcdate]
        frmopenrec.cbosupplier.Text = rs![supname]
        frmopenrec.cbodept.Text = rs![deptname]
        frmopenrec.GrnGrid.TextMatrix(i, 0) = rs![sno]
        frmopenrec.GrnGrid.TextMatrix(i, 1) = rs![itemname]
        frmopenrec.GrnGrid.TextMatrix(i, 2) = rs![colour]
        frmopenrec.GrnGrid.TextMatrix(i, 3) = rs![sizes]
        frmopenrec.GrnGrid.TextMatrix(i, 4) = Format(rs![qty], "0.000")
        frmopenrec.GrnGrid.TextMatrix(i, 5) = rs![uom]
        frmopenrec.txtremarks.Text = rs![remarks]
        
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
    cboopengrnnofilter.Clear
    cbosupfilter.Clear
    cbodept.Clear
    cboopengrnnofilter.Visible = False
    cbosupfilter.Visible = False
    cbodept.Visible = False
    Call grnmainload
End Sub

Private Sub cmdprint_Click()
    Dim rscomp As New ADODB.Recordset
    Dim rsgrn As New ADODB.Recordset
    Dim cn As New ADODB.Connection
    Dim rssup As New ADODB.Recordset
    Dim rsItem As New ADODB.Recordset
    Dim irc As Integer
    Dim filenum As Integer
    
    If GrnMainGrid.Rows = 1 Then
    MsgBox "No Record Found", vbInformation, "Information"
    GrnMainGrid.SetFocus
    ElseIf Trim(txteditgrn.Text) = "" Then
    MsgBox "Please Select The Open GRN No ", vbCritical, "GRN No Selection Error"
    txteditgrn.SetFocus
    Else
    
    cn.Open "Provider=Microsoft.jet.oledb.4.0; Data Source=" & App.path & "\Database\Data.mdb"
    rsgrn.Open "SELECT * FROM opengrnstatus_details where opengrnno ='" & txteditgrn.Text & "'", cn, 1, 3
    rssup.Open "select * from sup_details where supname ='" & GrnMainGrid.TextMatrix(GrnMainGrid.Row, 3) & "'", cn, 1, 3
    rscomp.Open "select * from com_details ", cn, 1, 3
    
    
    
    ChDrive App.path
    ChDir App.path
    
    Kill App.path & "\Reports\PrintOpenGrn.txt"
    Open App.path & "\Reports\PrintOpenGrn.txt" For Output As #1
   
    Print #1,
    Print #1, Tab(42); "GOODS RECEIPT";
    Print #1,
    Print #1, String(100, "-");
    Print #1, Tab(2); "GRN No: "; rsgrn![opengrnno]; Tab(75); "GRN Date :"; rsgrn![opengrndate];
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
    Print #1, Tab(2); "S.No"; Tab(8); "Item Name"; Tab(25); "Colour"; Tab(40); "Size"; Tab(59); "Rec.Qty", Tab(71); "UOM"; Tab(85); "PO No"; 'Tab(92); "Tot.Amt";
    
    Print #1,
    Print #1, String(100, "-");
    rsItem.Open "Select * from opengrn_details where opengrnno =" & txteditgrn.Text, cn, 1, 3
    i = 1
    If rsItem.BOF = False Then rsItem.MoveFirst
    While rsItem.EOF = False
    Print #1,
    Print #1, rsItem![sno]; Tab(8); rsItem![itemname]; Tab(25); rsItem![colour]; Tab(40); rsItem![sizes]; Tab(59); Format(rsItem![qty], "0.000"); Tab(71); rsItem![uom];
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
    frmopengrnreport.Show
    End If
End Sub

Private Sub Form_Load()
    Set cn = New ADODB.Connection
    Set ww = New ADODB.Recordset
    Set rs = New ADODB.Recordset
    cn.Open "Provider=Microsoft.jet.oledb.4.0; Data Source=" & App.path & "\Database\Data.mdb"
    ww.Open "Select * From opengrn_details", cn, adOpenKeyset, adLockOptimistic
    cn.CursorLocation = adUseClient
    
    Call grnmainload
    Call frmgrnmainopengrid
    
    cboopengrnnofilter.Visible = False
    cbosupfilter.Visible = False
    cbodept.Visible = False
End Sub
Private Sub GrnMainGrid_Click()
    If GrnMainGrid.Rows > 1 Then
    If GrnMainGrid.Col = 1 Or GrnMainGrid.Col = 2 Or GrnMainGrid.Col = 3 Or GrnMainGrid.Col = 4 Or GrnMainGrid.Col = 5 Or GrnMainGrid.Col = 6 Then
    txteditgrn.Text = GrnMainGrid.TextMatrix(GrnMainGrid.Row, 1)
    End If
    End If
        If GrnMainGrid.Col = 1 And GrnMainGrid.Row = 1 Then
        cboopengrnnofilter.Visible = True
        Else
        cboopengrnnofilter.Visible = False
        End If
        If GrnMainGrid.Col = 3 And GrnMainGrid.Row = 1 Then
        cbosupfilter.Visible = True
        Else
        cbosupfilter.Visible = False
        End If
        If GrnMainGrid.Col = 5 And GrnMainGrid.Row = 1 Then
        cbodept.Visible = True
        Else
        cbodept.Visible = False
        End If
End Sub
Sub grnmainload()
    Dim i As Integer
    Dim rs As New ADODB.Recordset
    rs.Open "Select * from opengrnstatus_details where opengrnstatus = 'Open'", cn, adOpenKeyset, adLockOptimistic
    If rs.RecordCount = 0 Then
    MsgBox "No Records Found", vbInformation, "Information"
    Else
    i = 1
    If rs.BOF = False Then rs.MoveFirst
    While rs.EOF = False
    GrnMainGrid.Rows = GrnMainGrid.Rows + 1
    GrnMainGrid.TextMatrix(i, 0) = i
    GrnMainGrid.TextMatrix(i, 1) = rs![opengrnno]
    GrnMainGrid.TextMatrix(i, 2) = rs![opengrndate]
    GrnMainGrid.TextMatrix(i, 3) = rs![supname]
    GrnMainGrid.TextMatrix(i, 4) = rs![opendcno]
    GrnMainGrid.TextMatrix(i, 5) = rs![deptname]
    rs.MoveNext
    i = i + 1
    Wend
    GrnMainGrid.Rows = rs.RecordCount + 1
    End If
End Sub
Sub opengrnnofilter()
    Dim i As Integer
    Dim rs As New ADODB.Recordset
    rs.Open "Select * from opengrnstatus_details where opengrnstatus='Open' AND opengrnno =' " & cboopengrnnofilter.Text & "'", cn, adOpenKeyset, adLockOptimistic
    i = 1
    If rs.BOF = False Then rs.MoveFirst
    While rs.EOF = False
    GrnMainGrid.Rows = GrnMainGrid.Rows + 1
    GrnMainGrid.TextMatrix(i, 0) = i
    GrnMainGrid.TextMatrix(i, 1) = rs![opengrnno]
    GrnMainGrid.TextMatrix(i, 2) = rs![opengrndate]
    GrnMainGrid.TextMatrix(i, 3) = rs![supname]
    GrnMainGrid.TextMatrix(i, 4) = rs![opendcno]
    GrnMainGrid.TextMatrix(i, 5) = rs![deptname]
    rs.MoveNext
    i = i + 1
    Wend
    GrnMainGrid.Rows = rs.RecordCount + 1
End Sub
Sub opengrnsupfilter()
    Dim i As Integer
    Dim rs As New ADODB.Recordset
    rs.Open "Select * from opengrnstatus_details where opengrnstatus='Open' AND supname = '" & cbosupfilter.Text & "'", cn, adOpenKeyset, adLockOptimistic
    i = 1
    If rs.BOF = False Then rs.MoveFirst
    While rs.EOF = False
    GrnMainGrid.Rows = GrnMainGrid.Rows + 1
    GrnMainGrid.TextMatrix(i, 0) = i
    GrnMainGrid.TextMatrix(i, 1) = rs![opengrnno]
    GrnMainGrid.TextMatrix(i, 2) = rs![opengrndate]
    GrnMainGrid.TextMatrix(i, 3) = rs![supname]
    GrnMainGrid.TextMatrix(i, 4) = rs![opendcno]
    GrnMainGrid.TextMatrix(i, 5) = rs![deptname]
    rs.MoveNext
    i = i + 1
    Wend
    GrnMainGrid.Rows = rs.RecordCount + 1
End Sub
Sub opengrndeptfilter()
    Dim i As Integer
    Dim rs As New ADODB.Recordset
    rs.Open "Select * from opengrnstatus_details where opengrnstatus ='open' AND deptname = '" & cbodept.Text & "'", cn, adOpenKeyset, adLockOptimistic
    i = 1
    If rs.BOF = False Then rs.MoveFirst
    While rs.EOF = False
    GrnMainGrid.Rows = GrnMainGrid.Rows + 1
    GrnMainGrid.TextMatrix(i, 0) = i
    GrnMainGrid.TextMatrix(i, 1) = rs![opengrnno]
    GrnMainGrid.TextMatrix(i, 2) = rs![opengrndate]
    GrnMainGrid.TextMatrix(i, 3) = rs![supname]
    GrnMainGrid.TextMatrix(i, 4) = rs![opendcno]
    GrnMainGrid.TextMatrix(i, 5) = rs![deptname]
    rs.MoveNext
    i = i + 1
    Wend
    GrnMainGrid.Rows = rs.RecordCount + 1
End Sub

