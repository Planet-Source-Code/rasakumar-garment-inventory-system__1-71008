VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmopenrec 
   BackColor       =   &H00EDDDD1&
   Caption         =   " * Open Receipts *"
   ClientHeight    =   7935
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9585
   Icon            =   "frmopenrec.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   11115
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtqty 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6480
      TabIndex        =   28
      Top             =   2460
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtremarks 
      Appearance      =   0  'Flat
      Height          =   735
      Left            =   4320
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   26
      ToolTipText     =   " Enter The Remarks "
      Top             =   6720
      Width           =   10815
   End
   Begin VB.TextBox txtcomid 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   405
      Left            =   8400
      TabIndex        =   25
      Top             =   7800
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton cmddeleterow 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Delete Item"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   24
      Tag             =   " "
      ToolTipText     =   "To Use Delete Items  "
      Top             =   6840
      Width           =   1335
   End
   Begin VB.CommandButton cmdaddrow 
      BackColor       =   &H00FFC0C0&
      Caption         =   "ADD Item"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   23
      ToolTipText     =   "  To Use Add Items "
      Top             =   6840
      Width           =   1215
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
      Left            =   10200
      TabIndex        =   22
      ToolTipText     =   " Select The Department Name "
      Top             =   840
      Width           =   3375
   End
   Begin VB.ComboBox cboitem 
      Appearance      =   0  'Flat
      ForeColor       =   &H000040C0&
      Height          =   315
      Left            =   240
      TabIndex        =   21
      ToolTipText     =   " Select The Item Name "
      Top             =   1680
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.ComboBox cbocolour 
      Appearance      =   0  'Flat
      ForeColor       =   &H000040C0&
      Height          =   315
      Left            =   2040
      TabIndex        =   20
      ToolTipText     =   " Select The Colour "
      Top             =   1680
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.ComboBox cbosize 
      Appearance      =   0  'Flat
      ForeColor       =   &H000040C0&
      Height          =   315
      Left            =   3480
      TabIndex        =   19
      ToolTipText     =   " Select The Size "
      Top             =   1680
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.ComboBox txtuom 
      Appearance      =   0  'Flat
      ForeColor       =   &H000040C0&
      Height          =   315
      Left            =   5040
      TabIndex        =   18
      ToolTipText     =   " Select The Unit of Measurement "
      Top             =   1680
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox txtgrnno 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      TabIndex        =   10
      ToolTipText     =   " Your GRN No "
      Top             =   360
      Width           =   1575
   End
   Begin VB.ComboBox cbosupplier 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10200
      TabIndex        =   8
      ToolTipText     =   " Select The Supplier Name "
      Top             =   360
      Width           =   3375
   End
   Begin VB.TextBox txtdc 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      TabIndex        =   7
      ToolTipText     =   " Enter The Supplier DC NO "
      Top             =   720
      Width           =   1575
   End
   Begin VB.TextBox txtid 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   6840
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   7800
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox txttype 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   405
      Left            =   9840
      TabIndex        =   4
      Text            =   "Open GRN "
      Top             =   7800
      Visible         =   0   'False
      Width           =   1575
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
      Height          =   615
      Left            =   7080
      MaskColor       =   &H00C0FFC0&
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "  Exit Window  "
      Top             =   9480
      UseMaskColor    =   -1  'True
      Width           =   1455
   End
   Begin VB.CommandButton cmdsave 
      BackColor       =   &H00FFC0C0&
      Caption         =   "&Save GRN"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5520
      MaskColor       =   &H00C0FFC0&
      Style           =   1  'Graphical
      TabIndex        =   2
      Tag             =   " "
      ToolTipText     =   " To Use Add Open GRN "
      Top             =   9480
      UseMaskColor    =   -1  'True
      Width           =   1455
   End
   Begin VB.CommandButton cmdupdate 
      BackColor       =   &H00FFC0C0&
      Caption         =   "&Update GRN"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5520
      MaskColor       =   &H00C0FFC0&
      Style           =   1  'Graphical
      TabIndex        =   1
      Tag             =   " "
      ToolTipText     =   " To use Update PO "
      Top             =   9480
      UseMaskColor    =   -1  'True
      Width           =   1455
   End
   Begin VB.TextBox txttot 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   5400
      TabIndex        =   0
      Top             =   7800
      Visible         =   0   'False
      Width           =   1455
   End
   Begin MSComCtl2.DTPicker dt2 
      Height          =   375
      Left            =   5040
      TabIndex        =   6
      Top             =   720
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   58261505
      CurrentDate     =   39481
   End
   Begin MSComCtl2.DTPicker dt1 
      Height          =   375
      Left            =   5040
      TabIndex        =   9
      Top             =   360
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   58261505
      CurrentDate     =   39481
   End
   Begin MSFlexGridLib.MSFlexGrid GrnGrid 
      Height          =   5175
      Left            =   120
      TabIndex        =   11
      ToolTipText     =   "  Note : # Indicate Place Click The First Row of Grid To  Open the Filter Options "
      Top             =   1440
      Width           =   15015
      _ExtentX        =   26485
      _ExtentY        =   9128
      _Version        =   393216
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
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Remarks"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   4
      Left            =   3120
      TabIndex        =   27
      Top             =   6720
      Width           =   1215
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H000000FF&
      Height          =   855
      Left            =   120
      Top             =   6720
      Width           =   2895
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H000000FF&
      Height          =   1215
      Index           =   1
      Left            =   7080
      Top             =   120
      Width           =   8055
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H000000FF&
      Height          =   1215
      Index           =   0
      Left            =   120
      Top             =   120
      Width           =   6855
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "OPEN GRN NO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   0
      Left            =   600
      TabIndex        =   17
      Top             =   360
      Width           =   1575
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "GRN Date"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   0
      Left            =   3720
      TabIndex        =   16
      Top             =   360
      Width           =   1335
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Supplier Name "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   1
      Left            =   8520
      TabIndex        =   15
      Top             =   360
      Width           =   1695
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Department"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   2
      Left            =   8520
      TabIndex        =   14
      Top             =   840
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "DC No"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   1
      Left            =   600
      TabIndex        =   13
      Top             =   720
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "DC Date"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   2
      Left            =   3720
      TabIndex        =   12
      Top             =   720
      Width           =   1335
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   2
      Height          =   855
      Left            =   5400
      Top             =   9360
      Width           =   3255
   End
End
Attribute VB_Name = "frmopenrec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim op As Variant
Dim cn As ADODB.Connection
Dim rs As ADODB.Recordset
Dim ww As ADODB.Recordset
Private Sub cbocolour_Click()
    If GrnGrid.Col = 2 Then
        Me.GrnGrid.Text = Me.cbocolour.Text
         Call cbocolour_dropdown
            GrnGrid.Col = 3
         Call GrnGrid_Click
    End If
End Sub

Private Sub cbocolour_KeyPress(KeyAscii As Integer)
     KeyAscii = 0
End Sub
Private Sub cbodept_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub cboitem_Click()
  If GrnGrid.Col = 1 Then
        Me.GrnGrid.Text = Me.cboitem.Text
         Call cboitem_dropdown
            GrnGrid.Col = 2
         Call GrnGrid_Click
 End If
End Sub
Private Sub cboitem_KeyPress(KeyAscii As Integer)
     KeyAscii = 0
End Sub

Private Sub cbosize_Click()
     If GrnGrid.Col = 3 Then
        Me.GrnGrid.Text = Me.cbosize.Text
         Call cbosize_dropdown
            GrnGrid.Col = 4
         Call GrnGrid_Click
    End If
End Sub
Private Sub cbosize_KeyPress(KeyAscii As Integer)
     KeyAscii = 0
End Sub

Private Sub cbosupplier_DropDown()
    On Error GoTo X
    cbosupplier.Clear
        Set rs = cn.Execute("select supname from sup_details order by supname")
        rs.MoveFirst
        Do While Not rs.EOF()
        cbosupplier.additem (rs(0))
        rs.MoveNext
        Loop
        cbosupplier.SetFocus
X:
End Sub
Private Sub cbosupplier_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub cmddeleterow_Click()
     If GrnGrid.Rows <= 1 Then
     MsgBox "No Record Found", vbInformation, "Information"
     Else
     GrnGrid.Rows = GrnGrid.Rows - 1
     GrnGrid.Row = GrnGrid.Rows - 1
     GrnGrid.TextMatrix(GrnGrid.Row, 0) = GrnGrid.Rows - 1
     End If
End Sub

Private Sub cmdsave_Click()
    If Trim(cbosupplier.Text) = "" Then
    MsgBox "Please Select The Supplier Name ", vbCritical, "Supplier Name Error "
    cbosupplier.SetFocus
    ElseIf Trim(cbodept.Text) = "" Then
    MsgBox "Please Enter The Department Name ", vbCritical, "Department Name Error"
    cbodept.SetFocus
    ElseIf Trim(txtdc.Text) = "" Then
    MsgBox "Please Enter the Supplier Dc No", vbCritical, "Supplier Dc No Error"
    txtdc.SetFocus
    ElseIf Trim(GrnGrid.Text) = "" Then
    MsgBox "Please Enter All Data", vbCritical, "Data Enter Error"
    Else
    Dim rs As New ADODB.Recordset
    Dim rsstatus As New ADODB.Recordset
    rs.Open "Select * from opengrn_details where opengrnno = " & txtgrnno.Text, cn, 1, 3
    rsstatus.Open "Select * from opengrnstatus_details where opengrnno= '" & txtgrnno.Text & "'", cn, adOpenKeyset, adLockOptimistic
        
        If rsstatus.RecordCount = 0 Then
        rsstatus.AddNew
        rsstatus![opengrnno] = txtgrnno.Text
        rsstatus![opengrndate] = dt1.Value
        rsstatus![opengrnstatus] = "Open"
        rsstatus![supname] = cbosupplier.Text
        rsstatus![deptname] = cbodept.Text
        rsstatus![opendcno] = txtdc.Text
        rsstatus.Update
        rsstatus.Close
        End If
        
    
    If rs.RecordCount = 0 Then
    For i = 1 To GrnGrid.Rows - 1
    rs.AddNew
    rs![opengrnno] = txtgrnno.Text
    rs![opengrndate] = dt1.Value
    rs![dcno] = txtdc.Text
    rs![dcdate] = dt2.Value
    rs![supname] = cbosupplier.Text
    rs![deptname] = cbodept.Text
    rs![sno] = GrnGrid.TextMatrix(i, 0)
    rs![itemname] = GrnGrid.TextMatrix(i, 1)
    rs![colour] = GrnGrid.TextMatrix(i, 2)
    rs![sizes] = GrnGrid.TextMatrix(i, 3)
    rs![qty] = GrnGrid.TextMatrix(i, 4)
    rs![uom] = GrnGrid.TextMatrix(i, 5)
    rs![remarks] = txtremarks.Text
    
    Next i
    rs.Update
    rs.Close
    MsgBox "One Record Save Successfully", vbInformation, "Information"
    Unload Me
    frmopengrnmain.Show
    End If
    End If
    
End Sub
Private Sub cmdaddrow_Click()
     If GrnGrid.Rows > 10 Then
     MsgBox "Only 10 Item Allowed ", vbCritical, "Exceed Row "
     Else
     GrnGrid.Rows = GrnGrid.Rows + 1
     GrnGrid.Row = GrnGrid.Rows - 1
     GrnGrid.TextMatrix(GrnGrid.Row, 0) = GrnGrid.Rows - 1
     End If
End Sub
Private Sub cmdexit_Click()
    op = MsgBox("Are You Sure To Close ?", vbYesNo + vbQuestion, "Confirm Close ?")
    If op = vbYes Then
    Unload Me
    frmopengrnmain.Show
    Else
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
    rsgrn.Open "SELECT * FROM opengrnmain_details where opengrnno = " & GrnMainGrid.TextMatrix(GrnMainGrid.Row, 1), cn, 1, 3
    rssup.Open "select * from sup_details where supname ='" & GrnMainGrid.TextMatrix(GrnMainGrid.Row, 3) & "'", cn, 1, 3
    rscomp.Open "select * from com_details ", cn, 1, 3
    
    If Trim(txteditgrn.Text) = "" Then
    MsgBox "Please Select The Open GRN No ", vbCritical, "GRN No Selection Error"
    txteditgrn.SetFocus
    Else
    
    ChDrive App.path
    ChDir App.path
    
    Kill App.path & "\Reports\PrintOpenGrn.txt"
    Open App.path & "\Reports\PrintOpenGrn.txt" For Output As #1
   
     
    Print #1,
    Print #1, Tab(42); "OPEN GOODS RECEIPT";
    Print #1,
    Print #1, String(100, "-");
    Print #1, Tab(2); "OPEN GRN No: "; rsgrn![opengrnno]; Tab(70); "OPEN GRN Date :"; rsgrn![opengrndate];
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
    Print #1, rsItem![sno]; Tab(8); rsItem![itemname]; Tab(25); rsItem![colour]; Tab(40); rsItem![sizes]; Tab(59); Format(rsItem![qty], "0.000"); Tab(71); rsItem![uom]; Tab(85);
    rsItem.MoveNext
    i = i + 1
    Wend
    
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
Private Sub cmdupdate_Click()
    If Trim(cbosupplier.Text) = "" Then
    MsgBox "Please Select The Supplier Name ", vbCritical, "Supplier Name Error "
    cbosupplier.SetFocus
    ElseIf Trim(cbodept.Text) = "" Then
    MsgBox "Please Enter The Department Name ", vbCritical, "Department Name Error"
    cbodept.SetFocus
    ElseIf Trim(txtdc.Text) = "" Then
    MsgBox "Please Enter the Supplier Dc No", vbCritical, "Supplier Dc No Error"
    txtdc.SetFocus
    ElseIf Trim(GrnGrid.Text) = "" Then
    MsgBox "Please Enter All Data", vbCritical, "Data Enter Error"
    Else
    Dim rs As New ADODB.Recordset
    rs.Open "Select * from opengrn_details where opengrnno = " & frmopenrec.txtgrnno.Text, cn, 1, 3
    If rs.RecordCount <> 0 Then
    i = 1
    If rs.BOF = False Then rs.MoveFirst
    While rs.EOF = False
    rs![opengrnno] = txtgrnno.Text
    rs![opengrndate] = dt1.Value
    rs![dcno] = txtdc.Text
    rs![dcdate] = dt2.Value
    rs![supname] = cbosupplier.Text
    rs![deptname] = cbodept.Text
    rs![sno] = GrnGrid.TextMatrix(i, 0)
    rs![itemname] = GrnGrid.TextMatrix(i, 1)
    rs![colour] = GrnGrid.TextMatrix(i, 2)
    rs![sizes] = GrnGrid.TextMatrix(i, 3)
    rs![qty] = GrnGrid.TextMatrix(i, 4)
    rs![uom] = GrnGrid.TextMatrix(i, 5)
    rs![types] = txttype.Text
    rs![remarks] = txtremarks.Text
    rs.Update
    rs.MoveNext
    i = i + 1
    Wend
    MsgBox "One Record Updated Successfully", vbInformation, "Information"
    Unload Me
    End If
    End If
    frmopengrnmain.Show
End Sub
Private Sub Form_Load()
    Set cn = New ADODB.Connection
    Set ww = New ADODB.Recordset
    Set rs = New ADODB.Recordset
    cn.Open "Provider=Microsoft.jet.oledb.4.0; Data Source=" & App.path & "\Database\Data.mdb"
    ww.Open "Select * From opengrn_details", cn, adOpenKeyset, adLockOptimistic
    
    Call frmgrnopen
    Call txtgrnno_Gotfocus

    dt1.Value = Date
    dt2.Value = Date
   
    i = 1
    GrnGrid.TextMatrix(i, 0) = 1
    
End Sub
Private Sub Form_Unload(Cancel As Integer)
  If Cancel = 0 Then
   frmopengrnmain.Show
   Else
   Cancel = 1
   End If
End Sub

Private Sub GrnGrid_Click()
    If cboitem.Visible = True Then cboitem.Visible = False
    If cbocolour.Visible = True Then cbocolour.Visible = False
    If cbosize.Visible = True Then cbosize.Visible = False
    If txtuom.Visible = True Then txtuom.Visible = False
    If txtqty.Visible = True Then txtqty.Visible = False
    If GrnGrid.Col = 1 Then
    Me.cboitem.Visible = True
    CurrentRow = Me.GrnGrid.Row
    Me.cboitem.Width = Me.GrnGrid.CellWidth - 10
    Me.cboitem.Left = Me.GrnGrid.CellLeft + Me.GrnGrid.Left
    Me.cboitem.Top = Me.GrnGrid.CellTop + Me.GrnGrid.Top
    donotchange = True
    Me.cboitem.Text = Me.GrnGrid.Text
    donotchange = False
    Me.cboitem.SetFocus
    ElseIf GrnGrid.Col = 2 Then
    Me.cbocolour.Visible = True
    CurrentRow = Me.GrnGrid.Row
    Me.cbocolour.Width = Me.GrnGrid.CellWidth - 10
    Me.cbocolour.Left = Me.GrnGrid.CellLeft + Me.GrnGrid.Left
    Me.cbocolour.Top = Me.GrnGrid.CellTop + Me.GrnGrid.Top
    donotchange = True
    Me.cbocolour.Text = Me.GrnGrid.Text
    donotchange = False
    Me.cbocolour.SetFocus
    ElseIf GrnGrid.Col = 3 Then
    Me.cbosize.Visible = True
    CurrentRow = Me.GrnGrid.Row
    Me.cbosize.Width = Me.GrnGrid.CellWidth - 10
    Me.cbosize.Left = Me.GrnGrid.CellLeft + Me.GrnGrid.Left
    Me.cbosize.Top = Me.GrnGrid.CellTop + Me.GrnGrid.Top
    donotchange = True
    Me.cbosize.Text = Me.GrnGrid.Text
    donotchange = False
    Me.cbosize.SetFocus
    ElseIf GrnGrid.Col = 4 Then
    Me.txtqty.Visible = True
    CurrentRow = Me.GrnGrid.Row
    Me.txtqty.Width = Me.GrnGrid.CellWidth - 10
    Me.txtqty.Left = Me.GrnGrid.CellLeft + Me.GrnGrid.Left
    Me.txtqty.Top = Me.GrnGrid.CellTop + Me.GrnGrid.Top
    donotchange = True
    Me.txtqty.Text = Format(Me.GrnGrid.Text, "0.000")
    donotchange = False
    Me.txtqty.SetFocus
    ElseIf GrnGrid.Col = 5 Then
    Me.txtuom.Visible = True
    CurrentRow = Me.GrnGrid.Row
    Me.txtuom.Width = Me.GrnGrid.CellWidth - 10
    Me.txtuom.Left = Me.GrnGrid.CellLeft + Me.GrnGrid.Left
    Me.txtuom.Top = Me.GrnGrid.CellTop + Me.GrnGrid.Top
    donotchange = True
    Me.txtuom.Text = Me.GrnGrid.Text
    donotchange = False
    Me.txtuom.SetFocus
    Else
   
End If
End Sub
Private Sub cbodept_dropdown()
    If Trim(cbosupplier.Text = "") Then
    MsgBox "Please Select the Supplier Name ", vbCritical, "Supplier Name Error"
    cbosupplier.SetFocus
    Else
    On Error GoTo X
    cbodept.Clear
        Set rs = cn.Execute("select deptname from dept_details order by deptname")
        rs.MoveFirst
        Do While Not rs.EOF()
        cbodept.additem (rs(0))
        rs.MoveNext
        Loop
        cbodept.SetFocus
X:
End If
End Sub
Private Sub cboitem_dropdown()
    If Trim(cbodept.Text = "") Then
    MsgBox "Please Select the Department Name ", vbCritical, "Department Name Error"
    cbodept.SetFocus
    Else
    On Error GoTo X
    cboitem.Clear
        Set rs = cn.Execute("select itemname from item_details where deptname ='" & cbodept.Text & "'")
        rs.MoveFirst
        Do While Not rs.EOF()
        cboitem.additem (rs(0))
        rs.MoveNext
        Loop
        cboitem.SetFocus
X:
End If
End Sub
Private Sub cbocolour_dropdown()
    If Trim(cbodept.Text = "") Then
    MsgBox "Please Select the Department Name ", vbCritical, "Department Name Error"
    Else
    On Error GoTo X
    cbocolour.Clear
        Set rs = cn.Execute("select colour from colour_details order by colour")
        rs.MoveFirst
        Do While Not rs.EOF()
        cbocolour.additem (rs(0))
        rs.MoveNext
        Loop
        cbocolourt.SetFocus
X:
End If
End Sub
Private Sub cbosize_dropdown()
    If Trim(cbodept.Text = "") Then
    MsgBox "Please Select the Department Name ", vbCritical, "Department Name Error"
    cbodept.SetFocus
    Else
    On Error GoTo X
    cbosize.Clear
        Set rs = cn.Execute("select sizes from size_details order by sizes")
        rs.MoveFirst
        Do While Not rs.EOF()
        cbosize.additem (rs(0))
        rs.MoveNext
        Loop
        cbosize.SetFocus
X:
End If
End Sub

Private Sub txtgrnno_Gotfocus()
    Dim rs As New ADODB.Recordset
    Dim a As String
    rs.Open "Select * from opengrn_details", cn, adOpenKeyset, adLockOptimistic
        If rs.RecordCount = 0 Then
        txtgrnno.Text = 1
        rs.Close
        Else
        Dim rsrs As New ADODB.Recordset
        rsrs.Open "Select max(opengrnno)as exp1 from opengrn_details", cn, adOpenKeyset, adLockOptimistic
        txtgrnno = rsrs![exp1] + 1
        rsrs.Close
        End If
    SendKeys "{tab}"
End Sub
Private Sub txtqty_OnEnter()
    If GrnGrid.Col = 4 Then
        Me.GrnGrid.Text = Format(Me.txtqty.Text, "0.000")
         GrnGrid.Col = 5
    Call GrnGrid_Click
    End If
End Sub
Private Sub txtqty_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Then
    ElseIf (KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 46 Then
    ElseIf KeyAscii = 13 And GrnGrid.Col = 4 Then
    Me.GrnGrid.Text = Format(Me.txtqty.Text, "0.000")
    GrnGrid.Col = 5
    GrnGrid.SetFocus
    Call GrnGrid_Click
    Else
    KeyAscii = 0
    End If
End Sub
Private Sub txtuom_Click()
If GrnGrid.Col = 5 Then
        Me.GrnGrid.Text = Me.txtuom.Text
         Call txtuom_dropdown
         txtuom.Visible = False
        GrnGrid.Col = 1
         GrnGrid.SetFocus
    End If
End Sub
Private Sub txtuom_dropdown()
    If Trim(cbodept.Text = "") Then
    MsgBox "Please Select the Department Name ", vbCritical, "Department Name Error"
    cbodept.SetFocus
    Else
    On Error GoTo X
    txtuom.Clear
    Set rs = cn.Execute("select uom from uom_details order by uom")
    rs.MoveFirst
    Do While Not rs.EOF()
    txtuom.additem (rs(0))
    rs.MoveNext
    Loop
    txtuom.SetFocus
X:
End If
End Sub
Private Sub txtuom_KeyPress(KeyAscii As Integer)
     KeyAscii = 0
End Sub
