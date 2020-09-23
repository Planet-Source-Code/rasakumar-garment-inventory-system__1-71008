VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmDeliverys 
   BackColor       =   &H00EDDDD1&
   Caption         =   " * Delivery Challan * "
   ClientHeight    =   9120
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11145
   Icon            =   "frmDeliverys.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   11115
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.ComboBox cbosup 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2880
      Style           =   2  'Dropdown List
      TabIndex        =   30
      ToolTipText     =   " Select the Supplier "
      Top             =   240
      Width           =   3015
   End
   Begin VB.OptionButton supopt 
      BackColor       =   &H00EDDDD1&
      Caption         =   "Supplier GRN Based"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1320
      TabIndex        =   29
      ToolTipText     =   " Supplier GRN Based To Delivery "
      Top             =   240
      Width           =   1455
   End
   Begin VB.OptionButton allgrnopt 
      BackColor       =   &H00EDDDD1&
      Caption         =   "All GRN Based"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   28
      ToolTipText     =   " All GRN Based To Delivery "
      Top             =   240
      Value           =   -1  'True
      Width           =   1215
   End
   Begin VB.ComboBox cbogrnno 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2040
      Style           =   2  'Dropdown List
      TabIndex        =   27
      ToolTipText     =   " Select the GRN Number "
      Top             =   720
      Width           =   3855
   End
   Begin VB.TextBox txtremarks 
      Appearance      =   0  'Flat
      Height          =   735
      Left            =   1320
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   24
      ToolTipText     =   " Enter The Remarks "
      Top             =   9600
      Width           =   10335
   End
   Begin VB.TextBox txttypes 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3360
      TabIndex        =   23
      Top             =   5040
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSComCtl2.DTPicker dt1 
      Height          =   375
      Left            =   7560
      TabIndex        =   22
      Top             =   720
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   57999361
      CurrentDate     =   39536
   End
   Begin VB.TextBox txttots 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2280
      TabIndex        =   4
      Text            =   "0"
      Top             =   5280
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtrec 
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
      Height          =   270
      Left            =   2880
      TabIndex        =   20
      Top             =   2475
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtbal 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2280
      TabIndex        =   18
      Top             =   5040
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtids 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1200
      TabIndex        =   19
      Top             =   5280
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdexits 
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
      Left            =   13440
      MaskColor       =   &H00C0FFC0&
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "  Exit Window  "
      Top             =   9720
      UseMaskColor    =   -1  'True
      Width           =   1575
   End
   Begin VB.CommandButton cmdsave 
      BackColor       =   &H00FFC0C0&
      Caption         =   "&Save DC"
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
      Left            =   11880
      MaskColor       =   &H00C0FFC0&
      Style           =   1  'Graphical
      TabIndex        =   16
      Tag             =   " "
      ToolTipText     =   " To Use Save the DC "
      Top             =   9720
      UseMaskColor    =   -1  'True
      Width           =   1455
   End
   Begin VB.TextBox txtid 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1200
      TabIndex        =   15
      Top             =   5040
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.ListBox lstid 
      Appearance      =   0  'Flat
      Height          =   420
      Left            =   6480
      TabIndex        =   14
      Top             =   5040
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.CommandButton cmddeleteitem 
      BackColor       =   &H00FFC0C0&
      Caption         =   "De&lete Item"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   13920
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   " To Delete  The Item For Delivery "
      Top             =   5040
      Width           =   1215
   End
   Begin VB.CommandButton cmdadditem 
      BackColor       =   &H00FFC0C0&
      Caption         =   "A&dd Item"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   12600
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   " To Add The Item For Delivery "
      Top             =   5040
      Width           =   1215
   End
   Begin VB.ComboBox cbodept 
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
      Left            =   11040
      TabIndex        =   7
      ToolTipText     =   " Select The  Department "
      Top             =   720
      Width           =   3975
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
      Left            =   11040
      TabIndex        =   6
      ToolTipText     =   " Select The Supplier Name "
      Top             =   240
      Width           =   3975
   End
   Begin VB.TextBox txtdcno 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7560
      TabIndex        =   5
      Top             =   240
      Width           =   1695
   End
   Begin VB.TextBox txttype 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   120
      TabIndex        =   3
      Top             =   5280
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtgrnnoedit 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Top             =   5040
      Visible         =   0   'False
      Width           =   1095
   End
   Begin MSFlexGridLib.MSFlexGrid DetailsGrid 
      Height          =   735
      Left            =   240
      TabIndex        =   1
      ToolTipText     =   " "
      Top             =   4080
      Width           =   14775
      _ExtentX        =   26061
      _ExtentY        =   1296
      _Version        =   393216
      Rows            =   1
      Cols            =   7
      FixedCols       =   0
      BackColor       =   13171709
      BackColorFixed  =   12542735
      ForeColorFixed  =   16777215
      BackColorSel    =   11790056
      ForeColorSel    =   12647934
      BackColorBkg    =   15588820
      GridColorFixed  =   4194368
      GridLines       =   2
      GridLinesFixed  =   1
      ScrollBars      =   0
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
   Begin MSFlexGridLib.MSFlexGrid DeliveryMainGrid 
      Height          =   3615
      Left            =   240
      TabIndex        =   11
      ToolTipText     =   " "
      Top             =   5760
      Width           =   14775
      _ExtentX        =   26061
      _ExtentY        =   6376
      _Version        =   393216
      Rows            =   1
      Cols            =   10
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
   Begin MSFlexGridLib.MSFlexGrid DeliveryGrid 
      Height          =   2775
      Left            =   240
      TabIndex        =   0
      ToolTipText     =   "  "
      Top             =   1200
      Width           =   14775
      _ExtentX        =   26061
      _ExtentY        =   4895
      _Version        =   393216
      Rows            =   1
      Cols            =   11
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
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "GRN No"
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
      Left            =   240
      TabIndex        =   26
      Top             =   720
      Width           =   1815
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
      Left            =   120
      TabIndex        =   25
      Top             =   9600
      Width           =   1215
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FF0000&
      Height          =   735
      Left            =   11760
      Top             =   9600
      Width           =   3375
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H000000FF&
      Height          =   495
      Left            =   120
      Top             =   5040
      Width           =   12375
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "DC Date"
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
      Index           =   3
      Left            =   6000
      TabIndex        =   21
      Top             =   720
      Width           =   1575
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FF00FF&
      Height          =   3855
      Left            =   120
      Top             =   5640
      Width           =   15015
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Department"
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
      Index           =   2
      Left            =   9360
      TabIndex        =   10
      Top             =   720
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Supplier Name"
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
      Index           =   1
      Left            =   9360
      TabIndex        =   9
      Top             =   240
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "DC NO"
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
      Index           =   0
      Left            =   6000
      TabIndex        =   8
      Top             =   240
      Width           =   1575
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H000000FF&
      Height          =   4815
      Index           =   0
      Left            =   120
      Top             =   120
      Width           =   15015
   End
End
Attribute VB_Name = "frmDeliverys"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim op As Variant
Dim cn As ADODB.Connection
Dim rs As ADODB.Recordset
Dim ww As ADODB.Recordset
Private Sub allgrnopt_Click()
    cbosup.Visible = False
    cbosup.Clear
    DeliveryGrid.Rows = 1
    DeliveryMainGrid.Rows = 1
    DetailsGrid.Rows = 1
    
End Sub
Private Sub cbodept_dropdown()
    Dim rs As New ADODB.Recordset
    On Error GoTo X
        cbodept.Clear
        Set rs = cn.Execute("select deptname from dept_details group by deptname")
        rs.MoveFirst
        Do While Not rs.EOF()
        cbodept.additem (rs(0))
        rs.MoveNext
        Loop
        cbodept.SetFocus
X:
End Sub
Private Sub cbodept_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub
Private Sub cbogrnno_Click()
    Call deliverygridload
    Call totalgrngrid
    Call DetailsGridload
   ' Call Detailsreceiveqtyload
End Sub
Private Sub cbogrnno_dropdown()
If allgrnopt.Value = True Then
    Dim rs As New ADODB.Recordset
    On Error GoTo X
        cbogrnno.Clear
        Set rs = cn.Execute("select grnno from grnstatus_details where grnstatus ='Open' Order by grnno")
        rs.MoveFirst
        Do While Not rs.EOF()
        cbogrnno.additem (rs(0))
        rs.MoveNext
        Loop
        cbogrnno.SetFocus
X:
ElseIf supopt.Value = True Then
    Dim rss As New ADODB.Recordset
    On Error GoTo y
        cbogrnno.Clear
        Set rss = cn.Execute("select grnno from grnstatus_details where grnstatus ='Open' AND supname ='" & cbosup.Text & "'")
        rss.MoveFirst
        Do While Not rss.EOF()
        cbogrnno.additem (rss(0))
        rss.MoveNext
        Loop
        cbogrnno.SetFocus
y:
End If
End Sub
Private Sub cbosup_dropdown()
    Dim rs As New ADODB.Recordset
    On Error GoTo X
        cbosup.Clear
        Set rs = cn.Execute("select supname from grnstatus_details where grnstatus ='Open' group by supname")
        rs.MoveFirst
        Do While Not rs.EOF()
        cbosup.additem (rs(0))
        rs.MoveNext
        Loop
        cbosup.SetFocus
X:
End Sub
Private Sub cbosupplier_DropDown()
    Dim rs As New ADODB.Recordset
    On Error GoTo X
        cbosupplier.Clear
        Set rs = cn.Execute("select supname from sup_details group by supname")
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

Private Sub cmdadditem_Click()
    If Trim(txtid.Text) = "" Then
        MsgBox "Please Enter The Grn Valid data", vbCritical, "Error"
        DeliveryMainGrid.SetFocus
    ElseIf Val(DeliveryGrid.TextMatrix(DeliveryGrid.Row, 8)) <= 0 Then
        MsgBox "Please Enter The Valid Qty", vbCritical, "Qty Error"
        DeliveryGrid.Col = 8
        DeliveryGrid.SetFocus
    ElseIf Val(txtrec.Text) <= 0 Then
        MsgBox "Please Enter Valid Received Quantity ", vbCritical, "Received Qty Error"
        DeliveryGrid.SetFocus
    ElseIf DeliveryMainGrid.Rows > 1 Then
        Call dupicates
    Else
        lstid.additem txtid.Text
        DeliveryMainGrid.Rows = DeliveryMainGrid.Rows + 1
        DeliveryMainGrid.Row = DeliveryMainGrid.Rows - 1
        Call DeliveryMainGridload
    End If
End Sub
Private Sub cmddeleteitem_Click()
    If Trim(txtids.Text) = "" And DeliveryMainGrid.Rows > 1 Then
        MsgBox "Please Select The Itemname For Deletion", vbInformation, "Delete Row"
        DeliveryMainGrid.SetFocus
    ElseIf DeliveryMainGrid.Rows > 2 Then
            DeliveryMainGrid.RemoveItem (DeliveryMainGrid.Row)
            DeliveryMainGrid.TextMatrix(DeliveryMainGrid.Row, 0) = DeliveryMainGrid.Rows - 1
                For i = 0 To lstid.ListCount - 1
                If lstid.List(i) = txtids.Text Then
                    lstid.RemoveItem (i)
                End If
                Next i
                txtids.Text = ""
    Else
        MsgBox "There Must be Atleast One Entry, So Cannot Delete! (or) Try Again", vbCritical
    End If
End Sub
Private Sub cmdexits_Click()
    op = MsgBox("Are You Sure To Close ?", vbYesNo + vbQuestion, "Confirm Close ?")
    If op = vbYes Then
        Unload Me
        frmdeliverymain.Show
    Else
    End If
End Sub
Private Sub cmdsave_Click()
    Dim i As Integer
    If Trim(cbosupplier.Text) = "" Then
        MsgBox "Please Select the Supplier Name ", vbCritical, "Supplier Name Error"
        cbosupplier.SetFocus
    ElseIf Trim(cbodept.Text) = "" Then
         MsgBox "Please Select Department ", vbCritical, "Department Name Error"
         cbodept.SetFocus
    ElseIf Val(txttots.Text) <= 0 Then
         MsgBox "Please Enter The Valid Received Quantity Details", vbCritical, "Received Quantity Error"
         DeliveryGrid.Col = 8
         DeliveryGrid.SetFocus
    ElseIf DeliveryMainGrid.Rows < 1 Then
        MsgBox "Please Add The Received Items", vbCritical, "Received Qty Error"
        DeliveryGrid.SetFocus
    Else
        Dim rs As New ADODB.Recordset
        rs.Open "select * from delivery_details where dcno =' " & txtdcno.Text & "'", cn, adOpenKeyset, adLockOptimistic
        If rs.RecordCount = 0 Then
            For i = 1 To DeliveryMainGrid.Rows - 1
            rs.AddNew
            rs![dcno] = txtdcno.Text
            rs![dcdate] = dt1.Value
            rs![supname] = cbosupplier.Text
            rs![deptname] = cbodept.Text
'            rs![deliverytype] = txttype.Text
            rs![sno] = DeliveryMainGrid.TextMatrix(i, 0)
            rs![grnno] = DeliveryMainGrid.TextMatrix(i, 1)
            rs![grndate] = DeliveryMainGrid.TextMatrix(i, 2)
            rs![itemname] = DeliveryMainGrid.TextMatrix(i, 3)
            rs![colour] = DeliveryMainGrid.TextMatrix(i, 4)
            rs![sizes] = DeliveryMainGrid.TextMatrix(i, 5)
            rs![stockqty] = DeliveryMainGrid.TextMatrix(i, 6)
            rs![uom] = DeliveryMainGrid.TextMatrix(i, 7)
            rs![delqty] = DeliveryMainGrid.TextMatrix(i, 8)
            rs![grnid] = DeliveryMainGrid.TextMatrix(i, 9)
            rs![grnnos] = DeliveryMainGrid.TextMatrix(i, 1)
            
            rs![remarks] = txtremarks.Text
            Next i
            rs.Update
            rs.Close
            MsgBox "One Record Save Successfully ", vbInformation, "Information"
            Unload Me
            frmdeliverymain.Show
        Else
        MsgBox "Already Exists", vbCritical, "Error"
    End If
    End If
End Sub
Private Sub DeliveryGrid_Click()
If txtrec.Visible = True Then txtrec.Visible = False
        If DeliveryGrid.Col = 8 Then
            Me.txtrec.Visible = True
            CurrentRow = Me.DeliveryGrid.Row
            Me.txtrec.Width = Me.DeliveryGrid.CellWidth - 10
            Me.txtrec.Left = Me.DeliveryGrid.CellLeft + Me.DeliveryGrid.Left
            Me.txtrec.Top = Me.DeliveryGrid.CellTop + Me.DeliveryGrid.Top
            donotchange = True
            Me.txtrec.Text = Me.DeliveryGrid.Text
            donotchange = False
            txtrec.SetFocus
   End If
If DeliveryGrid.Rows > 1 Then
If DeliveryGrid.Col = 1 Or DeliveryGrid.Col = 2 Or DeliveryGrid.Col = 3 Or DeliveryGrid.Col = 4 Or DeliveryGrid.Col = 5 Or DeliveryGrid.Col = 6 Or DeliveryGrid.Col = 7 Or DeliveryGrid.Col = 8 Or DeliveryGrid.Col = 9 Or DeliveryGrid.Col = 10 Or DeliveryGrid.Col = 11 Or DeliveryGrid.Col = 12 And DeliveryGrid.Rows = 1 Then
        txtgrnnoedit.Text = DeliveryGrid.TextMatrix(DeliveryGrid.Row, 1)
        txtid.Text = DeliveryGrid.TextMatrix(DeliveryGrid.Row, 9)
        txtbal.Text = DeliveryGrid.TextMatrix(DeliveryGrid.Row, 6)
End If
End If
    Call DetailsGridload
    Call Detailsreceiveqtyload
End Sub
Private Sub DeliveryMainGrid_Click()
    If DeliveryMainGrid.Rows > 1 Then
    If DeliveryMainGrid.Col = 7 Or DeliveryMainGrid.Col = 4 Or DeliveryMainGrid.Col = 5 Or DeliveryMainGrid.Col = 6 Or DeliveryMainGrid.Col = 8 Or DeliveryMainGrid.Col = 9 Or DeliveryMainGrid.Col = 3 Or DeliveryMainGrid.Col = 0 Or DeliveryMainGrid.Col = 1 Or DeliveryMainGrid.Col = 2 Or DeliveryMainGrid.Col = 3 Then
         txtids.Text = DeliveryMainGrid.TextMatrix(DeliveryMainGrid.Row, 9)
    End If
    End If
End Sub
Private Sub Form_Load()
    Set cn = New ADODB.Connection
    Set ww = New ADODB.Recordset
    Set rs = New ADODB.Recordset
    cn.Open "Provider=Microsoft.jet.oledb.4.0; Data Source=" & App.path & "\Database\Data.mdb"
    ww.Open "Select * From opengrn_details", cn, adOpenKeyset, adLockOptimistic
    cn.CursorLocation = adUseClient
    cbosup.Visible = False
    Call deliverygridloaditem
    Call deliverygridload
    Call txttypeload
    Call txtdcno_GotFocus
    'Call totalgrngrid
End Sub
Private Function deliverygridload()
    Dim i As Integer
    Dim rs As New ADODB.Recordset
    Dim rs1 As New ADODB.Recordset
    rs.Open "select * from grn_details where grnno= '" & Trim(cbogrnno.Text) & "'", cn, adOpenKeyset, adLockOptimistic
    i = 1
    If rs.BOF = False Then rs.MoveFirst
    While rs.EOF = False
        DeliveryGrid.Rows = rs.RecordCount + 1
        DeliveryGrid.TextMatrix(i, 0) = i
        DeliveryGrid.TextMatrix(i, 1) = rs![grnno]
        DeliveryGrid.TextMatrix(i, 2) = rs![grndate]
        DeliveryGrid.TextMatrix(i, 3) = rs![itemname]
        DeliveryGrid.TextMatrix(i, 4) = rs![colour]
        DeliveryGrid.TextMatrix(i, 5) = rs![sizes]
        DeliveryGrid.TextMatrix(i, 6) = Format(rs![recqty], "0.000")
        DeliveryGrid.TextMatrix(i, 7) = rs![uom]
        DeliveryGrid.TextMatrix(i, 9) = rs![grnid]
        DeliveryGrid.TextMatrix(i, 8) = 0
        rs.MoveNext
        i = i + 1
        Wend
        DeliveryGrid.Rows = rs.RecordCount + 1
End Function
Private Function DetailsGridload()
        Dim rs As New ADODB.Recordset
        rs.Open "select * from grndetails_details where grnno= '" & txtgrnnoedit.Text & "'", cn, adOpenKeyset, adLockOptimistic
        i = 1
        If rs.BOF = False Then rs.MoveFirst
        While rs.EOF = False
        DetailsGrid.Rows = DetailsGrid.Rows + 1
        DetailsGrid.TextMatrix(i, 0) = rs![pono]
        DetailsGrid.TextMatrix(i, 1) = rs![podates]
        DetailsGrid.TextMatrix(i, 2) = rs![grnno]
        DetailsGrid.TextMatrix(i, 3) = rs![grndate]
        DetailsGrid.TextMatrix(i, 4) = rs![supname]
        DetailsGrid.TextMatrix(i, 5) = rs![deptname]
        rs.MoveNext
        i = i + 1
        Wend
        DetailsGrid.Rows = rs.RecordCount + 1
End Function
Private Function Detailsreceiveqtyload()
    Dim rs As New ADODB.Recordset
    Dim exp1 As String
    rs.Open "Select * from grn_details", cn, adOpenKeyset, adLockOptimistic
        If rs.RecordCount = 0 Then
        DetailsGrid.TextMatrix(DetailsGrid.Row, 6) = 0
        rs.Close
    Else
    If DeliveryGrid.Rows > 1 Then
    Dim rsrs As New ADODB.Recordset
    rsrs.Open "Select sum(recqty)as exp1 from grn_details where grnno = '" & txtgrnnoedit.Text & "'", cn, adOpenKeyset, adLockOptimistic
       DetailsGrid.TextMatrix(DetailsGrid.Row, 6) = Format(rsrs![exp1], "0.000")
       rsrs.Close
    End If
    End If
    SendKeys "{tab}"
End Function
Private Function txttypeload()
        txttype.Text = "GRN(PO) Against Delivery"
        txttypes.Text = 1
End Function
Private Function dupicates()
    additem = True
 If Len(Trim(txtid.Text)) = 0 Then txtid.SetFocus: Exit Function
        lstid.Text = txtid.Text
    If lstid.ListIndex > -1 Then
        MsgBox "This Quantity Already Received! So Cann't Added", vbInformation, "Received Qty Error"
    Exit Function
    End If
        lstid.additem txtid.Text
        DeliveryMainGrid.Rows = DeliveryMainGrid.Rows + 1
        DeliveryMainGrid.Row = DeliveryMainGrid.Rows - 1
        Call DeliveryMainGridload
        additem = False
End Function
Private Function DeliveryMainGridload()
If Trim(txtid.Text) = "" Then
        MsgBox "Please Enter The Issue Qty Details", vbCritical, "Issue Qty Error"
    Else
        DeliveryMainGrid.TextMatrix(DeliveryMainGrid.Row, 0) = DeliveryMainGrid.Rows - 1
        DeliveryMainGrid.TextMatrix(DeliveryMainGrid.Row, 1) = (DeliveryGrid.TextMatrix(DeliveryGrid.Row, 1))
        DeliveryMainGrid.TextMatrix(DeliveryMainGrid.Row, 2) = (DeliveryGrid.TextMatrix(DeliveryGrid.Row, 2))
        DeliveryMainGrid.TextMatrix(DeliveryMainGrid.Row, 3) = (DeliveryGrid.TextMatrix(DeliveryGrid.Row, 3))
        DeliveryMainGrid.TextMatrix(DeliveryMainGrid.Row, 4) = (DeliveryGrid.TextMatrix(DeliveryGrid.Row, 4))
        DeliveryMainGrid.TextMatrix(DeliveryMainGrid.Row, 5) = (DeliveryGrid.TextMatrix(DeliveryGrid.Row, 5))
        DeliveryMainGrid.TextMatrix(DeliveryMainGrid.Row, 6) = (DeliveryGrid.TextMatrix(DeliveryGrid.Row, 6))
        DeliveryMainGrid.TextMatrix(DeliveryMainGrid.Row, 7) = (DeliveryGrid.TextMatrix(DeliveryGrid.Row, 7))
        DeliveryMainGrid.TextMatrix(DeliveryMainGrid.Row, 8) = Format(DeliveryGrid.TextMatrix(DeliveryGrid.Row, 8), "0.000")
        DeliveryMainGrid.TextMatrix(DeliveryMainGrid.Row, 9) = DeliveryGrid.TextMatrix(DeliveryGrid.Row, 9)
End If
End Function
Private Sub Form_Unload(Cancel As Integer)
    Unload Me
    frmdeliverymain.Show
End Sub
Private Sub supopt_Click()
    cbosup.Visible = True
    DeliveryGrid.Rows = 1
    DeliveryMainGrid.Rows = 1
    DetailsGrid.Rows = 1
    cbosup.Clear
End Sub
Private Sub txtdcno_GotFocus()
    Dim rs As New ADODB.Recordset
    Dim a As String
    rs.Open "Select * from delivery_details", cn, adOpenKeyset, adLockOptimistic
        If rs.RecordCount = 0 Then
        txtdcno.Text = 1
        rs.Close
    Else
    Dim rsrs As New ADODB.Recordset
    rsrs.Open "Select max(dcno)as exp1 from delivery_details", cn, adOpenKeyset, adLockOptimistic
       txtdcno = rsrs![exp1] + 1
       rsrs.Close
    End If
    SendKeys "{tab}"
End Sub
Private Sub txtrec_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Then
    ElseIf (KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 46 Then
    ElseIf KeyAscii = 13 And DeliveryGrid.Col = 8 Then
        Call txtrec_OnEnter
        DeliveryGrid.Col = 10
        DeliveryGrid.SetFocus
        txtrec.Visible = False
    Else
        KeyAscii = 0
End If
End Sub
Private Sub txtrec_OnEnter()
If Val(txtrec.Text) > Val(txtbal.Text) Then
        MsgBox "Please Check The Stock Quantity Against GRN No!", vbCritical, "Stock Quantity Error"
        txtrec.Text = "0"
        txtrec.SetFocus
    Else
        Me.DeliveryGrid.Text = Format(Me.txtrec.Text, "0.000")
   'Call calculates
    Call txttotcalculate
End If
End Sub
Private Function txttotcalculate()
    Dim i As Integer
    txttots = 0
    For i = 0 To DeliveryGrid.Rows - 1
    txttots.Text = Format(Val(txttots.Text) + Val(DeliveryGrid.TextMatrix(i, 8)), "0.000")
    Next i
End Function
Private Function totalgrngrid()
        Dim rs As New ADODB.Recordset
        Set rs = cn.Execute("select * from grntotal_details where grnno = '" & cbogrnno.Text & "'")
        i = 1
        If rs.BOF = False Then rs.MoveFirst
        While rs.EOF = False
        DeliveryGrid.TextMatrix(i, 10) = Format(rs![SumOfdelqty], "0.000")
        DeliveryGrid.TextMatrix(i, 6) = Format(Val(DeliveryGrid.TextMatrix(i, 6)) - Val(DeliveryGrid.TextMatrix(i, 10)), "0.000")
        rs.MoveNext
        i = i + 1
        Wend
End Function
