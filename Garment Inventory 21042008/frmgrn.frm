VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmgrn 
   BackColor       =   &H00EDDDD1&
   Caption         =   " *Goods Receipt Details  ( GRN ) *"
   ClientHeight    =   9450
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10485
   Icon            =   "frmgrn.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   11115
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
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
      Height          =   255
      Left            =   4440
      TabIndex        =   32
      Top             =   2520
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtremarks 
      Appearance      =   0  'Flat
      Height          =   735
      Left            =   1320
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   30
      Top             =   9600
      Width           =   10215
   End
   Begin VB.TextBox txttype 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   7320
      TabIndex        =   29
      Text            =   "PO Against GRN"
      Top             =   5160
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox txtgrs 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10080
      Locked          =   -1  'True
      TabIndex        =   28
      Top             =   360
      Width           =   1575
   End
   Begin VB.TextBox txttots 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   10800
      TabIndex        =   26
      Top             =   5400
      Visible         =   0   'False
      Width           =   1215
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
      Height          =   495
      Left            =   11760
      MaskColor       =   &H00C0FFC0&
      Style           =   1  'Graphical
      TabIndex        =   25
      Tag             =   " "
      ToolTipText     =   " To Use Add PO "
      Top             =   9720
      UseMaskColor    =   -1  'True
      Width           =   1575
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
      TabIndex        =   24
      ToolTipText     =   "  Exit Window  "
      Top             =   9720
      UseMaskColor    =   -1  'True
      Width           =   1575
   End
   Begin VB.TextBox txtids 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   10800
      TabIndex        =   23
      Top             =   5160
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtstock 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   9840
      TabIndex        =   21
      Top             =   5400
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox txttot 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   9840
      TabIndex        =   20
      Top             =   5160
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox txtbal 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   8640
      TabIndex        =   19
      Top             =   5400
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtid 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   8640
      TabIndex        =   18
      Top             =   5160
      Visible         =   0   'False
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
      TabIndex        =   17
      ToolTipText     =   " To Use Add The GRN Item "
      Top             =   5160
      Width           =   1215
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
      TabIndex        =   16
      ToolTipText     =   "  To Use Delete  The GRN Item "
      Top             =   5160
      Width           =   1215
   End
   Begin VB.TextBox txtdept 
      Alignment       =   2  'Center
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
      Height          =   375
      Left            =   4680
      TabIndex        =   13
      Top             =   720
      Width           =   1935
   End
   Begin MSComCtl2.DTPicker dt2 
      Height          =   375
      Left            =   12960
      TabIndex        =   11
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
      Format          =   20250625
      CurrentDate     =   39481
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
      Left            =   10080
      TabIndex        =   9
      ToolTipText     =   " Enter The Supplier DC NO "
      Top             =   720
      Width           =   1575
   End
   Begin VB.ComboBox cbopono 
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
      Left            =   1320
      TabIndex        =   7
      ToolTipText     =   " Select The PO No "
      Top             =   720
      Width           =   1815
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
      Left            =   1920
      TabIndex        =   4
      ToolTipText     =   " Select The Supplier Name "
      Top             =   240
      Width           =   4695
   End
   Begin MSComCtl2.DTPicker dt1 
      Height          =   375
      Left            =   12960
      TabIndex        =   2
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
      Format          =   20250625
      CurrentDate     =   39481
   End
   Begin VB.TextBox txtgrnno 
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
      Left            =   10080
      TabIndex        =   0
      ToolTipText     =   " Your GRN No "
      Top             =   360
      Width           =   1575
   End
   Begin MSFlexGridLib.MSFlexGrid GrnMainGrid 
      Height          =   3735
      Left            =   120
      TabIndex        =   15
      Top             =   5760
      Width           =   15015
      _ExtentX        =   26485
      _ExtentY        =   6588
      _Version        =   393216
      Rows            =   1
      Cols            =   12
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
   Begin VB.ListBox lstid 
      Appearance      =   0  'Flat
      Height          =   615
      Left            =   4920
      TabIndex        =   22
      Top             =   5040
      Visible         =   0   'False
      Width           =   1935
   End
   Begin MSFlexGridLib.MSFlexGrid GrnEditGrid 
      Height          =   3615
      Left            =   120
      TabIndex        =   27
      Top             =   1440
      Width           =   15015
      _ExtentX        =   26485
      _ExtentY        =   6376
      _Version        =   393216
      Rows            =   1
      Cols            =   13
      FixedCols       =   0
      BackColor       =   16777215
      BackColorFixed  =   12542735
      ForeColorFixed  =   16777215
      BackColorSel    =   12250607
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
   Begin MSFlexGridLib.MSFlexGrid GrnGrid 
      Height          =   3615
      Left            =   120
      TabIndex        =   14
      ToolTipText     =   "  Note : # Indicate Place Click The First Row of Grid To  Open the Filter Options "
      Top             =   1440
      Width           =   15015
      _ExtentX        =   26485
      _ExtentY        =   6376
      _Version        =   393216
      Rows            =   1
      Cols            =   12
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
   Begin VB.Shape Shape1 
      BorderColor     =   &H000000FF&
      Height          =   975
      Left            =   120
      Top             =   8520
      Width           =   15015
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
      TabIndex        =   31
      Top             =   9600
      Width           =   1215
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FF0000&
      Height          =   735
      Left            =   11640
      Top             =   9600
      Width           =   3495
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H000000FF&
      Height          =   495
      Left            =   120
      Top             =   5160
      Width           =   12375
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H000000FF&
      Height          =   1215
      Index           =   1
      Left            =   120
      Top             =   120
      Width           =   8295
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H000000FF&
      Height          =   1215
      Index           =   0
      Left            =   8520
      Top             =   120
      Width           =   6615
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "PO No "
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
      Left            =   240
      TabIndex        =   12
      Top             =   720
      Width           =   1095
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
      Index           =   2
      Left            =   11640
      TabIndex        =   10
      Top             =   720
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "DC No"
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
      Left            =   9120
      TabIndex        =   8
      Top             =   720
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Department"
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
      Left            =   3240
      TabIndex        =   6
      Top             =   720
      Width           =   1455
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Supplier Name "
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
      Left            =   240
      TabIndex        =   5
      Top             =   240
      Width           =   1695
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "GRN Date"
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
      Left            =   11640
      TabIndex        =   3
      Top             =   360
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "GRN NO"
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
      Left            =   9120
      TabIndex        =   1
      Top             =   360
      Width           =   975
   End
End
Attribute VB_Name = "frmgrn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim op As Variant
Dim cn As ADODB.Connection
Dim rs As ADODB.Recordset
Dim ww As ADODB.Recordset
Dim additem As Boolean
Dim ms As Variant
Private Sub cbopono_Click()
    Call txtdept_GotFocus
    Call grngridload
    Call totalgrngrid
    For i = 1 To GrnGrid.Rows - 1
    GrnGrid.Col = 9
    GrnGrid.Row = i
    GrnGrid.CellBackColor = &HEDDDD1
    Next i
End Sub
Private Sub cbopono_dropdown()
    Dim rsloaddcno As New ADODB.Recordset
    Set rsloaddcno = cn.Execute("SELECT pono FROM postatus_details where postatus = 'Open' AND supname ='" & Trim(cbosupplier.Text & "'"))
    cbopono.Clear
    If rsloaddcno.BOF = False Then rsloaddcno.MoveFirst
    While rsloaddcno.EOF = False
        cbopono.additem rsloaddcno![pono]
        rsloaddcno.MoveNext
    Wend
End Sub
Private Sub cbopono_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub cbosupplier_Click()
    GrnMainGrid.Rows = 1
    GrnGrid.Rows = 1
    lstid.Clear
End Sub
Private Sub cbosupplier_DropDown()
    Dim rs As New ADODB.Recordset
        On Error GoTo X
        cbosupplier.Clear
        Set rs = cn.Execute("select supname from postatus_details where postatus = 'Open' group by supname")
        If rs.RecordCount = 0 Then
        MsgBox "No Records Found", vbInformation, "Information"
        Else
        rs.MoveFirst
        Do While Not rs.EOF()
        cbosupplier.additem (rs(0))
        rs.MoveNext
        Loop
        cbosupplier.SetFocus
X:
End If
End Sub
Private Sub cbosupplier_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub cmdexits_Click()
    ms = MsgBox("Are You Sure To Close ?", vbYesNo + vbQuestion, "Confirm Close ?")
    If ms = vbYes Then
    Unload Me
    frmgrnmain.Show
    Else
    End If
End Sub
Private Sub cmdsave_Click()
    Dim i As Integer
    If Trim(cbosupplier.Text) = "" Then
    MsgBox "Please Select the Supplier Name ", vbCritical, "Supplier Name Error"
    cbosupplier.SetFocus
    ElseIf Trim(cbopono.Text) = "" Then
    MsgBox "Please Select The PO NO ", vbCritical, "Po No Error"
    cbopono.SetFocus
    ElseIf Trim(txtdc.Text) = "" Then
    MsgBox "Please Enter The Supplier DC No ", vbCritical, "Dc No Error"
    txtdc.SetFocus
    ElseIf Val(txttots.Text) <= 0 Then
    MsgBox "Please Enter The Valid Received Quantity Details", vbCritical, "Received Quantity Error"
    GrnGrid.Col = 7
    GrnGrid.SetFocus
    ElseIf GrnMainGrid.Rows <= 1 Then
    MsgBox "Please Add The Received Items", vbCritical, "Received Qty Error"
    GrnGrid.SetFocus
    Else
        Dim rs As New ADODB.Recordset
        Dim rsstatus As New ADODB.Recordset
        
        rs.Open "select * from grn_details where grnno =' " & txtgrnno.Text & "'", cn, adOpenKeyset, adLockOptimistic
        rsstatus.Open "Select * from grnstatus_details where grnno= '" & txtgrnno.Text & "'", cn, adOpenKeyset, adLockOptimistic
        
        If rsstatus.RecordCount = 0 Then
        rsstatus.AddNew
        rsstatus![grnno] = txtgrnno.Text
        rsstatus![grndate] = dt1.Value
        rsstatus![grnstatus] = "Open"
        rsstatus![supname] = cbosupplier.Text
        rsstatus![deptname] = txtdept.Text
        rsstatus![dcno] = txtdc.Text
        rsstatus.Update
        rsstatus.Close
        End If
        
        If rs.RecordCount = 0 Then
        For i = 1 To GrnMainGrid.Rows - 1
        rs.AddNew
        rs![grnno] = txtgrnno.Text
        rs![grndate] = dt1.Value
        rs![dcno] = txtdc.Text
        rs![dcdate] = dt2.Value
        rs![supname] = cbosupplier.Text
        rs![deptname] = txtdept.Text
        rs![pono] = cbopono.Text
        rs![sno] = GrnMainGrid.TextMatrix(i, 0)
        rs![pono] = GrnMainGrid.TextMatrix(i, 1)
        rs![podates] = GrnMainGrid.TextMatrix(i, 2)
        rs![itemname] = GrnMainGrid.TextMatrix(i, 3)
        rs![colour] = GrnMainGrid.TextMatrix(i, 4)
        rs![sizes] = GrnMainGrid.TextMatrix(i, 5)
        rs![poqty] = GrnMainGrid.TextMatrix(i, 6)
        rs![uom] = GrnMainGrid.TextMatrix(i, 7)
        rs![balqty] = GrnMainGrid.TextMatrix(i, 8)
        rs![recqty] = GrnMainGrid.TextMatrix(i, 9)
        rs![poid] = GrnMainGrid.TextMatrix(i, 10)
        rs![types] = txttype.Text
        rs![remarks] = txtremarks.Text
        Next i
        
        rs.Update
        rs.Close
        MsgBox "One Record Save Successfully ", vbInformation, "Information"
        Unload Me
        frmgrnmain.Show
        Else
        MsgBox "Already Exists", vbCritical, "Error"
    End If
    End If
End Sub
Private Sub cmdadditem_Click()
    If Trim(txtid.Text) = "" Then
    MsgBox "Please Enter The Grn Valid data", vbCritical, "Error"
    GrnGrid.SetFocus
    ElseIf Val(txtrec.Text) <= 0 Then
    MsgBox "Please Enter Valid Received Quantity ", vbCritical, "Received Qty Error"
    GrnGrid.SetFocus
    ElseIf Val(GrnGrid.TextMatrix(GrnGrid.Row, 9)) <= 0 Then
    MsgBox "Please Enter The Valid Qty", vbCritical, "Qty Error"
    txtrec.SetFocus
    
    ElseIf GrnMainGrid.Rows > 1 Then
    Call dupicates
    Else
    lstid.additem txtid.Text
    GrnMainGrid.Rows = GrnMainGrid.Rows + 1
    GrnMainGrid.Row = GrnMainGrid.Rows - 1
    Call grnmaingridload
    End If
End Sub
Private Sub cmdadditem_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
    If GrnMainGrid.Rows = 1 Then
    lstid.Clear
    End If
End Sub
Private Sub cmddeleteitem_Click()
    If Trim(txtids.Text) = "" And GrnMainGrid.Rows > 1 Then
    MsgBox "Please Select The Itemname For Deletion", vbInformation, "Delete Row"
    GrnMainGrid.SetFocus
    ElseIf GrnMainGrid.Rows > 2 Then
            GrnMainGrid.RemoveItem (GrnMainGrid.Row)
            GrnMainGrid.TextMatrix(GrnMainGrid.Row, 0) = GrnMainGrid.Rows - 1
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
Private Sub cmdupdate_Click()
        Dim i As Integer
        Dim rs As New ADODB.Recordset
        
        If Trim(txtdc.Text) = "" Then
        MsgBox "Please Enter The Supplier Dc No", vbCritical, "Supplier Dc No Error"
        txtdc.SetFocus
        ElseIf Val(txtedittot.Text) <= 0 Then
        MsgBox "Please Enter The Valid Received Qty Details ", vbCritical, "Received Qty Error"
        txtedittot.SetFocus
        Else
        rs.Open "Select * from grn_details where grnno = '" & txtgrs.Text & "'", cn, 1, 3
        i = 1
        If rs.RecordCount <> 0 Then
         If rs.BOF = False Then rs.MoveFirst
         While rs.EOF = False
        'rs.AddNew
        rs![grnno] = txtgrs.Text
        rs![grndate] = dt1.Value
        rs![dcno] = txtdc.Text
        rs![dcdate] = dt2.Value
        rs![supname] = cbosupplier.Text
        rs![deptname] = txtdept.Text
        rs![pono] = cbopono.Text
        rs![sno] = GrnEditGrid.TextMatrix(i, 0)
        rs![pono] = GrnEditGrid.TextMatrix(i, 1)
        rs![podates] = GrnEditGrid.TextMatrix(i, 2)
        rs![itemname] = GrnEditGrid.TextMatrix(i, 3)
        rs![colour] = GrnEditGrid.TextMatrix(i, 4)
        rs![sizes] = GrnEditGrid.TextMatrix(i, 5)
        rs![poqty] = GrnEditGrid.TextMatrix(i, 6)
        rs![uom] = GrnEditGrid.TextMatrix(i, 7)
        rs![balqty] = GrnEditGrid.TextMatrix(i, 8)
        rs![recqty] = GrnEditGrid.TextMatrix(i, 9)
        rs![poid] = GrnEditGrid.TextMatrix(i, 10)
        rs![types] = txttype.Text
        rs![remarks] = txtremarks.Text
        rs.Update
        rs.MoveNext
        
        i = i + 1
        Wend
        rs.Close
        
        MsgBox "One Record Save Successfully ", vbInformation, "Information"
        Unload Me
        frmgrnmain.Show
        frmmain.WindowState = 2
        Else
        MsgBox "Already Exists", vbCritical, "Error"
    End If
        
        End If
End Sub

Private Sub Form_Load()
        
    Set cn = New ADODB.Connection
    Set ww = New ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    cn.Open "Provider=Microsoft.jet.oledb.4.0; Data Source=" & App.path & "\Database\Data.mdb"
    ww.Open "Select * From grn_details", cn, adOpenKeyset, adLockOptimistic
    cn.CursorLocation = adUseClient
     
    Call frmgrngrid
    Call frmgrnmaingrid
    Call txtgrnno_Gotfocus
    
    GrnMainGrid.Height = 3735
    GrnMainGrid.Width = 15015
End Sub
Private Sub Form_Unload(Cancel As Integer)
   If Cancel = 0 Then
   frmgrnmain.Show
   Else
   Cancel = 1
   End If
End Sub

Private Sub GrnGrid_Click()
    If txtrec.Visible = True Then txtrec.Visible = False
    If GrnGrid.Col = 9 Then
    Me.txtrec.Visible = True
    CurrentRow = Me.GrnGrid.Row
    Me.txtrec.Width = Me.GrnGrid.CellWidth - 10
    Me.txtrec.Left = Me.GrnGrid.CellLeft + Me.GrnGrid.Left
    Me.txtrec.Top = Me.GrnGrid.CellTop + Me.GrnGrid.Top
    donotchange = True
    Me.txtrec.Text = Me.GrnGrid.Text
    donotchange = False
    Me.txtrec.SetFocus
    End If
    If GrnGrid.Col = 8 Or GrnGrid.Col = 7 Or GrnGrid.Col = 2 Or GrnGrid.Col = 4 Or GrnGrid.Col = 5 Or GrnGrid.Col = 6 Or GrnGrid.Col = 8 Or GrnGrid.Col = 9 Or GrnGrid.Col = 3 Then
        txtbal.Text = GrnGrid.TextMatrix(GrnGrid.Row, 8)
        txtid.Text = GrnGrid.TextMatrix(GrnGrid.Row, 10)
        txtstock.Text = GrnGrid.TextMatrix(GrnGrid.Row, 8)
    End If
End Sub
Private Sub GrnMainGrid_Click()
    If GrnMainGrid.Rows > 1 Then
    If GrnMainGrid.Col = 8 Or GrnMainGrid.Col = 7 Or GrnMainGrid.Col = 2 Or GrnMainGrid.Col = 4 Or GrnMainGrid.Col = 5 Or GrnMainGrid.Col = 6 Or GrnMainGrid.Col = 8 Or GrnMainGrid.Col = 9 Or GrnMainGrid.Col = 3 Or GrnMainGrid.Col = 10 Or GrnMainGrid.Col = 11 Or GrnMainGrid.Col = 0 Or GrnMainGrid.Col = 1 Or GrnMainGrid.Col = 2 Or GrnMainGrid.Col = 3 Then
         txtids.Text = GrnMainGrid.TextMatrix(GrnMainGrid.Row, 10)
    End If
    End If
End Sub
Private Sub txtdept_GotFocus()
    Dim rs As New ADODB.Recordset
    Dim a As String
    rs.Open "Select * from postatus_details", cn, adOpenKeyset, adLockOptimistic
        If rs.RecordCount = 0 Then
        txtdept.Text = ""
        rs.Close
        Else
        Dim rsrs As New ADODB.Recordset
        rsrs.Open "Select deptname as exp1 from postatus_details where pono='" & cbopono.Text & "'", cn, adOpenKeyset, adLockOptimistic
        txtdept = rsrs![exp1]
        rsrs.Close
        End If
    SendKeys "{tab}"
End Sub
Private Function grngridload()
    Dim i As Integer
    Dim rs As New ADODB.Recordset
    Dim rs1 As New ADODB.Recordset
    rs.Open "select * from po_details where pono= " & Trim(cbopono.Text), cn, adOpenKeyset, adLockOptimistic
    i = 1
    If rs.BOF = False Then rs.MoveFirst
    While rs.EOF = False
    GrnGrid.Rows = rs.RecordCount + 1
    GrnGrid.TextMatrix(i, 0) = i
    GrnGrid.TextMatrix(i, 1) = rs![pono]
    GrnGrid.TextMatrix(i, 2) = rs![podate]
    GrnGrid.TextMatrix(i, 3) = rs![itemname]
    GrnGrid.TextMatrix(i, 4) = rs![colour]
    GrnGrid.TextMatrix(i, 5) = rs![sizes]
    GrnGrid.TextMatrix(i, 6) = Format(rs![qty], "0.000")
    GrnGrid.TextMatrix(i, 7) = rs![uom]
    GrnGrid.TextMatrix(i, 8) = Format(rs![qty], "0.000")
    GrnGrid.TextMatrix(i, 9) = Format("0.000")
    GrnGrid.TextMatrix(i, 10) = rs![id]
    rs.MoveNext
    i = i + 1
    Wend
    GrnGrid.Rows = rs.RecordCount + 1
End Function
Private Function grnmaingridload()
    If Trim(txtid.Text) = "" Then
    MsgBox "Please Enter The Issue Qty Details", vbCritical, "Issue Qty Error"
    Else
        GrnMainGrid.TextMatrix(GrnMainGrid.Row, 0) = GrnMainGrid.Rows - 1
        GrnMainGrid.TextMatrix(GrnMainGrid.Row, 1) = (GrnGrid.TextMatrix(GrnGrid.Row, 1))
        GrnMainGrid.TextMatrix(GrnMainGrid.Row, 2) = (GrnGrid.TextMatrix(GrnGrid.Row, 2))
        GrnMainGrid.TextMatrix(GrnMainGrid.Row, 3) = (GrnGrid.TextMatrix(GrnGrid.Row, 3))
        GrnMainGrid.TextMatrix(GrnMainGrid.Row, 4) = (GrnGrid.TextMatrix(GrnGrid.Row, 4))
        GrnMainGrid.TextMatrix(GrnMainGrid.Row, 5) = (GrnGrid.TextMatrix(GrnGrid.Row, 5))
        GrnMainGrid.TextMatrix(GrnMainGrid.Row, 6) = (GrnGrid.TextMatrix(GrnGrid.Row, 6))
        GrnMainGrid.TextMatrix(GrnMainGrid.Row, 7) = (GrnGrid.TextMatrix(GrnGrid.Row, 7))
        GrnMainGrid.TextMatrix(GrnMainGrid.Row, 8) = Format(GrnGrid.TextMatrix(GrnGrid.Row, 8), "0.000")
        GrnMainGrid.TextMatrix(GrnMainGrid.Row, 9) = Format(GrnGrid.TextMatrix(GrnGrid.Row, 9), "0.000")
        GrnMainGrid.TextMatrix(GrnMainGrid.Row, 10) = (GrnGrid.TextMatrix(GrnGrid.Row, 10))
    End If
End Function
Private Function calculates()
    Dim i As Integer
    txttot = 0
    For i = 0 To GrnGrid.Rows - 1
    txttot.Text = Format(Val(txttot.Text) + Val(GrnGrid.TextMatrix(i, 9)), "0.000")
    Next i
End Function
Private Sub txtrec_changes()
    If Val(txtrec.Text) > Val(txtstock.Text) Or Val(txtrec.Text) < 0 Then
    MsgBox "Please Check The Stock Quantity Against GRN No!", vbCritical, "Stock Quantity Error"
    txtrec.Text = "0"
    txtrec.SetFocus
    Else
    Me.GrnGrid.Text = Format(Me.txtrec.Text, "0.000")
    Call calculates
    Call txttotcalculate
    End If
End Sub
Private Sub txtgrnno_Gotfocus()
     Dim rs As New ADODB.Recordset
        Dim a As String
        rs.Open "Select * from grn_details", cn, adOpenKeyset, adLockOptimistic
        If rs.RecordCount = 0 Then
        txtgrnno.Text = 1
        rs.Close
        Else
        Dim rsrs As New ADODB.Recordset
        rsrs.Open "Select max(grnno) as exp1 from grn_details ", cn, adOpenKeyset, adLockOptimistic
        txtgrnno = rsrs![exp1] + 1
        rsrs.Close
        End If
   SendKeys "{tab}"
End Sub
Private Sub txtrec_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Then
    ElseIf (KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 46 Then
    ElseIf KeyAscii = 13 And GrnGrid.Col = 9 Then
    Call txtrec_changes
    GrnGrid.Col = 10
    GrnGrid.SetFocus
    txtrec.Visible = False
Else
KeyAscii = 0
End If
End Sub
Private Function dupicates()
    additem = True
    If Len(Trim(txtid.Text)) = 0 Then txtid.SetFocus: Exit Function
    lstid.Text = txtid.Text
    If lstid.ListIndex > -1 Then
    MsgBox "This Quantity Already Received! So Cann't Added", vbInformation, "Received Qty Error"
    Exit Function
    End If
    lstid.additem txtid.Text
    GrnMainGrid.Rows = GrnMainGrid.Rows + 1
    GrnMainGrid.Row = GrnMainGrid.Rows - 1
    Call grnmaingridload
    additem = False
End Function
Private Function txttotcalculate()
    Dim i As Integer
    txttots = 0
    For i = 0 To GrnGrid.Rows - 1
    txttots.Text = Format(Val(txttots.Text) + Val(GrnGrid.TextMatrix(i, 9)), "0.000")
    Next i
End Function
Sub totalgrngrid()
    Dim i As Integer
    Dim rs As New ADODB.Recordset
    rs.Open "Select * from pomas_details where pono = '" & cbopono.Text & "'", cn, adOpenKeyset, adLockOptimistic
    i = 1
    If rs.BOF = False Then rs.MoveFirst
    While rs.EOF = False
        GrnGrid.TextMatrix(i, 11) = Format(rs![SumOfrecqty], "0.000")
        GrnGrid.TextMatrix(i, 8) = Format(Val(GrnGrid.TextMatrix(i, 6)) - Val(GrnGrid.TextMatrix(i, 11)), "0.000")
        rs.MoveNext
        i = i + 1
    Wend
End Sub

