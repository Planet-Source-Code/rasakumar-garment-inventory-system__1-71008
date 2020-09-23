VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frminvoice 
   BackColor       =   &H00EDDDD1&
   Caption         =   " * Invoice Details *"
   ClientHeight    =   9135
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10830
   Icon            =   "frminvoice.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   11115
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtnetamt 
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
      Left            =   12600
      Locked          =   -1  'True
      TabIndex        =   36
      Top             =   7200
      Width           =   2535
   End
   Begin VB.TextBox txttaxamt 
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
      Left            =   12600
      TabIndex        =   35
      Top             =   6360
      Width           =   2535
   End
   Begin VB.TextBox txtword 
      Appearance      =   0  'Flat
      Height          =   1575
      Left            =   12600
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   32
      Top             =   7920
      Width           =   2535
   End
   Begin VB.TextBox txttotamt 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Left            =   9840
      Locked          =   -1  'True
      TabIndex        =   31
      Top             =   9000
      Width           =   2535
   End
   Begin VB.TextBox txtrate 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   4080
      TabIndex        =   30
      Top             =   2520
      Visible         =   0   'False
      Width           =   975
   End
   Begin MSComCtl2.DTPicker dt2 
      Height          =   375
      Left            =   6360
      TabIndex        =   29
      Top             =   720
      Width           =   2055
      _ExtentX        =   3625
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
      CurrentDate     =   39547
   End
   Begin VB.TextBox txtsupbillno 
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
      Left            =   6360
      TabIndex        =   28
      ToolTipText     =   " Enter The Supplier Bill Number "
      Top             =   240
      Width           =   2055
   End
   Begin VB.TextBox txtgrnnoedit 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   8640
      TabIndex        =   17
      Top             =   4200
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txttype 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   8640
      TabIndex        =   16
      Top             =   4440
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtinvoiceno 
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
      Left            =   1800
      TabIndex        =   15
      Top             =   240
      Width           =   2055
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
      Left            =   12720
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   " To Add invoice Item "
      Top             =   5160
      Width           =   1095
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
      ToolTipText     =   " To Delete  invoice Item "
      Top             =   5160
      Width           =   1095
   End
   Begin VB.ListBox lstid 
      Appearance      =   0  'Flat
      Height          =   420
      Left            =   12960
      TabIndex        =   12
      Top             =   4200
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.TextBox txtid 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   9720
      TabIndex        =   11
      Top             =   4200
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdsave 
      BackColor       =   &H00FFC0C0&
      Caption         =   "&Save Invoice"
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
      TabIndex        =   10
      Tag             =   " "
      ToolTipText     =   " To Use Save Invoice "
      Top             =   9720
      UseMaskColor    =   -1  'True
      Width           =   1455
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
      TabIndex        =   9
      ToolTipText     =   "  Exit Window  "
      Top             =   9720
      UseMaskColor    =   -1  'True
      Width           =   1575
   End
   Begin VB.TextBox txtids 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   9720
      TabIndex        =   8
      Top             =   4440
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtbal 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   10800
      TabIndex        =   7
      Top             =   4200
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtqty 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   2880
      TabIndex        =   6
      Top             =   2475
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txttots 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   10800
      TabIndex        =   5
      Text            =   "0"
      Top             =   4440
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txttypes 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   11880
      TabIndex        =   3
      Top             =   4200
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtremarks 
      Appearance      =   0  'Flat
      Height          =   735
      Left            =   1320
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      ToolTipText     =   " Enter The Remarks "
      Top             =   9600
      Width           =   10335
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
      Left            =   11160
      TabIndex        =   1
      ToolTipText     =   " Select the GRN No "
      Top             =   720
      Width           =   3855
   End
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
      Left            =   11160
      TabIndex        =   0
      ToolTipText     =   " Select the Supplier Name "
      Top             =   240
      Width           =   3855
   End
   Begin MSComCtl2.DTPicker dt1 
      Height          =   375
      Left            =   1800
      TabIndex        =   4
      Top             =   720
      Width           =   2055
      _ExtentX        =   3625
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
   Begin MSFlexGridLib.MSFlexGrid InvoiceDetailsGrid 
      Height          =   735
      Left            =   240
      TabIndex        =   18
      ToolTipText     =   "  Note : # Indicate Place Click The First Row of Grid To  Open the Filter Options "
      Top             =   4080
      Width           =   14775
      _ExtentX        =   26061
      _ExtentY        =   1296
      _Version        =   393216
      Rows            =   1
      Cols            =   6
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
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid InvoiceMainGrid 
      Height          =   3735
      Left            =   240
      TabIndex        =   19
      ToolTipText     =   " "
      Top             =   5160
      Width           =   12135
      _ExtentX        =   21405
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
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid InvoiceGrid 
      Height          =   2775
      Left            =   240
      TabIndex        =   20
      Top             =   1200
      Width           =   14775
      _ExtentX        =   26061
      _ExtentY        =   4895
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
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H000000FF&
      Height          =   735
      Left            =   12600
      Top             =   5040
      Width           =   2535
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Net Amount ( Rs)"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   3
      Left            =   12600
      TabIndex        =   38
      Top             =   6840
      Width           =   2535
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Tax Amount ( Rs )"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   2
      Left            =   12600
      TabIndex        =   37
      Top             =   6000
      Width           =   2535
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Total Amount ( Rs)"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   1
      Left            =   7320
      TabIndex        =   34
      Top             =   9000
      Width           =   2535
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Amount in Words"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   12600
      TabIndex        =   33
      Top             =   7680
      Width           =   2535
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Sup.Bill.Date"
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
      Left            =   4800
      TabIndex        =   27
      Top             =   720
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Sup.Bill.No"
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
      Left            =   4800
      TabIndex        =   26
      Top             =   240
      Width           =   1575
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
      Index           =   5
      Left            =   9360
      TabIndex        =   25
      Top             =   240
      Width           =   1815
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H000000FF&
      Height          =   4815
      Index           =   0
      Left            =   120
      Top             =   120
      Width           =   15015
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Invoice No"
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
      Left            =   240
      TabIndex        =   24
      Top             =   240
      Width           =   1575
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FF00FF&
      Height          =   4455
      Left            =   120
      Top             =   5040
      Width           =   12375
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Invoice Date"
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
      TabIndex        =   23
      Top             =   720
      Width           =   1575
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FF0000&
      Height          =   735
      Left            =   11760
      Top             =   9600
      Width           =   3375
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
      TabIndex        =   22
      Top             =   9600
      Width           =   1215
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
      Left            =   9360
      TabIndex        =   21
      Top             =   720
      Width           =   1815
   End
End
Attribute VB_Name = "frminvoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim op As Variant
Dim rs As New ADODB.Recordset
Dim cn As New ADODB.Connection
Private Sub cbogrnno_Click()
    Call invoicegridload
    Call invoicedetailsgridload
    Call totalgrngrid
End Sub
Private Sub cbogrnno_dropdown()
    If Trim(cbosup.Text) = "" Then
    MsgBox "Please Select The Supplier Name ", vbCritical, "Supplier Name Selection Error"
    Else
    Dim rs As New ADODB.Recordset
    On Error GoTo X
        cbogrnno.Clear
        Set rs = cn.Execute("select grnno from grnstatus_details where grnstatus ='Open' and supname = '" & cbosup.Text & "'")
        rs.MoveFirst
        Do While Not rs.EOF()
        cbogrnno.additem (rs(0))
        rs.MoveNext
        Loop
        cbogrnno.SetFocus
X:
End If
End Sub
Private Sub cbogrnno_KeyPress(KeyAscii As Integer)
 KeyAscii = 0
End Sub

Private Sub cbosup_dropdown()
    Dim rs As New ADODB.Recordset
    On Error GoTo X
        cbosup.Clear
        Set rs = cn.Execute("select supname from grnstatus_details where grnstatus ='Open' Group by supname")
        rs.MoveFirst
        Do While Not rs.EOF()
        cbosup.additem (rs(0))
        rs.MoveNext
        Loop
        cbosup.SetFocus
X:
End Sub
Private Sub cbosup_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub cmdadditem_Click()
    If Trim(txtid.Text) = "" Then
        MsgBox "Please Enter The Grn Valid data", vbCritical, "Error"
        invoiceMainGrid.SetFocus
    ElseIf Val(InvoiceGrid.TextMatrix(InvoiceGrid.Row, 8)) <= 0 Then
        MsgBox "Please Enter The Valid Qty", vbCritical, "Qty Error"
        InvoiceGrid.Col = 8
        InvoiceGrid.SetFocus
    ElseIf Val(txtqty.Text) <= 0 Then
        MsgBox "Please Enter Valid Invoiced Quantity ", vbCritical, "Received Qty Error"
        InvoiceGrid.SetFocus
    ElseIf invoiceMainGrid.Rows > 1 Then
        Call dupicates
    Else
        lstid.additem txtid.Text
        invoiceMainGrid.Rows = invoiceMainGrid.Rows + 1
        invoiceMainGrid.Row = invoiceMainGrid.Rows - 1
        Call invoicemaingridload
        Call calstotalamt
    End If
End Sub

Private Sub cmdexits_Click()
    op = MsgBox("Are You Sure To Close ?", vbQuestion + vbYesNo, "Confirm Close ?")
        If op = vbYes Then
            Unload Me
            frminvoicemain.Show
        Else
        End If
End Sub
Private Sub cmdsave_Click()
    If Trim(txtsupbillno.Text) = "" Then
        MsgBox "Please Rnter The Supplier Bill Number", vbCritical, "Bill No Error"
        txtsupbillno.SetFocus
    ElseIf invoiceMainGrid.Rows <= 1 Then
        MsgBox "Please Enter The Valid Invoice Qty Details", vbCritical, "Invoice Qty Error"
    Else
        Dim rs As New ADODB.Recordset
        rs.Open "Select * from invoice_details where invoiceno= ' " & txtinvoiceno.Text & "'", cn, 1, 3
        If rs.RecordCount = 0 Then
            For i = 1 To invoiceMainGrid.Rows - 1
            rs.AddNew
            rs![invoiceno] = txtinvoiceno.Text
            rs![invoicedate] = dt1.Value
            rs![supbillno] = txtsupbillno.Text
            rs![supbilldate] = dt2.Value
            rs![supname] = cbosup.Text
            rs![grnno] = cbogrnno.Text
            rs![sno] = invoiceMainGrid.TextMatrix(i, 0)
            rs![itemname] = invoiceMainGrid.TextMatrix(i, 1)
            rs![colour] = invoiceMainGrid.TextMatrix(i, 2)
            rs![sizes] = invoiceMainGrid.TextMatrix(i, 3)
            rs![grnqty] = invoiceMainGrid.TextMatrix(i, 4)
            rs![uom] = invoiceMainGrid.TextMatrix(i, 5)
            rs![invoiceqty] = invoiceMainGrid.TextMatrix(i, 6)
            rs![invoicerate] = invoiceMainGrid.TextMatrix(i, 7)
            rs![totalamtgrid] = invoiceMainGrid.TextMatrix(i, 8)
            rs![grnid] = invoiceMainGrid.TextMatrix(i, 9)
            rs![totalamt] = txttotamt.Text
            rs![taxamt] = txttaxamt.Text
            rs![netamt] = txtnetamt.Text
            rs![words] = txtword.Text
            rs![remarks] = txtremarks.Text
            Next i
            rs.Update
            rs.Close
            MsgBox "One Record Save Successfully ", vbInformation, "Information"
            Unload Me
            frminvoicemain.Show
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
    ww.Open "Select * From invoice_details", cn, adOpenKeyset, adLockOptimistic
    
    Call frmgrninvoiceload
    Call txtinvoiceno_Gotfocus
End Sub
Private Sub InvoiceGrid_Click()
    If txtqty.Visible = True Then txtqty.Visible = False
    If txtrate.Visible = True Then txtrate.Visible = False
    
        If InvoiceGrid.Col = 6 Then
            Me.txtqty.Visible = True
            CurrentRow = Me.InvoiceGrid.Row
            Me.txtqty.Width = Me.InvoiceGrid.CellWidth - 10
            Me.txtqty.Left = Me.InvoiceGrid.CellLeft + Me.InvoiceGrid.Left
            Me.txtqty.Top = Me.InvoiceGrid.CellTop + Me.InvoiceGrid.Top
            donotchange = True
            Me.txtqty.Text = Me.InvoiceGrid.Text
            donotchange = False
            txtqty.SetFocus
        ElseIf InvoiceGrid.Col = 7 Then
            Me.txtrate.Visible = True
            CurrentRow = Me.InvoiceGrid.Row
            Me.txtrate.Width = Me.InvoiceGrid.CellWidth - 10
            Me.txtrate.Left = Me.InvoiceGrid.CellLeft + Me.InvoiceGrid.Left
            Me.txtrate.Top = Me.InvoiceGrid.CellTop + Me.InvoiceGrid.Top
            donotchange = True
            Me.txtrate.Text = Me.InvoiceGrid.Text
            donotchange = False
            txtrate.SetFocus
   End If
   If InvoiceGrid.Col = 8 Or InvoiceGrid.Col = 7 Or InvoiceGrid.Col = 2 Or InvoiceGrid.Col = 4 Or InvoiceGrid.Col = 5 Or InvoiceGrid.Col = 6 Or InvoiceGrid.Col = 8 Or InvoiceGrid.Col = 9 Or InvoiceGrid.Col = 3 Then
        txtbal.Text = InvoiceGrid.TextMatrix(InvoiceGrid.Row, 4)
        txtid.Text = InvoiceGrid.TextMatrix(InvoiceGrid.Row, 9)
       ' txtstock.Text = InvoiceGrid.TextMatrix(InvoiceGrid.Row, 8)
    End If
End Sub

Private Sub txtinvoiceno_Gotfocus()
    Dim rs As New ADODB.Recordset
    Dim a As String
    rs.Open "Select * from invoice_details", cn, adOpenKeyset, adLockOptimistic
        If rs.RecordCount = 0 Then
        txtinvoiceno.Text = 1
        rs.Close
    Else
    Dim rsrs As New ADODB.Recordset
    rsrs.Open "Select max(invoiceno)as exp1 from invoice_details", cn, adOpenKeyset, adLockOptimistic
       txtinvoiceno.Text = rsrs![exp1] + 1
       rsrs.Close
    End If
    SendKeys "{tab}"
End Sub
Private Function invoicegridload()
    Dim i As Integer
    Dim rs As New ADODB.Recordset
    Dim rs1 As New ADODB.Recordset
    rs.Open "select * from grn_details where grnno= '" & Trim(cbogrnno.Text) & "'", cn, adOpenKeyset, adLockOptimistic
    i = 1
    If rs.BOF = False Then rs.MoveFirst
    While rs.EOF = False
        InvoiceGrid.Rows = rs.RecordCount + 1
        InvoiceGrid.TextMatrix(i, 0) = i
        InvoiceGrid.TextMatrix(i, 1) = rs![itemname]
        InvoiceGrid.TextMatrix(i, 2) = rs![colour]
        InvoiceGrid.TextMatrix(i, 3) = rs![sizes]
        InvoiceGrid.TextMatrix(i, 4) = Format(rs![recqty], "0.000")
        InvoiceGrid.TextMatrix(i, 5) = rs![uom]
        InvoiceGrid.TextMatrix(i, 9) = rs![grnid]
        InvoiceGrid.TextMatrix(i, 11) = rs![grnno]
        InvoiceGrid.TextMatrix(i, 6) = 0
        InvoiceGrid.TextMatrix(i, 7) = Format("0.00")
        InvoiceGrid.TextMatrix(i, 8) = Format("0.00")
        rs.MoveNext
        i = i + 1
        Wend
        InvoiceGrid.Rows = rs.RecordCount + 1
End Function
Private Function invoicedetailsgridload()
        Dim rs As New ADODB.Recordset
        rs.Open "select * from grndetails_details where grnno= '" & cbogrnno.Text & "'", cn, adOpenKeyset, adLockOptimistic
        i = 1
        If rs.BOF = False Then rs.MoveFirst
        While rs.EOF = False
        InvoiceDetailsGrid.Rows = InvoiceDetailsGrid.Rows + 1
        InvoiceDetailsGrid.TextMatrix(i, 0) = rs![pono]
        InvoiceDetailsGrid.TextMatrix(i, 1) = rs![podates]
        InvoiceDetailsGrid.TextMatrix(i, 2) = rs![grnno]
        InvoiceDetailsGrid.TextMatrix(i, 3) = rs![grndate]
        InvoiceDetailsGrid.TextMatrix(i, 4) = rs![supname]
        InvoiceDetailsGrid.TextMatrix(i, 5) = rs![deptname]
        rs.MoveNext
        i = i + 1
        Wend
        InvoiceDetailsGrid.Rows = rs.RecordCount + 1
End Function
Private Sub txtqty_Change()
    If Val(txtqty.Text) > Val(InvoiceGrid.TextMatrix(InvoiceGrid.Row, 4)) Then
    MsgBox "Please Check The Stock Quantity Against GRN No!", vbCritical, "Stock Quantity Error"
    txtqty.Text = "0"
    txtqty.SetFocus
    Else
    Me.InvoiceGrid.Text = Format(Me.txtqty.Text, "0.000")
    End If
End Sub
Private Sub txtqty_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Then
    ElseIf (KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 46 Then
    ElseIf KeyAscii = 13 And InvoiceGrid.Col = 6 Then
        Me.InvoiceGrid.Text = Format(Me.txtqty.Text, "0.000")
        Call calculatetotalamout
        InvoiceGrid.Col = 7
        InvoiceGrid_Click
        InvoiceGrid.SetFocus
        txtrate.SetFocus
        txtqty.Visible = False
    Else
        KeyAscii = 0
End If
End Sub
Private Sub txtrate_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Then
    ElseIf (KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 46 Then
    ElseIf KeyAscii = 13 And InvoiceGrid.Col = 7 Then
        Me.InvoiceGrid.Text = Format(Me.txtrate.Text, "0.00")
        Call calculatetotalamout
        InvoiceGrid.Col = 8
        InvoiceGrid_Click
        InvoiceGrid.SetFocus
        txtrate.Visible = False
    Else
        KeyAscii = 0
End If
End Sub
Private Function calculatetotalamout()
    InvoiceGrid.TextMatrix(InvoiceGrid.Row, 8) = Format(Val(InvoiceGrid.TextMatrix(InvoiceGrid.Row, 6)) * Val(InvoiceGrid.TextMatrix(InvoiceGrid.Row, 7)), "0.00")
End Function
Private Function invoicemaingridload()
    If Trim(txtid.Text) = "" Then
        MsgBox "Please Enter The Invoice Qty Details", vbCritical, "Issue Qty Error"
    Else
        invoiceMainGrid.TextMatrix(invoiceMainGrid.Row, 0) = invoiceMainGrid.Rows - 1
        invoiceMainGrid.TextMatrix(invoiceMainGrid.Row, 1) = (InvoiceGrid.TextMatrix(InvoiceGrid.Row, 1))
        invoiceMainGrid.TextMatrix(invoiceMainGrid.Row, 2) = (InvoiceGrid.TextMatrix(InvoiceGrid.Row, 2))
        invoiceMainGrid.TextMatrix(invoiceMainGrid.Row, 3) = (InvoiceGrid.TextMatrix(InvoiceGrid.Row, 3))
        invoiceMainGrid.TextMatrix(invoiceMainGrid.Row, 4) = (InvoiceGrid.TextMatrix(InvoiceGrid.Row, 4))
        invoiceMainGrid.TextMatrix(invoiceMainGrid.Row, 5) = (InvoiceGrid.TextMatrix(InvoiceGrid.Row, 5))
        invoiceMainGrid.TextMatrix(invoiceMainGrid.Row, 6) = (InvoiceGrid.TextMatrix(InvoiceGrid.Row, 6))
        invoiceMainGrid.TextMatrix(invoiceMainGrid.Row, 7) = (InvoiceGrid.TextMatrix(InvoiceGrid.Row, 7))
        invoiceMainGrid.TextMatrix(invoiceMainGrid.Row, 8) = Format(InvoiceGrid.TextMatrix(InvoiceGrid.Row, 8), "0.00")
        invoiceMainGrid.TextMatrix(invoiceMainGrid.Row, 9) = InvoiceGrid.TextMatrix(InvoiceGrid.Row, 9)
        invoiceMainGrid.TextMatrix(invoiceMainGrid.Row, 10) = InvoiceGrid.TextMatrix(InvoiceGrid.Row, 10)
        invoiceMainGrid.TextMatrix(invoiceMainGrid.Row, 11) = InvoiceGrid.TextMatrix(InvoiceGrid.Row, 11)
        
End If
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
        invoiceMainGrid.Rows = invoiceMainGrid.Rows + 1
        invoiceMainGrid.Row = invoiceMainGrid.Rows - 1
        Call invoicemaingridload
        additem = False
        Call calstotalamt
End Function
Private Function calstotalamt()
    Dim i As Integer
    txttotamt = 0
    For i = 0 To invoiceMainGrid.Rows - 1
    txttotamt.Text = Format(Val(txttotamt.Text) + Val(invoiceMainGrid.TextMatrix(i, 8)), "0.00")
    Next i
    txtnetamt.Text = Format(Val(txttotamt.Text) + Val(txttaxamt.Text), "0.00")
    txtword.Text = Others.towords(Val(txtnetamt.Text))
End Function
Private Sub txttaxamt_Change()
    If Len(txttaxamt.Text) > 8 Then
        MsgBox "Please Enter Valid Tax Amount", vbCritical, "Tax Amount Error"
    Else
        txtnetamt.Text = Format(Val(txttotamt.Text) + Val(txttaxamt.Text), "0.00")
        txtword.Text = Others.towords(Val(txtnetamt.Text))
    End If
End Sub
Sub totalgrngrid()
    Dim i As Integer
    Dim rs As New ADODB.Recordset
    rs.Open "Select * from invoicemas_details where grnno = '" & cbogrnno.Text & "'", cn, adOpenKeyset, adLockOptimistic
    i = 1
    If rs.BOF = False Then rs.MoveFirst
    While rs.EOF = False
        InvoiceGrid.TextMatrix(i, 10) = Format(rs![SumOfinvoiceqty], "0.000")
        InvoiceGrid.TextMatrix(i, 4) = Format(Val(InvoiceGrid.TextMatrix(i, 4)) - Val(InvoiceGrid.TextMatrix(i, 10)), "0.000")
        rs.MoveNext
        i = i + 1
    Wend
End Sub
Private Sub txttaxamt_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Then
    ElseIf (KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 46 Then
    ElseIf KeyAscii = 13 And InvoiceGrid.Col = 6 Then
    Else
    KeyAscii = 0
    End If
End Sub
