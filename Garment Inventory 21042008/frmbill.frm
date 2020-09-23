VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmpayment 
   BackColor       =   &H00EDDDD1&
   Caption         =   " * Payment * "
   ClientHeight    =   9075
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10935
   Icon            =   "frmbill.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   11115
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtdebit 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   2
      Left            =   10920
      TabIndex        =   30
      Top             =   5400
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtreason 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   1
      Left            =   9600
      TabIndex        =   29
      Top             =   5400
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txttot 
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
      Left            =   12360
      TabIndex        =   28
      Text            =   "0"
      Top             =   7200
      Width           =   2655
   End
   Begin VB.TextBox txtpays 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   6000
      TabIndex        =   27
      Top             =   1800
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
      Height          =   615
      Left            =   8280
      MaskColor       =   &H00C0FFC0&
      Style           =   1  'Graphical
      TabIndex        =   26
      ToolTipText     =   "  Exit Window  "
      Top             =   9360
      UseMaskColor    =   -1  'True
      Width           =   1695
   End
   Begin VB.CommandButton cmdsave 
      BackColor       =   &H00FFC0C0&
      Caption         =   "&Save Payment"
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
      Left            =   6360
      MaskColor       =   &H00C0FFC0&
      Style           =   1  'Graphical
      TabIndex        =   25
      Tag             =   " "
      ToolTipText     =   " To Use Add PO "
      Top             =   9360
      UseMaskColor    =   -1  'True
      Width           =   1815
   End
   Begin VB.CommandButton cmddeletedbt 
      BackColor       =   &H00FFC0C0&
      Caption         =   "&Delete Debit"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   12240
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   6600
      Width           =   1335
   End
   Begin VB.CommandButton cmdadddebt 
      BackColor       =   &H00FFC0C0&
      Caption         =   "&Add Debit"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10920
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   6600
      Width           =   1215
   End
   Begin VB.TextBox txtword 
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
      Height          =   975
      Left            =   1680
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   22
      Top             =   7320
      Width           =   7455
   End
   Begin VB.TextBox txtremarks 
      Appearance      =   0  'Flat
      Height          =   2415
      Left            =   1680
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   21
      ToolTipText     =   " Enter the Remarks "
      Top             =   4800
      Width           =   7455
   End
   Begin VB.TextBox txtnetpay 
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
      Left            =   12360
      Locked          =   -1  'True
      TabIndex        =   16
      Text            =   "0"
      Top             =   7920
      Width           =   2655
   End
   Begin VB.TextBox txtdebt 
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
      Left            =   12360
      Locked          =   -1  'True
      TabIndex        =   15
      Text            =   "0"
      Top             =   7560
      Width           =   2655
   End
   Begin VB.TextBox txtcheque 
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
      Left            =   11160
      TabIndex        =   13
      ToolTipText     =   " Enter The Cheque Number "
      Top             =   720
      Width           =   3855
   End
   Begin VB.TextBox txtdd 
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
      Left            =   12120
      TabIndex        =   12
      ToolTipText     =   " Enter The DD Number "
      Top             =   240
      Width           =   2895
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00EDDDD1&
      Caption         =   "Cheque"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   10080
      TabIndex        =   11
      Top             =   720
      Width           =   975
   End
   Begin VB.OptionButton optdd 
      BackColor       =   &H00EDDDD1&
      Caption         =   "DD"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   11280
      TabIndex        =   10
      Top             =   240
      Width           =   735
   End
   Begin VB.OptionButton optadvance 
      BackColor       =   &H00EDDDD1&
      Caption         =   "Cash"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   10080
      TabIndex        =   9
      Top             =   240
      Width           =   1095
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
      Left            =   5760
      TabIndex        =   3
      ToolTipText     =   " Select The Supplier Name "
      Top             =   240
      Width           =   4215
   End
   Begin VB.ComboBox cboinvoiceno 
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
      Left            =   5760
      TabIndex        =   2
      ToolTipText     =   " Select The Invoice No "
      Top             =   720
      Width           =   4215
   End
   Begin VB.TextBox txtpayno 
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
      TabIndex        =   1
      Top             =   240
      Width           =   2055
   End
   Begin MSFlexGridLib.MSFlexGrid PaymentGrid 
      Height          =   3255
      Left            =   120
      TabIndex        =   0
      ToolTipText     =   "  Note : # Indicate Place Click The First Row of Grid To  Open the Filter Options "
      Top             =   1320
      Width           =   15015
      _ExtentX        =   26485
      _ExtentY        =   5741
      _Version        =   393216
      Rows            =   1
      Cols            =   8
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
      Format          =   58130433
      CurrentDate     =   39536
   End
   Begin MSFlexGridLib.MSFlexGrid DebtGrid 
      Height          =   1815
      Left            =   9360
      TabIndex        =   14
      ToolTipText     =   "  Note : # Indicate Place Click The First Row of Grid To  Open the Filter Options "
      Top             =   4680
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   3201
      _Version        =   393216
      Cols            =   3
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
   Begin VB.Shape Shape4 
      BorderColor     =   &H000000FF&
      Height          =   1335
      Left            =   9360
      Top             =   7080
      Width           =   5775
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Total Amount (Rs)"
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
      Index           =   8
      Left            =   9480
      TabIndex        =   31
      Top             =   7200
      Width           =   2895
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00FF0000&
      Height          =   855
      Left            =   6240
      Top             =   9240
      Width           =   3855
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H000000FF&
      Height          =   3735
      Left            =   120
      Top             =   4680
      Width           =   9135
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Amount In Words"
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
      Height          =   495
      Index           =   7
      Left            =   240
      TabIndex        =   20
      Top             =   7320
      Width           =   1455
   End
   Begin VB.Label Label1 
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
      Index           =   6
      Left            =   240
      TabIndex        =   19
      Top             =   4800
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Pay Amount ( Rs)"
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
      Left            =   9480
      TabIndex        =   18
      Top             =   7920
      Width           =   2895
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Debit Amount ( Rs)"
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
      Left            =   9480
      TabIndex        =   17
      Top             =   7560
      Width           =   2895
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H000000FF&
      Height          =   1095
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
      Index           =   4
      Left            =   3960
      TabIndex        =   8
      Top             =   720
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Payment Date"
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
      TabIndex        =   7
      Top             =   720
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Payment No"
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
      TabIndex        =   6
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
      Left            =   3960
      TabIndex        =   5
      Top             =   240
      Width           =   1815
   End
End
Attribute VB_Name = "frmpayment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim op As Variant
Dim cn As ADODB.Connection
Dim rs As ADODB.Recordset
Dim ww As ADODB.Recordset
Private Sub cboinvoiceno_Click()
    Call PaymentGridload
    Call PaymentGridloads
End Sub
Private Sub cboinvoiceno_dropdown()
On Error GoTo X
        cboinvoiceno.Clear
        Set rs = cn.Execute("select invoiceno from invoicemain_details where supname='" & cbosup.Text & "' Order by invoiceno")
        rs.MoveFirst
        Do While Not rs.EOF()
        cboinvoiceno.additem (rs(0))
        rs.MoveNext
        Loop
        cboinvoiceno.SetFocus
X:
End Sub
Private Sub cboinvoiceno_KeyPress(KeyAscii As Integer)
 KeyAscii = 0
End Sub

Private Sub cbosup_dropdown()
On Error GoTo X
        cbosup.Clear
        Set rs = cn.Execute("select supname from invoicemain_details Group by supname")
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

Private Sub cmdadddebt_Click()
        DebtGrid.Rows = DebtGrid.Rows + 1
        DebtGrid.Row = DebtGrid.Rows - 1
        DebtGrid.TextMatrix(DebtGrid.Row, 0) = DebtGrid.Rows - 1
End Sub
Private Sub cmddeletedbt_Click()
    If DebtGrid.Rows = 2 Then
        MsgBox "No Record Found", vbInformation, "Information"
        Call frmpaymentgridload
    Else
        DebtGrid.Rows = DebtGrid.Rows - 1
        DebtGrid.Row = DebtGrid.Rows - 1
        DebtGrid.TextMatrix(DebtGrid.Row, 0) = DebtGrid.Rows - 1
        Call totladebt
        Call netpaycals
    End If
End Sub
Private Sub cmdexits_Click()
    op = MsgBox("Are You Sure To Close ?", vbQuestion + vbYesNo, "Confirm To Close ?")
        If op = vbYes Then
            Unload Me
            frmpaymentmain.Show
        Else
        End If
End Sub
Private Sub cmdsave_Click()
    If PaymentGrid.Rows = 1 Then
        MsgBox "Please Enter The Valid Pay Amount Details", vbCritical, "Pay Amount Error"
        PaymentGrid.SetFocus
    ElseIf optdd.Value = True And Trim(txtdd.Text) = "" Then
        MsgBox "Please Enter The DD Number", vbCritical, "DD Number Error"
        txtdd.SetFocus
    ElseIf Option1.Value = True And Trim(txtcheque.Text) = "" Then
        MsgBox "please Enter The Cheque Number", vbCritical, "Cheque No Error"
        txtcheque.SetFocus
    ElseIf PaymentGrid.Rows > 1 And Val(txttot.Text) <= 0 Then
        MsgBox "Please Enter Valid Payment", vbCritical, "Payment Error"
        PaymentGrid.Col = 5
        PaymentGrid.SetFocus
    Else
        Dim rs As New ADODB.Recordset
        Dim rs1 As New ADODB.Recordset
        rs.Open "Select * from payment_details where payno= ' " & txtpayno.Text & "'", cn, 1, 3
        rs1.Open "Select * from debit_details where payno= ' " & txtpayno.Text & "'", cn, 1, 3
           If rs.RecordCount = 0 Then
            For i = 1 To PaymentGrid.Rows - 1
            rs.AddNew
            rs![payno] = txtpayno.Text
            rs![paydate] = dt1.Value
            rs![supname] = cbosup.Text
            rs![invoiceno] = cboinvoiceno.Text
            rs![sno] = PaymentGrid.TextMatrix(i, 0)
            rs![invoiceno] = PaymentGrid.TextMatrix(i, 1)
            rs![invoicedate] = PaymentGrid.TextMatrix(i, 2)
            rs![supbillno] = PaymentGrid.TextMatrix(i, 3)
            rs![invoiceamt] = PaymentGrid.TextMatrix(i, 4)
            rs![payamtgrid] = PaymentGrid.TextMatrix(i, 5)
            rs![words] = txtword.Text
            rs![remarks] = txtremarks.Text
            rs![advance] = optadvance.Value
            rs![dd] = optdd.Value
            rs![cheque] = Option1.Value
            rs![ddno] = txtdd.Text
            rs![chequeno] = txtcheque.Text
            rs![paynetamt] = txtnetpay.Text
            Next i
            rs.Update
            rs.Close
            If rs1.RecordCount = 0 Then
            For j = 1 To DebtGrid.Rows - 1
                rs1.AddNew
                rs1![payno] = txtpayno.Text
                rs1![paydate] = dt1.Value
                rs1![supname] = cbosup.Text
                rs1![sno] = DebtGrid.TextMatrix(j, 0)
                rs1![debtreason] = DebtGrid.TextMatrix(j, 1)
                rs1![debtamt] = DebtGrid.TextMatrix(j, 2)
                rs1![debttotamt] = txtdebt.Text
                rs1![paynetamt] = txtnetpay.Text
            Next j
            rs1.Update
            rs1.Close
            MsgBox "One Record Save Successfully ", vbInformation, "Information"
            Unload Me
            frmpaymentmain.Show
            Else
            MsgBox "Already Exists", vbCritical, "Error"
    End If
    End If
    End If
End Sub
Private Sub DebtGrid_Click()
    If txtreason(1).Visible = True Then txtreason(1).Visible = False
    If txtdebit(2).Visible = True Then txtdebit(2).Visible = False
    If DebtGrid.Col = 1 Then
    Me.txtreason(1).Visible = True
    CurrentRow = Me.PaymentGrid.Row
    Me.txtreason(1).Width = Me.DebtGrid.CellWidth - 10
    Me.txtreason(1).Left = Me.DebtGrid.CellLeft + Me.DebtGrid.Left
    Me.txtreason(1).Top = Me.DebtGrid.CellTop + Me.DebtGrid.Top
    donotchange = True
    Me.txtreason(1).Text = Me.DebtGrid.Text
    donotchange = False
    Me.txtreason(1).SetFocus
    ElseIf DebtGrid.Col = 2 Then
    Me.txtdebit(2).Visible = True
    CurrentRow = Me.PaymentGrid.Row
    Me.txtdebit(2).Width = Me.DebtGrid.CellWidth - 10
    Me.txtdebit(2).Left = Me.DebtGrid.CellLeft + Me.DebtGrid.Left
    Me.txtdebit(2).Top = Me.DebtGrid.CellTop + Me.DebtGrid.Top
    donotchange = True
    Me.txtdebit(2).Text = Me.DebtGrid.Text
    donotchange = False
    Me.txtdebit(2).SetFocus
    End If
End Sub
Private Sub Form_Load()
            
    Set cn = New ADODB.Connection
    Set ww = New ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    cn.Open "Provider=Microsoft.jet.oledb.4.0; Data Source=" & App.path & "\Database\Data.mdb"
    ww.Open "Select * From payment_details", cn, adOpenKeyset, adLockOptimistic
    cn.CursorLocation = adUseClient
    
    Call frmpaymentgridload
    Call txtinvoiceno_Gotfocus
    
    txtdd.Visible = False
    txtcheque.Visible = False
    optadvance.Value = True
    DebtGrid.TextMatrix(DebtGrid.Row, 0) = 1
    DebtGrid.TextMatrix(DebtGrid.Row, 1) = "Nil"
    DebtGrid.TextMatrix(DebtGrid.Row, 2) = "0.00"
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Unload Me
    frmpaymentmain.Show
End Sub

Private Sub optadvance_Click()
    txtdd.Visible = False
    txtcheque.Visible = False
End Sub
Private Sub optdd_Click()
    txtcheque.Visible = False
    txtdd.Visible = True
End Sub
Private Sub Option1_Click()
    txtdd.Visible = False
    txtcheque.Visible = True
End Sub
Private Sub txtinvoiceno_Gotfocus()
    Dim rs As New ADODB.Recordset
    Dim a As String
    rs.Open "Select * from payment_details", cn, adOpenKeyset, adLockOptimistic
        If rs.RecordCount = 0 Then
        txtpayno.Text = 1
        rs.Close
        Else
        Dim rsrs As New ADODB.Recordset
        rsrs.Open "Select max(payno)as exp1 from payment_details", cn, adOpenKeyset, adLockOptimistic
        txtpayno = rsrs![exp1] + 1
        rsrs.Close
        End If
    SendKeys "{tab}"
End Sub
Private Function PaymentGridload()
    Dim i As Integer
    Dim rs As New ADODB.Recordset
    rs.Open "select * from invoicemain_details where invoiceno= '" & Trim(cboinvoiceno.Text) & "'", cn, adOpenKeyset, adLockOptimistic
    i = 1
    If rs.BOF = False Then rs.MoveFirst
    While rs.EOF = False
         PaymentGrid.Rows = rs.RecordCount + 1
         PaymentGrid.TextMatrix(i, 0) = i
         PaymentGrid.TextMatrix(i, 1) = rs![invoiceno]
         PaymentGrid.TextMatrix(i, 2) = rs![invoicedate]
         PaymentGrid.TextMatrix(i, 3) = rs![supbillno]
         PaymentGrid.TextMatrix(i, 4) = Format(rs![netamt], "0.00")
         PaymentGrid.TextMatrix(i, 5) = Format("0.00")
         
        rs.MoveNext
        i = i + 1
        Wend
         PaymentGrid.Rows = rs.RecordCount + 1
End Function
Private Sub PaymentGrid_Click()
    If txtpays(0).Visible = True Then txtpays(0).Visible = False
    If PaymentGrid.Col = 5 Then
    Me.txtpays(0).Visible = True
    CurrentRow = Me.PaymentGrid.Row
    Me.txtpays(0).Width = Me.PaymentGrid.CellWidth - 10
    Me.txtpays(0).Left = Me.PaymentGrid.CellLeft + Me.PaymentGrid.Left
    Me.txtpays(0).Top = Me.PaymentGrid.CellTop + Me.PaymentGrid.Top
    donotchange = True
    Me.txtpays(0).Text = Me.PaymentGrid.Text
    donotchange = False
    Me.txtpays(0).SetFocus
    End If
    Call cals
End Sub
Private Sub txtdebit_Change(Index As Integer)
    If Val(txtdebit(2).Text) >= Val(txttot.Text) Then
        MsgBox "Please Enter The Valid Debit Amount", vbCritical, "Debit Amount Error"
        txtdebit(2).SetFocus
        DebtGrid.Col = 2
        DebtGrid.SetFocus
        DebtGrid_Click
    End If
End Sub
Private Sub txtdebit_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 8 Then
    ElseIf (KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 46 Then
    ElseIf KeyAscii = 13 And DebtGrid.Col = 2 Then
        Me.DebtGrid.Text = Format(Me.txtdebit(2).Text, "0.00")
        DebtGrid.Col = 2
        DebtGrid.SetFocus
        Call totladebt
        Call netpaycals
        txtdebit(2).Visible = False
    Else
        KeyAscii = 0
    End If
End Sub
Private Sub txtpays_Change(Index As Integer)
    If Val(txtpays(0).Text) > Val(PaymentGrid.TextMatrix(PaymentGrid.Row, 4)) Then
        MsgBox "Please Check Exceed Our Bill Amount", vbCritical, "Pay Amount Error"
        txtpays(0).Text = 0
        PaymentGrid.Col = 5
        PaymentGrid.SetFocus
        PaymentGrid_Click
    End If
End Sub
Private Sub txtpays_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 8 Then
    ElseIf (KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 46 Then
    ElseIf KeyAscii = 13 And PaymentGrid.Col = 5 Then
        Me.PaymentGrid.Text = Format(Me.txtpays(0).Text, "0.00")
        Call netpaycals
        Call cals
        PaymentGrid.Col = 6
        PaymentGrid.SetFocus
        txtpays(0).Visible = False
    Else
        KeyAscii = 0
    End If
End Sub
Private Function cals()
    Dim i As Integer
    txttotamt = 0
    For i = 0 To PaymentGrid.Rows - 1
         txttot.Text = Format(Val(PaymentGrid.TextMatrix(i, 5)), "0.00")
    Next i
    txtnetpay = 0
    txtnetpay.Text = Format(Val(txttot.Text) - Val(txtdebt.Text), "0.00")
    txtword.Text = Others.towords(Val(txtnetpay.Text))
End Function
Private Sub txtreason_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.DebtGrid.Text = Me.txtreason(1).Text
        DebtGrid.Col = 2
        DebtGrid.SetFocus
        txtdebit(2).Visible = True
        DebtGrid_Click
    End If
End Sub
Private Function totladebt()
    Dim i As Integer
    txtdebt = 0
    For i = 1 To DebtGrid.Rows - 1
    txtdebt.Text = Format(Val(txtdebt.Text) + Val(DebtGrid.TextMatrix(i, 2)), "0.00")
    Next i
End Function
Private Function netpaycals()
    txtnetpay = 0
    txtnetpay.Text = Format(Val(txttot.Text) - Val(txtdebt.Text), "0.00")
    txtword.Text = Others.towords(Val(txtnetpay.Text))
End Function
Private Function PaymentGridloads()
    Dim i As Integer
    Dim rs As New ADODB.Recordset
    rs.Open "Select * from paymas_details where invoiceno = '" & cboinvoiceno.Text & "'", cn, adOpenKeyset, adLockOptimistic
    i = 1
    If rs.BOF = False Then rs.MoveFirst
    While rs.EOF = False
        PaymentGrid.TextMatrix(i, 7) = Format(rs![SumOftotpayamt], "0.00")
        PaymentGrid.TextMatrix(i, 4) = Format(Val(PaymentGrid.TextMatrix(i, 4)) - Val(PaymentGrid.TextMatrix(i, 7)), "0.00")
        PaymentGrid.TextMatrix(i, 6) = rs![invoiceno]
        rs.MoveNext
        i = i + 1
    Wend
   
End Function
