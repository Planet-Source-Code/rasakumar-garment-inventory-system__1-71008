VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmopeninvoice 
   BackColor       =   &H00EDDDD1&
   Caption         =   " * Open Invoice *"
   ClientHeight    =   9345
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9915
   Icon            =   "frmopeninvoice.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   11115
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin MSComCtl2.DTPicker dt2 
      Height          =   375
      Left            =   6000
      TabIndex        =   38
      Top             =   720
      Width           =   1935
      _ExtentX        =   3413
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
      Format          =   58064897
      CurrentDate     =   39548
   End
   Begin VB.TextBox ponos 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
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
      Index           =   9
      Left            =   4320
      Locked          =   -1  'True
      TabIndex        =   37
      Text            =   "Sup.Bill.Date"
      Top             =   720
      Width           =   1695
   End
   Begin VB.TextBox txtsupbillno 
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
      Left            =   2520
      TabIndex        =   36
      Top             =   720
      Width           =   1815
   End
   Begin VB.TextBox ponos 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
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
      Index           =   8
      Left            =   240
      Locked          =   -1  'True
      TabIndex        =   35
      Text            =   "Sup.Bill.No"
      Top             =   720
      Width           =   2295
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
      Height          =   495
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   33
      Tag             =   " "
      ToolTipText     =   "To Use Delete Items  "
      Top             =   8280
      Width           =   1335
   End
   Begin VB.TextBox txtnet 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   12600
      Locked          =   -1  'True
      TabIndex        =   32
      ToolTipText     =   " Po Net Amount ( Tax Amount + Total Amount ) "
      Top             =   7440
      Width           =   2295
   End
   Begin VB.TextBox txttax 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   12600
      TabIndex        =   31
      ToolTipText     =   " Enter The Tax ( % )"
      Top             =   6720
      Width           =   2295
   End
   Begin VB.TextBox taxamt 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   12600
      Locked          =   -1  'True
      TabIndex        =   30
      ToolTipText     =   " Tax Amount ( Against Total Amount )"
      Top             =   7080
      Width           =   2295
   End
   Begin VB.TextBox txttot 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   12600
      Locked          =   -1  'True
      TabIndex        =   29
      ToolTipText     =   " Total Amount ( Without Tax )"
      Top             =   6360
      Width           =   2295
   End
   Begin VB.TextBox txtremarks 
      Appearance      =   0  'Flat
      Height          =   1455
      Left            =   2400
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   28
      ToolTipText     =   " Enter The Remarks "
      Top             =   6360
      Width           =   8295
   End
   Begin VB.TextBox txtinvoiceno 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000014&
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
      Left            =   2520
      Locked          =   -1  'True
      TabIndex        =   27
      ToolTipText     =   " Your PO No"
      Top             =   240
      Width           =   1815
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
      Height          =   495
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   26
      ToolTipText     =   "  To Use Add Items "
      Top             =   8280
      Width           =   1215
   End
   Begin VB.TextBox txtidt 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   13200
      TabIndex        =   25
      Top             =   3960
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox tw 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   9600
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   24
      Tag             =   " "
      ToolTipText     =   " Amount In Words "
      Top             =   8160
      Width           =   5295
   End
   Begin VB.TextBox ponos 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
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
      Index           =   0
      Left            =   240
      Locked          =   -1  'True
      TabIndex        =   22
      Text            =   "Invoice No"
      Top             =   240
      Width           =   2295
   End
   Begin VB.TextBox podates 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
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
      Left            =   4320
      Locked          =   -1  'True
      TabIndex        =   21
      Text            =   "Invoice Date"
      Top             =   240
      Width           =   1695
   End
   Begin VB.TextBox ponos 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
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
      Index           =   1
      Left            =   8760
      Locked          =   -1  'True
      TabIndex        =   20
      Text            =   "Supplier Name"
      Top             =   240
      Width           =   1695
   End
   Begin VB.TextBox ponos 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
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
      Index           =   2
      Left            =   8760
      Locked          =   -1  'True
      TabIndex        =   19
      Text            =   "Department"
      Top             =   720
      Width           =   1695
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
      Left            =   10440
      TabIndex        =   18
      ToolTipText     =   " Select the Supplier Name "
      Top             =   240
      Width           =   4215
   End
   Begin VB.ComboBox cbodept 
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
      Left            =   10440
      TabIndex        =   17
      ToolTipText     =   " Select The Department Name "
      Top             =   720
      Width           =   4215
   End
   Begin VB.ComboBox cboitem 
      Appearance      =   0  'Flat
      ForeColor       =   &H000040C0&
      Height          =   315
      Left            =   1320
      TabIndex        =   16
      ToolTipText     =   " Select The Item Name "
      Top             =   1680
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox ponos 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
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
      Index           =   3
      Left            =   10920
      Locked          =   -1  'True
      TabIndex        =   15
      Text            =   "Tot.Amt (Rs)"
      Top             =   6360
      Width           =   1695
   End
   Begin VB.TextBox ponos 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
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
      Index           =   4
      Left            =   10920
      Locked          =   -1  'True
      TabIndex        =   14
      Text            =   "Tax Amt (Rs)"
      Top             =   7080
      Width           =   1695
   End
   Begin VB.TextBox ponos 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
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
      Index           =   5
      Left            =   10920
      Locked          =   -1  'True
      TabIndex        =   13
      Text            =   "Tax (%)"
      Top             =   6720
      Width           =   1695
   End
   Begin VB.TextBox ponos 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
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
      Index           =   6
      Left            =   10920
      Locked          =   -1  'True
      TabIndex        =   12
      Text            =   "Net Amt (Rs)"
      Top             =   7440
      Width           =   1695
   End
   Begin VB.TextBox ponos 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
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
      Index           =   7
      Left            =   8160
      Locked          =   -1  'True
      TabIndex        =   11
      Text            =   "In Words"
      Top             =   8160
      Width           =   1455
   End
   Begin VB.ComboBox cbocolour 
      Appearance      =   0  'Flat
      ForeColor       =   &H000040C0&
      Height          =   315
      Left            =   2760
      TabIndex        =   10
      ToolTipText     =   " Select The Colour "
      Top             =   1680
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.ComboBox cbosize 
      Appearance      =   0  'Flat
      ForeColor       =   &H000040C0&
      Height          =   315
      Left            =   4200
      TabIndex        =   9
      ToolTipText     =   " Select The Size "
      Top             =   1680
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.ComboBox txtuom 
      Appearance      =   0  'Flat
      ForeColor       =   &H000040C0&
      Height          =   315
      Left            =   5640
      TabIndex        =   8
      ToolTipText     =   " Select The Unit of Measurement "
      Top             =   1680
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox ponos 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
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
      Index           =   13
      Left            =   360
      Locked          =   -1  'True
      TabIndex        =   7
      Text            =   "Remarks"
      Top             =   6360
      Width           =   2055
   End
   Begin VB.TextBox txtcomid 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   11880
      TabIndex        =   6
      Top             =   3600
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox txtsupid 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   13200
      TabIndex        =   5
      Top             =   3600
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton cmdsave 
      BackColor       =   &H00FFC0C0&
      Caption         =   "&Save Invoice"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5880
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   9360
      Width           =   1695
   End
   Begin VB.CommandButton cmdexit 
      BackColor       =   &H00FFC0C0&
      Caption         =   "&Exit"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7680
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   9360
      Width           =   1695
   End
   Begin VB.CommandButton cmdupdate 
      BackColor       =   &H00FFC0C0&
      Caption         =   "&Update Invoice"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5880
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   9360
      Width           =   1695
   End
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
      Left            =   7080
      TabIndex        =   1
      Top             =   2040
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtrate 
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
      Left            =   8280
      TabIndex        =   0
      Top             =   2040
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSComCtl2.DTPicker dt1 
      Height          =   375
      Left            =   6000
      TabIndex        =   23
      ToolTipText     =   " Your PO date "
      Top             =   240
      Width           =   1935
      _ExtentX        =   3413
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
      Format          =   58064897
      CurrentDate     =   39481
   End
   Begin MSFlexGridLib.MSFlexGrid InvoiceGrid 
      Height          =   4695
      Left            =   240
      TabIndex        =   34
      ToolTipText     =   "  Note : # Indicate Place Click The First Row of Grid To  Open the Filter Options "
      Top             =   1440
      Width           =   14775
      _ExtentX        =   26061
      _ExtentY        =   8281
      _Version        =   393216
      Cols            =   8
      FixedCols       =   0
      RowHeightMin    =   2
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
      BorderWidth     =   2
      Height          =   735
      Left            =   120
      Top             =   9240
      Width           =   15015
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H000000C0&
      Height          =   735
      Left            =   2640
      Top             =   8160
      Width           =   2895
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H000000FF&
      Height          =   975
      Index           =   0
      Left            =   240
      Top             =   8040
      Width           =   7695
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H000000FF&
      Height          =   1695
      Index           =   1
      Left            =   240
      Top             =   6240
      Width           =   14775
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H000000FF&
      Height          =   975
      Left            =   8040
      Top             =   8040
      Width           =   7095
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H000000FF&
      Height          =   1095
      Left            =   120
      Top             =   120
      Width           =   7935
   End
   Begin VB.Shape Shape6 
      BorderColor     =   &H000000FF&
      Height          =   1095
      Left            =   8160
      Top             =   120
      Width           =   6975
   End
   Begin VB.Shape Shape7 
      BorderColor     =   &H00400040&
      BorderWidth     =   2
      Height          =   7815
      Left            =   120
      Top             =   1320
      Width           =   15015
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      X1              =   10800
      X2              =   10800
      Y1              =   6240
      Y2              =   7920
   End
End
Attribute VB_Name = "frmopeninvoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Dim op As Variant
    Dim cn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim ww As ADODB.Recordset
    Dim i As Integer
Private Sub cbocolour_Click()
     If donotchange Then Exit Sub
        If InvoiceGrid.Col = 2 Then
            Me.InvoiceGrid.Text = Me.cbocolour.Text
            Call cbocolour_dropdown
            InvoiceGrid.Col = 3
            InvoiceGrid_Click
     End If
End Sub
Private Sub cbocolour_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call cbocolour_Click
    End If
End Sub
Private Sub cbodept_KeyPress(KeyAscii As Integer)
     KeyAscii = 0
End Sub

Private Sub cboitem_Click()
    If InvoiceGrid.Col = 1 Then
        Me.InvoiceGrid.Text = Me.cboitem.Text
         Call cboitem_dropdown
            InvoiceGrid.Col = 2
     Call InvoiceGrid_Click
 End If
End Sub
Private Sub cbocolour_dropdown()
    If Trim(cbodept.Text) = "" Then
    MsgBox "Please select the Department Name", vbCritical, "Department Name Error"
    cbodept.SetFocus
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
Private Sub cbodept_dropdown()
    If Trim(cbosupplier.Text) = "" Then
    MsgBox "Please Select The Supplier Name", vbCritical, "Supplier Select Error"
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
    If Trim(cbodept.Text) = "" Then
    MsgBox "Please Select the Department  Name ", vbCritical, "Department Name Error"
    cbodept.SetFocus
    ElseIf Trim(cbosupplier.Text) = "" Then
    MsgBox "Please Select The Supplier Name", vbCritical, "Supplier Name Error"
    Else
    On Error GoTo X
    cboitem.Clear
        Set rs = cn.Execute("select itemname from item_details ")
        rs.MoveFirst
        Do While Not rs.EOF()
        cboitem.additem (rs(0))
        rs.MoveNext
        Loop
        cboitem.SetFocus
X:
End If
End Sub
Private Sub cboitem_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
    If KeyAscii = 13 Then
    Call cboitem_Click
    End If
End Sub
Private Sub cboitem_OnEnter()
 If InvoiceGrid.Col = 1 Then
 Me.InvoiceGrid.Text = Me.cboitem.Text
 End If
 Call cboitem_dropdown
 Call txtuom_gotfocus
End Sub
Private Sub cbosize_Click()
    If donotchange Then Exit Sub
    If InvoiceGrid.Col = 3 Then
    Me.InvoiceGrid.Text = Me.cbosize.Text
    Call cbosize_gotfocus
    InvoiceGrid.Col = 4
    Call InvoiceGrid_Click
    End If
End Sub
Private Sub cbosize_gotfocus()
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
End Sub
Private Sub cbosize_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    Call cbosize_Click
    End If
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

Private Sub cmdsave_Click()
    Dim i As Integer
    If Trim(cbosupplier.Text) = "" Then
        MsgBox "Please Enter The Supplier Name", vbCritical, "Supplier Error"
        cbosupplier.SetFocus
    ElseIf Trim(cbodept.Text) = "" Then
        MsgBox "Please Enter the department Name", vbCritical, "Department Error"
        cbodept.SetFocus
    ElseIf Trim(InvoiceGrid.Text) = "" Then
        MsgBox "Please Enter All data", vbCritical, "Error"
        InvoiceGrid.SetFocus
        InvoiceGrid.Col = 1
    Else
        Dim rs As New ADODB.Recordset
        rs.Open "Select * from openinvoice_details where invoiceno= '" & txtinvoiceno.Text & "'", cn, adOpenKeyset, adLockOptimistic
    If rs.RecordCount = 0 Then
    For i = 1 To InvoiceGrid.Rows - 1
        rs.AddNew
        rs![invoiceno] = txtinvoiceno.Text
        rs![invoicedate] = dt1.Value
        rs![supbillno] = txtsupbillno.Text
        rs![supbilldate] = dt2.Value
        rs![supname] = cbosupplier.Text
        rs![deptname] = cbodept.Text
        rs![sno] = InvoiceGrid.TextMatrix(i, 0)
        rs![itemname] = InvoiceGrid.TextMatrix(i, 1)
        rs![colour] = InvoiceGrid.TextMatrix(i, 2)
        rs![sizes] = InvoiceGrid.TextMatrix(i, 3)
        rs![invoiceqty] = InvoiceGrid.TextMatrix(i, 4)
        rs![uom] = InvoiceGrid.TextMatrix(i, 5)
        rs![invoicerate] = InvoiceGrid.TextMatrix(i, 6)
        rs![totalamtgrid] = InvoiceGrid.TextMatrix(i, 7)
        rs![totalamt] = txttot.Text
        rs![taxes] = txttax.Text
        rs![taxamt] = taxamt.Text
        rs![netamt] = txtnet.Text
        rs![words] = tw.Text
        rs![remarks] = txtremarks.Text
        Next i
        rs.Update
        rs.Close
        MsgBox "One Record Save Successfully", vbInformation, "Information"
        Unload Me
        frmopeninvoicemain.Show
        Else
            MsgBox "This Invoice No Already Exists", vbCritical, "Invalid Error"
       End If
End If
End Sub
Private Sub cmdaddrow_Click()
    If InvoiceGrid.Rows > 15 Then
    MsgBox "Only 15 Items Allowed ", vbCritical, "Exceed Row "
    Else
     InvoiceGrid.Rows = InvoiceGrid.Rows + 1
     InvoiceGrid.Row = InvoiceGrid.Rows - 1
     InvoiceGrid.TextMatrix(InvoiceGrid.Row, 0) = InvoiceGrid.Rows - 1
     
    If InvoiceGrid.Row > 0 Then
    Call visibles
    End If
    End If
End Sub
Private Sub cmddeleterow_Click()
    If InvoiceGrid.Rows <= 1 Then
     MsgBox "No Record Found So Cann't Delete!", vbCritical, "Dellete Item Error"
    Else
        op = MsgBox("Are You to Delete ?", vbYesNo + vbQuestion, "Delete Row")
    If op = vbYes Then
        InvoiceGrid.Rows = InvoiceGrid.Rows - 1
End If
End If
    Call cals
    Call cals1
    Call visibles
End Sub
Private Sub cmdexit_Click()
    op = MsgBox("Are You Sure To Close ?", vbYesNo + vbQuestion, "Confirm Close ?")
        If op = vbYes Then
        Unload Me
        frmopeninvoicemain.Show
    Else
    End If
End Sub
Private Sub Form_Load()
    Dim i As Integer
    Set cn = New ADODB.Connection
    Set ww = New ADODB.Recordset
    cn.Open "Provider=Microsoft.jet.oledb.4.0; Data Source=" & App.path & "\Database\Data.mdb"
    ww.Open "Select * From openinvoice_details", cn, adOpenKeyset, adLockOptimistic
    cn.CursorLocation = adUseClient
   
    Call openinvoicegridalign
   
    Call txtuom_Click
    Call txtinvoiceno_Gotfocus
    
    Call visibles
    Call txttax_Change
    
    dt1.Value = Date
    dt2.Value = Date

    i = 1
    InvoiceGrid.TextMatrix(i, 0) = 1
    
End Sub
Private Sub Form_Unload(Cancel As Integer)
  If Cancel = 0 Then
    frmopeninvoicemain.Show
  Else
   Cancel = 1
   End If
End Sub
Private Sub InvoiceGrid_Click()
    If cboitem.Visible = True Then cboitem.Visible = False
    If cbocolour.Visible = True Then cbocolour.Visible = False
    If cbosize.Visible = True Then cbosize.Visible = False
    If txtuom.Visible = True Then txtuom.Visible = False
    If txtqty.Visible = True Then txtqty.Visible = False
    If txtrate.Visible = True Then txtrate.Visible = False
    
    If InvoiceGrid.Col = 1 Then
        Me.cboitem.Visible = True
        CurrentRow = Me.InvoiceGrid.Row
        Me.cboitem.Width = Me.InvoiceGrid.CellWidth - 10
        Me.cboitem.Left = Me.InvoiceGrid.CellLeft + Me.InvoiceGrid.Left
        Me.cboitem.Top = Me.InvoiceGrid.CellTop + Me.InvoiceGrid.Top
        donotchange = True
        Me.cboitem.Text = Me.InvoiceGrid.Text
        donotchange = False
        Me.cboitem.SetFocus
    ElseIf InvoiceGrid.Col = 2 Then
        Me.cbocolour.Visible = True
        CurrentRow = Me.InvoiceGrid.Row
        Me.cbocolour.Width = Me.InvoiceGrid.CellWidth - 10
        Me.cbocolour.Left = Me.InvoiceGrid.CellLeft + Me.InvoiceGrid.Left
        Me.cbocolour.Top = Me.InvoiceGrid.CellTop + Me.InvoiceGrid.Top
        donotchange = True
        Me.cbocolour.Text = Me.InvoiceGrid.Text
        donotchange = False
        Me.cbocolour.SetFocus
    ElseIf InvoiceGrid.Col = 3 Then
        Me.cbosize.Visible = True
        CurrentRow = Me.InvoiceGrid.Row
        Me.cbosize.Width = Me.InvoiceGrid.CellWidth - 10
        Me.cbosize.Left = Me.InvoiceGrid.CellLeft + Me.InvoiceGrid.Left
        Me.cbosize.Top = Me.InvoiceGrid.CellTop + Me.InvoiceGrid.Top
        donotchange = True
        Me.cbosize.Text = Me.InvoiceGrid.Text
        donotchange = False
        Me.cbosize.SetFocus
    ElseIf InvoiceGrid.Col = 4 Then
        Me.txtqty.Visible = True
        CurrentRow = Me.InvoiceGrid.Row
        Me.txtqty.Width = Me.InvoiceGrid.CellWidth - 10
        Me.txtqty.Left = Me.InvoiceGrid.CellLeft + Me.InvoiceGrid.Left
        Me.txtqty.Top = Me.InvoiceGrid.CellTop + Me.InvoiceGrid.Top
        donotchange = True
        Me.txtqty.Text = Me.InvoiceGrid.Text
        donotchange = False
        Me.txtqty.SetFocus
    ElseIf InvoiceGrid.Col = 5 Or InvoiceGrid.Col = 1 Then
        Call txtuom_gotfocus
        Me.txtuom.Visible = True
        CurrentRow = Me.InvoiceGrid.Row
        Me.txtuom.Width = Me.InvoiceGrid.CellWidth - 10
        Me.txtuom.Left = Me.InvoiceGrid.CellLeft + Me.InvoiceGrid.Left
        Me.txtuom.Top = Me.InvoiceGrid.CellTop + Me.InvoiceGrid.Top
        donotchange = True
        Me.txtuom.Text = Me.InvoiceGrid.Text
        donotchange = False
        Me.txtuom.SetFocus
    ElseIf InvoiceGrid.Col = 6 Then
        Me.txtrate.Visible = True
        CurrentRow = Me.InvoiceGrid.Row
        Me.txtrate.Width = Me.InvoiceGrid.CellWidth - 10
        Me.txtrate.Left = Me.InvoiceGrid.CellLeft + Me.InvoiceGrid.Left
        Me.txtrate.Top = Me.InvoiceGrid.CellTop + Me.InvoiceGrid.Top
        donotchange = True
        Me.txtrate.Text = Me.InvoiceGrid.Text
        donotchange = False
        Me.txtrate.SetFocus
    ElseIf InvoiceGrid.Col = 8 And InvoiceGrid.Rows > 1 Then
        op = MsgBox("Are You to Delete ?", vbYesNo + vbQuestion, "Delete Row")
    If op = vbYes Then
        InvoiceGrid.Rows = InvoiceGrid.Rows - 1
        txtitems.Caption = InvoiceGrid.Rows - 1
        Call cals
        Call cals1
        Call visibles
    Else
    End If
    End If
End Sub
Private Sub taxamt_change()
    taxamt.Text = Format(taxamt.Text, "0.00")
End Sub
Private Sub txtnet_Change()
    txtnet.Text = Format(Me.txtnet.Text, "0.00")
End Sub
Private Sub txtinvoiceno_Gotfocus()
    Dim rs As New ADODB.Recordset
    Dim exp1 As Variant
    rs.Open "Select invoiceno from openinvoice_details", cn, adOpenKeyset, adLockOptimistic
    If rs.RecordCount = 0 Then
    txtinvoiceno.Text = 1
    rs.Close
    Else
    Dim rsrs As New ADODB.Recordset
    rsrs.Open "Select max(invoiceno) as exp1 from openinvoice_details", cn, adOpenKeyset, adLockOptimistic
    txtinvoiceno = rsrs![exp1] + 1
    rsrs.Close
    End If
    SendKeys "{tab}"
End Sub
Private Sub txtqty_Change()
    If donotchange Then Exit Sub
    If InvoiceGrid.Col = 4 Then
    Me.InvoiceGrid.Text = Format(Me.txtqty.Text, "0.000")
    Call calculations
    Call cals
    Call cals1
    End If
End Sub
Private Sub txtqty_Click()
    If donotchange Then Exit Sub
    If InvoiceGrid.Col = 4 Then
    Me.InvoiceGrid.Text = Format(Me.txtqty.Text, "0.000")
    End If
End Sub
Private Sub txtqty_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Then
    ElseIf (KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 46 Then
    ElseIf KeyAscii = 13 And InvoiceGrid.Col = 4 Then
    Me.InvoiceGrid.Text = Format(Me.txtqty.Text, "0.000")
    InvoiceGrid.Col = 5
    InvoiceGrid_Click
    Call txtqty_Change
    tw.Text = Others.towords(Val(txtnet.Text))
    Else
    KeyAscii = 0
    End If
End Sub
Private Sub txtrate_Change()
     If donotchange Then Exit Sub
     If InvoiceGrid.Col = 6 Then
     Me.InvoiceGrid.Text = Format(Me.txtrate.Text, "0.00")
     Call calculations
     Call cals
     Call cals1
     End If
End Sub
Private Sub txtrate_Click()
    If donotchange Then Exit Sub
    If InvoiceGrid.Col = 6 Then
    Me.InvoiceGrid.Text = Format(Me.txtrate.Text, "0.00")
    End If
End Sub
Private Sub txtrate_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Then
    ElseIf (KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 46 Then
    ElseIf KeyAscii = 13 And InvoiceGrid.Col = 6 Then
    Me.InvoiceGrid.Text = Format(Me.txtrate.Text, "0.00")
    InvoiceGrid.Col = 7
    txtrate.Visible = False
    InvoiceGrid.SetFocus
    InvoiceGrid_Click
    Call cals
    tw.Text = Others.towords(Val(txtnet.Text))
    Else
    KeyAscii = 0
    End If
End Sub
Private Sub txttax_Change()
'    tw.Text = Others.towords(Val(txtnet.Text))
    If Trim(txttax.Text) = "" Or Len(txttax.Text) <= 0 Then
    txttax.Text = "0"
    Call cals1
    ElseIf Trim(txttax.Text) > 100 Then
    MsgBox "Invalid Tax Percenatage ", vbCritical, "Tax Error"
    txttax.Text = "0"
    txttax.SetFocus
    Call cals1
    Else
    Call cals1
    End If
End Sub
Private Sub txttax_GotFocus()
    Call cals1
End Sub
Public Sub txttax_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Then
    ElseIf (KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 46 Then
    ElseIf KeyAscii = 13 And InvoiceGrid.Col = 6 Then
    tw.Text = Others.towords(Val(txtnet.Text))
    Else
    KeyAscii = 0
    End If
End Sub
Private Sub txttot_GotFocus()
    Call cals
End Sub
Private Sub txtuom_Click()
    If donotchange Then Exit Sub
    If InvoiceGrid.Col = 5 Then
    Me.InvoiceGrid.Text = Me.txtuom.Text
    Call txtuom_gotfocus
    InvoiceGrid.Col = 6
    InvoiceGrid_Click
    End If
End Sub
Private Sub txtuom_gotfocus()
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
End Sub
Sub calculations()
    For i = 1 To InvoiceGrid.Rows - 1
       InvoiceGrid.TextMatrix(i, 7) = Format((Val(InvoiceGrid.TextMatrix(i, 4)) * Val(InvoiceGrid.TextMatrix(i, 6))), "0.00")
    Next i
End Sub
Sub cals()
    txttot = 0
    For i = 0 To InvoiceGrid.Rows - 1
    txttot.Text = Format(Val(txttot.Text) + Val(InvoiceGrid.TextMatrix(i, 7)), "0.00")
    Next i
End Sub
Sub cals1()
    taxamt.Text = (Val(txttot.Text) * Val(txttax.Text) / 100)
    txtnet.Text = Val(taxamt.Text) + Val(txttot.Text)
End Sub
Private Sub txtuom_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call txtuom_Click
    End If
End Sub

