VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmpos 
   BackColor       =   &H00EDDDD1&
   Caption         =   " * Purchase Order *"
   ClientHeight    =   10890
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
   Icon            =   "frmpos.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   11115
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
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
      Left            =   8040
      TabIndex        =   34
      Top             =   1680
      Visible         =   0   'False
      Width           =   1095
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
      Left            =   6720
      TabIndex        =   33
      Top             =   1680
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdupdate 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Update"
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
      TabIndex        =   32
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
      TabIndex        =   31
      Top             =   9360
      Width           =   1695
   End
   Begin VB.CommandButton cmdsave 
      BackColor       =   &H00FFC0C0&
      Caption         =   "&Save PO"
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
      TabIndex        =   30
      Top             =   9360
      Width           =   1695
   End
   Begin VB.TextBox txtsupid 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   13200
      TabIndex        =   28
      Top             =   3600
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox txtcomid 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   11880
      TabIndex        =   27
      Top             =   3600
      Visible         =   0   'False
      Width           =   1455
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
      TabIndex        =   26
      Text            =   "Remarks"
      Top             =   6120
      Width           =   2055
   End
   Begin VB.ComboBox txtuom 
      Appearance      =   0  'Flat
      ForeColor       =   &H000040C0&
      Height          =   315
      Left            =   5640
      TabIndex        =   25
      ToolTipText     =   " Select The Unit of Measurement "
      Top             =   1320
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.ComboBox cbosize 
      Appearance      =   0  'Flat
      ForeColor       =   &H000040C0&
      Height          =   315
      Left            =   4200
      TabIndex        =   24
      ToolTipText     =   " Select The Size "
      Top             =   1320
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.ComboBox cbocolour 
      Appearance      =   0  'Flat
      ForeColor       =   &H000040C0&
      Height          =   315
      Left            =   2760
      TabIndex        =   23
      ToolTipText     =   " Select The Colour "
      Top             =   1320
      Visible         =   0   'False
      Width           =   1455
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
      Index           =   7
      Left            =   8160
      Locked          =   -1  'True
      TabIndex        =   22
      Text            =   "In Words"
      Top             =   8040
      Width           =   1455
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
      TabIndex        =   21
      Text            =   "Net Amt (Rs)"
      Top             =   7320
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
      TabIndex        =   20
      Text            =   "Tax (%)"
      Top             =   6600
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
      TabIndex        =   19
      Text            =   "Tax Amt (Rs)"
      Top             =   6960
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
      Index           =   3
      Left            =   10920
      Locked          =   -1  'True
      TabIndex        =   18
      Text            =   "Tot.Amt (Rs)"
      Top             =   6120
      Width           =   1695
   End
   Begin VB.ComboBox cboitem 
      Appearance      =   0  'Flat
      ForeColor       =   &H000040C0&
      Height          =   315
      Left            =   1320
      TabIndex        =   17
      ToolTipText     =   " Select The Item Name "
      Top             =   1320
      Visible         =   0   'False
      Width           =   1335
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
      Left            =   12240
      TabIndex        =   16
      ToolTipText     =   " Select The Department Name "
      Top             =   240
      Width           =   2775
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
      Left            =   7560
      TabIndex        =   15
      ToolTipText     =   " Select the Supplier Name "
      Top             =   240
      Width           =   3255
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
      Left            =   10920
      Locked          =   -1  'True
      TabIndex        =   14
      Text            =   "Department"
      Top             =   240
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
      Index           =   1
      Left            =   5880
      Locked          =   -1  'True
      TabIndex        =   13
      Text            =   "Supplier Name"
      Top             =   240
      Width           =   1695
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
      Left            =   3120
      Locked          =   -1  'True
      TabIndex        =   12
      Text            =   "PO Date"
      Top             =   240
      Width           =   975
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
      TabIndex        =   11
      Text            =   "PO NO"
      Top             =   240
      Width           =   975
   End
   Begin MSComCtl2.DTPicker dt 
      Height          =   375
      Left            =   4080
      TabIndex        =   10
      ToolTipText     =   " Your PO date "
      Top             =   240
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
      Format          =   58130433
      CurrentDate     =   39481
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
      Height          =   855
      Left            =   9600
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   7
      Tag             =   " "
      ToolTipText     =   " Amount In Words "
      Top             =   8040
      Width           =   5295
   End
   Begin VB.TextBox txtidt 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   13200
      TabIndex        =   9
      Top             =   3960
      Visible         =   0   'False
      Width           =   1455
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
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "  To Use Add Items "
      Top             =   8160
      Width           =   1215
   End
   Begin VB.TextBox txtpono 
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
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   6
      ToolTipText     =   " Your PO No"
      Top             =   240
      Width           =   1815
   End
   Begin VB.TextBox txtremarks 
      Appearance      =   0  'Flat
      Height          =   1575
      Left            =   2400
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      ToolTipText     =   " Enter The Remarks "
      Top             =   6120
      Width           =   8295
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
      TabIndex        =   4
      ToolTipText     =   " Total Amount ( Without Tax )"
      Top             =   6120
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
      TabIndex        =   3
      ToolTipText     =   " Tax Amount ( Against Total Amount )"
      Top             =   6960
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
      TabIndex        =   2
      ToolTipText     =   " Enter The Tax ( % )"
      Top             =   6600
      Width           =   2295
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
      TabIndex        =   1
      ToolTipText     =   " Po Net Amount ( Tax Amount + Total Amount ) "
      Top             =   7320
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
      Height          =   615
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   0
      Tag             =   " "
      ToolTipText     =   "To Use Delete Items  "
      Top             =   8160
      Width           =   1335
   End
   Begin MSFlexGridLib.MSFlexGrid PoGrid 
      Height          =   4935
      Left            =   240
      TabIndex        =   29
      ToolTipText     =   "  Note : # Indicate Place Click The First Row of Grid To  Open the Filter Options "
      Top             =   960
      Width           =   14775
      _ExtentX        =   26061
      _ExtentY        =   8705
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
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      X1              =   10800
      X2              =   10800
      Y1              =   6000
      Y2              =   7800
   End
   Begin VB.Shape Shape7 
      BorderColor     =   &H00400040&
      BorderWidth     =   2
      Height          =   8295
      Left            =   120
      Top             =   840
      Width           =   15015
   End
   Begin VB.Shape Shape6 
      BorderColor     =   &H000000FF&
      Height          =   615
      Left            =   5760
      Top             =   120
      Width           =   9375
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H000000FF&
      Height          =   615
      Left            =   120
      Top             =   120
      Width           =   5535
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H000000FF&
      Height          =   1095
      Left            =   8040
      Top             =   7920
      Width           =   6975
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H000000FF&
      Height          =   1815
      Index           =   1
      Left            =   240
      Top             =   6000
      Width           =   14775
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H000000FF&
      Height          =   1095
      Index           =   0
      Left            =   240
      Top             =   7920
      Width           =   7695
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H000000C0&
      Height          =   855
      Left            =   2640
      Top             =   8040
      Width           =   2895
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      Height          =   735
      Left            =   120
      Top             =   9240
      Width           =   15015
   End
End
Attribute VB_Name = "frmpos"
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
        If PoGrid.Col = 2 Then
            Me.PoGrid.Text = Me.cbocolour.Text
            Call cbocolour_dropdown
            PoGrid.Col = 3
            PoGrid_Click
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
    If PoGrid.Col = 1 Then
        Me.PoGrid.Text = Me.cboitem.Text
         Call cboitem_dropdown
            PoGrid.Col = 2
     Call PoGrid_Click
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
Private Sub cboitem_KeyPress(KeyAscii As Integer)
     KeyAscii = 0
    If KeyAscii = 13 Then
    Call cboitem_Click
    End If
End Sub
Private Sub cboitem_OnEnter()
 If PoGrid.Col = 1 Then
 Me.PoGrid.Text = Me.cboitem.Text
 End If
 Call cboitem_dropdown
 Call txtuom_gotfocus
End Sub
Private Sub cbosize_Click()
    If donotchange Then Exit Sub
    If PoGrid.Col = 3 Then
    Me.PoGrid.Text = Me.cbosize.Text
    Call cbosize_gotfocus
    PoGrid.Col = 4
    Call PoGrid_Click
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
Private Sub cbosupplier_Click()
    Call txtsupid_GotFocus
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
    ElseIf Trim(PoGrid.Text) = "" Then
    MsgBox "Please Enter All data", vbCritical, "Error"
    PoGrid.SetFocus
    PoGrid.Col = 1
    Else
    Dim rs As New ADODB.Recordset
    Dim rsstatus As New ADODB.Recordset
    rs.Open "Select * from po_details where pono= " & txtpono.Text, cn, adOpenKeyset, adLockOptimistic
    rsstatus.Open "Select * from postatus_details where pono= '" & txtpono.Text & "'", cn, adOpenKeyset, adLockOptimistic
    If rsstatus.RecordCount = 0 Then
        rsstatus.AddNew
        rsstatus![pono] = txtpono.Text
        rsstatus![podate] = dt.Value
        rsstatus![postatus] = "Open"
        rsstatus![supname] = cbosupplier.Text
        rsstatus![deptname] = cbodept.Text
        rsstatus![netamt] = txtnet.Text
        rsstatus.Update
        rsstatus.Close
    End If
    If rs.RecordCount = 0 Then
    For i = 1 To PoGrid.Rows - 1
        rs.AddNew
        rs![pono] = txtpono.Text
        rs![compname] = txtcomid.Text
        rs![podate] = dt.Value
        rs![supname] = cbosupplier.Text
        rs![deptname] = cbodept.Text
        rs![sno] = PoGrid.TextMatrix(i, 0)
        rs![itemname] = PoGrid.TextMatrix(i, 1)
        rs![colour] = PoGrid.TextMatrix(i, 2)
        rs![sizes] = PoGrid.TextMatrix(i, 3)
        rs![qty] = Format(PoGrid.TextMatrix(i, 4), "0.000")
        rs![uom] = PoGrid.TextMatrix(i, 5)
        rs![rates] = Format(PoGrid.TextMatrix(i, 6), "0.00")
        rs![totamts] = Format(PoGrid.TextMatrix(i, 7), "0.00")
        rs![totamt] = Format(txttot.Text, "0.00")
        rs![tax] = txttax.Text
        rs![taxamt] = Format(taxamt.Text, "0.00")
        rs![netamt] = Format(txtnet.Text, "0.00")
        rs![remarks] = txtremarks.Text
        rs![words] = tw.Text
        Next i
            rs.Update
            rs.Close
            MsgBox "One Record Save Successfully", vbInformation, "Information"
            Unload Me
            frmpomain.Show
        Else
            MsgBox "This Company Already Exists", vbCritical, "Invalid Error"
       End If
End If
End Sub
Private Sub cmdaddrow_Click()
    If PoGrid.Rows > 10 Then
    MsgBox "Only 10 Item Allowed ", vbCritical, "Exceed Row "
    Else
     PoGrid.Rows = PoGrid.Rows + 1
     PoGrid.Row = PoGrid.Rows - 1
     PoGrid.TextMatrix(PoGrid.Row, 0) = PoGrid.Rows - 1
     
    If PoGrid.Row > 0 Then
    Call visibles
    End If
    End If
End Sub
Private Sub cmddeleterow_Click()
    If PoGrid.Rows <= 1 Then
     MsgBox "No Record Found So Cann't Delete!", vbCritical, "Dellete Item Error"
    Else
    op = MsgBox("Are You to Delete ?", vbYesNo + vbQuestion, "Delete Row")
    If op = vbYes Then
    PoGrid.Rows = PoGrid.Rows - 1
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
    frmpomain.Show
    Else
    End If
End Sub
Private Sub cmdupdate_Click()
    Dim rs1 As New ADODB.Recordset
    rs1.Open "Select * from grn_details where pono= '" & txtpono.Text & "'", cn, 1, 3
    If rs1.RecordCount > 1 Then
    MsgBox "Already Some Of Qty Received Against This PO No! So Cann't Update", vbCritical, "Update Error"
    frmpos.Show
    frmpos.cmdupdate.Enabled = False
    Else
    If Trim(cbosupplier.Text) = "" Then
    MsgBox "Please Enter The Supplier Name", vbCritical, "Supplier Error"
    cbosupplier.SetFocus
    ElseIf Trim(cbodept.Text) = "" Then
    MsgBox "Please Enter the department Name", vbCritical, "Department Error"
    cbodept.SetFocus
    ElseIf Trim(PoGrid.Text) = "" Then
    MsgBox "Please Enter All data", vbCritical, "Error"
    Else
    Dim rs As New ADODB.Recordset
    rs.Open "Select * from po_details where pono= " & txtpono.Text, cn, adOpenStatic, adLockPessimistic
    If rs.RecordCount <> 0 Then
    For i = 1 To rs.RecordCount
    If rs.BOF = False Then rs.MoveFirst
    While rs.EOF = False
        rs![pono] = txtpono.Text
        rs![compname] = txtcomid.Text
        rs![podate] = dt.Value
        rs![supname] = cbosupplier.Text
        rs![deptname] = cbodept.Text
        rs![sno] = PoGrid.TextMatrix(i, 0)
        rs![itemname] = PoGrid.TextMatrix(i, 1)
        rs![colour] = PoGrid.TextMatrix(i, 2)
        rs![sizes] = PoGrid.TextMatrix(i, 3)
        rs![qty] = PoGrid.TextMatrix(i, 4)
        rs![uom] = PoGrid.TextMatrix(i, 5)
        rs![rates] = PoGrid.TextMatrix(i, 6)
        rs![totamts] = PoGrid.TextMatrix(i, 7)
        rs![totamt] = txttot.Text
        rs![tax] = txttax.Text
        rs![taxamt] = taxamt.Text
        rs![netamt] = txtnet.Text
        rs![remarks] = txtremarks.Text
        rs![words] = tw.Text
        rs.Update
        rs.MoveNext
        i = i + 1
        Wend
        Next i
            rs.Close
            MsgBox "One Record Save Successfully", vbInformation, "Information"
            Unload Me
            frmpomain.Show
        Else
            MsgBox "This Company Already Exists", vbCritical, "Invalid Error"
        End If
        End If
End If
End Sub

Private Sub Form_Load()
    Dim i As Integer
    Set cn = New ADODB.Connection
    Set ww = New ADODB.Recordset
    cn.Open "Provider=Microsoft.jet.oledb.4.0; Data Source=" & App.path & "\Database\Data.mdb"
    ww.Open "Select * From po_details", cn, adOpenKeyset, adLockOptimistic
    cn.CursorLocation = adUseClient
   
    Call pogridalign
   
    Call txtuom_Click
    Call txtpono_gotfocus
    Call txtcomid_GotFocus
    Call visibles
    Call txttax_Change
    
    dt.Value = Date

    i = 1
    PoGrid.TextMatrix(i, 0) = 1
    
End Sub
Private Sub Form_Unload(Cancel As Integer)
  If Cancel = 0 Then
   frmpomain.Show
   Else
   Cancel = 1
   End If
End Sub

Private Sub PoGrid_Click()
    If cboitem.Visible = True Then cboitem.Visible = False
    If cbocolour.Visible = True Then cbocolour.Visible = False
    If cbosize.Visible = True Then cbosize.Visible = False
    If txtuom.Visible = True Then txtuom.Visible = False
    If txtqty.Visible = True Then txtqty.Visible = False
    If txtrate.Visible = True Then txtrate.Visible = False
    If PoGrid.Col = 1 Then
    
    Me.cboitem.Visible = True
    CurrentRow = Me.PoGrid.Row
    Me.cboitem.Width = Me.PoGrid.CellWidth - 10
    Me.cboitem.Left = Me.PoGrid.CellLeft + Me.PoGrid.Left
    Me.cboitem.Top = Me.PoGrid.CellTop + Me.PoGrid.Top
    donotchange = True
    Me.cboitem.Text = Me.PoGrid.Text
    donotchange = False
    Me.cboitem.SetFocus
    ElseIf PoGrid.Col = 2 Then
    Me.cbocolour.Visible = True
    CurrentRow = Me.PoGrid.Row
    Me.cbocolour.Width = Me.PoGrid.CellWidth - 10
    Me.cbocolour.Left = Me.PoGrid.CellLeft + Me.PoGrid.Left
    Me.cbocolour.Top = Me.PoGrid.CellTop + Me.PoGrid.Top
    donotchange = True
    Me.cbocolour.Text = Me.PoGrid.Text
    donotchange = False
    Me.cbocolour.SetFocus
    ElseIf PoGrid.Col = 3 Then
    Me.cbosize.Visible = True
    CurrentRow = Me.PoGrid.Row
    Me.cbosize.Width = Me.PoGrid.CellWidth - 10
    Me.cbosize.Left = Me.PoGrid.CellLeft + Me.PoGrid.Left
    Me.cbosize.Top = Me.PoGrid.CellTop + Me.PoGrid.Top
    donotchange = True
    Me.cbosize.Text = Me.PoGrid.Text
    donotchange = False
    Me.cbosize.SetFocus
    ElseIf PoGrid.Col = 4 Then
    Me.txtqty.Visible = True
    CurrentRow = Me.PoGrid.Row
    Me.txtqty.Width = Me.PoGrid.CellWidth - 10
    Me.txtqty.Left = Me.PoGrid.CellLeft + Me.PoGrid.Left
    Me.txtqty.Top = Me.PoGrid.CellTop + Me.PoGrid.Top
    donotchange = True
    Me.txtqty.Text = Me.PoGrid.Text
    donotchange = False
    Me.txtqty.SetFocus
    ElseIf PoGrid.Col = 5 Or PoGrid.Col = 1 Then
    Call txtuom_gotfocus
    Me.txtuom.Visible = True
    CurrentRow = Me.PoGrid.Row
    Me.txtuom.Width = Me.PoGrid.CellWidth - 10
    Me.txtuom.Left = Me.PoGrid.CellLeft + Me.PoGrid.Left
    Me.txtuom.Top = Me.PoGrid.CellTop + Me.PoGrid.Top
    donotchange = True
    Me.txtuom.Text = Me.PoGrid.Text
    donotchange = False
    Me.txtuom.SetFocus
    ElseIf PoGrid.Col = 6 Then
    Me.txtrate.Visible = True
    CurrentRow = Me.PoGrid.Row
    Me.txtrate.Width = Me.PoGrid.CellWidth - 10
    Me.txtrate.Left = Me.PoGrid.CellLeft + Me.PoGrid.Left
    Me.txtrate.Top = Me.PoGrid.CellTop + Me.PoGrid.Top
    donotchange = True
    Me.txtrate.Text = Me.PoGrid.Text
    donotchange = False
    Me.txtrate.SetFocus
    ElseIf PoGrid.Col = 8 And PoGrid.Rows > 1 Then
    op = MsgBox("Are You to Delete ?", vbYesNo + vbQuestion, "Delete Row")
    If op = vbYes Then
    PoGrid.Rows = PoGrid.Rows - 1
    txtitems.Caption = PoGrid.Rows - 1
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
Private Sub txtcomid_GotFocus()
    Dim rs As New ADODB.Recordset
    Dim exp1 As String
    rs.Open "Select compname from com_details", cn, adOpenKeyset, adLockOptimistic
    If rs.RecordCount = 0 Then
    txtcomid.Text = ""
    rs.Close
    Else
    Dim rsrs As New ADODB.Recordset
    rsrs.Open "Select compname as exp1 from com_details", cn, adOpenKeyset, adLockOptimistic
    txtcomid = rsrs![exp1]
    rsrs.Close
    End If
    SendKeys "{tab}"
End Sub
Private Sub txtnet_Change()
    txtnet.Text = Format(Me.txtnet.Text, "0.00")
End Sub
Private Sub txtpono_gotfocus()
    Dim rs As New ADODB.Recordset
    Dim exp1 As Variant
    rs.Open "Select pono from po_details", cn, adOpenKeyset, adLockOptimistic
    If rs.RecordCount = 0 Then
    txtpono.Text = 1
    rs.Close
    Else
    Dim rsrs As New ADODB.Recordset
    rsrs.Open "Select max(pono) as exp1 from po_details", cn, adOpenKeyset, adLockOptimistic
    txtpono = rsrs![exp1] + 1
    rsrs.Close
    End If
    SendKeys "{tab}"
End Sub
Private Sub txtqty_Change()
    If donotchange Then Exit Sub
    If PoGrid.Col = 4 Then
    Me.PoGrid.Text = Format(Me.txtqty.Text, "0.000")
    Call calculations
    Call cals
    Call cals1
    End If
End Sub
Private Sub txtqty_Click()
    If donotchange Then Exit Sub
    If PoGrid.Col = 4 Then
    Me.PoGrid.Text = Format(Me.txtqty.Text, "0.000")
    End If
End Sub
Private Sub txtqty_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Then
    ElseIf (KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 46 Then
    ElseIf KeyAscii = 13 And PoGrid.Col = 4 Then
    Me.PoGrid.Text = Format(Me.txtqty.Text, "0.000")
    PoGrid.Col = 5
    PoGrid_Click
    Call txtqty_Change
    tw.Text = Others.towords(Val(txtnet.Text))
    Else
    KeyAscii = 0
    End If
End Sub
Private Sub txtrate_Change()
     If donotchange Then Exit Sub
     If PoGrid.Col = 6 Then
     Me.PoGrid.Text = Format(Me.txtrate.Text, "0.00")
     tw.Text = Others.towords(Val(txtnet.Text))
     Call calculations
     Call cals
     Call cals1
     End If
End Sub
Private Sub txtrate_Click()
    If donotchange Then Exit Sub
    If PoGrid.Col = 6 Then
    Me.PoGrid.Text = Format(Me.txtrate.Text, "0.00")
    End If
End Sub
Private Sub txtrate_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Then
    ElseIf (KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 46 Then
    ElseIf KeyAscii = 13 And PoGrid.Col = 6 Then
    Me.PoGrid.Text = Format(Me.txtrate.Text, "0.00")
    tw.Text = Others.towords(Val(txtnet.Text))
    PoGrid.Col = 7
    PoGrid.SetFocus
    txtrate.Visible = False
    PoGrid_Click
    Else
    KeyAscii = 0
    End If
End Sub
Private Sub txtsupid_GotFocus()
    Dim rs As New ADODB.Recordset
    Dim exp1 As Variant
    rs.Open "Select supid from sup_details", cn, adOpenKeyset, adLockOptimistic
    If rs.RecordCount = 0 Then
    txtsupid.Text = ""
    rs.Close
    Else
    Dim rsrs As New ADODB.Recordset
    rsrs.Open "Select supid as exp1 from sup_details where supname='" & cbosupplier.Text & "'", cn, adOpenKeyset, adLockOptimistic
    txtsupid = rsrs![exp1]
    rsrs.Close
    End If
    SendKeys "{tab}"
End Sub
Private Sub txttax_Change()
    tw.Text = Others.towords(Val(txtnet.Text))
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
    ElseIf KeyAscii = 13 Then
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
    If PoGrid.Col = 5 Then
    Me.PoGrid.Text = Me.txtuom.Text
    Call txtuom_gotfocus
    PoGrid.Col = 6
    PoGrid_Click
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
    For i = 1 To PoGrid.Rows - 1
       PoGrid.TextMatrix(i, 7) = Format((Val(PoGrid.TextMatrix(i, 4)) * Val(PoGrid.TextMatrix(i, 6))), "0.00")
    Next i
End Sub
Sub cals()
    txttot = 0
    For i = 0 To PoGrid.Rows - 1
    txttot.Text = Format(Val(txttot.Text) + Val(PoGrid.TextMatrix(i, 7)), "0.00")
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
