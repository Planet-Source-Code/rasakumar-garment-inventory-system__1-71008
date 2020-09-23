VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmopendelivery 
   BackColor       =   &H00EDDDD1&
   Caption         =   " * Open Delivery *"
   ClientHeight    =   9405
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12375
   Icon            =   "frmopendelivery.frx":0000
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
      Height          =   225
      Left            =   4920
      TabIndex        =   24
      ToolTipText     =   " Enter The Delivery Qty "
      Top             =   1980
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txteditdc 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   480
      TabIndex        =   16
      Top             =   8640
      Visible         =   0   'False
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
      Height          =   615
      Left            =   6360
      MaskColor       =   &H00C0FFC0&
      Style           =   1  'Graphical
      TabIndex        =   14
      Tag             =   " "
      ToolTipText     =   " To Use Add DC"
      Top             =   9480
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
      Height          =   615
      Left            =   7920
      MaskColor       =   &H00C0FFC0&
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "  Exit Window  "
      Top             =   9480
      UseMaskColor    =   -1  'True
      Width           =   1455
   End
   Begin VB.TextBox txttype 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   405
      Left            =   9840
      TabIndex        =   12
      Text            =   "Open GRN "
      Top             =   8520
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox txttot 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   6840
      Locked          =   -1  'True
      TabIndex        =   11
      Text            =   "0"
      Top             =   8520
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.ComboBox cbosupplier 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   7320
      TabIndex        =   10
      ToolTipText     =   " Select The Supplier Name "
      Top             =   240
      Width           =   3615
   End
   Begin VB.TextBox txtdcno 
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
      Left            =   1440
      TabIndex        =   9
      ToolTipText     =   " Your GRN No "
      Top             =   240
      Width           =   1575
   End
   Begin VB.ComboBox txtuom 
      Appearance      =   0  'Flat
      ForeColor       =   &H000040C0&
      Height          =   315
      Left            =   5040
      TabIndex        =   8
      ToolTipText     =   " Select The Unit of Measurement "
      Top             =   1200
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.ComboBox cbosize 
      Appearance      =   0  'Flat
      ForeColor       =   &H000040C0&
      Height          =   315
      Left            =   3480
      TabIndex        =   7
      ToolTipText     =   " Select The Size "
      Top             =   1200
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.ComboBox cbocolour 
      Appearance      =   0  'Flat
      ForeColor       =   &H000040C0&
      Height          =   315
      Left            =   2040
      TabIndex        =   6
      ToolTipText     =   " Select The Colour "
      Top             =   1200
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.ComboBox cboitem 
      Appearance      =   0  'Flat
      ForeColor       =   &H000040C0&
      Height          =   315
      Left            =   240
      TabIndex        =   5
      ToolTipText     =   " Select The Item Name "
      Top             =   1200
      Visible         =   0   'False
      Width           =   1335
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
      Left            =   12120
      TabIndex        =   4
      ToolTipText     =   " Select the Department "
      Top             =   240
      Width           =   2895
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
      TabIndex        =   3
      ToolTipText     =   "  To Use Add Items "
      Top             =   7560
      Width           =   1215
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
      TabIndex        =   2
      Tag             =   " "
      ToolTipText     =   "To Use Delete Items  "
      Top             =   7560
      Width           =   1335
   End
   Begin VB.TextBox txtcomid 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   405
      Left            =   8400
      TabIndex        =   1
      Top             =   8520
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox txtremarks 
      Appearance      =   0  'Flat
      Height          =   855
      Left            =   4320
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      ToolTipText     =   " Enter the Remarks"
      Top             =   7440
      Width           =   10815
   End
   Begin MSComCtl2.DTPicker dt1 
      Height          =   375
      Left            =   4440
      TabIndex        =   17
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
      Format          =   58064897
      CurrentDate     =   39481
   End
   Begin MSFlexGridLib.MSFlexGrid DeliveryGrid 
      Height          =   6495
      Left            =   120
      TabIndex        =   18
      ToolTipText     =   "  Note : # Indicate Place Click The First Row of Grid To  Open the Filter Options "
      Top             =   840
      Width           =   15015
      _ExtentX        =   26485
      _ExtentY        =   11456
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
      Left            =   6360
      MaskColor       =   &H00C0FFC0&
      Style           =   1  'Graphical
      TabIndex        =   15
      Tag             =   " "
      ToolTipText     =   " To use Update PO "
      Top             =   9480
      UseMaskColor    =   -1  'True
      Width           =   1455
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   2
      Height          =   855
      Left            =   6240
      Top             =   9360
      Width           =   3255
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
      Left            =   11040
      TabIndex        =   23
      Top             =   240
      Width           =   1095
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
      Left            =   6000
      TabIndex        =   22
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "OPEN DC Date"
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
      Left            =   3000
      TabIndex        =   21
      Top             =   240
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "OPEN DC NO"
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
      Left            =   240
      TabIndex        =   20
      Top             =   240
      Width           =   1215
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H000000FF&
      Height          =   615
      Index           =   1
      Left            =   120
      Top             =   120
      Width           =   15015
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H000000FF&
      Height          =   855
      Left            =   120
      Top             =   7440
      Width           =   2895
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H000000FF&
      Height          =   855
      Left            =   120
      Top             =   8400
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
      Left            =   3120
      TabIndex        =   19
      Top             =   7440
      Width           =   1215
   End
End
Attribute VB_Name = "frmopendelivery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim op As Variant
Dim cn As ADODB.Connection
Dim rs As ADODB.Recordset
Dim ww As ADODB.Recordset
Private Sub cbocolour_Click()
    If DeliveryGrid.Col = 2 Then
        Me.DeliveryGrid.Text = Me.cbocolour.Text
         Call cbocolour_dropdown
            DeliveryGrid.Col = 3
     Call DeliveryGrid_Click
 End If
End Sub

Private Sub cbocolour_dropdown()
Dim rs As New ADODB.Recordset
        On Error GoTo X
        cbocolour.Clear
        Set rs = cn.Execute("select colour from colour_details group by colour")
        rs.MoveFirst
        Do While Not rs.EOF()
        cbocolour.additem (rs(0))
        rs.MoveNext
        Loop
        cbocolour.SetFocus
X:
End Sub

Private Sub cbocolour_KeyPress(KeyAscii As Integer)
 KeyAscii = 0
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

Private Sub cboitem_Click()
If DeliveryGrid.Col = 1 Then
        Me.DeliveryGrid.Text = Me.cboitem.Text
         Call cboitem_dropdown
            DeliveryGrid.Col = 2
     Call DeliveryGrid_Click
 End If
End Sub
Private Sub cboitem_dropdown()
Dim rs As New ADODB.Recordset
        On Error GoTo X
        cboitem.Clear
        Set rs = cn.Execute("select itemname from item_details group by itemname")
        rs.MoveFirst
        Do While Not rs.EOF()
        cboitem.additem (rs(0))
        rs.MoveNext
        Loop
        cboitem.SetFocus
X:
End Sub
Private Sub cboitem_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub cbosize_Click()
    If DeliveryGrid.Col = 3 Then
        Me.DeliveryGrid.Text = Me.cbosize.Text
         Call cbosize_dropdown
            DeliveryGrid.Col = 4
     Call DeliveryGrid_Click
 End If
End Sub
Private Sub cbosize_dropdown()
Dim rs As New ADODB.Recordset
        On Error GoTo X
        cbosize.Clear
        Set rs = cn.Execute("select sizes from size_details group by sizes")
        rs.MoveFirst
        Do While Not rs.EOF()
        cbosize.additem (rs(0))
        rs.MoveNext
        Loop
        cbosize.SetFocus
X:
End Sub

Private Sub cbosize_KeyPress(KeyAscii As Integer)
 KeyAscii = 0
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

Private Sub cmdaddrow_Click()
     If DeliveryGrid.Rows > 20 Then
     MsgBox "Only 10 Item Allowed ", vbCritical, "Exceed Row "
     Else
     DeliveryGrid.Rows = DeliveryGrid.Rows + 1
     DeliveryGrid.Row = DeliveryGrid.Rows - 1
     DeliveryGrid.TextMatrix(DeliveryGrid.Row, 0) = DeliveryGrid.Rows - 1
     End If
End Sub
Private Sub cmddeleterow_Click()
     If DeliveryGrid.Rows <= 1 Then
     MsgBox "No Record Found", vbInformation, "Information"
     Call frmgrnopendeliveryload
     Else
     DeliveryGrid.Rows = DeliveryGrid.Rows - 1
     DeliveryGrid.Row = DeliveryGrid.Rows - 1
     DeliveryGrid.TextMatrix(DeliveryGrid.Row, 0) = DeliveryGrid.Rows - 1
     End If
End Sub

Private Sub cmdexit_Click()
    op = MsgBox("Aru You Sure To Close ?", vbQuestion + vbYesNo, "Confirm Close ?")
    If op = vbYes Then
        Unload Me
        frmopensdeliverymain.Show
    Else
    End If
End Sub
Private Sub cmdsave_Click()
    If Trim(cbosupplier.Text) = "" Then
    MsgBox "Please Select The Supplier Name ", vbCritical, "Supplier Name Error "
    cbosupplier.SetFocus
    ElseIf Trim(cbodept.Text) = "" Then
    MsgBox "Please Enter The Department Name ", vbCritical, "Department Name Error"
    cbodept.SetFocus
    ElseIf Val(txttot.Text) <= 0 Then
    MsgBox "Please Enter The Valid Issue Quantity Details", vbCritical, "Issue Qty Error"
    DeliveryGrid.Col = 4
    DeliveryGrid.SetFocus
    ElseIf Trim(DeliveryGrid.Text) = "" Then
    MsgBox "Please Enter All Data", vbCritical, "Data Enter Error"
    Else
        Dim rs As New ADODB.Recordset
        Dim rsstatus As New ADODB.Recordset
        
        rs.Open "select * from opendelivery_details where opdcno =' " & txtdcno.Text & "'", cn, adOpenKeyset, adLockOptimistic
               
        If rs.RecordCount = 0 Then
        For i = 1 To DeliveryGrid.Rows - 1
        rs.AddNew
        rs![opdcno] = txtdcno.Text
        rs![opdcdate] = dt1.Value
        rs![supname] = cbosupplier.Text
        rs![deptname] = cbodept.Text
        rs![sno] = DeliveryGrid.TextMatrix(i, 0)
        rs![itemname] = DeliveryGrid.TextMatrix(i, 1)
        rs![colour] = DeliveryGrid.TextMatrix(i, 2)
        rs![Size] = DeliveryGrid.TextMatrix(i, 3)
        rs![deliveryqty] = DeliveryGrid.TextMatrix(i, 4)
        rs![uom] = DeliveryGrid.TextMatrix(i, 5)
        rs![remarks] = txtremarks.Text
        Next i
        
        rs.Update
        rs.Close
        MsgBox "One Record Save Successfully ", vbInformation, "Information"
        Unload Me
        frmopensdeliverymain.Show
        Else
        MsgBox "Already Exists", vbCritical, "Error"
    End If
    End If
        
    
End Sub
Private Sub cmdupdate_Click()
    If Trim(cbosupplier.Text) = "" Then
    MsgBox "Please Select The Supplier Name ", vbCritical, "Supplier Name Error "
    cbosupplier.SetFocus
    ElseIf Trim(cbodept.Text) = "" Then
    MsgBox "Please Enter The Department Name ", vbCritical, "Department Name Error"
    cbodept.SetFocus
    ElseIf Val(txttot.Text) <= 0 Then
    MsgBox "Please Enter The Valid Issue Quantity Details", vbCritical, "Issue Qty Error"
    DeliveryGrid.Col = 4
    DeliveryGrid.SetFocus
    ElseIf Trim(DeliveryGrid.Text) = "" Then
    MsgBox "Please Enter All Data", vbCritical, "Data Enter Error"
    Else
        Dim rs As New ADODB.Recordset
        rs.Open "Select * from opendelivery_details where opdcno = '" & txtdcno.Text & "'", cn, 1, 3
        If rs.RecordCount <> 0 Then
        i = 1
        If rs.BOF = False Then rs.MoveFirst
        While rs.EOF = False
        rs![opdcno] = txtdcno.Text
        rs![opdcdate] = dt1.Value
        rs![supname] = cbosupplier.Text
        rs![deptname] = cbodept.Text
        rs![sno] = DeliveryGrid.TextMatrix(i, 0)
        rs![itemname] = DeliveryGrid.TextMatrix(i, 1)
        rs![colour] = DeliveryGrid.TextMatrix(i, 2)
        rs![Size] = DeliveryGrid.TextMatrix(i, 3)
        rs![deliveryqty] = DeliveryGrid.TextMatrix(i, 4)
        rs![uom] = DeliveryGrid.TextMatrix(i, 5)
        rs![remarks] = txtremarks.Text
        rs.Update
        rs.MoveNext
        i = i + 1
        Wend
        MsgBox "One Record Updated Successfully", vbInformation, "Information"
        Unload Me
        frmopensdeliverymain.Show
        End If
        End If
        
End Sub
Private Sub DeliveryGrid_Click()
    If cboitem.Visible = True Then cboitem.Visible = False
    If cbocolour.Visible = True Then cbocolour.Visible = False
    If cbosize.Visible = True Then cbosize.Visible = False
    If txtuom.Visible = True Then txtuom.Visible = False
    If txtqty.Visible = True Then txtqty.Visible = False
    
    
    If DeliveryGrid.Col = 1 Then
    Me.cboitem.Visible = True
    CurrentRow = Me.DeliveryGrid.Row
    Me.cboitem.Width = Me.DeliveryGrid.CellWidth - 10
    Me.cboitem.Left = Me.DeliveryGrid.CellLeft + Me.DeliveryGrid.Left
    Me.cboitem.Top = Me.DeliveryGrid.CellTop + Me.DeliveryGrid.Top
    donotchange = True
    Me.cboitem.Text = Me.DeliveryGrid.Text
    donotchange = False
    Me.cboitem.SetFocus
    ElseIf DeliveryGrid.Col = 2 Then
    Me.cbocolour.Visible = True
    CurrentRow = Me.DeliveryGrid.Row
    Me.cbocolour.Width = Me.DeliveryGrid.CellWidth - 10
    Me.cbocolour.Left = Me.DeliveryGrid.CellLeft + Me.DeliveryGrid.Left
    Me.cbocolour.Top = Me.DeliveryGrid.CellTop + Me.DeliveryGrid.Top
    donotchange = True
    Me.cbocolour.Text = Me.DeliveryGrid.Text
    donotchange = False
    Me.cbocolour.SetFocus
    ElseIf DeliveryGrid.Col = 3 Then
    Me.cbosize.Visible = True
    CurrentRow = Me.DeliveryGrid.Row
    Me.cbosize.Width = Me.DeliveryGrid.CellWidth - 10
    Me.cbosize.Left = Me.DeliveryGrid.CellLeft + Me.DeliveryGrid.Left
    Me.cbosize.Top = Me.DeliveryGrid.CellTop + Me.DeliveryGrid.Top
    donotchange = True
    Me.cbosize.Text = Me.DeliveryGrid.Text
    donotchange = False
    Me.cbosize.SetFocus
    ElseIf DeliveryGrid.Col = 4 Then
    Me.txtqty.Visible = True
    CurrentRow = Me.DeliveryGrid.Row
    Me.txtqty.Width = Me.DeliveryGrid.CellWidth - 10
    Me.txtqty.Left = Me.DeliveryGrid.CellLeft + Me.DeliveryGrid.Left
    Me.txtqty.Top = Me.DeliveryGrid.CellTop + Me.DeliveryGrid.Top
    donotchange = True
    Me.txtqty.Text = Me.DeliveryGrid.Text
    donotchange = False
    Me.txtqty.SetFocus
    ElseIf DeliveryGrid.Col = 5 Or DeliveryGrid.Col = 1 Then
    Call txtuom_dropdown
    Me.txtuom.Visible = True
    CurrentRow = Me.DeliveryGrid.Row
    Me.txtuom.Width = Me.DeliveryGrid.CellWidth - 10
    Me.txtuom.Left = Me.DeliveryGrid.CellLeft + Me.DeliveryGrid.Left
    Me.txtuom.Top = Me.DeliveryGrid.CellTop + Me.DeliveryGrid.Top
    donotchange = True
    Me.txtuom.Text = Me.DeliveryGrid.Text
    donotchange = False
    Me.txtuom.SetFocus
    End If
End Sub
Private Sub Form_Load()
    Set cn = New ADODB.Connection
    Set ww = New ADODB.Recordset
    Set rs = New ADODB.Recordset
    cn.Open "Provider=Microsoft.jet.oledb.4.0; Data Source=" & App.path & "\Database\Data.mdb"
    ww.Open "Select * From opendelivery_details", cn, adOpenKeyset, adLockOptimistic
    Call calculates
    Call frmgrnopendeliveryload
    Call txtgrnno_Gotfocus
    DeliveryGrid.TextMatrix(DeliveryGrid.Row, 0) = DeliveryGrid.Rows - 1
End Sub
Private Sub txtgrnno_Gotfocus()
    Dim rs As New ADODB.Recordset
    Dim a As String
    rs.Open "Select * from opendelivery_details", cn, adOpenKeyset, adLockOptimistic
        If rs.RecordCount = 0 Then
        txtdcno.Text = 1
        rs.Close
    Else
    Dim rsrs As New ADODB.Recordset
    rsrs.Open "Select max(opdcno)as exp1 from opendelivery_details", cn, adOpenKeyset, adLockOptimistic
       txtdcno = rsrs![exp1] + 1
       rsrs.Close
    End If
    SendKeys "{tab}"
    
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
Call calculates
End Sub

Private Sub Form_Unload(Cancel As Integer)
        Unload Me
        frmopensdeliverymain.Show
End Sub
Private Sub txtqty_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Then
    ElseIf (KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 46 Then
    ElseIf KeyAscii = 13 And DeliveryGrid.Col = 4 Then
        Me.DeliveryGrid.Text = Format(Me.txtqty.Text, "0.000")
         Call calculates
'         Call cbocolour_dropdown
            DeliveryGrid.Col = 5
     Call DeliveryGrid_Click
     Else
     KeyAscii = 0
     End If
End Sub
Private Sub txtqty_OnEnter()
    If DeliveryGrid.Col = 4 And KeyAscii = 13 Then
        Me.DeliveryGrid.Text = Format(Me.txtqty.Text, "0.000")
            Call cbocolour_dropdown
            DeliveryGrid.Col = 5
     Call DeliveryGrid_Click
 End If
End Sub
Private Sub txtuom_Click()
     If DeliveryGrid.Col = 5 Then
        Me.DeliveryGrid.Text = Me.txtuom.Text
         Call txtuom_dropdown
            DeliveryGrid.Col = 1
      DeliveryGrid.SetFocus
      txtuom.Visible = False
     End If
End Sub
Private Sub txtuom_dropdown()
Dim rs As New ADODB.Recordset
        On Error GoTo X
        txtuom.Clear
        Set rs = cn.Execute("select uom from uom_details group by uom")
        rs.MoveFirst
        Do While Not rs.EOF()
        txtuom.additem (rs(0))
        rs.MoveNext
        Loop
        txtuom.SetFocus
X:
End Sub
Private Function calculates()
    Dim i As Integer
    txttot = 0
    For i = 0 To DeliveryGrid.Rows - 1
    txttot.Text = Format(Val(txttot.Text) + Val(DeliveryGrid.TextMatrix(i, 4)), "0.000")
    Next i
End Function

Private Sub txtuom_KeyPress(KeyAscii As Integer)
 KeyAscii = 0
End Sub
