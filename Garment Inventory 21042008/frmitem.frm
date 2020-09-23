VERSION 5.00
Object = "{B69D5E45-990C-4D4D-906E-FF041974C40B}#1.0#0"; "osenxpsuite2005.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmitem 
   BackColor       =   &H00EDDDD1&
   BorderStyle     =   1  'Fixed Single
   Caption         =   " * Item Details *"
   ClientHeight    =   9585
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10860
   Icon            =   "frmitem.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9585
   ScaleWidth      =   10860
   Begin VB.ComboBox cbofilter1 
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
      Height          =   360
      Left            =   6840
      Style           =   2  'Dropdown List
      TabIndex        =   16
      ToolTipText     =   " Filter The Single Departments "
      Top             =   3240
      Width           =   2895
   End
   Begin VB.ComboBox cbodept 
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
      Height          =   360
      Left            =   5400
      TabIndex        =   5
      ToolTipText     =   " Select the Department Name "
      Top             =   1800
      Width           =   3615
   End
   Begin VB.TextBox txtedit 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
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
      Left            =   2880
      TabIndex        =   13
      Top             =   720
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.CommandButton cmdexit 
      BackColor       =   &H00FFC0C0&
      Caption         =   "E&xit"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
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
      TabIndex        =   12
      ToolTipText     =   " Exit Window "
      Top             =   2520
      UseMaskColor    =   -1  'True
      Width           =   1695
   End
   Begin VB.CommandButton cmdadd 
      BackColor       =   &H00FFC0C0&
      Caption         =   "&Add Item"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      MaskColor       =   &H00C0FFC0&
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   " To Use Add Item "
      Top             =   2520
      UseMaskColor    =   -1  'True
      Width           =   1575
   End
   Begin VB.CommandButton cmdedit 
      BackColor       =   &H00FFC0C0&
      Caption         =   "&Edit Item "
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1920
      MaskColor       =   &H00C0FFC0&
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   " To Use Edit Item "
      Top             =   2520
      UseMaskColor    =   -1  'True
      Width           =   1695
   End
   Begin VB.CommandButton cmdupdate 
      BackColor       =   &H00FFC0C0&
      Caption         =   "&Update  Item"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3720
      MaskColor       =   &H00C0FFC0&
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   " To Use Update Item "
      Top             =   2520
      UseMaskColor    =   -1  'True
      Width           =   1695
   End
   Begin VB.CommandButton cmddelete 
      BackColor       =   &H00FFC0C0&
      Caption         =   "&Delete Item"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5520
      MaskColor       =   &H00C0FFC0&
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   " To Use Delete Item "
      Top             =   2520
      UseMaskColor    =   -1  'True
      Width           =   1575
   End
   Begin VB.CommandButton cmdprint 
      BackColor       =   &H00FFC0C0&
      Caption         =   "&Print Item"
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
      Height          =   495
      Left            =   7200
      MaskColor       =   &H00C0FFC0&
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   " To Use Print Item "
      Top             =   2520
      UseMaskColor    =   -1  'True
      Width           =   1695
   End
   Begin VB.ComboBox cbofilter 
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
      Height          =   360
      ItemData        =   "frmitem.frx":0442
      Left            =   840
      List            =   "frmitem.frx":0444
      Style           =   2  'Dropdown List
      TabIndex        =   6
      ToolTipText     =   " Filter The Single Items "
      Top             =   3240
      Width           =   3495
   End
   Begin VB.CommandButton cmdlist 
      BackColor       =   &H00FFC0C0&
      Caption         =   "&List All Items "
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
      Left            =   4440
      MaskColor       =   &H00C0FFC0&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3240
      UseMaskColor    =   -1  'True
      Width           =   1815
   End
   Begin VB.ComboBox cbouom 
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
      Height          =   360
      Left            =   5400
      TabIndex        =   4
      ToolTipText     =   " Select The Unit of Measurement "
      Top             =   1440
      Width           =   3615
   End
   Begin VB.TextBox txtitem 
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
      Left            =   5400
      TabIndex        =   1
      ToolTipText     =   " Enter The Item Name "
      Top             =   1080
      Width           =   3615
   End
   Begin VB.TextBox txtid 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
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
      Left            =   5400
      TabIndex        =   14
      Top             =   720
      Visible         =   0   'False
      Width           =   3615
   End
   Begin osenxpsuite2005.OsenXPLabel OsenXPLabel3 
      Height          =   615
      Index           =   0
      Left            =   3000
      TabIndex        =   17
      Top             =   1080
      Width           =   300
      _ExtentX        =   529
      _ExtentY        =   1085
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "*"
      ForeColor       =   255
      BackStyle       =   0
   End
   Begin osenxpsuite2005.OsenXPLabel OsenXPLabel3 
      Height          =   615
      Index           =   1
      Left            =   3000
      TabIndex        =   18
      Top             =   1440
      Width           =   300
      _ExtentX        =   529
      _ExtentY        =   1085
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "*"
      ForeColor       =   255
      BackStyle       =   0
   End
   Begin osenxpsuite2005.OsenXPLabel OsenXPLabel3 
      Height          =   615
      Index           =   2
      Left            =   3000
      TabIndex        =   19
      Top             =   1800
      Width           =   300
      _ExtentX        =   529
      _ExtentY        =   1085
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "*"
      ForeColor       =   255
      BackStyle       =   0
   End
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   5655
      Left            =   0
      TabIndex        =   20
      ToolTipText     =   "  Note : # Indicate Place Click The First Row of Grid To  Open the Filter Options "
      Top             =   3600
      Width           =   10815
      _ExtentX        =   19076
      _ExtentY        =   9975
      _Version        =   393216
      Cols            =   4
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
      BackColor       =   &H00EDDDD1&
      Caption         =   "ITEM  MASTER"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   495
      Left            =   4560
      TabIndex        =   21
      Top             =   120
      Width           =   3135
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "       Name of Department"
      BeginProperty Font 
         Name            =   "Arial"
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
      Left            =   2880
      TabIndex        =   15
      Top             =   1800
      Width           =   2535
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00400000&
      BorderWidth     =   2
      Height          =   735
      Left            =   120
      Top             =   2400
      Width           =   10695
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Unit of Measurement "
      BeginProperty Font 
         Name            =   "Arial"
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
      Left            =   2880
      TabIndex        =   2
      Top             =   1440
      Width           =   2535
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Item Name "
      BeginProperty Font 
         Name            =   "Arial"
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
      Left            =   2880
      TabIndex        =   0
      Top             =   1080
      Width           =   2535
   End
End
Attribute VB_Name = "frmitem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim op As Variant
Dim cn As ADODB.Connection
Dim rs As ADODB.Recordset
Dim ww As ADODB.Recordset
Private Sub cbodept_Click()
        cmdadd.SetFocus
End Sub
Private Sub cbodept_dropdown()
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
End Sub
Private Sub cbodept_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub cbofilter_Click()
    Call filter
    cmdlist.Visible = True
    Call frmitemgrid
End Sub
Private Sub cbofilter_DropDown()
    On Error GoTo X
        If Len(cbofilter1.Text) <= 0 Then
        cbofilter.Clear
        Set rs = cn.Execute("select itemname from item_details order by itemname")
        rs.MoveFirst
        cbofilter.additem "                     <Ascending>"
        cbofilter.additem "                     <Decending>"
        Do While Not rs.EOF()
        cbofilter.additem (rs(0))
        rs.MoveNext
        Loop
        cbofilter.SetFocus
X:
ElseIf Len(cbofilter1.Text) > 0 Then
        On Error GoTo y
        cbofilter.Clear
        Set rs = cn.Execute("select itemname from item_details where deptname= '" & cbofilter1.Text & "'")
        rs.MoveFirst
        cbofilter.additem "                     <Ascending>"
        cbofilter.additem "                     <Decending>"
        Do While Not rs.EOF()
        cbofilter.additem (rs(0))
        rs.MoveNext
        Loop
        cbofilter.SetFocus
y:
End If
End Sub
Private Sub cbofilter1_Click()
    Call filter1
    Call frmitemgrid
    Dim rs As New ADODB.Recordset
    rs.Open "SELECT * From item_details WHERE deptname='" & cbofilter1.Text & "'", cn, adOpenStatic, adLockPessimistic
    cmdlist.Visible = True
    Grid.Row = 0
    Grid.CellBackColor = RGB(85, 194, 154)
End Sub
Private Sub cbofilter1_DropDown()
    Dim rs As New ADODB.Recordset
    On Error GoTo X
     cbofilter1.Clear
        Set rs = cn.Execute("select deptname from dept_details order by deptname")
        rs.MoveFirst
        Do While Not rs.EOF()
         cbofilter1.additem (rs(0))
         rs.MoveNext
         Loop
         cbofilter1.SetFocus
X:
End Sub
Public Sub cbouom_DropDown()
    Dim rs As New ADODB.Recordset
        On Error GoTo X
        cbouom.Clear
        Set rs = cn.Execute("select uom from uom_details order by uom")
        rs.MoveFirst
        Do While Not rs.EOF()
        cbouom.additem (rs(0))
        rs.MoveNext
        Loop
        cbouom.SetFocus
X:
End Sub
Private Sub cbouom_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub
Private Sub cmdadd_Click()
    If Trim(txtitem.Text) = "" Then
    MsgBox "Please Enter The Item Name", vbCritical, "Item Name Error"
    txtitem.SetFocus
    ElseIf Trim(cbouom.Text) = "" Then
    MsgBox "Please Select the Unit of Measurement ", vbCritical, "UOM Error"
    cbouom.SetFocus
    ElseIf Trim(cbodept.Text) = "" Then
    MsgBox "Please Select the Department Name ", vbCritical, "Department Name Error"
    cbodept.SetFocus
    Else
    Dim rs As New ADODB.Recordset
    rs.Open "Select * from item_details where itemid= " & txtid.Text, cn, adOpenKeyset, adLockOptimistic
        If rs.RecordCount = 0 Then
        rs.AddNew
        rs![itemid] = txtid.Text
        rs![itemname] = txtitem.Text
        rs![uom] = cbouom.Text
        rs![deptname] = cbodept.Text
        rs.Update
        rs.Clone
        MsgBox "One Record Save Successfully", vbInformation, "Information"
        Unload Me
        frmitem.Show
        
        Else
        MsgBox "This Item Already Exists", vbCritical, "Invalid Item"
        
        End If
     End If
End Sub
Private Sub cmddelete_Click()
         If Trim(txtedit.Text) = "" Then
         MsgBox "Please Select The Item Name", vbCritical, "Selecting Error"
         Else
         If MsgBox("Are You Sure Delete This Record  " & txtedit.Text & " ? ", vbQuestion + vbYesNo, "Confirm To Delete") = vbYes Then
         Dim rs As New ADODB.Recordset
             rs.Open "Select * from item_details where itemname ='" & txtedit.Text & "'", cn, adOpenKeyset, adLockOptimistic
             If rs.RecordCount <> 0 Then
             rs.Delete
             rs.Requery
             rs.Close
             MsgBox "One Record Deleted Successfully", vbInformation, "Information"
             Unload Me
             frmitem.Show
             
             Else
             MsgBox "Please Select The Item Name ", vbCritical, "Invalid"
             End If
         Else
         End If
    End If
End Sub
Private Sub cmdedit_Click()
        cmdadd.Enabled = False
        cmdupdate.Enabled = True
        On Error Resume Next
        Dim X As Double
        X = Val(txtedit.Text)
        Dim rs As New ADODB.Recordset
        rs.Open "Select * from item_details where itemname ='" & txtedit.Text & "'", cn, adOpenKeyset, adLockOptimistic
        If rs.RecordCount <> 0 Then
            txtid.Text = rs![itemid]
            txtitem.Text = rs![itemname]
            cbouom.Text = rs![uom]
            cbodept.Text = rs![deptname]
        rs.Close
        Else
        MsgBox "Please Select The Item Name  ", vbCritical, "Invalid"
        cmdadd.Enabled = True
        cmdupdate.Enabled = False
        End If
End Sub
Private Sub cmdexit_Click()
    op = MsgBox("Are You Sure To Close ?", vbYesNo + vbQuestion, "Confirm Close ?")
    If op = vbYes Then
    Unload Me
    Else
    End If
End Sub
Private Sub cmdlist_Click()
    Call gridloads
    cbofilter.Clear
    cbofilter1.Clear
    Grid.Col = 1
    Grid.Row = 0
    Grid.CellBackColor = &HBF630F
    Grid.Col = 3
    Grid.Row = 0
    Grid.CellBackColor = &HBF630F
    cmdlist.Visible = False
    cbofilter.Visible = False
    cbofilter1.Visible = False
End Sub
Private Sub cmdupdate_Click()
            If Trim(txtitem.Text) = "" Then
            MsgBox "Item Name Is Empty ", vbCritical, "Company Name Error"
            txtitem.SetFocus
            ElseIf Trim(cbouom.Text) = "" Then
            MsgBox "UOM Address is Empty ", vbCritical, "Address Error"
            cbouom.SetFocus
            Else
            Dim rs As New ADODB.Recordset
                rs.Open "Select *  from item_details where itemid=" & txtid.Text, cn, adOpenKeyset, adLockOptimistic
                    If rs.RecordCount <> 0 Then
                    rs![itemid] = txtid.Text
                    rs![itemname] = txtitem.Text
                    rs![uom] = cbouom.Text
                    rs![deptname] = cbodept.Text
                    rs.Update
                    rs.Close
                    MsgBox "One Record Updated Successfully", vbInformation, "Information"
                    Unload Me
                    frmitem.Show
                    
                    cmdadd.Enabled = True
                    cmdupdate.Enabled = False
                    Else
                    MsgBox "Already This Item Exists", vbCritical, "Invalid Item"
                    
                    cmdadd.Enabled = True
                    cmdupdate.Enabled = False
                    End If
         End If
End Sub
Private Sub Form_Load()
    Dim i As Integer
    Set cn = New ADODB.Connection
    Set ww = New ADODB.Recordset
    cn.Open "Provider=Microsoft.jet.oledb.4.0; Data Source=" & App.path & "\Database\Data.mdb"
    ww.Open "Select * From item_details", cn, adOpenKeyset, adLockOptimistic
    Call txtid_GotFocus
    txtitem.TabIndex = 4
    cmdlist.Visible = False
    Call frmitemgrid
    cn.CursorLocation = adUseClient
    cbofilter.Visible = False
    cbofilter1.Visible = False
    Call gridloads
    
    cmdupdate.Enabled = False
End Sub
Private Sub Grid_Click()
    If Grid.Col = 1 Then
    txtedit.Text = Grid
    End If
    If Grid.Col = 1 And Grid.Row = 1 Then
    cbofilter.Visible = True
    Else
    cbofilter.Visible = False
    End If
    If Grid.Col = 3 And Grid.Row = 1 Then
    cbofilter1.Visible = True
    Else
    cbofilter1.Visible = False
    End If
End Sub
Private Sub Grid_DblClick()
    If Grid.Col = 1 Then
    txtedit.Text = Grid
    Call cmdedit_Click
    End If
    If Grid.Col = 1 And Grid.Row = 1 Then
    cbofilter.Visible = True
    Else
    cbofilter.Visible = False
    End If
End Sub
Private Sub txtid_GotFocus()
Dim rs As New ADODB.Recordset
rs.Open "Select * from item_details", cn, adOpenKeyset, adLockOptimistic
    If rs.RecordCount = 0 Then
    txtid.Text = 1
    rs.Close
    Else
    Dim rsrs As New ADODB.Recordset
    rsrs.Open "Select max(itemid)as exp1 from item_details", cn, adOpenKeyset, adLockOptimistic
    txtid = rsrs![exp1] + 1
    rsrs.Close
    End If
SendKeys "{tab}"
End Sub
Sub gridloads()
    Dim i As Integer
    Dim rs As New ADODB.Recordset
    rs.Open "Select * from item_details", cn, adOpenKeyset, adLockOptimistic
    i = 1
    If rs.BOF = False Then rs.MoveFirst
    While rs.EOF = False
    Grid.Rows = Grid.Rows + 1
    Grid.TextMatrix(i, o) = i
    Grid.TextMatrix(i, 1) = rs![itemname]
    Grid.TextMatrix(i, 2) = rs![uom]
    Grid.TextMatrix(i, 3) = rs![deptname]
    rs.MoveNext
    i = i + 1
    Wend
    Grid.Rows = rs.RecordCount + 1
End Sub
Sub filter()
    Dim i As Integer
    Dim rs As New ADODB.Recordset
    rs.Open "Select * from item_details where itemname='" & cbofilter.Text & "'", cn, adOpenKeyset, adLockOptimistic
        For i = 1 To rs.RecordCount
                Grid.TextMatrix(i, 0) = i
                Grid.TextMatrix(i, 1) = rs![itemname]
                Grid.TextMatrix(i, 2) = rs![uom]
                Grid.TextMatrix(i, 3) = rs![deptname]
        Grid.Rows = rs.RecordCount + 1
        Grid.Row = 0
        Grid.CellBackColor = RGB(85, 194, 154)
        Next i
        If Trim(cbofilter.Text) = "<Ascending>" Then
        Grid.Sort = flexSortGenericAscending
        Grid.Row = 0
        Grid.CellBackColor = RGB(85, 194, 154)
        ElseIf Trim(cbofilter.Text) = "<Decending>" Then
        Grid.Sort = flexSortGenericDescending
        Grid.Row = 0
        Grid.CellBackColor = RGB(85, 194, 154)
        End If
End Sub
Sub filter1()
    Dim j As Integer
    Dim rs As New ADODB.Recordset
    rs.Open "select * from item_details where deptname = '" & Trim(cbofilter1.Text) & "'", cn, adOpenKeyset, adLockOptimistic
    i = 1
    If rs.BOF = False Then rs.MoveFirst
    While rs.EOF = False
    Grid.Rows = Grid.Rows + 1
    Grid.TextMatrix(i, 0) = i
    Grid.TextMatrix(i, 1) = rs![itemname]
    Grid.TextMatrix(i, 2) = rs![uom]
    Grid.TextMatrix(i, 3) = rs![deptname]
    rs.MoveNext
    i = i + 1
    Wend
    Grid.Rows = rs.RecordCount + 1
End Sub
Private Sub txtitem_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
