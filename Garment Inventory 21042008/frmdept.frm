VERSION 5.00
Object = "{B69D5E45-990C-4D4D-906E-FF041974C40B}#1.0#0"; "osenxpsuite2005.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmdept 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " * Department Details *"
   ClientHeight    =   9060
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10920
   Icon            =   "frmdept.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9060
   ScaleWidth      =   10920
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   360
      Top             =   360
   End
   Begin VB.TextBox txtedit 
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
      Left            =   2520
      TabIndex        =   0
      Top             =   720
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.CommandButton cmdexit 
      BackColor       =   &H00FFC0C0&
      Caption         =   "E&xit"
      Height          =   495
      Left            =   9000
      MaskColor       =   &H00C0FFC0&
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   " Exit Window "
      Top             =   2160
      UseMaskColor    =   -1  'True
      Width           =   1695
   End
   Begin VB.CommandButton cmdadd 
      BackColor       =   &H00FFC0C0&
      Caption         =   "&Add Department"
      Height          =   495
      Left            =   240
      MaskColor       =   &H00C0FFC0&
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "To Use Add Department "
      Top             =   2160
      UseMaskColor    =   -1  'True
      Width           =   1575
   End
   Begin VB.CommandButton cmdedit 
      BackColor       =   &H00FFC0C0&
      Caption         =   "&Edit Department"
      Height          =   495
      Left            =   1920
      MaskColor       =   &H00C0FFC0&
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   " To Use Edit Department "
      Top             =   2160
      UseMaskColor    =   -1  'True
      Width           =   1695
   End
   Begin VB.CommandButton cmdupdate 
      BackColor       =   &H00FFC0C0&
      Caption         =   "&Update Department"
      Height          =   495
      Left            =   3720
      MaskColor       =   &H00C0FFC0&
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   " To Use Update Department "
      Top             =   2160
      UseMaskColor    =   -1  'True
      Width           =   1695
   End
   Begin VB.CommandButton cmddelete 
      BackColor       =   &H00FFC0C0&
      Caption         =   "&Delete Department"
      Height          =   495
      Left            =   5520
      MaskColor       =   &H00C0FFC0&
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   " To Use Delete Department "
      Top             =   2160
      UseMaskColor    =   -1  'True
      Width           =   1575
   End
   Begin VB.CommandButton cmdprint 
      BackColor       =   &H00FFC0C0&
      Caption         =   "&Print Department"
      Height          =   495
      Left            =   7200
      MaskColor       =   &H00C0FFC0&
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   " To Use Print Department "
      Top             =   2160
      UseMaskColor    =   -1  'True
      Width           =   1695
   End
   Begin VB.ComboBox cbofilter 
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
      ItemData        =   "frmdept.frx":0442
      Left            =   840
      List            =   "frmdept.frx":0444
      TabIndex        =   11
      Text            =   "             < Select Department  >"
      Top             =   2880
      Width           =   3495
   End
   Begin VB.CommandButton cmdlist 
      BackColor       =   &H00FFC0C0&
      Caption         =   "&List All Depts"
      Height          =   375
      Left            =   4440
      MaskColor       =   &H00C0FFC0&
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   2880
      UseMaskColor    =   -1  'True
      Width           =   1695
   End
   Begin VB.TextBox txtdept 
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
      Left            =   4560
      TabIndex        =   3
      ToolTipText     =   " Enter The Department Name "
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
      Left            =   4680
      TabIndex        =   1
      Top             =   720
      Visible         =   0   'False
      Width           =   3615
   End
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   5535
      Left            =   0
      TabIndex        =   12
      Top             =   3240
      Width           =   10815
      _ExtentX        =   19076
      _ExtentY        =   9763
      _Version        =   393216
      BackColor       =   16777215
      ForeColorFixed  =   8388608
      ForeColorSel    =   12647934
      GridColorFixed  =   8421504
      GridLines       =   2
      MergeCells      =   1
      AllowUserResizing=   3
      FormatString    =   ""
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
   Begin osenxpsuite2005.OsenXPLabel OsenXPLabel1 
      Height          =   660
      Left            =   3120
      TabIndex        =   13
      Top             =   360
      Width           =   4770
      _ExtentX        =   8414
      _ExtentY        =   1164
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontBold        =   -1  'True
      Caption         =   "Department  Details "
      ForeColor       =   0
      BackStyle       =   0
   End
   Begin osenxpsuite2005.OsenXPLabel OsenXPLabel2 
      Height          =   285
      Left            =   4320
      TabIndex        =   14
      Top             =   8760
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   503
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontBold        =   -1  'True
      Caption         =   "Note : # Indicate Place Click The First Row of Grid To Open the Filter Options"
      ForeColor       =   0
      BackStyle       =   0
   End
   Begin osenxpsuite2005.OsenXPLabel OsenXPLabel5 
      Height          =   255
      Left            =   2280
      TabIndex        =   15
      Top             =   8760
      Width           =   1380
      _ExtentX        =   2434
      _ExtentY        =   450
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontBold        =   -1  'True
      Caption         =   "OsenXPLabel5"
      ForeColor       =   16711680
      BackStyle       =   0
   End
   Begin osenxpsuite2005.OsenXPLabel OsenXPLabel4 
      Height          =   255
      Left            =   0
      TabIndex        =   16
      Top             =   8760
      Width           =   2340
      _ExtentX        =   4128
      _ExtentY        =   450
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontBold        =   -1  'True
      Caption         =   "Total No of Departments :"
      ForeColor       =   0
      BackStyle       =   0
   End
   Begin osenxpsuite2005.OsenXPLabel OsenXPLabel3 
      Height          =   615
      Index           =   1
      Left            =   2280
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
   Begin VB.Shape Shape1 
      BorderColor     =   &H00400000&
      BorderWidth     =   2
      Height          =   735
      Left            =   120
      Top             =   2040
      Width           =   10695
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Department Name"
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
      Left            =   2280
      TabIndex        =   2
      Top             =   1080
      Width           =   2295
   End
End
Attribute VB_Name = "frmdept"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim op As Variant
Dim cn As ADODB.Connection
Dim rs As ADODB.Recordset
Dim ww As ADODB.Recordset
Private Sub cbofilter_Click()
    Call gridsets
    cmdlist.Visible = True
End Sub
Private Sub cbofilter_DropDown()
On Error GoTo x
        cbofilter.Clear
        Set rs = cn.Execute("select deptname from dept_details order by deptname")
        rs.MoveFirst
        cbofilter.AddItem "               <Ascending>"
        cbofilter.AddItem "               <Decending>"
        Do While Not rs.EOF()
        cbofilter.AddItem (rs(0))
        rs.MoveNext
        Loop
        cbofilter.SetFocus
x:
End Sub
Private Sub cmdadd_Click()
        If Trim(txtdept.Text) = "" Then
        MsgBox "Please Enter The Department Name", vbCritical, "Department Name Error"
        txtdept.SetFocus
        Else
            Dim rs As New ADODB.Recordset
            rs.Open "Select * from dept_details where deptid= " & txtid.Text, cn, adOpenKeyset, adLockOptimistic
                If rs.RecordCount = 0 Then
                rs.AddNew
                rs![deptid] = txtid.Text
                rs![deptname] = txtdept.Text
                rs.Update
                rs.Clone
                MsgBox "One Record Save Successfully", vbInformation, "Information"
                Unload Me
                frmdept.Show
                OsenXPLabel5.Caption = Grid.Rows - 2
                Else
                MsgBox "This Department Already Exists", vbCritical, "Invalid Department"
                OsenXPLabel5.Caption = Grid.Rows - 2
                End If
        End If
End Sub
Private Sub cmddelete_Click()
        If Trim(txtedit.Text) = "" Then
        MsgBox "Please Select The Department Name", vbCritical, "Selecting Error"
        Else
             If MsgBox("Are You Sure Delete This Record No " & txtedit.Text & " ? ", vbQuestion + vbYesNo, "Confirm To Delete") = vbYes Then
             Dim rs As New ADODB.Recordset
             rs.Open "Select * from dept_details where deptname ='" & txtedit.Text & "'", cn, adOpenKeyset, adLockOptimistic
             If rs.RecordCount <> 0 Then
                rs.Delete
                rs.Requery
                rs.Close
            MsgBox "One Record Deleted Successfully", vbInformation, "Information"
            Unload Me
            frmdept.Show
            OsenXPLabel5.Caption = Grid.Rows - 2
            Else
            MsgBox "Please Select The Department Name ", vbCritical, "Invalid Department"
            OsenXPLabel5.Caption = Grid.Rows - 2
            End If
        Else
        End If
   End If
End Sub
Private Sub cmdedit_Click()
      cmdadd.Enabled = False
      cmdupdate.Enabled = True
      On Error Resume Next
      Dim x As Double
             x = Val(txtedit.Text)
             Dim rs As New ADODB.Recordset
             rs.Open "Select * from dept_details where deptname ='" & txtedit.Text & "'", cn, adOpenKeyset, adLockOptimistic
                     If rs.RecordCount <> 0 Then
                     txtid.Text = rs![deptid]
                     txtdept.Text = rs![deptname]
                     rs.Close
                     Else
                     MsgBox "Please Select The Department Name  ", vbCritical, "Invalid Department"
                     cmdadd.Enabled = True
                     cmdupdate.Enabled = False
                     End If
End Sub
Private Sub cmdexit_Click()
    op = MsgBox("Are You Sure To Close ?", vbYesNo + vbQuestion, "Close ?")
    If op = vbYes Then
    Unload Me
    Else
    End If
End Sub
Private Sub cmdlist_Click()
    Unload Me
    frmdept.Show
End Sub
Private Sub cmdupdate_Click()
        If Trim(txtdept.Text) = "" Then
        MsgBox "Department Name Is Empty ", vbCritical, "Department Name Error"
        txtdept.SetFocus
        Else
        Dim rs As New ADODB.Recordset
        rs.Open "Select *  from dept_details where deptid=" & txtid.Text, cn, adOpenKeyset, adLockOptimistic
             If rs.RecordCount <> 0 Then
             rs![deptid] = txtid.Text
             rs![deptname] = txtdept.Text
             rs.Update
             rs.Close
             MsgBox "One Record Updated Successfully", vbInformation, "Information"
             Unload Me
             frmdept.Show
             cmdadd.Enabled = True
             cmdupdate.Enabled = False
             OsenXPLabel5.Caption = Grid.Rows - 2
             Else
             MsgBox "Already This Department Exists", vbCritical, "Invalid Department"
             OsenXPLabel5.Caption = Grid.Rows - 2
             cmdadd.Enabled = False
             cmdupdate.Enabled = True
         End If
        End If
End Sub
Private Sub Form_Load()
    Dim i As Integer
    Set cn = New ADODB.Connection
    Set ww = New ADODB.Recordset
    cn.Open "Provider=Microsoft.jet.oledb.4.0; Data Source=" & App.Path & "\Database\Data.mdb"
    ww.Open "Select * From dept_details", cn, adOpenKeyset, adLockOptimistic
    Call txtid_GotFocus
    txtdept.TabIndex = 1
    cmdlist.Visible = False
    Call frmdeptgrid
    frmdept.BackColor = RGB(184, 210, 210)
    cn.CursorLocation = adUseClient
    cbofilter.Visible = False
    Call gridload
    OsenXPLabel5.Caption = Grid.Rows - 2
    txtdept.TabIndex = 4
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
End Sub
Private Sub Grid_dblClick()
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
Private Sub Grid_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Grid.Row = 1 Then
    Grid.ToolTipText = " Click First Row of Grid To Open The Filter Options"
    End If
End Sub

Private Sub Timer1_Timer()
  frmdept.OsenXPLabel2.ForeColor = QBColor(Rnd * 15)
End Sub
Private Sub txtdept_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
Private Sub txtid_GotFocus()
    Dim rs As New ADODB.Recordset
    Dim a As String
    rs.Open "Select * from dept_details", cn, adOpenKeyset, adLockOptimistic
        If rs.RecordCount = 0 Then
        txtid.Text = 1
        rs.Close
        Else
        Dim rsrs As New ADODB.Recordset
        rsrs.Open "Select max(deptid)as exp1 from dept_details", cn, adOpenKeyset, adLockOptimistic
        txtid = rsrs![exp1] + 1
        rsrs.Close
        End If
    SendKeys "{tab}"
End Sub
Sub gridload()
    Dim i As Integer
    Dim rs As New ADODB.Recordset
    rs.Open "Select * from dept_details", cn, adOpenKeyset, adLockOptimistic
    i = 1
    If rs.BOF = False Then rs.MoveFirst
    While rs.EOF = False
        Grid.Rows = Grid.Rows + 1
        Grid.TextMatrix(i, 1) = rs![deptname]
        rs.MoveNext
        i = i + 1
    Wend
End Sub
Sub gridsets()
    Dim i As Integer
    Dim rs As New ADODB.Recordset
    rs.Open "Select * from dept_details where deptname='" & cbofilter.Text & "'", cn, adOpenKeyset, adLockOptimistic
        For i = 1 To rs.RecordCount
        Grid.TextMatrix(i, 1) = rs![deptname]
        Grid.Rows = rs.RecordCount + 1
        Grid.Row = 0
        Grid.CellBackColor = RGB(85, 194, 154)
Next i
        If Trim(cbofilter.Text) = "<Ascending>" Then
        Grid.Sort = flexSortGenericAscending
        Grid.Row = 0
        Grid.CellBackColor = RGB(85, 194, 154)
        cmdlist.Visible = True
        ElseIf Trim(cbofilter.Text) = "<Decending>" Then
        Grid.Sort = flexSortGenericDescending
        Grid.Row = 0
        Grid.CellBackColor = RGB(85, 194, 154)
        cmdlist.Visible = True
        ElseIf Trim(cbofilter.Text) = "<All>" Then
        Else
        cmdlist.Visible = False
        End If
End Sub




