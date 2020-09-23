VERSION 5.00
Object = "{B69D5E45-990C-4D4D-906E-FF041974C40B}#1.0#0"; "osenxpsuite2005.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmcolour 
   BackColor       =   &H00EDDDD1&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "* Colour Details *"
   ClientHeight    =   9315
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11070
   Icon            =   "frmcolour.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9315
   ScaleWidth      =   11070
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "List All Colours"
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
      Left            =   4080
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3360
      UseMaskColor    =   -1  'True
      Width           =   1815
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
      Left            =   2760
      TabIndex        =   0
      Top             =   840
      Visible         =   0   'False
      Width           =   2175
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
      TabIndex        =   8
      ToolTipText     =   " Exit Window"
      Top             =   2640
      UseMaskColor    =   -1  'True
      Width           =   1695
   End
   Begin VB.CommandButton cmdadd 
      BackColor       =   &H00FFC0C0&
      Caption         =   "&Add Colour"
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
      TabIndex        =   3
      ToolTipText     =   " To Use Add Colour "
      Top             =   2640
      UseMaskColor    =   -1  'True
      Width           =   1575
   End
   Begin VB.CommandButton cmdedit 
      BackColor       =   &H00FFC0C0&
      Caption         =   "&Edit Colour"
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
      TabIndex        =   4
      ToolTipText     =   " To Use Edit Colour "
      Top             =   2640
      UseMaskColor    =   -1  'True
      Width           =   1695
   End
   Begin VB.CommandButton cmdupdate 
      BackColor       =   &H00FFC0C0&
      Caption         =   "&Update Colour"
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
      TabIndex        =   5
      ToolTipText     =   " To Use Update Colour "
      Top             =   2640
      UseMaskColor    =   -1  'True
      Width           =   1695
   End
   Begin VB.CommandButton cmddelete 
      BackColor       =   &H00FFC0C0&
      Caption         =   "&Delete Colour"
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
      TabIndex        =   6
      ToolTipText     =   " To Use Delete Colour "
      Top             =   2640
      UseMaskColor    =   -1  'True
      Width           =   1575
   End
   Begin VB.CommandButton cmdprint 
      BackColor       =   &H00FFC0C0&
      Caption         =   "&Print Colour"
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
      ToolTipText     =   " To use Print Colour "
      Top             =   2640
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
      ItemData        =   "frmcolour.frx":0442
      Left            =   600
      List            =   "frmcolour.frx":0444
      Style           =   2  'Dropdown List
      TabIndex        =   9
      ToolTipText     =   " Filter The Single Colour "
      Top             =   3360
      Width           =   3375
   End
   Begin VB.TextBox txtcolour 
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
      Left            =   4920
      TabIndex        =   2
      ToolTipText     =   " Enter The Colour Name "
      Top             =   1200
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
      Left            =   4920
      TabIndex        =   1
      Top             =   840
      Visible         =   0   'False
      Width           =   3615
   End
   Begin osenxpsuite2005.OsenXPLabel OsenXPLabel3 
      Height          =   615
      Left            =   2760
      TabIndex        =   13
      Top             =   1200
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
      Height          =   5535
      Left            =   0
      TabIndex        =   11
      ToolTipText     =   "  Note : # Indicate Place Click The First Row of Grid To  Open the Filter Options "
      Top             =   3720
      Width           =   11055
      _ExtentX        =   19500
      _ExtentY        =   9763
      _Version        =   393216
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
      Caption         =   "COLOUR MASTER"
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
      Left            =   4320
      TabIndex        =   14
      Top             =   120
      Width           =   3135
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00400000&
      BorderWidth     =   2
      Height          =   735
      Left            =   120
      Top             =   2520
      Width           =   10695
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Name of Colour"
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
      Left            =   2760
      TabIndex        =   12
      Top             =   1200
      Width           =   2175
   End
End
Attribute VB_Name = "frmcolour"
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
    Command1.Visible = True
End Sub
Private Sub cbofilter_DropDown()
On Error GoTo X
    cbofilter.Clear
        Set rs = cn.Execute("select colour from colour_details order by colour")
        rs.MoveFirst
        cbofilter.additem "                     <Ascending>"
        cbofilter.additem "                     <Decending>"
        Do While Not rs.EOF()
        cbofilter.additem (rs(0))
        rs.MoveNext
        Loop
        cbofilter.SetFocus
X:
End Sub
Private Sub cmdadd_Click()
    If Trim(txtcolour.Text) = "" Then
    MsgBox "Please Enter The Name of Colour", vbCritical, "Colour Name Error"
    txtcolour.SetFocus
    Else
    Dim rs As New ADODB.Recordset
    rs.Open "Select * from colour_details where colour= '" & Trim(txtcolour.Text) & "'", cn, adOpenKeyset, adLockOptimistic
        If rs.RecordCount = 0 Then
        rs.AddNew
        rs![colourid] = txtid.Text
        rs![colour] = txtcolour.Text
        rs.Update
        rs.Clone
        MsgBox "One Record Save Successfully", vbInformation, "Information"
        Unload Me
        frmcolour.Show
        Else
        MsgBox "This Colour Already Exists", vbCritical, "Invalid Colour"
        End If
    End If
End Sub
Private Sub cmddelete_Click()
   If Trim(txtedit.Text) = "" Then
                MsgBox "Please Select The Colour", vbCritical, "Selecting Error"
                Else
                If MsgBox("Are You Sure Delete This Record  " & txtedit.Text & " ? ", vbQuestion + vbYesNo, "Confirm To Delete") = vbYes Then
                Dim rs As New ADODB.Recordset
                    rs.Open "Select * from colour_details where colour ='" & txtedit.Text & "'", cn, adOpenKeyset, adLockOptimistic
                    If rs.RecordCount <> 0 Then
                    rs.Delete
                    rs.Requery
                    rs.Close
                    MsgBox "One Record Deleted Successfully", vbInformation, "Information"
                    Unload Me
                    frmcolour.Show
                    Else
                    MsgBox "Please Slect The Colour ", vbCritical, "Invalid Colour"
                    
                    End If
                Else
                End If
            End If
End Sub
Private Sub cmdedit_Click()
    cmdupdate.Enabled = True
    cmdadd.Enabled = False
    On Error Resume Next
    Dim X As Double
             X = Val(txtedit.Text)
             Dim rs As New ADODB.Recordset
             rs.Open "Select * from colour_details where colour ='" & txtedit.Text & "'", cn, adOpenKeyset, adLockOptimistic
             If rs.RecordCount <> 0 Then
                     txtid.Text = rs![colourid]
                     txtcolour.Text = rs![colour]
                rs.Close
                Else
                MsgBox "Please Select The Colour   ", vbCritical, "Invalid Colour"
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
Private Sub cmdupdate_Click()
            If Trim(txtcolour.Text) = "" Then
            MsgBox "Colour Name Is Empty ", vbCritical, "Colour Name  Error"
            txtcolour.SetFocus
            Else
            Dim rs As New ADODB.Recordset
                rs.Open "Select *  from colour_details where colourid=" & txtid.Text, cn, adOpenKeyset, adLockOptimistic
                    If rs.RecordCount <> 0 Then
                    rs![colourid] = txtid.Text
                    rs![colour] = txtcolour.Text
                    rs.Update
                    rs.Close
                    MsgBox "One Record Updated Successfully", vbInformation, "Information"
                    Unload Me
                    frmcolour.Show
                    
                    Else
                    MsgBox "Already This ColourExists", vbCritical, "Invalid Colour"
                    
           End If
           End If
End Sub
Private Sub Command1_Click()
    cbofilter.Clear
    Call gridload
    Grid.Col = 1
    Grid.Row = 0
    Grid.CellBackColor = &HBF630F
    Command1.Visible = False
    cbofilter.Visible = False
End Sub
Private Sub Form_Load()
    Set cn = New ADODB.Connection
    Set ww = New ADODB.Recordset
    Set rs = New ADODB.Recordset
    cn.Open "Provider=Microsoft.jet.oledb.4.0; Data Source=" & App.path & "\Database\Data.mdb"
    ww.Open "Select * From colour_details", cn, adOpenKeyset, adLockOptimistic
    Call txtid_GotFocus
    txtcolour.TabIndex = 3
    cn.CursorLocation = adUseClient
    cmdupdate.Enabled = False
    cbofilter.Visible = False
    Call frmcolourload
    Call gridload
    Command1.Visible = False
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

Private Sub txtcolour_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
    cmdadd.SetFocus
    End If
End Sub
Private Sub txtid_GotFocus()
    Dim rs As New ADODB.Recordset
    Dim a As String
    rs.Open "Select * from colour_details", cn, adOpenKeyset, adLockOptimistic
        If rs.RecordCount = 0 Then
        txtid.Text = 1
        rs.Close
    Else
    Dim rsrs As New ADODB.Recordset
    rsrs.Open "Select max(colourid)as exp1 from colour_details", cn, adOpenKeyset, adLockOptimistic
       txtid = rsrs![exp1] + 1
       rsrs.Close
    End If
    SendKeys "{tab}"
End Sub
Sub gridload()
    Dim i As Integer
    Dim rs As New ADODB.Recordset
    rs.Open "Select * from colour_details", cn, adOpenKeyset, adLockOptimistic
        i = 1
        If rs.BOF = False Then rs.MoveFirst
        While rs.EOF = False
        Grid.Rows = Grid.Rows + 1
        Grid.TextMatrix(i, 0) = i
        Grid.TextMatrix(i, 1) = rs![colour]
        rs.MoveNext
        i = i + 1
        Wend
        Grid.Rows = rs.RecordCount + 1
End Sub
Sub gridsets()
    Dim i As Integer
    Dim rs As New ADODB.Recordset
    rs.Open "Select * from colour_details where colour='" & cbofilter.Text & "'", cn, adOpenKeyset, adLockOptimistic
        For i = 1 To rs.RecordCount
        Grid.TextMatrix(i, 1) = rs![colour]
        Grid.Rows = rs.RecordCount + 1
        Grid.Row = 0
        Grid.CellBackColor = RGB(85, 194, 154)
Next i
        If Trim(cbofilter.Text) = "<Ascending>" Then
        Grid.Sort = flexSortGenericAscending
        Grid.Row = 0
        Grid.CellBackColor = RGB(85, 194, 154)
        Command1.Visible = True
        ElseIf Trim(cbofilter.Text) = "<Decending>" Then
        Grid.Sort = flexSortGenericDescending
        Grid.Row = 0
        Grid.CellBackColor = RGB(85, 194, 154)
        Command1.Visible = True
        ElseIf Trim(cbofilter.Text) = "<All>" Then
        Else
        Command1.Visible = False
        End If
End Sub

