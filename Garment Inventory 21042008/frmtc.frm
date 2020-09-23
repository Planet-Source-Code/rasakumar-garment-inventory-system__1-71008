VERSION 5.00
Object = "{B69D5E45-990C-4D4D-906E-FF041974C40B}#1.0#0"; "osenxpsuite2005.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmtc 
   BackColor       =   &H00EDDDD1&
   BorderStyle     =   1  'Fixed Single
   Caption         =   " * Terms & Condtions *"
   ClientHeight    =   9480
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11010
   Icon            =   "frmtc.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9480
   ScaleWidth      =   11010
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
      Left            =   5160
      TabIndex        =   8
      Top             =   480
      Visible         =   0   'False
      Width           =   3615
   End
   Begin VB.TextBox txtcon 
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
      Left            =   5160
      TabIndex        =   7
      ToolTipText     =   " Enter The Condition "
      Top             =   840
      Width           =   3615
   End
   Begin VB.CommandButton cmdprint 
      BackColor       =   &H00FFC0C0&
      Caption         =   "&Print TC"
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
      TabIndex        =   6
      ToolTipText     =   " To use Print Condition  "
      Top             =   1800
      UseMaskColor    =   -1  'True
      Width           =   1695
   End
   Begin VB.CommandButton cmddelete 
      BackColor       =   &H00FFC0C0&
      Caption         =   "&Delete TC"
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
      TabIndex        =   5
      ToolTipText     =   " To Use Delete Condition  "
      Top             =   1800
      UseMaskColor    =   -1  'True
      Width           =   1575
   End
   Begin VB.CommandButton cmdupdate 
      BackColor       =   &H00FFC0C0&
      Caption         =   "&Update TC"
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
      TabIndex        =   4
      ToolTipText     =   " To Use Update Condition "
      Top             =   1800
      UseMaskColor    =   -1  'True
      Width           =   1695
   End
   Begin VB.CommandButton cmdedit 
      BackColor       =   &H00FFC0C0&
      Caption         =   "&Edit TC"
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
      TabIndex        =   3
      ToolTipText     =   " To Use Edit Condition "
      Top             =   1800
      UseMaskColor    =   -1  'True
      Width           =   1695
   End
   Begin VB.CommandButton cmdadd 
      BackColor       =   &H00FFC0C0&
      Caption         =   "&Add TC"
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
      TabIndex        =   2
      ToolTipText     =   " To Use Add Condition "
      Top             =   1800
      UseMaskColor    =   -1  'True
      Width           =   1575
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
      TabIndex        =   1
      ToolTipText     =   " Exit Window"
      Top             =   1800
      UseMaskColor    =   -1  'True
      Width           =   1695
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
      Left            =   3000
      TabIndex        =   0
      Top             =   480
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   -960
      Top             =   960
   End
   Begin osenxpsuite2005.OsenXPLabel OsenXPLabel3 
      Height          =   615
      Left            =   3000
      TabIndex        =   9
      Top             =   840
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
      Height          =   6615
      Left            =   120
      TabIndex        =   10
      ToolTipText     =   "  Note : # Indicate Place Click The First Row of Grid To  Open the Filter Options "
      Top             =   2520
      Width           =   10815
      _ExtentX        =   19076
      _ExtentY        =   11668
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
      Caption         =   "TC MASTER"
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
      Left            =   4680
      TabIndex        =   12
      Top             =   240
      Width           =   2295
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Condtions"
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
      Left            =   3000
      TabIndex        =   11
      Top             =   840
      Width           =   2175
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00400000&
      BorderWidth     =   2
      Height          =   735
      Left            =   120
      Top             =   1680
      Width           =   10695
   End
End
Attribute VB_Name = "frmtc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim op As Variant
Dim cn As ADODB.Connection
Dim rs As ADODB.Recordset
Dim ww As ADODB.Recordset
Private Sub cmddelete_Click()
    If Trim(txtedit.Text) = "" Then
    MsgBox "Please Select The Conditions", vbCritical, "Selecting Error"
    Else
    If MsgBox("Are You Sure Delete This Record  " & txtedit.Text & " ? ", vbQuestion + vbYesNo, "Confirm To Delete") = vbYes Then
    Dim rs As New ADODB.Recordset
    rs.Open "Select * from condit_details where con ='" & txtedit.Text & "'", cn, adOpenKeyset, adLockOptimistic
    If rs.RecordCount <> 0 Then
    rs.Delete
    rs.Requery
    rs.Close
    MsgBox "One Record Deleted Successfully", vbInformation, "Information"
    Unload Me
    frmtc.Show
    Else
    MsgBox "Please Select The Condition ", vbCritical, "Invalid Condition"
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
             rs.Open "Select * from condit_details where con ='" & txtedit.Text & "'", cn, adOpenKeyset, adLockOptimistic
             If rs.RecordCount <> 0 Then
                     txtid.Text = rs![conid]
                     txtcon.Text = rs![con]
                rs.Close
                Else
                MsgBox "Please Select The Condition   ", vbCritical, "Invalid Condition"
                cmdadd.Enabled = True
                cmdupdate.Enabled = False
             End If
End Sub

Private Sub cmdupdate_Click()
    If Trim(txtcon.Text) = "" Then
    MsgBox "Please Enter The Conditions", vbCritical, "Colour Name Error"
    txtcon.SetFocus
    Else
    Dim rs As New ADODB.Recordset
    rs.Open "Select * from condit_details where conid= " & txtid.Text, cn, adOpenKeyset, adLockOptimistic
    If rs.RecordCount <> 0 Then
        rs![conid] = txtid.Text
        rs![con] = txtcon.Text
        rs.Update
        rs.Clone
        MsgBox "One Record Save Successfully", vbInformation, "Information"
        Unload Me
        frmtc.Show
        Else
        MsgBox "This Condtions Already Exists", vbCritical, "Invalid Condition"
        End If
    End If
End Sub
Private Sub Form_Load()
    Set cn = New ADODB.Connection
    Set ww = New ADODB.Recordset
    Set rs = New ADODB.Recordset
    cn.Open "Provider=Microsoft.jet.oledb.4.0; Data Source=" & App.path & "\Database\Data.mdb"
    ww.Open "Select * From condit_details", cn, adOpenKeyset, adLockOptimistic
    Call txtid_GotFocus
    txtcon.TabIndex = 3
    cn.CursorLocation = adUseClient
    Call frntcs
    cmdupdate.Enabled = False
    Call gridload
End Sub
Private Sub Grid_Click()
     If Grid.Col = 1 Then
       txtedit.Text = Grid
     End If
End Sub
Private Sub Grid_DblClick()
    If Grid.Col = 1 Then
        txtedit.Text = Grid
        Call cmdedit_Click
    End If
End Sub
Private Sub txtid_GotFocus()
    Dim rs As New ADODB.Recordset
    Dim exp1  As String
    rs.Open "Select * from condit_details", cn, adOpenKeyset, adLockOptimistic
        If rs.RecordCount = 0 Then
        txtid.Text = 1
        rs.Close
    Else
    Dim rsrs As New ADODB.Recordset
    rsrs.Open "Select max(conid)as exp1 from condit_details", cn, adOpenKeyset, adLockOptimistic
       txtid = rsrs![exp1] + 1
       rsrs.Close
    End If
    SendKeys "{tab}"
End Sub
Sub gridload()
    Dim i As Integer
    Dim rs As New ADODB.Recordset
    rs.Open "Select * from condit_details", cn, adOpenKeyset, adLockOptimistic
        i = 1
        If rs.BOF = False Then rs.MoveFirst
        While rs.EOF = False
        Grid.Rows = Grid.Rows + 1
        Grid.TextMatrix(i, 0) = i
        Grid.TextMatrix(i, 1) = rs![con]
        rs.MoveNext
        i = i + 1
        Wend
        Grid.Rows = rs.RecordCount + 1
End Sub
Private Sub cmdadd_Click()
    If Trim(txtcon.Text) = "" Then
    MsgBox "Please Enter The Conditions", vbCritical, "Colour Name Error"
    txtcon.SetFocus
    Else
    Dim rs As New ADODB.Recordset
    rs.Open "Select * from condit_details where conid= " & txtid.Text, cn, adOpenKeyset, adLockOptimistic
    If rs.RecordCount = 0 Then
        rs.AddNew
        rs![conid] = txtid.Text
        rs![con] = txtcon.Text
        rs.Update
        rs.Clone
        MsgBox "One Record Save Successfully", vbInformation, "Information"
        Unload Me
        frmtc.Show
        Else
        MsgBox "This Condtions Already Exists", vbCritical, "Invalid Condition"
        End If
    End If
End Sub
