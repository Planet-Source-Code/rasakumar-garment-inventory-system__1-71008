VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmusers 
   BackColor       =   &H00EDDDD1&
   Caption         =   " * User Details *"
   ClientHeight    =   8610
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9570
   Icon            =   "frmusers.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8610
   ScaleWidth      =   9570
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdupdate 
      BackColor       =   &H00FFC0C0&
      Caption         =   "&Update User"
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
      Left            =   6240
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   3360
      Width           =   1455
   End
   Begin VB.TextBox txtedit 
      Height          =   375
      Left            =   360
      TabIndex        =   17
      Top             =   1320
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CheckBox optpwd 
      Appearance      =   0  'Flat
      BackColor       =   &H00EDDDD1&
      Caption         =   " Show Password"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   240
      TabIndex        =   15
      Top             =   3480
      Width           =   2415
   End
   Begin MSComCtl2.DTPicker dt1 
      Height          =   375
      Left            =   10560
      TabIndex        =   14
      Top             =   2640
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      _Version        =   393216
      Format          =   58785793
      CurrentDate     =   39552
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
      Height          =   495
      Left            =   11040
      MaskColor       =   &H00C0FFC0&
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "  Exit Window  "
      Top             =   3360
      UseMaskColor    =   -1  'True
      Width           =   1695
   End
   Begin VB.CommandButton cmdadd 
      BackColor       =   &H00FFC0C0&
      Caption         =   "&Add User"
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
      Left            =   3120
      MaskColor       =   &H00C0FFC0&
      Style           =   1  'Graphical
      TabIndex        =   12
      Tag             =   " "
      ToolTipText     =   " To Use Add PO "
      Top             =   3360
      UseMaskColor    =   -1  'True
      Width           =   1455
   End
   Begin VB.CommandButton cmdedit 
      BackColor       =   &H00FFC0C0&
      Caption         =   "&Edit User"
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
      Left            =   4680
      MaskColor       =   &H00C0FFC0&
      Style           =   1  'Graphical
      TabIndex        =   11
      Tag             =   " "
      ToolTipText     =   " To Use View  PO "
      Top             =   3360
      UseMaskColor    =   -1  'True
      Width           =   1455
   End
   Begin VB.CommandButton cmddelete 
      BackColor       =   &H00FFC0C0&
      Caption         =   "&Delete User"
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
      Left            =   7800
      MaskColor       =   &H00C0FFC0&
      Style           =   1  'Graphical
      TabIndex        =   10
      Tag             =   " "
      ToolTipText     =   " To use Delete PO  "
      Top             =   3360
      UseMaskColor    =   -1  'True
      Width           =   1455
   End
   Begin VB.CommandButton cmdprint 
      BackColor       =   &H00FFC0C0&
      Caption         =   "&Print User"
      Enabled         =   0   'False
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
      Left            =   9360
      MaskColor       =   &H00C0FFC0&
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "  To use Print PO  "
      Top             =   3360
      UseMaskColor    =   -1  'True
      Width           =   1575
   End
   Begin VB.TextBox txtuserid 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7200
      TabIndex        =   7
      Top             =   1320
      Width           =   2895
   End
   Begin VB.TextBox txtPassword2 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      IMEMode         =   3  'DISABLE
      Left            =   7200
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   2520
      Width           =   2925
   End
   Begin VB.TextBox txtPassword1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      IMEMode         =   3  'DISABLE
      Left            =   7200
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   2160
      Width           =   2925
   End
   Begin VB.TextBox txtusername 
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
      Left            =   7200
      TabIndex        =   1
      Top             =   1680
      Width           =   2895
   End
   Begin MSFlexGridLib.MSFlexGrid loginGrid 
      Height          =   6015
      Left            =   240
      TabIndex        =   6
      ToolTipText     =   "  Note : # Indicate Place Click The First Row of Grid To  Open the Filter Options "
      Top             =   4200
      Width           =   14775
      _ExtentX        =   26061
      _ExtentY        =   10610
      _Version        =   393216
      Rows            =   1
      Cols            =   5
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
   Begin VB.Label Label1 
      BackColor       =   &H00EDDDD1&
      Caption         =   "User Details"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   495
      Index           =   2
      Left            =   240
      TabIndex        =   16
      Top             =   360
      Width           =   7335
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H000000FF&
      Height          =   735
      Index           =   1
      Left            =   120
      Top             =   240
      Width           =   15015
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00008000&
      BorderColor     =   &H000000C0&
      BorderWidth     =   2
      Height          =   735
      Index           =   1
      Left            =   120
      Top             =   3240
      Width           =   15015
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H000000FF&
      Height          =   6255
      Left            =   120
      Top             =   4080
      Width           =   15015
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "User ID"
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
      Left            =   5400
      TabIndex        =   8
      Top             =   1320
      Width           =   1815
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H000000FF&
      Height          =   1815
      Index           =   0
      Left            =   5280
      Top             =   1200
      Width           =   4935
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Retype Password"
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
      Left            =   5400
      TabIndex        =   5
      Top             =   2520
      Width           =   1815
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Password"
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
      Left            =   5400
      TabIndex        =   4
      Top             =   2160
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "User Name"
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
      Index           =   0
      Left            =   5400
      TabIndex        =   3
      Top             =   1680
      Width           =   1815
   End
End
Attribute VB_Name = "frmusers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cn As New ADODB.Connection
Dim ww As New ADODB.Recordset
Dim op As Variant
Dim rs As New ADODB.Recordset
Dim path As String
Private Sub cmdadd_Click()
    If Trim(txtusername.Text) = "" Then
        MsgBox "Please Enter The User Name", vbCritical, "Error"
        txtusername.SetFocus
    ElseIf Trim(txtPassword1.Text) = "" Then
        MsgBox "Please Enter The Password", vbCritical, "Error"
        txtPassword1.SetFocus
    ElseIf txtPassword2.Text <> txtPassword1 Then
        MsgBox "Please Re-Type The Correct Password", vbCritical, "Error"
        txtPassword2.SetFocus
    Else
        rs.Open "Select * from user_details where userid=" & txtuserid.Text, cn, adOpenKeyset, adLockOptimistic
    If rs.RecordCount = 0 Then
        rs.AddNew
        rs![userid] = txtuserid.Text
        rs![UserName] = txtusername.Text
        rs![pwd1] = txtPassword1.Text
        rs![pwd2] = txtPassword2.Text
        rs![Date] = dt1.Value
        rs.Update
        rs.Close
        MsgBox "One Record Saved Successfully", vbInformation, "Information"
        Unload Me
        frmusers.Show
    Else
        MsgBox "Id Exists", vbCritical, "Invalid"
    End If
    End If
End Sub
Private Sub cmddelete_Click()
        If Trim(txtedit.Text) = "" Then
         MsgBox "Please Select The User ID", vbCritical, "Selecting Error"
         Else
         If MsgBox("Are You Sure Delete This User ID  " & txtedit.Text & " ? ", vbQuestion + vbYesNo, "Confirm To Delete") = vbYes Then
         Dim rs As New ADODB.Recordset
             rs.Open "Select * from user_details where userid =" & txtedit.Text, cn, adOpenKeyset, adLockOptimistic
             If rs.RecordCount <> 0 Then
             rs.Delete
             rs.Requery
             rs.Close
             MsgBox "One Record Deleted Successfully", vbInformation, "Information"
             Unload Me
             frmusers.Show
             Else
             MsgBox "Please Select The User ID ", vbCritical, "Invalid User ID"
             End If
         End If
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
            If Trim(txtusername.Text) = "" Then
            MsgBox "Please Enter The User Name", vbCritical, "Error"
            txtusername.SetFocus
             ElseIf Trim(txtPassword1.Text) = "" Then
            MsgBox "Please Enter The Password", vbCritical, "Error"
            txtPassword1.SetFocus
            ElseIf txtPassword2.Text <> txtPassword1 Then
            MsgBox "Please Re-Type The Correct Password", vbCritical, "Error"
            txtPassword2.SetFocus
            Else
            Dim rs As New ADODB.Recordset
               rs.Open "Select * from user_details where userid=" & txtuserid.Text, cn, adOpenKeyset, adLockOptimistic
                    If rs.RecordCount <> 0 Then
                    rs![userid] = txtuserid.Text
                    rs![UserName] = txtusername.Text
                    rs![pwd1] = txtPassword1.Text
                    rs![pwd2] = txtPassword2.Text
                    rs.Update
                    rs.Close
                    MsgBox "One Record Updated Successfully", vbInformation, "Information"
                    Unload Me
                    frmusers.Show
                    Else
                    MsgBox "Invalid Record Update", vbCritical, "Invalid"
                   End If
                  End If
End Sub
Private Sub cmdedit_Click()
    If Trim(txtedit.Text) = "" Then
        MsgBox "Please Select The User ID", vbCritical, "User ID Error"
        loginGrid.Col = 1
        loginGrid.SetFocus
    Else
    cmdupdate.Enabled = True
    Dim i As Integer
    Dim rs As New ADODB.Recordset
    rs.Open "select *from user_details where userid = " & txtedit.Text, cn, 1, 3
    i = 1
    If rs.BOF = False Then rs.MoveFirst
    While rs.EOF = False
        txtuserid.Text = rs![userid]
        txtusername.Text = rs![UserName]
        txtPassword1.Text = rs![pwd1]
        txtPassword2.Text = rs![pwd2]
        dt1.Value = rs![Date]
        rs.MoveNext
        i = i + 1
        Wend
       
        End If
End Sub

Private Sub Form_Load()
    Set cn = New ADODB.Connection
    Set ww = New ADODB.Recordset
    cn.Open "Provider=Microsoft.jet.oledb.4.0; Data Source=" & App.path & "\Database\Data.mdb"
    ww.Open "Select * from user_details", cn, adOpenKeyset, adLockOptimistic
    
    Call txtuserid_GotFocus
    dt1.Value = Now()
    Call frmusersgridload
    Call logingridloads
    For i = 1 To loginGrid.Rows - 1
        loginGrid.Col = 3
        loginGrid.Row = i
        loginGrid.CellFontName = "Wingdings"
    Next i
    cmdupdate.Enabled = False
    cmdprint.Enabled = False
    
End Sub
Private Sub loginGrid_Click()
    If loginGrid.Col = 0 Or loginGrid.Col = 1 Or loginGrid.Col = 2 Or loginGrid.Col = 3 Or loginGrid.Col = 4 Then
         txtedit.Text = loginGrid.TextMatrix(loginGrid.Row, 1)
    End If
End Sub
Private Sub optpwd_Click()
    If optpwd.Value = 1 Then
    For i = 1 To loginGrid.Rows - 1
        loginGrid.Col = 3
        loginGrid.Row = i
        loginGrid.CellFontName = "Verdana"
    Next i
    ElseIf optpwd.Value = 0 Then
    For i = 1 To loginGrid.Rows - 1
        loginGrid.Col = 3
        loginGrid.Row = i
        loginGrid.CellFontName = "Wingdings"
    Next i
    End If
End Sub

Private Sub txtPassword1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    txtPassword2.SetFocus
    End If
End Sub
Private Sub txtPassword2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    cmdadd.SetFocus
    End If
End Sub
Private Sub txtuserid_GotFocus()
    Dim rs As New ADODB.Recordset
    rs.Open "Select * from user_details", cn, adOpenKeyset, adLockOptimistic
    If rs.RecordCount = 0 Then
    txtuserid.Text = 1
    rs.Close
    Else
    Dim rsrs As New ADODB.Recordset
    rsrs.Open "Select max(userid)as exp1 from user_details", cn, adOpenKeyset, adLockOptimistic
    txtuserid = rsrs![exp1] + 1
    rsrs.Close
    End If
SendKeys "{tab}"
End Sub
Private Sub txtusername_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
    txtPassword1.SetFocus
    End If
End Sub
Private Function logingridloads()
    Dim i As Integer
    Dim rs As New ADODB.Recordset
    rs.Open "Select * from user_details", cn, adOpenKeyset, adLockOptimistic
    If rs.RecordCount = 0 Then
    MsgBox "No Record Found", vbInformation, "Information"
    Else
    i = 1
    If rs.BOF = False Then rs.MoveFirst
    While rs.EOF = False
    loginGrid.Rows = loginGrid.Rows + 1
    loginGrid.TextMatrix(i, 0) = i
    loginGrid.TextMatrix(i, 1) = rs![userid]
    loginGrid.TextMatrix(i, 2) = rs![UserName]
    loginGrid.TextMatrix(i, 3) = rs![pwd1]
    loginGrid.TextMatrix(i, 4) = Format(rs![Date], "dd/mmm/yyyy")
    rs.MoveNext
    i = i + 1
    Wend
    loginGrid.Rows = rs.RecordCount + 1
    End If
End Function
