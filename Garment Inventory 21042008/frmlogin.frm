VERSION 5.00
Begin VB.Form frmlogin 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   ClientHeight    =   8295
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10995
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmlogin.frx":0000
   ScaleHeight     =   8295
   ScaleWidth      =   10995
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer6 
      Interval        =   6000
      Left            =   360
      Top             =   3360
   End
   Begin VB.Timer Timer5 
      Interval        =   5000
      Left            =   360
      Top             =   2760
   End
   Begin VB.Timer Timer4 
      Interval        =   4000
      Left            =   360
      Top             =   2160
   End
   Begin VB.Timer Timer3 
      Interval        =   3000
      Left            =   360
      Top             =   1560
   End
   Begin VB.Timer Timer2 
      Interval        =   2000
      Left            =   360
      Top             =   960
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   360
      Top             =   480
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00FFC0C0&
      Caption         =   "&CANCEL"
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
      Left            =   9240
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   7320
      Width           =   1095
   End
   Begin VB.CommandButton cmdok 
      BackColor       =   &H00FFC0C0&
      Caption         =   "&OK"
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
      Left            =   8040
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   7320
      Width           =   1095
   End
   Begin VB.TextBox txtPassword 
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
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   7920
      PasswordChar    =   "*"
      TabIndex        =   1
      ToolTipText     =   " Enter The Correct Password "
      Top             =   6720
      Width           =   2565
   End
   Begin VB.ComboBox txtuser 
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
      Left            =   7920
      TabIndex        =   0
      ToolTipText     =   " Select The User Name "
      Top             =   6240
      Width           =   2535
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Height          =   8055
      Left            =   120
      Top             =   120
      Width           =   10815
   End
   Begin VB.Label userlbl 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   1
      Left            =   6480
      TabIndex        =   13
      Top             =   6720
      Width           =   1335
   End
   Begin VB.Label userlbl 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Left            =   6480
      TabIndex        =   12
      Top             =   6240
      Width           =   1335
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00000080&
      BackStyle       =   0  'Transparent
      Caption         =   "."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Index           =   5
      Left            =   6960
      TabIndex        =   11
      Top             =   4080
      Width           =   135
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00000080&
      BackStyle       =   0  'Transparent
      Caption         =   "."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Index           =   4
      Left            =   6720
      TabIndex        =   10
      Top             =   4080
      Width           =   135
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00000080&
      BackStyle       =   0  'Transparent
      Caption         =   "."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Index           =   3
      Left            =   6480
      TabIndex        =   9
      Top             =   4080
      Width           =   135
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00000080&
      BackStyle       =   0  'Transparent
      Caption         =   "."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Index           =   2
      Left            =   6240
      TabIndex        =   8
      Top             =   4080
      Width           =   135
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00000080&
      BackStyle       =   0  'Transparent
      Caption         =   "."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Index           =   1
      Left            =   6000
      TabIndex        =   7
      Top             =   4080
      Width           =   135
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00000080&
      BackStyle       =   0  'Transparent
      Caption         =   "."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Index           =   0
      Left            =   5760
      TabIndex        =   6
      Top             =   4080
      Width           =   135
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000080&
      BackStyle       =   0  'Transparent
      Caption         =   "Loading"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   4440
      TabIndex        =   5
      Top             =   4200
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000013&
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   10560
      MouseIcon       =   "frmlogin.frx":5CFA
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   240
      Width           =   255
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00008000&
      BorderWidth     =   2
      Height          =   1935
      Left            =   6360
      Top             =   6120
      Width           =   4215
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FFFFFF&
      Height          =   735
      Left            =   7920
      Top             =   7200
      Width           =   2535
   End
End
Attribute VB_Name = "frmlogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cn As ADODB.Connection
Dim ww As New ADODB.Recordset
Dim op As Variant
Dim rs As New ADODB.Recordset
Dim flag As Boolean
Private Sub cmdcancel_Click()
    End
End Sub
Private Sub cmdOk_Click()
    If txtuser = "" Then
    MsgBox "Please Select  your User Name", vbCritical, "Password Error"
    Exit Sub
    End If
    If txtPassword = "" Then
    MsgBox "Please type your Correct Password", vbCritical, "Password Error"
    Exit Sub
    End If
    
    Set rs = cn.Execute("Select * from user_details")
    
    If Not rs.EOF Then
    rs.MoveFirst
      Do While Not rs.EOF
      If (txtuser) = rs!UserName And (txtPassword) = rs!pwd1 Then
      flag = True
      End If
      If (txtuser) = rs!UserName And (txtPassword) <> rs!pwd2 Then
      flaguser = True
      End If
      rs.MoveNext
      Loop
    End If
    
    If flag = True Then
    Unload Me
    frmmain.Show
    Exit Sub
    End If
    
    If flaguser = True Then
    MsgBox "Password is incorrect! Please check if CapsLock is on.", vbCritical, "Error"
    txtPassword = ""
    flaguser = False
    Else
    MsgBox "The UserName that you are trying to login is not registered", vbCritical, "Error"
    txtuser = ""
    txtPassword = ""
    End If
End Sub
Private Sub Form_Load()
    Set cn = New ADODB.Connection
    Set ww = New ADODB.Recordset
    cn.Open "Provider=Microsoft.jet.oledb.4.0; Data Source=" & App.path & "\Database\Data.mdb"
    ww.Open "Select * from user_details", cn, adOpenKeyset, adLockOptimistic
    
    Label3(0).Visible = False
    Label3(1).Visible = False
    Label3(2).Visible = False
    Label3(3).Visible = False
    Label3(4).Visible = False
    Label3(5).Visible = False
    
    Shape1.Visible = False
    userlbl(0).Visible = False
    userlbl(1).Visible = False
    txtuser.Visible = False
    txtPassword.Visible = False
    Shape2.Visible = False
    cmdok.Visible = False
    cmdcancel.Visible = False
    
End Sub
Private Sub Label1_Click()
    End
End Sub
Private Sub Timer1_Timer()
    Label3(0).Visible = True
End Sub
Private Sub Timer2_Timer()
    Label3(1).Visible = True
End Sub
Private Sub Timer3_Timer()
    Label3(2).Visible = True
End Sub
Private Sub Timer4_Timer()
    Label3(3).Visible = True
End Sub
Private Sub Timer5_Timer()
    Label3(4).Visible = True
End Sub
Private Sub Timer6_Timer()
    Label3(5).Visible = True
    Shape1.Visible = True
    userlbl(0).Visible = True
    userlbl(1).Visible = True
    txtuser.Visible = True
    txtPassword.Visible = True
    Shape2.Visible = True
    cmdok.Visible = True
    cmdcancel.Visible = True
    Call notvisi
End Sub
Private Sub txtPassword_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    cmdok.SetFocus
    Else
    End If
End Sub
Private Sub txtuser_Click()
 txtPassword.SetFocus
End Sub
Private Sub txtuser_dropdown()
On Error GoTo X
    txtuser.Clear
    Set rs = cn.Execute("select username from user_details order by username")
    rs.MoveFirst
    Do While Not rs.EOF()
    txtuser.additem (rs(0))
    rs.MoveNext
    Loop
    txtuser.SetFocus
X:
End Sub
Sub notvisi()
    Label3(0).Visible = False
    Label3(1).Visible = False
    Label3(2).Visible = False
    Label3(3).Visible = False
    Label3(4).Visible = False
    Label3(5).Visible = False
    Label2.Visible = False
    
    Timer1.Enabled = False
    Timer2.Enabled = False
    Timer3.Enabled = False
    Timer4.Enabled = False
    Timer5.Enabled = False
    Timer6.Enabled = False
    Label2.Visible = False
End Sub

