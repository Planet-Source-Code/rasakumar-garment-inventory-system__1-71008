VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmpoclosure 
   BackColor       =   &H00EDDDD1&
   Caption         =   " * Purchase Order Closure *"
   ClientHeight    =   8820
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10725
   Icon            =   "frmpoclosure.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8820
   ScaleWidth      =   10725
   WindowState     =   2  'Maximized
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
      Left            =   8040
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   9600
      Width           =   1695
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
      Left            =   6240
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   9600
      Width           =   1695
   End
   Begin VB.PictureBox Checked 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   840
      Picture         =   "frmpoclosure.frx":0442
      ScaleHeight     =   255
      ScaleWidth      =   645
      TabIndex        =   3
      Top             =   120
      Visible         =   0   'False
      Width           =   675
   End
   Begin VB.PictureBox Unchecked 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      Picture         =   "frmpoclosure.frx":0D4A
      ScaleHeight     =   225
      ScaleWidth      =   705
      TabIndex        =   2
      Top             =   120
      Visible         =   0   'False
      Width           =   735
   End
   Begin MSFlexGridLib.MSFlexGrid PoMainGrid 
      Height          =   7095
      Left            =   240
      TabIndex        =   1
      ToolTipText     =   "  Note : # Indicate Place Click The First Row of Grid To  Open the Filter Options "
      Top             =   1080
      Width           =   14655
      _ExtentX        =   25850
      _ExtentY        =   12515
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
      FocusRect       =   2
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
   Begin VB.ListBox CheckedList 
      Height          =   2790
      Left            =   480
      TabIndex        =   0
      Top             =   5400
      Width           =   2775
   End
   Begin VB.TextBox txtedit 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   1080
      TabIndex        =   7
      Top             =   6120
      Width           =   1455
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Height          =   8175
      Left            =   120
      Top             =   120
      Width           =   15015
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00008000&
      BorderColor     =   &H000000C0&
      BorderWidth     =   2
      Height          =   735
      Left            =   120
      Top             =   9480
      Width           =   15015
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H000000FF&
      Height          =   975
      Index           =   0
      Left            =   120
      Top             =   8400
      Width           =   15015
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H000000FF&
      Height          =   735
      Index           =   1
      Left            =   240
      Top             =   240
      Width           =   14775
   End
   Begin VB.Label Label1 
      BackColor       =   &H00EDDDD1&
      Caption         =   "Purchase Order's  Closure"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   495
      Left            =   360
      TabIndex        =   6
      Top             =   360
      Width           =   5415
   End
End
Attribute VB_Name = "frmpoclosure"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim op As Variant
Dim cn As ADODB.Connection
Dim rs As ADODB.Recordset
Dim ww As ADODB.Recordset
Private Sub cmdupdate_Click()
    Dim rs As New ADODB.Recordset
    rs.Open "Select * from postatus_details ", cn, 1, 3
    If rs.RecordCount <> 0 Then
    For i = 1 To rs.RecordCount
    If rs.BOF = False Then rs.MoveFirst
    While rs.EOF = False
        rs![postatus] = PoMainGrid.TextMatrix(i, 7)
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
End Sub
Private Sub PoMainGrid_Click()
    Dim oldx, oldy, cell2text As String, strTextCheck As String
    oldx = PoMainGrid.Col
    oldy = PoMainGrid.Row
    If PoMainGrid.TextMatrix(1, 0) <> "" Then
        If PoMainGrid.Col = 6 And PoMainGrid.Row <> 0 Then
            If PoMainGrid.CellPicture = Checked Then
            PoMainGrid.CellForeColor = &H8000&
            PoMainGrid.TextMatrix(PoMainGrid.Row, 7) = "Open"
                Set PoMainGrid.CellPicture = Unchecked
                strTextCheck = PoMainGrid.TextMatrix(PoMainGrid.Row, 6)
    For i = 0 To CheckedList.ListCount - 1
        If CheckedList.List(i) = PoMainGrid.TextMatrix(PoMainGrid.Row, 1) Then
                CheckedList.RemoveItem i
        End If
    Next i
    Else
                Set PoMainGrid.CellPicture = Checked
                strTextCheck = PoMainGrid.TextMatrix(PoMainGrid.Row, 6)
                PoMainGrid.TextMatrix(PoMainGrid.Row, 7) = "Close"
                CheckedList.additem PoMainGrid.TextMatrix(PoMainGrid.Row, 1)
         End If
        End If
    End If
    PoMainGrid.Col = oldx
    PoMainGrid.Row = oldy
        
End Sub
Sub gridload()
    Dim i As Integer
    Dim rs As New ADODB.Recordset
    rs.Open "Select * from postatus_details", cn, adOpenKeyset, adLockOptimistic
    i = 1
    If rs.BOF = False Then rs.MoveFirst
    While rs.EOF = False
    PoMainGrid.Rows = PoMainGrid.Rows + 1
    PoMainGrid.TextMatrix(i, 0) = i
    PoMainGrid.TextMatrix(i, 1) = rs![pono]
    PoMainGrid.TextMatrix(i, 2) = rs![podate]
    PoMainGrid.TextMatrix(i, 3) = rs![supname]
    PoMainGrid.TextMatrix(i, 4) = Format(rs![netamt], "0.00")
    PoMainGrid.TextMatrix(i, 5) = rs![deptname]
    PoMainGrid.TextMatrix(i, 7) = rs![postatus]
    
    If PoMainGrid.TextMatrix(i, 7) = "Open" Then
    PoMainGrid.Col = 6: PoMainGrid.Row = i
    Set PoMainGrid.CellPicture = Unchecked.Picture
    ElseIf PoMainGrid.TextMatrix(i, 7) = "Close" Then
    PoMainGrid.Col = 6: PoMainGrid.Row = i
    Set PoMainGrid.CellPicture = Checked.Picture
    End If
    
    rs.MoveNext
    i = i + 1
    Wend
    PoMainGrid.Rows = rs.RecordCount + 1
End Sub
Private Sub Form_Load()

    Dim i As Integer
    Set cn = New ADODB.Connection
    Set ww = New ADODB.Recordset
    cn.Open "Provider=Microsoft.jet.oledb.4.0; Data Source=" & App.Path & "\Database\Data.mdb"
    ww.Open "Select * From postatus_details", cn, adOpenKeyset, adLockOptimistic
    cn.CursorLocation = adUseClient
    
    Call poclosuregridloads
    Call gridload
    
'    For i = 1 To PoMainGrid.Rows - 1
'    PoMainGrid.Col = 7: PoMainGrid.Row = i
'    PoMainGrid.CellForeColor = &H4000&
'    Next i
End Sub
