VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmdeliveryopenedreport 
   BackColor       =   &H00EDDDD1&
   Caption         =   " * Delivery Challan * ( Open )"
   ClientHeight    =   9240
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10410
   Icon            =   "frmdeliveryopened.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9240
   ScaleWidth      =   10410
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdprint 
      BackColor       =   &H00FFC0C0&
      Caption         =   "&Print"
      Height          =   495
      Left            =   240
      MaskColor       =   &H00C0FFC0&
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   " To use Print The  DC"
      Top             =   120
      UseMaskColor    =   -1  'True
      Width           =   855
   End
   Begin MSComDlg.CommonDialog cdg 
      Left            =   4920
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin RichTextLib.RichTextBox rtf 
      Height          =   9615
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   15135
      _ExtentX        =   26696
      _ExtentY        =   16960
      _Version        =   393217
      ReadOnly        =   -1  'True
      ScrollBars      =   3
      DisableNoScroll =   -1  'True
      Appearance      =   0
      AutoVerbMenu    =   -1  'True
      FileName        =   "D:\Inven\Reports\PrintPOGeneral.txt"
      TextRTF         =   $"frmdeliveryopened.frx":0442
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmdeliveryopenedreport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdprint_Click()
   cdg.ShowPrinter
    rtf.SelPrint (Printer.hDC)
End Sub
Private Sub Form_Load()
  rtf.LoadFile (App.path & "\Reports\Printdeliveryopened.txt")
'  With Selprinter
'    .Clear
'    For i = 0 To Printers.Count - 1
'        .AddItem Printers(i).DeviceName
'    Next
'    For i = 0 To .ListCount - 1
'        If .List(i) = PrnName Then
'            .ListIndex = i
'            Exit For
'        End If
'    Next
'End With
End Sub


