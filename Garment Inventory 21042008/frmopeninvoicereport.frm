VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmopeninvoicereport 
   BackColor       =   &H00EDDDD1&
   Caption         =   " * Open Invoice * "
   ClientHeight    =   9015
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10110
   Icon            =   "frmopeninvoicereport.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9015
   ScaleWidth      =   10110
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdprint 
      BackColor       =   &H00FFC0C0&
      Caption         =   "&Print"
      Height          =   495
      Left            =   120
      MaskColor       =   &H00C0FFC0&
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   " To use Print Colour "
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
      Width           =   15015
      _ExtentX        =   26485
      _ExtentY        =   16960
      _Version        =   393217
      ReadOnly        =   -1  'True
      ScrollBars      =   3
      DisableNoScroll =   -1  'True
      Appearance      =   0
      AutoVerbMenu    =   -1  'True
      FileName        =   "D:\Inven\Reports\PrintPOGeneral.txt"
      TextRTF         =   $"frmopeninvoicereport.frx":0442
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
Attribute VB_Name = "frmopeninvoicereport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdprint_Click()
   cdg.ShowPrinter
   
   ' rtf.SelPrint (Printer.hDC)
   
End Sub
Private Sub Form_Load()
  rtf.LoadFile (App.Path & "\Reports\Printopeninvoice.txt")
End Sub
