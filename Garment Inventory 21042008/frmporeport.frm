VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmporeport 
   BackColor       =   &H00EDDDD1&
   Caption         =   " * PO Printout * "
   ClientHeight    =   8820
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9540
   Icon            =   "frmporeport.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8820
   ScaleWidth      =   9540
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdprint 
      BackColor       =   &H00FFC0C0&
      Caption         =   "&Print"
      Height          =   495
      Left            =   120
      MaskColor       =   &H00C0FFC0&
      Style           =   1  'Graphical
      TabIndex        =   1
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
      TabIndex        =   0
      Top             =   720
      Width           =   15015
      _ExtentX        =   26485
      _ExtentY        =   16960
      _Version        =   393217
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   3
      DisableNoScroll =   -1  'True
      Appearance      =   0
      AutoVerbMenu    =   -1  'True
      FileName        =   "D:\Inven\Reports\PrintPOGeneral.txt"
      TextRTF         =   $"frmporeport.frx":0442
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
Attribute VB_Name = "frmporeport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdprint_Click()
   cdg.ShowPrinter
   
    rtf.SelPrint (Printer.hDC)
   
End Sub
Private Sub Form_Load()
  rtf.LoadFile (App.Path & "\Reports\PrintPOGeneral.txt")
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
''This will print the information on richtextbox to printer
'Private Sub PrintToPrinter()
'Dim intAsk As Integer
'On Error GoTo PrintError
'  If IsPrinterInstalled = False Then
'     MsgBox "There is no printer has been installed" & Chr(13) & _
'            "in your computer. Please install" & Chr(13) & _
'            "printer first!", vbExclamation, _
'            "Printer Not Install"
'     Exit Sub
'  Else
'  End If
'  If rtf.Text = "" Then
'     MsgBox "There is no information is being displayed to your screen this time!" & Chr(13) & _
'            "Please choose the category by clicking on menu Report above" & Chr(13) & _
'            "and then click on menu File->Print or button Print.", _
'            vbCritical, "No Result"
'     Exit Sub
'  End If
'
'  'Print rtfLap1.Text to printer
'  Printer.FontName = "Courier New"
'  Printer.FontSize = "9"
'  Printer.Print rtf.Text
'  Printer.EndDoc '<-- This will eject the paper till the end
'                 '    of paper
'
'  'If you don't want printer roll up the paper till the end,
'  'you can use the following code. The printer head will be stop
'  'just after printing the last line in richtexbox control
'  'Here is the code:
'  'Open Printer.Port For Input As #1
'  '   Printer.Print rtfLap1.Text
'  'Close #1
'
'  'Are you sure the report is correct? If so,
'  'clear richtexbox, if not yet, leave it...
'  If MsgBox("The information in your screen has been sent to the printer." & vbCrLf & _
'            "Are you sure you want to clear the information on your screen?", _
'            vbInformation + vbYesNo, "Print") = vbYes Then
'     rtf.Text = ""
'  End If
'  Exit Sub
'PrintError:
'    MsgBox "Error number: " & Err.Number & vbCrLf & _
'           "Description: " & Err.Description & "" & Chr(13) & _
'           "" & Chr(13) & _
'           "May be printer is still off or out of paper." & Chr(13) & _
'           "Please turn on your printer now or fill in " & Chr(13) & _
'           "the paper to printer. Then, try again.", _
'           vbCritical, "Printer Error"
'    Exit Sub
'End Sub
'Public Function IsPrinterInstalled() As Boolean
'On Error Resume Next
'Dim strDummy As String
'  strDummy = Printer.DeviceName
'  If Err.Number Then
'     IsPrinterInstalled = False
'  Else
'     IsPrinterInstalled = True
'  End If
'End Function
'
'
'
'
'
'
'
