VERSION 5.00
Object = "{B69D5E45-990C-4D4D-906E-FF041974C40B}#1.0#0"; "osenxpsuite2005.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmpo 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   " * Purchase Order *"
   ClientHeight    =   8505
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9555
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmpo.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8505
   ScaleWidth      =   9555
   WindowState     =   2  'Maximized
   Begin VB.ListBox List1 
      Height          =   780
      Left            =   720
      TabIndex        =   38
      Top             =   2520
      Width           =   1935
   End
   Begin VB.TextBox txtit 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   600
      TabIndex        =   37
      Top             =   5640
      Width           =   1455
   End
   Begin MSFlexGridLib.MSFlexGrid PoMainGrid 
      Height          =   3135
      Left            =   0
      TabIndex        =   36
      Top             =   6240
      Width           =   15135
      _ExtentX        =   26696
      _ExtentY        =   5530
      _Version        =   393216
      Cols            =   7
      ForeColor       =   2359332
      ForeColorFixed  =   8388608
      ForeColorSel    =   4194368
      GridColorFixed  =   4194368
      FocusRect       =   2
      GridLines       =   2
      MergeCells      =   1
      AllowUserResizing=   3
   End
   Begin VB.TextBox txtponos 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   600
      TabIndex        =   35
      Top             =   5280
      Width           =   1455
   End
   Begin VB.CommandButton cmdprint 
      BackColor       =   &H00FFC0C0&
      Caption         =   "&Print Item"
      Height          =   495
      Left            =   9240
      MaskColor       =   &H00C0FFC0&
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   5400
      UseMaskColor    =   -1  'True
      Width           =   1695
   End
   Begin VB.CommandButton cmddelete 
      BackColor       =   &H00FFC0C0&
      Caption         =   "&Delete Item"
      Height          =   495
      Left            =   7560
      MaskColor       =   &H00C0FFC0&
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   5400
      UseMaskColor    =   -1  'True
      Width           =   1575
   End
   Begin VB.CommandButton cmdupdate 
      BackColor       =   &H00FFC0C0&
      Caption         =   "&Update  Item"
      Height          =   495
      Left            =   5760
      MaskColor       =   &H00C0FFC0&
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   5400
      UseMaskColor    =   -1  'True
      Width           =   1695
   End
   Begin VB.CommandButton cmdedit 
      BackColor       =   &H00FFC0C0&
      Caption         =   "&Edit Item "
      Height          =   495
      Left            =   3960
      MaskColor       =   &H00C0FFC0&
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   5400
      UseMaskColor    =   -1  'True
      Width           =   1695
   End
   Begin VB.CommandButton cmdadd 
      BackColor       =   &H00FFC0C0&
      Caption         =   "&Add Item"
      Height          =   495
      Left            =   2280
      MaskColor       =   &H00C0FFC0&
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   5400
      UseMaskColor    =   -1  'True
      Width           =   1575
   End
   Begin VB.CommandButton cmdexit 
      BackColor       =   &H00FFC0C0&
      Caption         =   "&Exit"
      Height          =   495
      Left            =   11040
      MaskColor       =   &H00C0FFC0&
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   5400
      UseMaskColor    =   -1  'True
      Width           =   1695
   End
   Begin osenxpsuite2005.OsenXPLabel txtitems 
      Height          =   255
      Left            =   1560
      TabIndex        =   28
      Top             =   4680
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   450
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   ""
      ForeColor       =   0
      BackStyle       =   0
   End
   Begin osenxpsuite2005.OsenXPLabel OsenXPLabel2 
      Height          =   300
      Left            =   240
      TabIndex        =   27
      Top             =   4680
      Width           =   1290
      _ExtentX        =   2275
      _ExtentY        =   529
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "No of Items  :"
      ForeColor       =   0
      BackStyle       =   0
   End
   Begin VB.CommandButton cmddeleterow 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Delete Item"
      Height          =   375
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   4200
      Width           =   1335
   End
   Begin VB.TextBox txtnet 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Height          =   375
      Left            =   13200
      Locked          =   -1  'True
      TabIndex        =   24
      Top             =   4680
      Width           =   1815
   End
   Begin VB.TextBox txttax 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   13560
      TabIndex        =   23
      Top             =   4200
      Width           =   1455
   End
   Begin VB.TextBox taxamt 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   10320
      Locked          =   -1  'True
      TabIndex        =   22
      Top             =   4680
      Width           =   1335
   End
   Begin VB.TextBox txttot 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   10320
      Locked          =   -1  'True
      TabIndex        =   21
      Top             =   4200
      Width           =   2295
   End
   Begin osenxpsuite2005.OsenXPTextBox txtrate 
      Height          =   285
      Left            =   6240
      TabIndex        =   20
      Top             =   1560
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   503
      Alignment       =   2
      BackColor       =   12648384
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   ""
      BorderColor     =   16777215
      NumberOnly      =   -1  'True
      BackColor       =   12648384
   End
   Begin osenxpsuite2005.OsenXPTextBox txtqty 
      Height          =   285
      Left            =   4800
      TabIndex        =   19
      Top             =   1560
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   503
      Alignment       =   2
      BackColor       =   12648384
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   ""
      BorderColor     =   16777215
      NumberOnly      =   -1  'True
      BackColor       =   12648384
   End
   Begin VB.TextBox txtremarks 
      Appearance      =   0  'Flat
      Height          =   855
      Left            =   4920
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   17
      Top             =   4200
      Width           =   3375
   End
   Begin osenxpsuite2005.OsenXPComboBox cbosupplier 
      Height          =   375
      Left            =   6720
      TabIndex        =   11
      Top             =   120
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      LBN             =   16777215
      LBS             =   10841658
      LBG1            =   16777215
      LBG2            =   14854529
      LAR             =   -1  'True
      LSGL            =   -1  'True
      LIH             =   18
      LIO             =   2
      LITL            =   2
      IMGLIST         =   ""
      BeginProperty HeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HeaderFontColor =   -2147483630
      ASURC           =   0   'False
   End
   Begin MSComCtl2.DTPicker dt 
      Height          =   375
      Left            =   3840
      TabIndex        =   9
      Top             =   120
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   20381697
      CurrentDate     =   39442
   End
   Begin osenxpsuite2005.OsenXPLabel OsenXPLabel1 
      Height          =   285
      Index           =   0
      Left            =   240
      TabIndex        =   8
      Top             =   120
      Width           =   660
      _ExtentX        =   1164
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
      Caption         =   "PO No"
      ForeColor       =   0
      BackStyle       =   0
   End
   Begin VB.TextBox txtpono 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   960
      TabIndex        =   7
      Top             =   120
      Width           =   1935
   End
   Begin osenxpsuite2005.OsenXPLabel deptlbl 
      Height          =   285
      Index           =   0
      Left            =   10560
      TabIndex        =   6
      Top             =   120
      Width           =   1590
      _ExtentX        =   2805
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
      Caption         =   "Department Name"
      ForeColor       =   0
      BackStyle       =   0
   End
   Begin osenxpsuite2005.OsenXPComboBox cbodept 
      Height          =   375
      Left            =   12120
      TabIndex        =   25
      Top             =   120
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      LBN             =   16777215
      LBS             =   10841658
      LBG1            =   16777215
      LBG2            =   14854529
      LAR             =   -1  'True
      LSGL            =   -1  'True
      LIH             =   18
      LIO             =   2
      IMGLIST         =   ""
      BeginProperty HeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HeaderFontColor =   -2147483630
      ASURC           =   0   'False
   End
   Begin VB.CommandButton cmdaddrow 
      BackColor       =   &H00FFC0C0&
      Caption         =   "ADD Item"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4200
      Width           =   1215
   End
   Begin osenxpsuite2005.OsenXPLabel OsenXPLabel1 
      Height          =   285
      Index           =   1
      Left            =   3000
      TabIndex        =   10
      Top             =   120
      Width           =   810
      _ExtentX        =   1429
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
      Caption         =   "PO Date"
      ForeColor       =   0
      BackStyle       =   0
   End
   Begin osenxpsuite2005.OsenXPLabel deptlbl 
      Height          =   285
      Index           =   1
      Left            =   5400
      TabIndex        =   12
      Top             =   120
      Width           =   1380
      _ExtentX        =   2434
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
      Caption         =   "Supplier Name "
      ForeColor       =   0
      BackStyle       =   0
   End
   Begin osenxpsuite2005.OsenXPComboBox cboitem 
      Height          =   255
      Left            =   1080
      TabIndex        =   2
      Top             =   1920
      Visible         =   0   'False
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   450
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      LBN             =   16777215
      LBS             =   10841658
      LBG1            =   16777215
      LBG2            =   14854529
      LAR             =   -1  'True
      LSGL            =   -1  'True
      LFS             =   12583104
      LIH             =   18
      LIO             =   2
      IMGLIST         =   ""
      BeginProperty HeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HeaderFontColor =   -2147483630
      ASURC           =   0   'False
   End
   Begin osenxpsuite2005.OsenXPComboBox cbocolour 
      Height          =   255
      Left            =   4560
      TabIndex        =   3
      Top             =   2040
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   450
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      LBN             =   16777215
      LBS             =   10841658
      LBG1            =   16777215
      LBG2            =   14854529
      LAR             =   -1  'True
      LSGL            =   -1  'True
      LFS             =   12583104
      LIH             =   18
      LIO             =   2
      LITL            =   2
      IMGLIST         =   ""
      BeginProperty HeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HeaderFontColor =   -2147483630
      ASURC           =   0   'False
   End
   Begin osenxpsuite2005.OsenXPComboBox cbosize 
      Height          =   255
      Left            =   7200
      TabIndex        =   4
      Top             =   2040
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   450
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      LBN             =   16777215
      LBS             =   10841658
      LBG1            =   16777215
      LBG2            =   14854529
      LAR             =   -1  'True
      LSGL            =   -1  'True
      LFS             =   12583104
      LIH             =   18
      LIO             =   2
      LITL            =   2
      IMGLIST         =   ""
      BeginProperty HeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HeaderFontColor =   -2147483630
      ASURC           =   0   'False
   End
   Begin osenxpsuite2005.OsenXPComboBox txtuom 
      Height          =   255
      Left            =   9240
      TabIndex        =   5
      Top             =   1920
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   450
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      LBN             =   16777215
      LBS             =   10841658
      LBG1            =   16777215
      LBG2            =   14854529
      LAR             =   -1  'True
      LSGL            =   -1  'True
      LFS             =   16711935
      LIH             =   18
      LIO             =   2
      LITL            =   2
      IMGLIST         =   ""
      BeginProperty HeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HeaderFontColor =   -2147483630
      ASURC           =   0   'False
   End
   Begin MSFlexGridLib.MSFlexGrid PoGrid 
      Height          =   3375
      Left            =   -120
      TabIndex        =   0
      Top             =   600
      Width           =   15255
      _ExtentX        =   26908
      _ExtentY        =   5953
      _Version        =   393216
      Cols            =   8
      ForeColorFixed  =   8388608
      FocusRect       =   2
      GridLinesFixed  =   1
      MergeCells      =   1
      FormatString    =   "  "
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
   Begin osenxpsuite2005.OsenXPLabel deptlbl 
      Height          =   285
      Index           =   2
      Left            =   8640
      TabIndex        =   13
      Top             =   4200
      Width           =   1665
      _ExtentX        =   2937
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
      Caption         =   "Total Amount (Rs.)"
      ForeColor       =   0
      BackStyle       =   0
   End
   Begin osenxpsuite2005.OsenXPLabel deptlbl 
      Height          =   285
      Index           =   3
      Left            =   12720
      TabIndex        =   14
      Top             =   4200
      Width           =   855
      _ExtentX        =   1508
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
      Caption         =   "Tax ( %)"
      ForeColor       =   0
      BackStyle       =   0
   End
   Begin osenxpsuite2005.OsenXPLabel deptlbl 
      Height          =   285
      Index           =   4
      Left            =   8760
      TabIndex        =   15
      Top             =   4680
      Width           =   1545
      _ExtentX        =   2725
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
      Caption         =   "Tax Amount (Rs.)"
      ForeColor       =   0
      BackStyle       =   0
   End
   Begin osenxpsuite2005.OsenXPLabel deptlbl 
      Height          =   285
      Index           =   5
      Left            =   11640
      TabIndex        =   16
      Top             =   4680
      Width           =   1530
      _ExtentX        =   2699
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
      Caption         =   "Net Amount (Rs.)"
      ForeColor       =   0
      BackStyle       =   0
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00400000&
      BorderWidth     =   2
      Height          =   735
      Left            =   2160
      Top             =   5280
      Width           =   10695
   End
   Begin VB.Shape Shape2 
      BorderWidth     =   2
      Height          =   1095
      Left            =   120
      Top             =   4080
      Width           =   8295
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Remarks"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   3600
      TabIndex        =   18
      Top             =   4200
      Width           =   1335
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   2
      Height          =   1095
      Left            =   8520
      Top             =   4080
      Width           =   6615
   End
End
Attribute VB_Name = "frmpo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim op As Variant

Dim cn As ADODB.Connection
Dim rs As ADODB.Recordset

Dim ww As ADODB.Recordset
Dim donotchange As Integer
Dim txttotal As Double

Dim i As Integer

Public pono As Integer
Public podate As String
Public supplier As String
Public dept As String
Public itemname As String
Public colour As String
Public sizes As String
Public qty As String
Public rate As Double
Public uom As String
Public tax As Double
'Public taxamt As Double
Public totamt As Double
Public netamt As Double
Public items As Integer

Private Sub cbocolour_Click()
     If donotchange Then Exit Sub
        If PoGrid.Col = 2 Then
            Me.PoGrid.Text = Me.cbocolour.Text
            Call cbocolour_gotfocus
            PoGrid.Col = 3
            PoGrid_Click
        End If
        
End Sub

Private Sub cbocolour_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call cbocolour_Click
End If
End Sub

Private Sub cbodept_Click()
Call cbodept_gotfocus
End Sub
Private Sub cboitem_click()
    If PoGrid.Col = 1 Then
        Me.PoGrid.Text = Me.cboitem.Text
         Call cboitem_gotfocus
            PoGrid.Col = 2
     Call PoGrid_Click
 End If
End Sub

Private Sub cbocolour_gotfocus()
On Error GoTo X
    cbocolour.Clear
        Set rs = cn.Execute("select colour from colour_details order by colour")
        
        rs.MoveFirst
        
        Do While Not rs.EOF()
        cbocolour.AddItem (rs(0))
       

        rs.MoveNext
        Loop
        cbocolourt.SetFocus
X:

End Sub

Private Sub cbodept_gotfocus()
On Error GoTo X
    cbodept.Clear
        Set rs = cn.Execute("select deptname from dept_details order by deptname")
        
        rs.MoveFirst
        
        Do While Not rs.EOF()
        cbodept.AddItem (rs(0))
       

        rs.MoveNext
        Loop
        cbodept.SetFocus
X:
End Sub
Private Sub cboitem_gotfocus()
On Error GoTo X
    cboitem.Clear
        Set rs = cn.Execute("select itemname from item_details where deptname ='" & cbodept.Text & "'")
        
        rs.MoveFirst
        
        Do While Not rs.EOF()
        cboitem.AddItem (rs(0))
       

        rs.MoveNext
        Loop
        cboitem.SetFocus
X:
End Sub


Private Sub cboitem_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Call cboitem_click
End If
End Sub

Private Sub cboitem_OnEnter()
 If PoGrid.Col = 1 Then
 Me.PoGrid.Text = Me.cboitem.Text
 End If
 Call cboitem_gotfocus
 Call txtuom_gotfocus
End Sub

Private Sub cbosize_Click()
    If donotchange Then Exit Sub
        If PoGrid.Col = 3 Then
         Me.PoGrid.Text = Me.cbosize.Text
         Call cbosize_gotfocus
         PoGrid.Col = 4
         Call PoGrid_Click
         
         End If
     
End Sub

Private Sub cbosize_gotfocus()
On Error GoTo X
    cbosize.Clear
        Set rs = cn.Execute("select sizes from size_details order by sizes")
        
        rs.MoveFirst
        
        Do While Not rs.EOF()
        cbosize.AddItem (rs(0))
       

        rs.MoveNext
        Loop
        cbosize.SetFocus
X:
End Sub

Private Sub cbosize_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call cbosize_Click
End If
End Sub

Private Sub cbosupplier_gotfocus()
On Error GoTo X
    cbosupplier.Clear
        Set rs = cn.Execute("select supname from sup_details order by supname")
        
        rs.MoveFirst
        
        Do While Not rs.EOF()
        cbosupplier.AddItem (rs(0))
       

        rs.MoveNext
        Loop
        cbosupplier.SetFocus
X:

End Sub

Private Sub cmdadd_Click()

Dim i As Integer
If Trim(cbosupplier.Text) = "" Then
    MsgBox "Please Enter The Supplier Name", vbCritical, "Supplier Error"
    cbosupplier.SetFocus
ElseIf Trim(cbodept.Text) = "" Then
    MsgBox "Please Enter the department Name", vbCritical, "Department Error"
    cbodept.SetFocus
ElseIf PoGrid.Text = "" Then
    MsgBox "Please Enter All data", vbCritical, "Error"
    PoGrid.SetFocus
    PoGrid.Col = 1
Else
    Dim rs As New ADODB.Recordset
    
    rs.Open "Select * from po_details where pono= " & txtpono.Text, cn, adOpenKeyset, adLockOptimistic
    
    If rs.RecordCount = 0 Then
    
    For i = 1 To PoGrid.Rows - 1
        rs.AddNew
        rs![pono] = txtpono.Text
        rs![podate] = dt.Value
        rs![supplier] = cbosupplier.Text
        rs![dept] = cbodept.Text
        rs![itemname] = PoGrid.TextMatrix(i, 1)
        rs![colour] = PoGrid.TextMatrix(i, 2)
        rs![Size] = PoGrid.TextMatrix(i, 3)
        rs![qty] = PoGrid.TextMatrix(i, 4)
        rs![uom] = PoGrid.TextMatrix(i, 5)
        rs![rates] = PoGrid.TextMatrix(i, 6)
        rs![totamts] = PoGrid.TextMatrix(i, 7)
        rs![totamt] = txttot.Text
        rs![tax] = txttax.Text
        rs![taxamt] = taxamt.Text
        rs![netamt] = txtnet.Text
        rs![remarks] = txtremarks.Text
        rs![items] = txtitems.Caption
        Next i
            rs.Update
            rs.Clone
            MsgBox "One Record Save Successfully", vbInformation, "Information"
            Unload Me
            frmpo.Show
        Else
            MsgBox "This Company Already Exists", vbCritical, "Invalid Error"
        End If
End If
End Sub
Private Sub cmdaddrow_Click()

If PoGrid.Row >= 10 Then
MsgBox "Only 10 Item Allowed ", vbCritical, "Exceed Row "
Else


 PoGrid.Rows = PoGrid.Rows + 1
  PoGrid.Row = PoGrid.Rows - 1
  txtitems.Caption = PoGrid.Rows - 1
   
If PoGrid.Row > 0 Then
    
   
    Call visibles

  

End If

End If
End Sub

Private Sub cmddeleterow_Click()

op = MsgBox("Are You to Delete ?", vbYesNo + vbQuestion, "Delete Row")
If op = vbYes Then
PoGrid.Rows = PoGrid.Rows - 1
txtitems.Caption = PoGrid.Rows - 1
End If
Call cals
Call cals1
Call visibles

End Sub

Private Sub cmdedit_Click()

cmdaddrow.Enabled = False

On Error Resume Next

    Dim X As String
    
     
     
             X = Val(txtponos.Text)
             
             Dim rs As New ADODB.Recordset
             
             rs.Open "Select * from po_details where pono =" & txtponos.Text, cn, adOpenKeyset, adLockOptimistic
             
             If rs.RecordCount <> 0 Then
  
            
               txtpono.Text = rs![pono]
               dt.Value = rs![podate]
               cbosupplier.Text = rs![supplier]
               cbodept.Text = rs![dept]
                     
               PoGrid.TextMatrix(i, 1) = rs![itemname]
               PoGrid.TextMatrix(i, 2) = rs![colour]
               PoGrid.TextMatrix(i, 3) = rs![Size]
               PoGrid.TextMatrix(i, 4) = rs![qty]
               PoGrid.TextMatrix(i, 5) = rs![uom]
               PoGrid.TextMatrix(i, 6) = rs![rates]
               PoGrid.TextMatrix(i, 7) = rs![totamts]
                     
                txttot.Text = rs![totamt]
                txttax.Text = rs![tax]
                taxamt.Text = rs![taxamt]
                txtnet = rs![netamt]
                txtremarks.Text = rs![remarks]
                txtitems.Caption = rs![items]
               
                
     

                     
            rs.Close
               Else
            MsgBox "Please Select The Item Name  ", vbCritical, "Invalid"
                
        End If
End Sub

Private Sub Form_Load()
Dim i As Integer

Set cn = New ADODB.Connection
Set ww = New ADODB.Recordset

cn.Open "Provider=Microsoft.jet.oledb.4.0; Data Source=" & App.Path & "\Database\Data.mdb"
ww.Open "Select * From po_details", cn, adOpenKeyset, adLockOptimistic



txtitems.Caption = PoGrid.Rows - 1

Call pogridalign
Call pogridmainalign
Call gridload
Call listload


Call cboitem_gotfocus
Call cbodept_gotfocus
Call txtuom_Click
Call cbosupplier_gotfocus
Call txtpono_gotfocus

frmpo.BackColor = RGB(184, 210, 210)



cn.CursorLocation = adUseClient



End Sub







Private Sub List1_GotFocus()
Call listload
End Sub

Private Sub PoGrid_Click()

If cboitem.Visible = True Then cboitem.Visible = False
If cbocolour.Visible = True Then cbocolour.Visible = False
If cbosize.Visible = True Then cbosize.Visible = False
If txtuom.Visible = True Then txtuom.Visible = False
If txtqty.Visible = True Then txtqty.Visible = False
If txtrate.Visible = True Then txtrate.Visible = False


If PoGrid.Col = 1 Then
txtitems.Caption = PoGrid.Rows - 1
Me.cboitem.Visible = True
CurrentRow = Me.PoGrid.Row
Me.cboitem.Width = Me.PoGrid.CellWidth - 10
Me.cboitem.Left = Me.PoGrid.CellLeft + Me.PoGrid.Left
Me.cboitem.Top = Me.PoGrid.CellTop + Me.PoGrid.Top

donotchange = True
Me.cboitem.Text = Me.PoGrid.Text
donotchange = False
Me.cboitem.SetFocus


ElseIf PoGrid.Col = 2 Then

Me.cbocolour.Visible = True
CurrentRow = Me.PoGrid.Row
Me.cbocolour.Width = Me.PoGrid.CellWidth - 10
Me.cbocolour.Left = Me.PoGrid.CellLeft + Me.PoGrid.Left
Me.cbocolour.Top = Me.PoGrid.CellTop + Me.PoGrid.Top

donotchange = True
Me.cbocolour.Text = Me.PoGrid.Text
donotchange = False
Me.cbocolour.SetFocus

ElseIf PoGrid.Col = 3 Then

Me.cbosize.Visible = True
CurrentRow = Me.PoGrid.Row
Me.cbosize.Width = Me.PoGrid.CellWidth - 10
Me.cbosize.Left = Me.PoGrid.CellLeft + Me.PoGrid.Left
Me.cbosize.Top = Me.PoGrid.CellTop + Me.PoGrid.Top

donotchange = True
Me.cbosize.Text = Me.PoGrid.Text
donotchange = False
Me.cbosize.SetFocus

ElseIf PoGrid.Col = 4 Then

Me.txtqty.Visible = True
CurrentRow = Me.PoGrid.Row
Me.txtqty.Width = Me.PoGrid.CellWidth - 10
Me.txtqty.Left = Me.PoGrid.CellLeft + Me.PoGrid.Left
Me.txtqty.Top = Me.PoGrid.CellTop + Me.PoGrid.Top

donotchange = True
Me.txtqty.Text = Me.PoGrid.Text
donotchange = False
Me.txtqty.SetFocus

ElseIf PoGrid.Col = 5 Or PoGrid.Col = 1 Then
Call txtuom_gotfocus
Me.txtuom.Visible = True
CurrentRow = Me.PoGrid.Row
Me.txtuom.Width = Me.PoGrid.CellWidth - 10
Me.txtuom.Left = Me.PoGrid.CellLeft + Me.PoGrid.Left
Me.txtuom.Top = Me.PoGrid.CellTop + Me.PoGrid.Top

donotchange = True
 Me.txtuom.Text = Me.PoGrid.Text
donotchange = False
Me.txtuom.SetFocus

ElseIf PoGrid.Col = 6 Then

Me.txtrate.Visible = True
CurrentRow = Me.PoGrid.Row
Me.txtrate.Width = Me.PoGrid.CellWidth - 10
Me.txtrate.Left = Me.PoGrid.CellLeft + Me.PoGrid.Left
Me.txtrate.Top = Me.PoGrid.CellTop + Me.PoGrid.Top

donotchange = True
Me.txtrate.Text = Me.PoGrid.Text
donotchange = False
Me.txtrate.SetFocus



ElseIf PoGrid.Col = 8 And PoGrid.Rows > 1 Then

op = MsgBox("Are You to Delete ?", vbYesNo + vbQuestion, "Delete Row")
If op = vbYes Then
PoGrid.Rows = PoGrid.Rows - 1
txtitems.Caption = PoGrid.Rows - 1
    
    Call cals
    Call cals1
    Call visibles
Else
End If
End If

End Sub



Private Sub PoMainGrid_Click()
If PoMainGrid.Col = 1 Then
        txtponos.Text = PoMainGrid.TextMatrix(PoMainGrid.Row, 1)
        txtit.Text = PoMainGrid.TextMatrix(PoMainGrid.Row, 6)
        Call listload
End If
End Sub

Private Sub taxamt_change()
    taxamt.Text = Format(taxamt.Text, "0.00")
End Sub
Private Sub txtnet_Change()
    txtnet.Text = Format(Me.txtnet.Text, "0.00")
End Sub

Private Sub txtpono_gotfocus()

Dim rs As New ADODB.Recordset
Dim exp1 As Variant
rs.Open "Select pono from po_details", cn, adOpenKeyset, adLockOptimistic

    If rs.RecordCount = 0 Then
    
    txtpono.Text = 1

    rs.Close
    Else
    
    Dim rsrs As New ADODB.Recordset
    rsrs.Open "Select max(pono) as exp1 from po_details", cn, adOpenKeyset, adLockOptimistic
    txtpono = rsrs![exp1] + 1
    
    rsrs.Close
    End If
    
    
    
    
    'j = Len(ostr)
    'dstr = ""
    'k = 1
    'For i = 1 To Len(ostr)
    'temp = Right(ostr, j)
    'If Left(temp, 1) <> "/" Then
    '    dstr = dstr + Left(temp, 1)
    'Else
'        part(k) = dstr
 '       k = k + 1
  '      dstr = ""
''    End If
 '   j = j - 1
'Next i
'part(3) = dstr

'If Val(Format(Date, "yy")) > Val(part(2)) Then
 '   LoadNumber = part(1) & "/" & Format(Date, "yy") & "/1"
'Else
 '   LoadNumber = part(1) & "/" & Format(Date, "yy") & "/" & (Val(part(3)) + 1)
'End If



    'If rs.RecordCount > 0 Then
    'If rs![pono] <> "0" Then
      ' txtpono.Text = Others.loadnumber(rs![pono])
           ' Else
        'txtpono.Text = "APO/" & Format(Date, "yy") & "/1"
    'End If
    'Else
    'txtpono.Text = "APO/" & Format(Date, "yy") & "/1"
    'End If


    

SendKeys "{tab}"

End Sub



Private Sub txtqty_Change()
If donotchange Then Exit Sub
 If PoGrid.Col = 4 Then
 Me.PoGrid.Text = Format(Me.txtqty.Text, "0.000")
 Call calculations
 Call cals
 Call cals1
 
 End If


End Sub

Private Sub txtqty_Click()
If donotchange Then Exit Sub
 If PoGrid.Col = 4 Then
 Me.PoGrid.Text = Format(Me.txtqty.Text, "0.000")
 End If
End Sub

Private Sub txtqty_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    PoGrid.Col = 5
    PoGrid_Click
    Call txtqty_Change
End If
End Sub

Private Sub txtrate_Change()
If donotchange Then Exit Sub
 If PoGrid.Col = 6 Then
 Me.PoGrid.Text = Format(Me.txtrate.Text, "0.00")
 Call calculations
 Call cals
 Call cals1
 
 
 End If
 
End Sub

Private Sub txtrate_Click()
If donotchange Then Exit Sub
 If PoGrid.Col = 6 Then
 Me.PoGrid.Text = Format(Me.txtrate.Text, "0.00")
 End If
End Sub
Private Sub txtrate_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    PoGrid.Col = 7
    PoGrid_Click
    op = MsgBox("Add One More Item", vbYesNo + vbQuestion, "Add Item")
            If op = vbYes Then
                 Call cmdaddrow_Click
            Else
            PoGrid.Col = 7
            End If
             
 End If
Call cals
End Sub
Private Sub txttax_Change()
If Trim(txttax.Text) = "" Then
    txttax.Text = ""
    Call cals1
    ElseIf Trim(txttax.Text) > 100 Then
    MsgBox "Invalid Tax Percenatage ", vbCritical, "Tax Error"
    txttax.Text = ""
    txttax.SetFocus
    Call cals1
    Else
    Call cals1
End If
End Sub
Private Sub txttax_GotFocus()
Call cals1
End Sub

Private Sub txttot_GotFocus()
Call cals
End Sub

Private Sub txtuom_Click()
If donotchange Then Exit Sub
 If PoGrid.Col = 5 Then
 Me.PoGrid.Text = Me.txtuom.Text
 Call txtuom_gotfocus
 PoGrid.Col = 6
 PoGrid_Click
 
 End If
 
End Sub

Private Sub txtuom_gotfocus()
On Error GoTo X
    txtuom.Clear
        Set rs = cn.Execute("select uom from uom_details order by uom")
        
        rs.MoveFirst
        
        Do While Not rs.EOF()
        txtuom.AddItem (rs(0))
       

        rs.MoveNext
        Loop
        txtuom.SetFocus
X:

End Sub

Sub calculations()
For i = 1 To PoGrid.Rows - 1
        
       PoGrid.TextMatrix(i, 7) = Format((Val(PoGrid.TextMatrix(i, 4)) * Val(PoGrid.TextMatrix(i, 6))), "0.00")
       
    Next i
End Sub
Sub cals()
txttot = 0
For i = 0 To PoGrid.Rows - 1
txttot.Text = Format(Val(txttot.Text) + Val(PoGrid.TextMatrix(i, 7)), "0.00")
Next i

End Sub
Sub cals1()
taxamt.Text = (Val(txttot.Text) * Val(txttax.Text) / 100)

txtnet.Text = Val(taxamt.Text) + Val(txttot.Text)
End Sub

Private Sub txtuom_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call txtuom_Click
End If
End Sub

Sub gridload()

Dim i As Integer
Dim rs As New ADODB.Recordset

    rs.Open "Select * from pomain_details", cn, adOpenKeyset, adLockOptimistic
    i = 1

        If rs.BOF = False Then rs.MoveFirst
            While rs.EOF = False
            PoMainGrid.Rows = PoMainGrid.Rows + 1
            PoMainGrid.TextMatrix(i, 1) = rs![pono]
            PoMainGrid.TextMatrix(i, 2) = rs![podate]
            PoMainGrid.TextMatrix(i, 3) = rs![supplier]
            PoMainGrid.TextMatrix(i, 4) = rs![netamt]
            PoMainGrid.TextMatrix(i, 5) = rs![dept]
            PoMainGrid.TextMatrix(i, 6) = rs![items]
    
    rs.MoveNext
    i = i + 1
Wend

End Sub
Sub listload()

On Error GoTo X
    List1.Clear
        Set rs = cn.Execute("select id from po_details where pono= " & txtponos.Text)
        
        rs.MoveFirst
        
        Do While Not rs.EOF()
        List1.AddItem (rs(0))
       

        rs.MoveNext
        Loop
        List1.SetFocus
X:


End Sub
