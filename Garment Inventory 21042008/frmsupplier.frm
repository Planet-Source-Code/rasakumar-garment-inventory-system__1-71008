VERSION 5.00
Object = "{B69D5E45-990C-4D4D-906E-FF041974C40B}#1.0#0"; "osenxpsuite2005.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmsupplier 
   BackColor       =   &H00EDDDD1&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "* Supllier Details *"
   ClientHeight    =   9645
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13140
   FillColor       =   &H80000000&
   Icon            =   "frmsupplier.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9645
   ScaleWidth      =   13140
   Begin VB.TextBox txtcity 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   2280
      TabIndex        =   38
      Top             =   2280
      Width           =   4095
   End
   Begin VB.CommandButton cmdadd 
      BackColor       =   &H00FFC0C0&
      Caption         =   "&Add Supplier"
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
      Left            =   1080
      MaskColor       =   &H00C0FFC0&
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   " To Use Add Supplier "
      Top             =   4080
      UseMaskColor    =   -1  'True
      Width           =   1575
   End
   Begin VB.CommandButton cmdedit 
      BackColor       =   &H00FFC0C0&
      Caption         =   "&Edit Supplier"
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
      Left            =   2760
      MaskColor       =   &H00C0FFC0&
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   " To Use Edit Supplier "
      Top             =   4080
      UseMaskColor    =   -1  'True
      Width           =   1695
   End
   Begin VB.CommandButton cmdupdate 
      BackColor       =   &H00FFC0C0&
      Caption         =   "&Update Supplier"
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
      Left            =   4560
      MaskColor       =   &H00C0FFC0&
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   " To Use Update Supplier "
      Top             =   4080
      UseMaskColor    =   -1  'True
      Width           =   1815
   End
   Begin VB.CommandButton cmddelete 
      BackColor       =   &H00FFC0C0&
      Caption         =   "&Delete Supplier"
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
      Left            =   6480
      MaskColor       =   &H00C0FFC0&
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   " To Use Delete Supplier "
      Top             =   4080
      UseMaskColor    =   -1  'True
      Width           =   1695
   End
   Begin VB.CommandButton cmdprint 
      BackColor       =   &H00FFC0C0&
      Caption         =   "&Print Supplier"
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
      Left            =   8280
      MaskColor       =   &H00C0FFC0&
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   " To Use Print Supplier "
      Top             =   4080
      UseMaskColor    =   -1  'True
      Width           =   1695
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
      Left            =   10080
      MaskColor       =   &H00C0FFC0&
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   " Exit Window "
      Top             =   4080
      UseMaskColor    =   -1  'True
      Width           =   1695
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
      Left            =   240
      TabIndex        =   22
      Top             =   1560
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox txtsup 
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
      Left            =   2280
      TabIndex        =   0
      ToolTipText     =   " Enter The Supplier Name "
      Top             =   720
      Width           =   4095
   End
   Begin VB.TextBox txtaddres 
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
      Height          =   975
      Left            =   2280
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Tag             =   " Enter The Supplier Address "
      Top             =   1200
      Width           =   4095
   End
   Begin VB.TextBox txtphone 
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
      Left            =   2280
      TabIndex        =   2
      ToolTipText     =   " Enter The Supplier Phone Number "
      Top             =   2640
      Width           =   4095
   End
   Begin VB.TextBox txtmobile 
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
      Left            =   2280
      TabIndex        =   3
      ToolTipText     =   " Enter The Supplier Mobile Number "
      Top             =   3000
      Width           =   4095
   End
   Begin VB.TextBox txtfax 
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
      Left            =   2280
      TabIndex        =   4
      ToolTipText     =   " Enter The Supplier Fax Number "
      Top             =   3360
      Width           =   4095
   End
   Begin VB.TextBox txtemail 
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
      Left            =   9000
      TabIndex        =   5
      ToolTipText     =   " Enter The Supplier Email - ID's "
      Top             =   720
      Width           =   3975
   End
   Begin VB.TextBox txtweb 
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
      Left            =   9000
      TabIndex        =   6
      ToolTipText     =   " Enter The Supplier Website "
      Top             =   1080
      Width           =   3975
   End
   Begin VB.TextBox txttin 
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
      Left            =   9000
      TabIndex        =   7
      ToolTipText     =   " Enter The Supplier TIN No "
      Top             =   1440
      Width           =   3975
   End
   Begin VB.TextBox txtcon 
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
      Left            =   9000
      TabIndex        =   9
      ToolTipText     =   " Enter The Supplier Contact Person "
      Top             =   2160
      Width           =   3975
   End
   Begin VB.TextBox txtconno 
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
      Left            =   9000
      TabIndex        =   10
      ToolTipText     =   " Enter The Supplier Contact Number "
      Top             =   2520
      Width           =   3975
   End
   Begin VB.TextBox txtremarks 
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
      Height          =   855
      Left            =   9000
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   11
      ToolTipText     =   " Enter The Remarks "
      Top             =   2880
      Width           =   3975
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1800
      TabIndex        =   21
      Top             =   1560
      Visible         =   0   'False
      Width           =   495
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
      ItemData        =   "frmsupplier.frx":0442
      Left            =   840
      List            =   "frmsupplier.frx":0444
      Style           =   2  'Dropdown List
      TabIndex        =   20
      ToolTipText     =   " Filter The Single Supplier Details "
      Top             =   4800
      Width           =   3495
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   240
      TabIndex        =   19
      Top             =   1800
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdlist 
      BackColor       =   &H00FFC0C0&
      Caption         =   "&List All Supplier"
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
      Left            =   4440
      MaskColor       =   &H00C0FFC0&
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   4800
      UseMaskColor    =   -1  'True
      Width           =   2055
   End
   Begin MSComCtl2.DTPicker dt 
      Height          =   375
      Left            =   9000
      TabIndex        =   8
      ToolTipText     =   " Select The Supplier TIN Date "
      Top             =   1800
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
      Format          =   58392577
      CurrentDate     =   39423
   End
   Begin osenxpsuite2005.OsenXPLabel OsenXPLabel3 
      Height          =   615
      Index           =   0
      Left            =   240
      TabIndex        =   35
      Top             =   480
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
   Begin osenxpsuite2005.OsenXPLabel OsenXPLabel3 
      Height          =   615
      Index           =   1
      Left            =   6840
      TabIndex        =   36
      Top             =   1440
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
      Height          =   4215
      Left            =   0
      TabIndex        =   37
      ToolTipText     =   "  Note : # Indicate Place Click The First Row of Grid To  Open the Filter Options "
      Top             =   5160
      Width           =   13095
      _ExtentX        =   23098
      _ExtentY        =   7435
      _Version        =   393216
      Cols            =   14
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
      Caption         =   "Supplier Master"
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
      Index           =   0
      Left            =   4800
      TabIndex        =   40
      Top             =   0
      Width           =   3015
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "City"
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
      Index           =   1
      Left            =   240
      TabIndex        =   39
      Top             =   2280
      Width           =   2055
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "                         Remarks"
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
      Height          =   855
      Left            =   6720
      TabIndex        =   34
      Top             =   2880
      Width           =   2295
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Cont.Number"
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
      Left            =   6720
      TabIndex        =   33
      Top             =   2520
      Width           =   2295
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Cont. Person"
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
      Left            =   6720
      TabIndex        =   32
      Top             =   2160
      Width           =   2295
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "TIN Date"
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
      Left            =   6720
      TabIndex        =   31
      Top             =   1800
      Width           =   2295
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "TIN No"
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
      Left            =   6720
      TabIndex        =   30
      Top             =   1440
      Width           =   2295
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Website"
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
      Left            =   6720
      TabIndex        =   29
      Top             =   1080
      Width           =   2295
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Email ID"
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
      Left            =   6720
      TabIndex        =   28
      Top             =   720
      Width           =   2295
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Fax Number"
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
      Left            =   240
      TabIndex        =   27
      Top             =   3360
      Width           =   2055
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Mobi.Number"
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
      Left            =   240
      TabIndex        =   26
      Top             =   3000
      Width           =   2055
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Ph.Number"
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
      Left            =   240
      TabIndex        =   25
      Top             =   2640
      Width           =   2055
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Address"
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
      Left            =   240
      TabIndex        =   24
      Top             =   1200
      Width           =   2055
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Supplier Name"
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
      Index           =   1
      Left            =   240
      TabIndex        =   23
      Top             =   720
      Width           =   2055
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00400000&
      BorderWidth     =   2
      Height          =   735
      Left            =   960
      Top             =   3960
      Width           =   10935
   End
   Begin VB.Shape Shape2 
      BorderWidth     =   2
      Height          =   3255
      Left            =   120
      Top             =   600
      Width           =   6375
   End
   Begin VB.Shape Shape3 
      BorderWidth     =   2
      Height          =   3255
      Left            =   6600
      Top             =   600
      Width           =   6495
   End
End
Attribute VB_Name = "frmsupplier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim op As Variant
Dim cn As ADODB.Connection
Dim rs As ADODB.Recordset
Dim ww As ADODB.Recordset
Private Sub cbofilter_Click()
    Call filter
    cmdlist.Visible = True
End Sub
Private Sub cbofilter_DropDown()
On Error GoTo X
        cbofilter.Clear
        Set rs = cn.Execute("select supname from sup_details order by supname")
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
            If Trim(txtsup.Text) = "" Then
            MsgBox "Please Enter the Supplier Name ", vbCritical, "Company Name Error"
            txtsup.SetFocus
            ElseIf Trim(txtaddres.Text) = "" Then
            MsgBox "Please Enter the Address ", vbCritical, "Address Error"
            txtaddres.SetFocus
            ElseIf Trim(txtcity.Text) = "" Then
            MsgBox "Please Enter the City & State Name ", vbCritical, "Address Error"
            txtcity.SetFocus
            ElseIf Trim(txttin.Text) = "" Then
            MsgBox "Please Enter the Supplier TIN Number", vbCritical, "TIN No Error"
            txttin.SetFocus
            Else
            Dim rs As New ADODB.Recordset
                rs.Open "Select *  from sup_details where supid=" & txtid.Text, cn, adOpenKeyset, adLockOptimistic
                    If rs.RecordCount = 0 Then
                    rs.AddNew
                    rs![supid] = txtid.Text
                    rs![supname] = txtsup.Text
                    rs![address] = txtaddres.Text
                    rs![city] = txtcity.Text
                    rs![phonenumber] = txtphone.Text
                    rs![mobilenumber] = txtmobile.Text
                    rs![faxnumber] = txtfax.Text
                    rs![email] = txtemail.Text
                    rs![website] = txtweb.Text
                    rs![tinno] = txttin.Text
                    rs![tindate] = dt.Value
                    rs![contactperson] = txtcon.Text
                    rs![contactnumber] = txtconno.Text
                    rs![remarks] = txtremarks.Text
                    rs.Update
                    rs.Close
                    MsgBox "One Record Saved Successfully", vbInformation, "Information"
                    Unload Me
                    frmsupplier.Show
                    Else
                    MsgBox "Already This Supplier Exists", vbCritical, "Invalid"
                    End If
            End If
End Sub
Private Sub cmddelete_Click()
                If Trim(Text1.Text) = "" Then
                MsgBox "Please Select The Company Name", vbCritical, "Selecting Error"
                Else
                If MsgBox("Are You Sure Delete This Record " & Text1.Text & " ? ", vbQuestion + vbYesNo, "Confirm To Delete") = vbYes Then
                Dim rs As New ADODB.Recordset
                    rs.Open "Select * from sup_details where supname ='" & Text1.Text & "'", cn, adOpenKeyset, adLockOptimistic
                    If rs.RecordCount <> 0 Then
                    rs.Delete
                    rs.Requery
                    rs.Close
                    MsgBox "One Record Deleted Successfully", vbInformation, "Information"
                    Unload Me
                    frmsupplier.Show
                    
                    Else
                    MsgBox "Please Select Supplier Name  ", vbCritical, "Invalid"
                    
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
    X = Val(Text1.Text)
    Dim rs As New ADODB.Recordset
    rs.Open "Select * from sup_details where supname ='" & Text1.Text & "'", cn, adOpenKeyset, adLockOptimistic
    If rs.RecordCount <> 0 Then
    txtid.Text = rs![supid]
    txtsup.Text = rs![supname]
    txtaddres.Text = rs![address]
    txtcity.Text = rs![city]
    txtphone.Text = rs![phonenumber]
    txtmobile.Text = rs![mobilenumber]
    txtfax.Text = rs![faxnumber]
    txtemail.Text = rs![email]
    txtweb.Text = rs![website]
    txttin.Text = rs![tinno]
    dt.Value = rs![tindate]
    txtcon.Text = rs![contactperson]
    txtconno.Text = rs![contactnumber]
    txtremarks.Text = rs![remarks]
    rs.Close
    
    Else
    MsgBox "Please Select The supplier Name  ", vbCritical, "Invalid"
    
    cmdupdate.Enabled = False
    cmdadd.Enabled = True
    End If
End Sub
Private Sub cmdexit_Click()
    op = MsgBox("Are You Sure To Close ?", vbYesNo + vbQuestion, "Confirm Close ?")
    If op = vbYes Then
    Unload Me
    Else
    End If
End Sub
Private Sub cmdlist_Click()
    Call gridload
    cbofilter.Clear
    Grid.Col = 1
    Grid.Row = 0
    Grid.CellBackColor = &HBF630F
    cmdlist.Visible = False
    cbofilter.Visible = False
End Sub
Private Sub cmdupdate_Click()
            If Trim(txtsup.Text) = "" Then
            MsgBox "Supplier Name Is Empty ", vbCritical, "Company Name Error"
            txtcomp.SetFocus
            ElseIf Trim(txtaddres.Text) = "" Then
            MsgBox "Supplier Address is Empty ", vbCritical, "Address Error"
            txtaddres.SetFocus
            ElseIf Trim(txtcity.Text) = "" Then
            MsgBox "Please Enter the City & State Name ", vbCritical, "Address Error"
            txtcity.SetFocus
            ElseIf Trim(txttin.Text) = "" Then
            MsgBox "Supplier TIN Number is Empty", vbCritical, "TIN No Error"
            txttin.SetFocus
            Else
            Dim rs As New ADODB.Recordset
                rs.Open "Select *  from sup_details where supid=" & txtid.Text, cn, adOpenKeyset, adLockOptimistic
                    If rs.RecordCount <> 0 Then
                    rs![supid] = txtid.Text
                    rs![supname] = txtsup.Text
                    rs![address] = txtaddres.Text
                    rs![city] = txtcity.Text
                    rs![phonenumber] = txtphone.Text
                    rs![mobilenumber] = txtmobile.Text
                    rs![faxnumber] = txtfax.Text
                    rs![email] = txtemail.Text
                    rs![website] = txtweb.Text
                    rs![tinno] = txttin.Text
                    rs![tindate] = dt.Value
                    rs![contactperson] = txtcon.Text
                    rs![contactnumber] = txtconno.Text
                    rs![remarks] = txtremarks.Text
                    rs.Update
                    rs.Close
                    MsgBox "One Record Updated Successfully", vbInformation, "Information"
                    Unload Me
                    frmsupplier.Show
                    
                    cmdupdate.Enabled = False
                    cmdadd.Enabled = True
                    Else
                    MsgBox "Already This Supplier  Exists", vbCritical, "Invalid"
                    
                    cmdupdate.Enabled = True
                    cmdadd.Enabled = False
                    End If
            End If
End Sub
Private Sub Form_Load()
    Dim i As Integer
    Set cn = New ADODB.Connection
    Set ww = New ADODB.Recordset
    cn.Open "Provider=Microsoft.jet.oledb.4.0; Data Source=" & App.path & "\Database\Data.mdb"
    ww.Open "Select * From sup_details", cn, adOpenKeyset, adLockOptimistic
    Call txtid_GotFocus
    cmdlist.Visible = False
    txtsup.TabIndex = 1
    cn.CursorLocation = adUseClient
    cbofilter.Visible = False
    Call frmsuppliergrid
    Call gridload
    cmdupdate.Enabled = False
End Sub
Private Sub Grid_Click()
    If Grid.Col = 1 Then
    Text1.Text = Grid
    End If
    If Grid.Col = 1 And Grid.Row = 1 Then
    cbofilter.Visible = True
    Else
    cbofilter.Visible = False
    End If
End Sub
Private Sub Grid_DblClick()
    If Grid.Col = 1 Then
    Text1.Text = Grid
    Call cmdedit_Click
    End If
    If Grid.Col = 1 And Grid.Row = 1 Then
    cbofilter.Visible = True
    Else
    cbofilter.Visible = False
    End If
End Sub

Private Sub txtemail_KeyPress(KeyAscii As Integer)
        KeyAscii = Asc(LCase(Chr(KeyAscii)))
        If KeyAscii = 13 Then
        txtweb.SetFocus
        End If
End Sub
Private Sub txtfax_KeyPress(KeyAscii As Integer)
        If KeyAscii = 8 Then
        ElseIf (KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 46 Then
        ElseIf KeyAscii = 13 Then
        txtemail.SetFocus
        Else
        KeyAscii = 0
        End If
End Sub
Private Sub txtid_GotFocus()
    Dim rs As New ADODB.Recordset
    rs.Open "Select * from sup_details", cn, adOpenKeyset, adLockOptimistic
        If rs.RecordCount = 0 Then
        txtid.Text = 1
        rs.Close
        Else
        Dim rsrs As New ADODB.Recordset
        rsrs.Open "Select max(supid)as exp1 from sup_details", cn, adOpenKeyset, adLockOptimistic
        txtid = rsrs![exp1] + 1
        rsrs.Close
        End If
    SendKeys "{tab}"
End Sub
Sub gridload()
    Dim i As Integer
    Dim rs As New ADODB.Recordset
    rs.Open "Select * from sup_details", cn, adOpenKeyset, adLockOptimistic
    i = 1
    If rs.BOF = False Then rs.MoveFirst
    While rs.EOF = False
    Grid.Rows = Grid.Rows + 1
    Grid.TextMatrix(i, 0) = i
    Grid.TextMatrix(i, 1) = rs![supname]
    Grid.TextMatrix(i, 2) = rs![address]
    Grid.TextMatrix(i, 3) = rs![city]
    Grid.TextMatrix(i, 4) = rs![phonenumber]
    Grid.TextMatrix(i, 5) = rs![mobilenumber]
    Grid.TextMatrix(i, 6) = rs![faxnumber]
    Grid.TextMatrix(i, 7) = rs![email]
    Grid.TextMatrix(i, 8) = rs![website]
    Grid.TextMatrix(i, 9) = rs![tinno]
    Grid.TextMatrix(i, 10) = rs![tindate]
    Grid.TextMatrix(i, 11) = rs![contactperson]
    Grid.TextMatrix(i, 12) = rs![contactnumber]
    Grid.TextMatrix(i, 13) = rs![remarks]
    rs.MoveNext
    i = i + 1
    Wend
    Grid.Rows = rs.RecordCount + 1
End Sub

Private Sub txtmobile_KeyPress(KeyAscii As Integer)
        If KeyAscii = 8 Then
        ElseIf (KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 46 Then
        ElseIf KeyAscii = 13 Then
        txtfax.SetFocus
        Else
        KeyAscii = 0
        End If
End Sub
Private Sub txtphone_KeyPress(KeyAscii As Integer)
        If KeyAscii = 8 Then
        ElseIf (KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 46 Then
        ElseIf KeyAscii = 13 Then
        txtmobile.SetFocus
        Else
        KeyAscii = 0
        End If
End Sub

Private Sub txtremarks_KeyPress(KeyAscii As Integer)
     If KeyAscii = 13 Then
     cmdadd.SetFocus
     End If
End Sub
Private Sub txtsup_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
Sub filter()
        Dim i As Integer
        Dim rs As New ADODB.Recordset
        rs.Open "Select * from sup_details where supname='" & cbofilter.Text & "'", cn, adOpenKeyset, adLockOptimistic
        For i = 1 To rs.RecordCount
        Grid.TextMatrix(i, 0) = i
        Grid.TextMatrix(i, 1) = rs![supname]
        Grid.TextMatrix(i, 2) = rs![address]
        Grid.TextMatrix(i, 3) = rs![city]
        Grid.TextMatrix(i, 4) = rs![phonenumber]
        Grid.TextMatrix(i, 5) = rs![mobilenumber]
        Grid.TextMatrix(i, 6) = rs![faxnumber]
        Grid.TextMatrix(i, 7) = rs![email]
        Grid.TextMatrix(i, 8) = rs![website]
        Grid.TextMatrix(i, 9) = rs![tinno]
        Grid.TextMatrix(i, 10) = rs![tindate]
        Grid.TextMatrix(i, 11) = rs![contactperson]
        Grid.TextMatrix(i, 12) = rs![contactnumber]
        Grid.TextMatrix(i, 13) = rs![remarks]
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
        Else
        cmdlist.Visible = False
        End If
End Sub
Private Sub txttin_KeyPress(KeyAscii As Integer)
        If KeyAscii = 8 Then
        ElseIf (KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 46 Then
        ElseIf KeyAscii = 13 Then
        txtfax.SetFocus
        Else
        KeyAscii = 0
        End If
End Sub
