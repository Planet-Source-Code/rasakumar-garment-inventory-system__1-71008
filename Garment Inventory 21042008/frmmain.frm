VERSION 5.00
Begin VB.MDIForm frmmain 
   BackColor       =   &H00EDDDD1&
   Caption         =   " * Purchase System *"
   ClientHeight    =   8325
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   10470
   Icon            =   "frmmain.frx":0000
   LinkTopic       =   "MDIForm1"
   Picture         =   "frmmain.frx":0442
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Menu Mas 
      Caption         =   " &Master"
      Index           =   1
      Begin VB.Menu Com 
         Caption         =   "Company Details"
         Index           =   1
         Begin VB.Menu comps 
            Caption         =   "Company "
         End
         Begin VB.Menu L9 
            Caption         =   "-"
         End
         Begin VB.Menu dept 
            Caption         =   "Department"
         End
      End
      Begin VB.Menu l1 
         Caption         =   "-"
      End
      Begin VB.Menu Sup 
         Caption         =   "&Supplier"
      End
      Begin VB.Menu L7 
         Caption         =   "-"
      End
      Begin VB.Menu itmes 
         Caption         =   "Item "
         Begin VB.Menu it 
            Caption         =   "Item"
         End
         Begin VB.Menu L4 
            Caption         =   "-"
         End
         Begin VB.Menu colou 
            Caption         =   "Colour"
         End
         Begin VB.Menu L3 
            Caption         =   "-"
         End
         Begin VB.Menu si 
            Caption         =   "Size"
         End
         Begin VB.Menu L5 
            Caption         =   "-"
         End
         Begin VB.Menu TC 
            Caption         =   "Terms And Condition"
         End
         Begin VB.Menu l2 
            Caption         =   "-"
         End
         Begin VB.Menu unit 
            Caption         =   "Unit of Measurement"
         End
      End
   End
   Begin VB.Menu pus 
      Caption         =   "   &Purchase"
      Begin VB.Menu purs 
         Caption         =   "Purchase Order"
      End
      Begin VB.Menu l25 
         Caption         =   "-"
      End
      Begin VB.Menu Good 
         Caption         =   "Goods Receipt"
         Begin VB.Menu op 
            Caption         =   "Open Receipt"
         End
         Begin VB.Menu po 
            Caption         =   "-"
         End
         Begin VB.Menu pofr 
            Caption         =   "Good Receipt ( PO Against )"
         End
      End
   End
   Begin VB.Menu Dekivery 
      Caption         =   "   &Delivery"
      Begin VB.Menu dels 
         Caption         =   "Delivery Challan"
         Begin VB.Menu Deli 
            Caption         =   "Delivery ( GRN Against )"
         End
         Begin VB.Menu L90 
            Caption         =   "-"
         End
         Begin VB.Menu Delit 
            Caption         =   "Delivery ( Open GRN Against )"
         End
      End
      Begin VB.Menu OPs 
         Caption         =   "Open Delivery"
      End
   End
   Begin VB.Menu invoi 
      Caption         =   "   &Invoice "
      Begin VB.Menu bill 
         Caption         =   "Bill Entry"
      End
      Begin VB.Menu L76 
         Caption         =   "-"
      End
      Begin VB.Menu openbill 
         Caption         =   "Open Bill Entry"
      End
   End
   Begin VB.Menu Pay 
      Caption         =   "   &Payment"
      Begin VB.Menu Pays 
         Caption         =   "Payment Entry"
      End
   End
   Begin VB.Menu Report 
      Caption         =   "   &Reports"
      Begin VB.Menu kuer 
         Caption         =   "Master"
         Begin VB.Menu Comss 
            Caption         =   "Company Details"
         End
         Begin VB.Menu er 
            Caption         =   "-"
         End
         Begin VB.Menu Supps 
            Caption         =   "Supplier Details"
         End
         Begin VB.Menu rere 
            Caption         =   "-"
         End
         Begin VB.Menu itemss 
            Caption         =   "Item Details"
         End
      End
      Begin VB.Menu ewr 
         Caption         =   "-"
      End
      Begin VB.Menu podf 
         Caption         =   "Puchase"
         Begin VB.Menu purcha 
            Caption         =   "Purchase Order Details"
         End
         Begin VB.Menu yuui 
            Caption         =   "-"
         End
         Begin VB.Menu Gofs 
            Caption         =   "Goods Receipt Details"
         End
         Begin VB.Menu popo 
            Caption         =   "-"
         End
         Begin VB.Menu dell 
            Caption         =   "Delivery Details"
         End
      End
      Begin VB.Menu dfg 
         Caption         =   "-"
      End
      Begin VB.Menu Pos 
         Caption         =   "PO Qty Vs GRN Qty"
      End
      Begin VB.Menu dsfsdf 
         Caption         =   "-"
      End
      Begin VB.Menu pur 
         Caption         =   "Purchase Order's"
      End
   End
   Begin VB.Menu tols 
      Caption         =   "   &Tools"
      Begin VB.Menu user 
         Caption         =   "User Setting"
      End
   End
   Begin VB.Menu Logs 
      Caption         =   " &Log Off"
   End
   Begin VB.Menu ex 
      Caption         =   "   &Exit"
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim opt As Variant
Private Sub bill_Click()
    frminvoicemain.Show
End Sub
Private Sub colou_Click()
    frmcolour.Show
End Sub
Private Sub comps_Click()
    frmcompany.Show
End Sub

Private Sub Comss_Click()
 comreport.Show
End Sub

Private Sub Deli_Click()
  frmdeliverymain.Show
End Sub

Private Sub Delit_Click()
    frmopendeliverymain.Show
End Sub
Private Sub dept_Click()
    frmdepts.Show
End Sub
Private Sub ex_Click()
    opt = MsgBox("Are You Sure To Confirm Close ?", vbYesNo + vbQuestion, "Confirm Close ?")
    If opt = vbYes Then
    End
    Else
    End If
End Sub
Private Sub it_Click()
    frmitem.Show
End Sub

Private Sub itemss_Click()
    itemreport.Show
End Sub

Private Sub Logs_Click()
    opt = MsgBox("Are You Sure To Log Off ?", vbYesNo + vbQuestion, "Confirm To Log Off ?")
    If opt = vbYes Then
    Unload Me
    frmlogin.Show
    Else
    End If
End Sub
Private Sub MDIForm_Unload(Cancel As Integer)
'  opt = MsgBox(" Are You Sure To Close?", vbYesNo + vbQuestion, "Confirm Close ?")
'   If opt = vbYes Then
'   Cancel = 0
'   Else
'   Cancel = 1
'   End If
End Sub
Private Sub op_Click()
    frmopengrnmain.Show
End Sub
Private Sub openbill_Click()
    frmopeninvoicemain.Show
End Sub
Private Sub OPs_Click()
    frmopensdeliverymain.Show
End Sub
Private Sub Pays_Click()
    frmpaymentmain.Show
End Sub
Private Sub pofr_Click()
    frmgrnmain.Show
End Sub
Private Sub Pos_Click()
    frmporeporting.Show
End Sub
Private Sub pur_Click()
    frmpurchasereporting.Show
End Sub
Private Sub purcha_Click()
    poreport.Show
End Sub
Private Sub purs_Click()
    frmpomain.Show
End Sub
Private Sub si_Click()
    frmsizes.Show
End Sub
Private Sub Sup_Click()
    frmsupplier.Show
End Sub
Private Sub Supps_Click()
    supreport.Show
End Sub
Private Sub TC_Click()
    frmtc.Show
End Sub
Private Sub unit_Click()
    frmuom.Show
End Sub
Private Sub user_Click()
    frmusers.Show
End Sub
