Attribute VB_Name = "Grid"
Sub pogridalign()
    frmpos.PoGrid.ColWidth(0) = 600
    frmpos.PoGrid.ColAlignment(0) = 3
    frmpos.PoGrid.TextMatrix(0, 0) = " S.No "
    frmpos.PoGrid.TextMatrix(0, 1) = "  * Item Name     "
    frmpos.PoGrid.ColWidth(1) = 3000
    frmpos.PoGrid.TextMatrix(0, 2) = "  * Colour     "
    frmpos.PoGrid.ColWidth(2) = 2000
    frmpos.PoGrid.TextMatrix(0, 3) = "  * Size   "
    frmpos.PoGrid.ColWidth(3) = 2200
    frmpos.PoGrid.ColAlignment(3) = 1
    frmpos.PoGrid.TextMatrix(0, 4) = "  Quantity "
    frmpos.PoGrid.ColWidth(4) = 1600
    frmpos.PoGrid.TextMatrix(0, 5) = "     UOM "
    frmpos.PoGrid.ColWidth(5) = 1650
    frmpos.PoGrid.TextMatrix(0, 6) = "  Rate (Rs) "
    frmpos.PoGrid.ColWidth(6) = 1500
    frmpos.PoGrid.TextMatrix(0, 7) = "  Tot.Amount (Rs) "
    frmpos.PoGrid.ColWidth(7) = 2000
End Sub
Sub visibles()
    frmpos.txtrate.Visible = False
    frmpos.txtqty.Visible = False
    frmpos.cboitem.Visible = False
    frmpos.cbocolour.Visible = False
    frmpos.cbosize.Visible = False
End Sub
Sub pogridmainalign()
    frmpomain.PoMainGrid.ColWidth(0) = 1200
    frmpomain.PoMainGrid.ColAlignment(0) = 3
    frmpomain.PoMainGrid.TextMatrix(0, 0) = "  S.No   "
    
    frmpomain.PoMainGrid.TextMatrix(0, 1) = "  # PO No     "
    frmpomain.PoMainGrid.ColWidth(1) = 2000
    
    frmpomain.PoMainGrid.TextMatrix(0, 2) = "   PO Date   "
    frmpomain.PoMainGrid.ColWidth(2) = 1600
    
    frmpomain.PoMainGrid.TextMatrix(0, 3) = "  # Supplier Name   "
    frmpomain.PoMainGrid.ColWidth(3) = 3500
    
    frmpomain.PoMainGrid.TextMatrix(0, 4) = "              Net Amt (Rs)"
    frmpomain.PoMainGrid.ColWidth(4) = 2600
    
    frmpomain.PoMainGrid.TextMatrix(0, 5) = " # Department "
    frmpomain.PoMainGrid.ColWidth(5) = 2800
    
    frmpomain.PoMainGrid.ColAlignment(1) = 3
    frmpomain.PoMainGrid.ColAlignment(2) = 3
    
End Sub
Sub frmcolourload()
    frmcolour.Grid.ColWidth(0) = 600
    frmcolour.Grid.ColAlignment(0) = 3
    frmcolour.Grid.TextMatrix(0, 0) = " S.No "
    frmcolour.Grid.ColWidth(1) = 3400
    frmcolour.Grid.TextMatrix(0, 1) = "  # Colour"
    frmcolour.Height = 9800
    frmcolour.Width = 11160
    frmcolour.Top = 2
    frmcolour.Left = 2
End Sub

Sub frmcompanygrid()
    frmcompany.Height = 9800
    frmcompany.Width = 13300
    frmcompany.Top = 2
    frmcompany.Left = 2
    frmcompany.Grid.ColWidth(0) = 800
    frmcompany.Grid.ColAlignment(0) = 3
    frmcompany.Grid.TextMatrix(0, 0) = " S.No "
    frmcompany.Grid.ColWidth(1) = 3500
    frmcompany.Grid.TextMatrix(0, 1) = " # Company Name "
    frmcompany.Grid.ColWidth(2) = 5000
    frmcompany.Grid.TextMatrix(0, 2) = " Address "
    frmcompany.Grid.ColWidth(3) = 2000
    frmcompany.Grid.TextMatrix(0, 3) = " City  & State "
    frmcompany.Grid.ColWidth(4) = 2500
    frmcompany.Grid.TextMatrix(0, 4) = " Phone Number "
    frmcompany.Grid.ColWidth(5) = 2500
    frmcompany.Grid.TextMatrix(0, 5) = " Mobile  Number "
    frmcompany.Grid.ColWidth(6) = 2500
    frmcompany.Grid.TextMatrix(0, 6) = " Fax Number "
    frmcompany.Grid.ColWidth(7) = 3000
    frmcompany.Grid.TextMatrix(0, 7) = " Email - ID "
    frmcompany.Grid.ColWidth(8) = 2500
    frmcompany.Grid.TextMatrix(0, 8) = " Website "
    frmcompany.Grid.ColWidth(9) = 2500
    frmcompany.Grid.TextMatrix(0, 9) = " Website "
    frmcompany.Grid.ColWidth(10) = 1500
    frmcompany.Grid.TextMatrix(0, 10) = " TIN Date "
    frmcompany.Grid.ColWidth(11) = 3000
    frmcompany.Grid.TextMatrix(0, 11) = " Contact Person "
    frmcompany.Grid.ColWidth(12) = 2500
    frmcompany.Grid.TextMatrix(0, 12) = " Contact No "
    frmcompany.Grid.ColWidth(13) = 6000
    frmcompany.Grid.TextMatrix(0, 13) = "Remarks "
End Sub
Sub frmdeptsgrid()
    frmdepts.Grid.ColWidth(0) = 600
    frmdepts.Grid.ColAlignment(0) = 3
    frmdepts.Grid.TextMatrix(0, 0) = " S.No"
    frmdepts.Grid.ColWidth(1) = 3450
    frmdepts.Grid.TextMatrix(0, 1) = "  # Name of Department"
    frmdepts.Height = 9800
    frmdepts.Width = 11050
    frmdepts.Top = 2
    frmdepts.Left = 2
End Sub
Sub frmitemgrid()
    frmitem.Grid.ColWidth(0) = 800
    frmitem.Grid.ColAlignment(0) = 3
    frmitem.Grid.TextMatrix(0, 0) = " S.No"
    frmitem.Grid.ColWidth(1) = 3450
    frmitem.Grid.TextMatrix(0, 1) = "  # Name of Items"
    frmitem.Grid.ColWidth(2) = 2500
    frmitem.Grid.TextMatrix(0, 2) = "  UOM "
    frmitem.Grid.ColWidth(3) = 3000
    frmitem.Grid.TextMatrix(0, 3) = " # Departments "
    frmitem.Height = 9700
    frmitem.Width = 11000
    frmitem.Top = 2
    frmitem.Left = 2
End Sub
Sub frmsizegrid()
    frmsizes.Grid.ColWidth(0) = 800
    frmsizes.Grid.ColAlignment(0) = 3
    frmsizes.Grid.TextMatrix(0, 0) = " S.No"
    frmsizes.Grid.ColWidth(1) = 3450
    frmsizes.Grid.TextMatrix(0, 1) = "  # Sizes"
    frmsizes.Grid.ColAlignment(1) = 1
    frmsizes.Height = 9700
    frmsizes.Width = 11150
    frmsizes.Top = 2
    frmsizes.Left = 2
End Sub
Sub frmsuppliergrid()
    frmsupplier.Height = 9800
    frmsupplier.Width = 13300
    frmsupplier.Top = 2
    frmsupplier.Left = 2
    frmsupplier.Grid.ColWidth(0) = 800
    frmsupplier.Grid.ColAlignment(0) = 3
    frmsupplier.Grid.TextMatrix(0, 0) = "S.No "
    frmsupplier.Grid.ColWidth(1) = 3500
    frmsupplier.Grid.TextMatrix(0, 1) = " # Supplier Name "
    frmsupplier.Grid.ColWidth(2) = 5000
    frmsupplier.Grid.TextMatrix(0, 2) = " Address "
    frmsupplier.Grid.ColWidth(3) = 2500
    frmsupplier.Grid.TextMatrix(0, 3) = " City & State "
    frmsupplier.Grid.ColWidth(4) = 2000
    frmsupplier.Grid.TextMatrix(0, 4) = " Phone Number "
    frmsupplier.Grid.ColWidth(5) = 2500
    frmsupplier.Grid.TextMatrix(0, 5) = " Mobile  Number "
    frmsupplier.Grid.ColWidth(6) = 2500
    frmsupplier.Grid.TextMatrix(0, 6) = " Fax Number "
    frmsupplier.Grid.ColWidth(7) = 3000
    frmsupplier.Grid.TextMatrix(0, 7) = " Email - ID "
    frmsupplier.Grid.ColWidth(8) = 2500
    frmsupplier.Grid.TextMatrix(0, 8) = " Website "
    frmsupplier.Grid.ColWidth(9) = 2500
    frmsupplier.Grid.TextMatrix(0, 9) = " TIN No "
    frmsupplier.Grid.ColWidth(10) = 1500
    frmsupplier.Grid.TextMatrix(0, 10) = " TIN Date "
    frmsupplier.Grid.ColWidth(11) = 3000
    frmsupplier.Grid.TextMatrix(0, 11) = " Contact Person "
    frmsupplier.Grid.ColWidth(12) = 2500
    frmsupplier.Grid.TextMatrix(0, 12) = " Contact No "
    frmsupplier.Grid.ColWidth(13) = 6000
    frmsupplier.Grid.TextMatrix(0, 13) = "Remarks "
End Sub
Sub frmuomgrid()
    frmuom.Grid.ColWidth(0) = 800
    frmuom.Grid.ColAlignment(0) = 3
    frmuom.Grid.TextMatrix(0, 0) = " S.No"
    frmuom.Grid.ColWidth(1) = 3450
    frmuom.Grid.TextMatrix(0, 1) = "  # Unit of Measurements "
    frmuom.Height = 9700
    frmuom.Width = 10950
    frmuom.Top = 2
    frmuom.Left = 2
End Sub
Sub frmgrngrid()
    frmgrn.GrnGrid.ColWidth(0) = 600
    frmgrn.GrnGrid.ColAlignment(0) = 3
    frmgrn.GrnGrid.TextMatrix(0, 0) = "S.No"
    
    frmgrn.GrnGrid.TextMatrix(0, 1) = "  * Po No     "
    frmgrn.GrnGrid.ColWidth(1) = 1000
    frmgrn.GrnGrid.ColAlignment(1) = 3
    
    frmgrn.GrnGrid.TextMatrix(0, 2) = "* Po Date    "
    frmgrn.GrnGrid.ColWidth(2) = 1300
    frmgrn.GrnGrid.ColAlignment(2) = 3
    
    frmgrn.GrnGrid.TextMatrix(0, 3) = "  * Item Name    "
    frmgrn.GrnGrid.ColWidth(3) = 2500
    
    frmgrn.GrnGrid.TextMatrix(0, 4) = "  * Colour  "
    frmgrn.GrnGrid.ColWidth(4) = 1800
    
    frmgrn.GrnGrid.TextMatrix(0, 5) = " * Size "
    frmgrn.GrnGrid.ColWidth(5) = 1600
    frmgrn.GrnGrid.ColAlignment(5) = 1
     
    frmgrn.GrnGrid.TextMatrix(0, 6) = "* Po Qty "
    frmgrn.GrnGrid.ColWidth(6) = 1300
    
    frmgrn.GrnGrid.TextMatrix(0, 7) = " UOM "
    frmgrn.GrnGrid.ColWidth(7) = 1400
    
    frmgrn.GrnGrid.TextMatrix(0, 8) = "  Balance Qty "
    frmgrn.GrnGrid.ColWidth(8) = 1500
    
    frmgrn.GrnGrid.TextMatrix(0, 9) = " Receive Qty "
    frmgrn.GrnGrid.ColWidth(9) = 1500
    
    frmgrn.GrnGrid.TextMatrix(0, 10) = " ID "
    frmgrn.GrnGrid.ColWidth(10) = 0
    frmgrn.GrnGrid.ColAlignment(10) = 3
    
    frmgrn.GrnGrid.TextMatrix(0, 11) = "Total Grn "
    frmgrn.GrnGrid.ColWidth(11) = 0
End Sub
Sub frmgrnmaingrid()
    frmgrn.GrnMainGrid.ColWidth(0) = 600
    frmgrn.GrnMainGrid.ColAlignment(0) = 3
    frmgrn.GrnMainGrid.TextMatrix(0, 0) = "S.No"
    
    frmgrn.GrnMainGrid.TextMatrix(0, 1) = "  * Po No     "
    frmgrn.GrnMainGrid.ColWidth(1) = 1000
    frmgrn.GrnMainGrid.ColAlignment(1) = 3
    
    frmgrn.GrnMainGrid.TextMatrix(0, 2) = "* Po Date    "
    frmgrn.GrnMainGrid.ColWidth(2) = 1300
    frmgrn.GrnMainGrid.ColAlignment(2) = 3
    
    frmgrn.GrnMainGrid.TextMatrix(0, 3) = "  * Item Name    "
    frmgrn.GrnMainGrid.ColWidth(3) = 2500
    
    frmgrn.GrnMainGrid.TextMatrix(0, 4) = "  * Colour  "
    frmgrn.GrnMainGrid.ColWidth(4) = 1800
    
    frmgrn.GrnMainGrid.TextMatrix(0, 5) = " * Size "
    frmgrn.GrnMainGrid.ColWidth(5) = 1600
    frmgrn.GrnMainGrid.ColAlignment(5) = 1
     
    frmgrn.GrnMainGrid.TextMatrix(0, 6) = "* Po Qty "
    frmgrn.GrnMainGrid.ColWidth(6) = 1300
    
    frmgrn.GrnMainGrid.TextMatrix(0, 7) = " UOM "
    frmgrn.GrnMainGrid.ColWidth(7) = 1400
    
    frmgrn.GrnMainGrid.TextMatrix(0, 8) = "  Balance Qty "
    frmgrn.GrnMainGrid.ColWidth(8) = 1500
    
    frmgrn.GrnMainGrid.TextMatrix(0, 9) = " Receive Qty "
    frmgrn.GrnMainGrid.ColWidth(9) = 1500
    
    frmgrn.GrnMainGrid.TextMatrix(0, 10) = " ID "
    frmgrn.GrnMainGrid.ColWidth(10) = 0
    frmgrn.GrnMainGrid.ColAlignment(10) = 3
    
    frmgrn.GrnMainGrid.TextMatrix(0, 11) = "Total Grn "
    frmgrn.GrnMainGrid.ColWidth(11) = 0
End Sub
Sub frntcs()
    frmtc.Grid.ColWidth(0) = 600
    frmtc.Grid.ColAlignment(0) = 3
    frmtc.Grid.TextMatrix(0, 0) = " S.No "
    frmtc.Grid.ColWidth(1) = 8000
    frmtc.Grid.TextMatrix(0, 1) = "  # Terms & Conditions "
    frmtc.Height = 9600
    frmtc.Width = 11160
    frmtc.Top = 2
    frmtc.Left = 2
End Sub
Sub frmgrnopen()
    frmopenrec.GrnGrid.ColWidth(0) = 700
    frmopenrec.GrnGrid.ColAlignment(0) = 3
    frmopenrec.GrnGrid.TextMatrix(0, 0) = "S.No"
    frmopenrec.GrnGrid.TextMatrix(0, 1) = "  * Item Name    "
    frmopenrec.GrnGrid.ColWidth(1) = 3300
    frmopenrec.GrnGrid.TextMatrix(0, 2) = "  * Colour  "
    frmopenrec.GrnGrid.ColWidth(2) = 2800
    frmopenrec.GrnGrid.TextMatrix(0, 3) = "* Size "
    frmopenrec.GrnGrid.ColWidth(3) = 2800
    frmopenrec.GrnGrid.ColAlignment(3) = 1
    frmopenrec.GrnGrid.TextMatrix(0, 4) = " Receive Qty "
    frmopenrec.GrnGrid.ColWidth(4) = 2200
    frmopenrec.GrnGrid.TextMatrix(0, 5) = " * UOM "
    frmopenrec.GrnGrid.ColWidth(5) = 2300
    
End Sub
Sub frmgrnmainopengrid()
    frmopengrnmain.GrnMainGrid.ColWidth(0) = 900
    frmopengrnmain.GrnMainGrid.ColAlignment(0) = 3
    frmopengrnmain.GrnMainGrid.TextMatrix(0, 0) = "S.No"
    
    frmopengrnmain.GrnMainGrid.TextMatrix(0, 1) = " # Open GRN No    "
    frmopengrnmain.GrnMainGrid.ColWidth(1) = 2000
    frmopengrnmain.GrnMainGrid.ColAlignment(1) = 3
    
    frmopengrnmain.GrnMainGrid.TextMatrix(0, 2) = " Open GRN Date  "
    frmopengrnmain.GrnMainGrid.ColWidth(2) = 1800
    frmopengrnmain.GrnMainGrid.ColAlignment(2) = 3
    
    frmopengrnmain.GrnMainGrid.TextMatrix(0, 3) = " # Supplier   "
    frmopengrnmain.GrnMainGrid.ColWidth(3) = 4000
    
    frmopengrnmain.GrnMainGrid.TextMatrix(0, 4) = "  Sup.DC No "
    frmopengrnmain.GrnMainGrid.ColWidth(4) = 1300
    frmopengrnmain.GrnMainGrid.ColAlignment(4) = 3
    
    frmopengrnmain.GrnMainGrid.TextMatrix(0, 5) = " # Department "
    frmopengrnmain.GrnMainGrid.ColWidth(5) = 2200
    
End Sub
Sub frmDeliveryss()
    frmDelivery.DeliveryGrid.ColWidth(0) = 700
    frmDelivery.DeliveryGrid.ColAlignment(0) = 3
    frmDelivery.DeliveryGrid.TextMatrix(0, 0) = "S.No"
    
    frmDelivery.DeliveryGrid.TextMatrix(0, 1) = "* GRN No    "
    frmDelivery.DeliveryGrid.ColWidth(1) = 1200
    frmDelivery.DeliveryGrid.ColAlignment(1) = 3
    
    frmDelivery.DeliveryGrid.TextMatrix(0, 2) = "* GRN Date  "
    frmDelivery.DeliveryGrid.ColWidth(2) = 1200
    frmDelivery.DeliveryGrid.ColAlignment(2) = 3
    
    frmDelivery.DeliveryGrid.TextMatrix(0, 3) = "* Item Name "
    frmDelivery.DeliveryGrid.ColWidth(3) = 2200
    frmDelivery.DeliveryGrid.ColAlignment(3) = 1
        
    frmDelivery.DeliveryGrid.TextMatrix(0, 4) = " * Colour"
    frmDelivery.DeliveryGrid.ColWidth(4) = 2000
    
    frmDelivery.DeliveryGrid.TextMatrix(0, 5) = " * Size "
    frmDelivery.DeliveryGrid.ColWidth(5) = 2000
    frmDelivery.DeliveryGrid.ColAlignment(5) = 1
    
    frmDelivery.DeliveryGrid.TextMatrix(0, 6) = " * Stock Qty "
    frmDelivery.DeliveryGrid.ColWidth(6) = 1600
    
    frmDelivery.DeliveryGrid.TextMatrix(0, 7) = " * UOM "
    frmDelivery.DeliveryGrid.ColWidth(7) = 1800
    
    frmDelivery.DeliveryGrid.TextMatrix(0, 8) = "  Issue Qty "
    frmDelivery.DeliveryGrid.ColWidth(8) = 1600
    
    frmDelivery.DeliveryGrid.TextMatrix(0, 9) = " GRN Id "
    frmDelivery.DeliveryGrid.ColWidth(9) = 1800
    frmDelivery.DeliveryGrid.ColAlignment(9) = 3
  
End Sub

Sub frmgrnmaingrids()
    frmgrnmain.GrnMainGrid.ColWidth(0) = 1200
    frmgrnmain.GrnMainGrid.ColAlignment(0) = 3
    frmgrnmain.GrnMainGrid.TextMatrix(0, 0) = "S.No"
    
    frmgrnmain.GrnMainGrid.TextMatrix(0, 1) = " # GRN No    "
    frmgrnmain.GrnMainGrid.ColWidth(1) = 1800
    frmgrnmain.GrnMainGrid.ColAlignment(1) = 3
    
    frmgrnmain.GrnMainGrid.TextMatrix(0, 2) = "  GRN Date  "
    frmgrnmain.GrnMainGrid.ColWidth(2) = 1600
    frmgrnmain.GrnMainGrid.ColAlignment(2) = 3
    
    frmgrnmain.GrnMainGrid.TextMatrix(0, 3) = " # Supplier   "
    frmgrnmain.GrnMainGrid.ColWidth(3) = 4000
    
    frmgrnmain.GrnMainGrid.TextMatrix(0, 4) = "  Sup.DC No "
    frmgrnmain.GrnMainGrid.ColWidth(4) = 2000
    frmgrnmain.GrnMainGrid.ColAlignment(4) = 3
    
    frmgrnmain.GrnMainGrid.TextMatrix(0, 5) = " # Department "
    frmgrnmain.GrnMainGrid.ColWidth(5) = 3000
    
    'frmgrnmain.WindowState = 2
    
'    frmgrnmain.GrnMainGrid.TextMatrix(0, 6) = " # PO No "
'    frmgrnmain.GrnMainGrid.ColWidth(6) = 1500
'    frmgrnmain.GrnMainGrid.ColAlignment(6) = 3
   
End Sub
Sub frmgrneditgrids()
    frmgrn.GrnEditGrid.ColWidth(0) = 600
    frmgrn.GrnEditGrid.ColAlignment(0) = 3
    frmgrn.GrnEditGrid.TextMatrix(0, 0) = "S.No"
    
    frmgrn.GrnEditGrid.TextMatrix(0, 1) = "  * Po No     "
    frmgrn.GrnEditGrid.ColWidth(1) = 950
    frmgrn.GrnEditGrid.ColAlignment(1) = 3
    
    frmgrn.GrnEditGrid.TextMatrix(0, 2) = "* Po Date    "
    frmgrn.GrnEditGrid.ColWidth(2) = 1250
    frmgrn.GrnEditGrid.ColAlignment(2) = 3
    
    frmgrn.GrnEditGrid.TextMatrix(0, 3) = "  * Item Name    "
    frmgrn.GrnEditGrid.ColWidth(3) = 2400
    
    frmgrn.GrnEditGrid.TextMatrix(0, 4) = "  * Colour  "
    frmgrn.GrnEditGrid.ColWidth(4) = 2000
    
    frmgrn.GrnEditGrid.TextMatrix(0, 5) = " * Size "
    frmgrn.GrnEditGrid.ColWidth(5) = 1700
    frmgrn.GrnEditGrid.ColAlignment(5) = 1
     
    frmgrn.GrnEditGrid.TextMatrix(0, 6) = "* Po Qty "
    frmgrn.GrnEditGrid.ColWidth(6) = 1500
    
    frmgrn.GrnEditGrid.TextMatrix(0, 7) = " UOM "
    frmgrn.GrnEditGrid.ColWidth(7) = 1300
    
    frmgrn.GrnEditGrid.TextMatrix(0, 8) = "  Balance Qty "
    frmgrn.GrnEditGrid.ColWidth(8) = 1500
    
    frmgrn.GrnEditGrid.TextMatrix(0, 9) = " Receive Qty "
    frmgrn.GrnEditGrid.ColWidth(9) = 1500
    
    frmgrn.GrnEditGrid.TextMatrix(0, 10) = " ID "
    frmgrn.GrnEditGrid.ColWidth(10) = 900
    frmgrn.GrnEditGrid.ColAlignment(10) = 3
    
    frmgrn.GrnEditGrid.TextMatrix(0, 11) = "Total Grn "
    frmgrn.GrnEditGrid.ColWidth(11) = 900
    
End Sub
Sub deliverygridloads()
    frmdeliveryselection.DeliveryGrid.ColWidth(0) = 1000
    frmdeliveryselection.DeliveryGrid.ColAlignment(0) = 3
    frmdeliveryselection.DeliveryGrid.TextMatrix(0, 0) = "S.No"
    
    frmdeliveryselection.DeliveryGrid.TextMatrix(0, 1) = " # GRN No    "
    frmdeliveryselection.DeliveryGrid.ColWidth(1) = 1500
    frmdeliveryselection.DeliveryGrid.ColAlignment(1) = 3
    
    frmdeliveryselection.DeliveryGrid.TextMatrix(0, 2) = "  GRN Date  "
    frmdeliveryselection.DeliveryGrid.ColWidth(2) = 1600
    frmdeliveryselection.DeliveryGrid.ColAlignment(2) = 3
    
    frmdeliveryselection.DeliveryGrid.TextMatrix(0, 3) = " # Supplier   "
    frmdeliveryselection.DeliveryGrid.ColWidth(3) = 4000
    
    frmdeliveryselection.DeliveryGrid.TextMatrix(0, 4) = "  Sup.DC No "
    frmdeliveryselection.DeliveryGrid.ColWidth(4) = 1800
    frmdeliveryselection.DeliveryGrid.ColAlignment(4) = 3
    
    frmdeliveryselection.DeliveryGrid.TextMatrix(0, 5) = " # Department "
    frmdeliveryselection.DeliveryGrid.ColWidth(5) = 3300
        
    frmdeliveryselection.DeliveryGrid.TextMatrix(0, 6) = "Select"
    frmdeliveryselection.DeliveryGrid.ColWidth(6) = 800
    frmdeliveryselection.DeliveryGrid.ColAlignment(6) = 1
End Sub
Sub deliverygridloaditem()
    frmDeliverys.DeliveryGrid.ColWidth(0) = 800
    frmDeliverys.DeliveryGrid.ColAlignment(0) = 3
    frmDeliverys.DeliveryGrid.TextMatrix(0, 0) = "S.No"
    
    frmDeliverys.DeliveryGrid.TextMatrix(0, 1) = "GRN No    "
    frmDeliverys.DeliveryGrid.ColWidth(1) = 1200
    frmDeliverys.DeliveryGrid.ColAlignment(1) = 3
    
    frmDeliverys.DeliveryGrid.TextMatrix(0, 2) = " GRN Date  "
    frmDeliverys.DeliveryGrid.ColWidth(2) = 1600
    frmDeliverys.DeliveryGrid.ColAlignment(2) = 3
    
    frmDeliverys.DeliveryGrid.TextMatrix(0, 3) = "Item Name   "
    frmDeliverys.DeliveryGrid.ColWidth(3) = 2200
    
    frmDeliverys.DeliveryGrid.TextMatrix(0, 4) = "Colour "
    frmDeliverys.DeliveryGrid.ColWidth(4) = 1800
    
    frmDeliverys.DeliveryGrid.TextMatrix(0, 5) = "Size "
    frmDeliverys.DeliveryGrid.ColWidth(5) = 1800
    
    frmDeliverys.DeliveryGrid.TextMatrix(0, 6) = " Stock Qty "
    frmDeliverys.DeliveryGrid.ColWidth(6) = 1700
    
    frmDeliverys.DeliveryGrid.TextMatrix(0, 7) = "UOM"
    frmDeliverys.DeliveryGrid.ColWidth(7) = 1600
    
    frmDeliverys.DeliveryGrid.TextMatrix(0, 8) = " Issue Qty "
    frmDeliverys.DeliveryGrid.ColWidth(8) = 1700
    
    frmDeliverys.DeliveryGrid.TextMatrix(0, 9) = " GRN ID "
    frmDeliverys.DeliveryGrid.ColWidth(9) = 0
    frmDeliverys.DeliveryGrid.ColAlignment(9) = 3
    
    frmDeliverys.DeliveryGrid.TextMatrix(0, 10) = "Tot Deliv "
    frmDeliverys.DeliveryGrid.ColWidth(10) = 0
    
    frmDeliverys.DetailsGrid.TextMatrix(0, 0) = "PO No "
    frmDeliverys.DetailsGrid.ColWidth(0) = 1200
    frmDeliverys.DetailsGrid.ColAlignment(0) = 3
    
    frmDeliverys.DetailsGrid.TextMatrix(0, 1) = "PO Date"
    frmDeliverys.DetailsGrid.ColWidth(1) = 1400
    frmDeliverys.DetailsGrid.ColAlignment(1) = 3
    
    frmDeliverys.DetailsGrid.TextMatrix(0, 2) = "GRN No "
    frmDeliverys.DetailsGrid.ColWidth(2) = 1500
    frmDeliverys.DetailsGrid.ColAlignment(2) = 3
    
    frmDeliverys.DetailsGrid.TextMatrix(0, 3) = "GRN Date "
    frmDeliverys.DetailsGrid.ColWidth(3) = 1500
    frmDeliverys.DetailsGrid.ColAlignment(3) = 3
    
    frmDeliverys.DetailsGrid.TextMatrix(0, 4) = "Supplier Name "
    frmDeliverys.DetailsGrid.ColWidth(4) = 3500
    
    frmDeliverys.DetailsGrid.TextMatrix(0, 5) = "Department "
    frmDeliverys.DetailsGrid.ColWidth(5) = 2200
    
    frmDeliverys.DetailsGrid.TextMatrix(0, 6) = "Received Qty "
    frmDeliverys.DetailsGrid.ColWidth(6) = 2500
    
    frmDeliverys.DeliveryMainGrid.ColWidth(0) = 800
    frmDeliverys.DeliveryMainGrid.ColAlignment(0) = 3
    frmDeliverys.DeliveryMainGrid.TextMatrix(0, 0) = "S.No"
    
    frmDeliverys.DeliveryMainGrid.TextMatrix(0, 1) = "GRN No    "
    frmDeliverys.DeliveryMainGrid.ColWidth(1) = 1200
    frmDeliverys.DeliveryMainGrid.ColAlignment(1) = 3
    
    frmDeliverys.DeliveryMainGrid.TextMatrix(0, 2) = " GRN Date  "
    frmDeliverys.DeliveryMainGrid.ColWidth(2) = 1600
    frmDeliverys.DeliveryMainGrid.ColAlignment(2) = 3
    
    frmDeliverys.DeliveryMainGrid.TextMatrix(0, 3) = "Item Name   "
    frmDeliverys.DeliveryMainGrid.ColWidth(3) = 2200
    
    frmDeliverys.DeliveryMainGrid.TextMatrix(0, 4) = "Colour "
    frmDeliverys.DeliveryMainGrid.ColWidth(4) = 1800
    
    frmDeliverys.DeliveryMainGrid.TextMatrix(0, 5) = "Size "
    frmDeliverys.DeliveryMainGrid.ColWidth(5) = 1800
    
    frmDeliverys.DeliveryMainGrid.TextMatrix(0, 6) = " Stock Qty "
    frmDeliverys.DeliveryMainGrid.ColWidth(6) = 1700
    
    frmDeliverys.DeliveryMainGrid.TextMatrix(0, 7) = "UOM"
    frmDeliverys.DeliveryMainGrid.ColWidth(7) = 1600
    
    frmDeliverys.DeliveryMainGrid.TextMatrix(0, 8) = " Issue Qty "
    frmDeliverys.DeliveryMainGrid.ColWidth(8) = 1700
    
    frmDeliverys.DeliveryMainGrid.TextMatrix(0, 9) = " GRN ID "
    frmDeliverys.DeliveryMainGrid.ColWidth(9) = 0
    frmDeliverys.DeliveryMainGrid.ColAlignment(9) = 3
    
End Sub
Sub frmdeliverymaingridloads()
    frmdeliverymain.DeliveryMainGrid.ColWidth(0) = 1000
    frmdeliverymain.DeliveryMainGrid.ColAlignment(0) = 3
    frmdeliverymain.DeliveryMainGrid.TextMatrix(0, 0) = "S.No"
    
    frmdeliverymain.DeliveryMainGrid.TextMatrix(0, 1) = " # DC No    "
    frmdeliverymain.DeliveryMainGrid.ColWidth(1) = 2000
    frmdeliverymain.DeliveryMainGrid.ColAlignment(1) = 3
    
    frmdeliverymain.DeliveryMainGrid.TextMatrix(0, 2) = "  DC Date  "
    frmdeliverymain.DeliveryMainGrid.ColWidth(2) = 2200
    frmdeliverymain.DeliveryMainGrid.ColAlignment(2) = 3
    
    frmdeliverymain.DeliveryMainGrid.TextMatrix(0, 3) = " # Supplier Name   "
    frmdeliverymain.DeliveryMainGrid.ColWidth(3) = 5000
    
    frmdeliverymain.DeliveryMainGrid.TextMatrix(0, 4) = " # Department  "
    frmdeliverymain.DeliveryMainGrid.ColWidth(4) = 3500
    frmdeliverymain.DeliveryMainGrid.ColAlignment(4) = 1
    
End Sub
Sub deliverygridopenloaditemd()
    frmopenDeliverys.DeliveryGrid.ColWidth(0) = 800
    frmopenDeliverys.DeliveryGrid.ColAlignment(0) = 3
    frmopenDeliverys.DeliveryGrid.TextMatrix(0, 0) = "S.No"
    
    frmopenDeliverys.DeliveryGrid.TextMatrix(0, 1) = "GRN No    "
    frmopenDeliverys.DeliveryGrid.ColWidth(1) = 1200
    frmopenDeliverys.DeliveryGrid.ColAlignment(1) = 3
    
    frmopenDeliverys.DeliveryGrid.TextMatrix(0, 2) = " GRN Date  "
    frmopenDeliverys.DeliveryGrid.ColWidth(2) = 1600
    frmopenDeliverys.DeliveryGrid.ColAlignment(2) = 3
    
    frmopenDeliverys.DeliveryGrid.TextMatrix(0, 3) = "Item Name   "
    frmopenDeliverys.DeliveryGrid.ColWidth(3) = 2200
    
    frmopenDeliverys.DeliveryGrid.TextMatrix(0, 4) = "Colour "
    frmopenDeliverys.DeliveryGrid.ColWidth(4) = 1800
    
    frmopenDeliverys.DeliveryGrid.TextMatrix(0, 5) = "Size "
    frmopenDeliverys.DeliveryGrid.ColWidth(5) = 1800
    
    frmopenDeliverys.DeliveryGrid.TextMatrix(0, 6) = " Stock Qty "
    frmopenDeliverys.DeliveryGrid.ColWidth(6) = 1700
    
    frmopenDeliverys.DeliveryGrid.TextMatrix(0, 7) = "UOM"
    frmopenDeliverys.DeliveryGrid.ColWidth(7) = 1600
    
    frmopenDeliverys.DeliveryGrid.TextMatrix(0, 8) = " Issue Qty "
    frmopenDeliverys.DeliveryGrid.ColWidth(8) = 1700
    
    frmopenDeliverys.DeliveryGrid.TextMatrix(0, 9) = " GRN ID "
    frmopenDeliverys.DeliveryGrid.ColWidth(9) = 0
    frmopenDeliverys.DeliveryGrid.ColAlignment(9) = 3
    
    frmopenDeliverys.DeliveryGrid.TextMatrix(0, 10) = "Tot Deliv "
    frmopenDeliverys.DeliveryGrid.ColWidth(10) = 0
    
    frmopenDeliverys.DetailsGrid.TextMatrix(0, 0) = "PO No "
    frmopenDeliverys.DetailsGrid.ColWidth(0) = 1200
    frmopenDeliverys.DetailsGrid.ColAlignment(0) = 3
    
    frmopenDeliverys.DetailsGrid.TextMatrix(0, 1) = "PO Date"
    frmopenDeliverys.DetailsGrid.ColWidth(1) = 1400
    frmopenDeliverys.DetailsGrid.ColAlignment(1) = 3
    
    frmopenDeliverys.DetailsGrid.TextMatrix(0, 2) = "GRN No "
    frmopenDeliverys.DetailsGrid.ColWidth(2) = 1200
    frmopenDeliverys.DetailsGrid.ColAlignment(2) = 3
    
    frmopenDeliverys.DetailsGrid.TextMatrix(0, 3) = "GRN Date "
    frmopenDeliverys.DetailsGrid.ColWidth(3) = 1400
    frmopenDeliverys.DetailsGrid.ColAlignment(3) = 3
    
    frmopenDeliverys.DetailsGrid.TextMatrix(0, 4) = "Supplier Name "
    frmopenDeliverys.DetailsGrid.ColWidth(4) = 2500
    
    frmopenDeliverys.DetailsGrid.TextMatrix(0, 5) = "Department "
    frmopenDeliverys.DetailsGrid.ColWidth(5) = 2000
    
    frmopenDeliverys.DetailsGrid.TextMatrix(0, 6) = "Received Qty "
    frmopenDeliverys.DetailsGrid.ColWidth(6) = 1800
    
    frmopenDeliverys.DetailsGrid.TextMatrix(0, 7) = "Type "
    frmopenDeliverys.DetailsGrid.ColWidth(7) = 3000
    frmopenDeliverys.DetailsGrid.ColAlignment(7) = 3
    
    
    frmopenDeliverys.DeliveryMainGrid.ColWidth(0) = 800
    frmopenDeliverys.DeliveryMainGrid.ColAlignment(0) = 3
    frmopenDeliverys.DeliveryMainGrid.TextMatrix(0, 0) = "S.No"
    
    frmopenDeliverys.DeliveryMainGrid.TextMatrix(0, 1) = "GRN No    "
    frmopenDeliverys.DeliveryMainGrid.ColWidth(1) = 1200
    frmopenDeliverys.DeliveryMainGrid.ColAlignment(1) = 3
    
    frmopenDeliverys.DeliveryMainGrid.TextMatrix(0, 2) = " GRN Date  "
    frmopenDeliverys.DeliveryMainGrid.ColWidth(2) = 1600
    frmopenDeliverys.DeliveryMainGrid.ColAlignment(2) = 3
    
    frmopenDeliverys.DeliveryMainGrid.TextMatrix(0, 3) = "Item Name   "
    frmopenDeliverys.DeliveryMainGrid.ColWidth(3) = 2200
    
    frmopenDeliverys.DeliveryMainGrid.TextMatrix(0, 4) = "Colour "
    frmopenDeliverys.DeliveryMainGrid.ColWidth(4) = 1800
    
    frmopenDeliverys.DeliveryMainGrid.TextMatrix(0, 5) = "Size "
    frmopenDeliverys.DeliveryMainGrid.ColWidth(5) = 1800
    
    frmopenDeliverys.DeliveryMainGrid.TextMatrix(0, 6) = " Stock Qty "
    frmopenDeliverys.DeliveryMainGrid.ColWidth(6) = 1700
    
    frmopenDeliverys.DeliveryMainGrid.TextMatrix(0, 7) = "UOM"
    frmopenDeliverys.DeliveryMainGrid.ColWidth(7) = 1600
    
    frmopenDeliverys.DeliveryMainGrid.TextMatrix(0, 8) = " Issue Qty "
    frmopenDeliverys.DeliveryMainGrid.ColWidth(8) = 1700
    
    frmopenDeliverys.DeliveryMainGrid.TextMatrix(0, 9) = " GRN ID "
    frmopenDeliverys.DeliveryMainGrid.ColWidth(9) = 0
    frmopenDeliverys.DeliveryMainGrid.ColAlignment(9) = 3
    
End Sub
Sub frmdeliverymaingridload()
    frmdeliverymain.DeliveryMainGrid.ColWidth(0) = 1000
    frmdeliverymain.DeliveryMainGrid.ColAlignment(0) = 3
    frmdeliverymain.DeliveryMainGrid.TextMatrix(0, 0) = "S.No"
    
    frmdeliverymain.DeliveryMainGrid.TextMatrix(0, 1) = " # DC No    "
    frmdeliverymain.DeliveryMainGrid.ColWidth(1) = 1600
    frmdeliverymain.DeliveryMainGrid.ColAlignment(1) = 3
    
    frmdeliverymain.DeliveryMainGrid.TextMatrix(0, 2) = "  DC Date  "
    frmdeliverymain.DeliveryMainGrid.ColWidth(2) = 1800
    frmdeliverymain.DeliveryMainGrid.ColAlignment(2) = 3
    
    frmdeliverymain.DeliveryMainGrid.TextMatrix(0, 3) = " # Supplier Name   "
    frmdeliverymain.DeliveryMainGrid.ColWidth(3) = 4000
    
    frmdeliverymain.DeliveryMainGrid.TextMatrix(0, 4) = " # Department  "
    frmdeliverymain.DeliveryMainGrid.ColWidth(4) = 2200
    frmdeliverymain.DeliveryMainGrid.ColAlignment(4) = 1
    
    frmdeliverymain.DeliveryMainGrid.TextMatrix(0, 5) = " # Delivery Type "
    frmdeliverymain.DeliveryMainGrid.ColWidth(5) = 3500
        
    frmdeliverymain.DeliveryMainGrid.TextMatrix(0, 6) = " # Delivery Typeno "
    frmdeliverymain.DeliveryMainGrid.ColWidth(6) = 1200
End Sub
Sub frmdeliverygridloadmains()
    frmopendeliverymain.DeliveryMainGrid.ColWidth(0) = 1000
    frmopendeliverymain.DeliveryMainGrid.ColAlignment(0) = 3
    frmopendeliverymain.DeliveryMainGrid.TextMatrix(0, 0) = "S.No"
    
    frmopendeliverymain.DeliveryMainGrid.TextMatrix(0, 1) = " # DC No    "
    frmopendeliverymain.DeliveryMainGrid.ColWidth(1) = 2000
    frmopendeliverymain.DeliveryMainGrid.ColAlignment(1) = 3
    
    frmopendeliverymain.DeliveryMainGrid.TextMatrix(0, 2) = "  DC Date  "
    frmopendeliverymain.DeliveryMainGrid.ColWidth(2) = 2000
    frmopendeliverymain.DeliveryMainGrid.ColAlignment(2) = 3
    
    frmopendeliverymain.DeliveryMainGrid.TextMatrix(0, 3) = " # Supplier Name   "
    frmopendeliverymain.DeliveryMainGrid.ColWidth(3) = 4500
    
    frmopendeliverymain.DeliveryMainGrid.TextMatrix(0, 4) = " # Department  "
    frmopendeliverymain.DeliveryMainGrid.ColWidth(4) = 3000
    frmopendeliverymain.DeliveryMainGrid.ColAlignment(4) = 1
        
End Sub
Sub poclosuregridloads()
    frmpoclosure.PoMainGrid.ColWidth(0) = 1000
    frmpoclosure.PoMainGrid.ColAlignment(0) = 3
    frmpoclosure.PoMainGrid.TextMatrix(0, 0) = "S.No"
    
    frmpoclosure.PoMainGrid.TextMatrix(0, 1) = " # PO No    "
    frmpoclosure.PoMainGrid.ColWidth(1) = 1300
    frmpoclosure.PoMainGrid.ColAlignment(1) = 3
    
    frmpoclosure.PoMainGrid.TextMatrix(0, 2) = "  PO Date  "
    frmpoclosure.PoMainGrid.ColWidth(2) = 1600
    frmpoclosure.PoMainGrid.ColAlignment(2) = 3
    
    frmpoclosure.PoMainGrid.TextMatrix(0, 3) = " # Supplier   "
    frmpoclosure.PoMainGrid.ColWidth(3) = 3500
    
    frmpoclosure.PoMainGrid.TextMatrix(0, 4) = "  Net Amt (Rs) "
    frmpoclosure.PoMainGrid.ColWidth(4) = 1800
    
    
    frmpoclosure.PoMainGrid.TextMatrix(0, 5) = " # Department "
    frmpoclosure.PoMainGrid.ColWidth(5) = 2500
            
    frmpoclosure.PoMainGrid.TextMatrix(0, 6) = "Close"
    frmpoclosure.PoMainGrid.ColWidth(6) = 710
    frmpoclosure.PoMainGrid.ColAlignment(6) = 1
    
    frmpoclosure.PoMainGrid.TextMatrix(0, 7) = "Status  "
    frmpoclosure.PoMainGrid.ColWidth(7) = 1500
    frmpoclosure.PoMainGrid.ColAlignment(7) = 1
End Sub
Sub grnclosuregridloads()
    frmgrnclosure.GrnMainGrid.ColWidth(0) = 1000
    frmgrnclosure.GrnMainGrid.ColAlignment(0) = 3
    frmgrnclosure.GrnMainGrid.TextMatrix(0, 0) = "S.No"
    
    frmgrnclosure.GrnMainGrid.TextMatrix(0, 1) = " # GRN No    "
    frmgrnclosure.GrnMainGrid.ColWidth(1) = 1500
    frmgrnclosure.GrnMainGrid.ColAlignment(1) = 3
    
    frmgrnclosure.GrnMainGrid.TextMatrix(0, 2) = " GRN Date  "
    frmgrnclosure.GrnMainGrid.ColWidth(2) = 1800
    frmgrnclosure.GrnMainGrid.ColAlignment(2) = 3
    
    frmgrnclosure.GrnMainGrid.TextMatrix(0, 3) = " # Supplier   "
    frmgrnclosure.GrnMainGrid.ColWidth(3) = 4500
      
    frmgrnclosure.GrnMainGrid.TextMatrix(0, 4) = " # Department "
    frmgrnclosure.GrnMainGrid.ColWidth(4) = 3000
            
    frmgrnclosure.GrnMainGrid.TextMatrix(0, 5) = "Close"
    frmgrnclosure.GrnMainGrid.ColWidth(5) = 720
    frmgrnclosure.GrnMainGrid.ColAlignment(5) = 1
    
    frmgrnclosure.GrnMainGrid.TextMatrix(0, 6) = "Status  "
    frmgrnclosure.GrnMainGrid.ColWidth(6) = 1400
    frmgrnclosure.GrnMainGrid.ColAlignment(6) = 1
End Sub
Sub opengrnclosuregridloads()
    frmopengrnclosure.GrnMainGrid.ColWidth(0) = 1000
    frmopengrnclosure.GrnMainGrid.ColAlignment(0) = 3
    frmopengrnclosure.GrnMainGrid.TextMatrix(0, 0) = "S.No"
    
    frmopengrnclosure.GrnMainGrid.TextMatrix(0, 1) = " # Open GRN No    "
    frmopengrnclosure.GrnMainGrid.ColWidth(1) = 1500
    frmopengrnclosure.GrnMainGrid.ColAlignment(1) = 3
    
    frmopengrnclosure.GrnMainGrid.TextMatrix(0, 2) = " Open GRN Date  "
    frmopengrnclosure.GrnMainGrid.ColWidth(2) = 1800
    frmopengrnclosure.GrnMainGrid.ColAlignment(2) = 3
    
    frmopengrnclosure.GrnMainGrid.TextMatrix(0, 3) = " # Supplier   "
    frmopengrnclosure.GrnMainGrid.ColWidth(3) = 4500
      
    frmopengrnclosure.GrnMainGrid.TextMatrix(0, 4) = " # Department "
    frmopengrnclosure.GrnMainGrid.ColWidth(4) = 3000
            
    frmopengrnclosure.GrnMainGrid.TextMatrix(0, 5) = "Close"
    frmopengrnclosure.GrnMainGrid.ColWidth(5) = 720
    frmopengrnclosure.GrnMainGrid.ColAlignment(5) = 1
    
    frmopengrnclosure.GrnMainGrid.TextMatrix(0, 6) = "Status  "
    frmopengrnclosure.GrnMainGrid.ColWidth(6) = 1400
    frmopengrnclosure.GrnMainGrid.ColAlignment(6) = 1
End Sub
Sub frmgrnopendeliveryload()
    frmopendelivery.DeliveryGrid.ColWidth(0) = 800
    frmopendelivery.DeliveryGrid.ColAlignment(0) = 3
    frmopendelivery.DeliveryGrid.TextMatrix(0, 0) = "S.No"
    
    frmopendelivery.DeliveryGrid.TextMatrix(0, 1) = "  * Item Name    "
    frmopendelivery.DeliveryGrid.ColWidth(1) = 3500
    frmopendelivery.DeliveryGrid.TextMatrix(0, 2) = "  * Colour  "
    frmopendelivery.DeliveryGrid.ColWidth(2) = 2600
    
    frmopendelivery.DeliveryGrid.TextMatrix(0, 3) = "* Size "
    frmopendelivery.DeliveryGrid.ColWidth(3) = 2600
    
    frmopendelivery.DeliveryGrid.ColAlignment(3) = 1
    frmopendelivery.DeliveryGrid.TextMatrix(0, 4) = " Issue Qty "
    frmopendelivery.DeliveryGrid.ColWidth(4) = 2500
    
    frmopendelivery.DeliveryGrid.TextMatrix(0, 5) = " * UOM "
    frmopendelivery.DeliveryGrid.ColWidth(5) = 2500
    
    End Sub
Sub frmopendeliverygridmain()
    frmopensdeliverymain.DeliveryMainGrid.ColWidth(0) = 1200
    frmopensdeliverymain.DeliveryMainGrid.ColAlignment(0) = 3
    frmopensdeliverymain.DeliveryMainGrid.TextMatrix(0, 0) = "S.No"
    
    frmopensdeliverymain.DeliveryMainGrid.TextMatrix(0, 1) = " # OPEN DC No "
    frmopensdeliverymain.DeliveryMainGrid.ColWidth(1) = 2000
    frmopensdeliverymain.DeliveryMainGrid.ColAlignment(1) = 3
    
    frmopensdeliverymain.DeliveryMainGrid.TextMatrix(0, 2) = " OPEN DC Date  "
    frmopensdeliverymain.DeliveryMainGrid.ColWidth(2) = 2000
    frmopensdeliverymain.DeliveryMainGrid.ColAlignment(2) = 3
    
    frmopensdeliverymain.DeliveryMainGrid.TextMatrix(0, 3) = " # Supplier   "
    frmopensdeliverymain.DeliveryMainGrid.ColWidth(3) = 4500
           
    frmopensdeliverymain.DeliveryMainGrid.TextMatrix(0, 4) = " # Department "
    frmopensdeliverymain.DeliveryMainGrid.ColWidth(4) = 2500
    
End Sub
Sub frmgrninvoiceload()
    frminvoice.InvoiceGrid.ColWidth(0) = 800
    frminvoice.InvoiceGrid.ColAlignment(0) = 3
    frminvoice.InvoiceGrid.TextMatrix(0, 0) = "S.No"
    
    frminvoice.InvoiceGrid.TextMatrix(0, 1) = "  * Item Name    "
    frminvoice.InvoiceGrid.ColWidth(1) = 2100
    
    frminvoice.InvoiceGrid.TextMatrix(0, 2) = "  * Colour  "
    frminvoice.InvoiceGrid.ColWidth(2) = 1600
    
    frminvoice.InvoiceGrid.TextMatrix(0, 3) = "* Size "
    frminvoice.InvoiceGrid.ColWidth(3) = 1600
    frminvoice.InvoiceGrid.ColAlignment(3) = 1
    
    frminvoice.InvoiceGrid.TextMatrix(0, 4) = " * GRN Qty "
    frminvoice.InvoiceGrid.ColWidth(4) = 1500
    
    frminvoice.InvoiceGrid.TextMatrix(0, 5) = " * UOM "
    frminvoice.InvoiceGrid.ColWidth(5) = 1600
    
    frminvoice.InvoiceGrid.TextMatrix(0, 6) = " * Invoiced Qty "
    frminvoice.InvoiceGrid.ColWidth(6) = 1500
    
    frminvoice.InvoiceGrid.TextMatrix(0, 7) = " * Invoice Rate (Rs)"
    frminvoice.InvoiceGrid.ColWidth(7) = 1500
    
    frminvoice.InvoiceGrid.TextMatrix(0, 8) = "  Total Amt (Rs)"
    frminvoice.InvoiceGrid.ColWidth(8) = 1500
    
    frminvoice.InvoiceGrid.TextMatrix(0, 9) = " * GRN ID"
    frminvoice.InvoiceGrid.ColWidth(9) = 0
    
    frminvoice.InvoiceGrid.TextMatrix(0, 10) = " * Total Invoice Qty"
    frminvoice.InvoiceGrid.ColWidth(10) = 0
    
    frminvoice.InvoiceGrid.TextMatrix(0, 11) = " * GRN No"
    frminvoice.InvoiceGrid.ColWidth(11) = 0
    
    frminvoice.InvoiceDetailsGrid.TextMatrix(0, 0) = "PO No "
    frminvoice.InvoiceDetailsGrid.ColWidth(0) = 1200
    frminvoice.InvoiceDetailsGrid.ColAlignment(0) = 3
    
    frminvoice.InvoiceDetailsGrid.TextMatrix(0, 1) = "PO Date"
    frminvoice.InvoiceDetailsGrid.ColWidth(1) = 1400
    frminvoice.InvoiceDetailsGrid.ColAlignment(1) = 3
    
    frminvoice.InvoiceDetailsGrid.TextMatrix(0, 2) = "GRN No "
    frminvoice.InvoiceDetailsGrid.ColWidth(2) = 1200
    frminvoice.InvoiceDetailsGrid.ColAlignment(2) = 3
    
    frminvoice.InvoiceDetailsGrid.TextMatrix(0, 3) = "GRN Date "
    frminvoice.InvoiceDetailsGrid.ColWidth(3) = 1400
    frminvoice.InvoiceDetailsGrid.ColAlignment(3) = 3
    
    frminvoice.InvoiceDetailsGrid.TextMatrix(0, 4) = "Supplier Name "
    frminvoice.InvoiceDetailsGrid.ColWidth(4) = 2500
    
    frminvoice.InvoiceDetailsGrid.TextMatrix(0, 5) = "Department "
    frminvoice.InvoiceDetailsGrid.ColWidth(5) = 2000
    
    
    frminvoice.invoiceMainGrid.ColWidth(0) = 600
    frminvoice.invoiceMainGrid.ColAlignment(0) = 3
    frminvoice.invoiceMainGrid.TextMatrix(0, 0) = "S.No"
    
    frminvoice.invoiceMainGrid.TextMatrix(0, 1) = "  * Item Name    "
    frminvoice.invoiceMainGrid.ColWidth(1) = 1800
    
    frminvoice.invoiceMainGrid.TextMatrix(0, 2) = "  * Colour  "
    frminvoice.invoiceMainGrid.ColWidth(2) = 1250
    
    frminvoice.invoiceMainGrid.TextMatrix(0, 3) = "* Size "
    frminvoice.invoiceMainGrid.ColWidth(3) = 1250
    frminvoice.invoiceMainGrid.ColAlignment(3) = 1
    
    frminvoice.invoiceMainGrid.TextMatrix(0, 4) = " GRN Qty "
    frminvoice.invoiceMainGrid.ColWidth(4) = 1400
    
    frminvoice.invoiceMainGrid.TextMatrix(0, 5) = " * UOM "
    frminvoice.invoiceMainGrid.ColWidth(5) = 1200
    
    frminvoice.invoiceMainGrid.TextMatrix(0, 6) = "* Invoiced Qty "
    frminvoice.invoiceMainGrid.ColWidth(6) = 1400
    
    frminvoice.invoiceMainGrid.TextMatrix(0, 7) = " * Invoice Rate (Rs)"
    frminvoice.invoiceMainGrid.ColWidth(7) = 1500
    
    frminvoice.invoiceMainGrid.TextMatrix(0, 8) = " Total Amt(Rs)"
    frminvoice.invoiceMainGrid.ColWidth(8) = 1400
    
    frminvoice.invoiceMainGrid.TextMatrix(0, 9) = " * GRN ID"
    frminvoice.invoiceMainGrid.ColWidth(9) = 0
    
    frminvoice.invoiceMainGrid.TextMatrix(0, 10) = " * Total Invoice Qty"
    frminvoice.invoiceMainGrid.ColWidth(10) = 0
    
    frminvoice.invoiceMainGrid.TextMatrix(0, 11) = " GRN No"
    frminvoice.invoiceMainGrid.ColWidth(11) = 0
End Sub
Sub frminvoicemaingridload()
    frminvoicemain.invoiceMainGrid.ColWidth(0) = 1000
    frminvoicemain.invoiceMainGrid.TextMatrix(0, 0) = "S.No"
    frminvoicemain.invoiceMainGrid.ColAlignment(0) = 3
    
    frminvoicemain.invoiceMainGrid.TextMatrix(0, 1) = " # Invoice No"
    frminvoicemain.invoiceMainGrid.ColAlignment(1) = 3
    frminvoicemain.invoiceMainGrid.ColWidth(1) = 2000
    
    frminvoicemain.invoiceMainGrid.TextMatrix(0, 2) = "* Invoice Date"
    frminvoicemain.invoiceMainGrid.ColAlignment(2) = 3
    frminvoicemain.invoiceMainGrid.ColWidth(2) = 2000
    
    frminvoicemain.invoiceMainGrid.TextMatrix(0, 3) = " #  Sup.Bill.No"
    frminvoicemain.invoiceMainGrid.ColAlignment(3) = 3
    frminvoicemain.invoiceMainGrid.ColWidth(3) = 2000
    
    frminvoicemain.invoiceMainGrid.TextMatrix(0, 4) = "#  Supplier Name"
    frminvoicemain.invoiceMainGrid.ColAlignment(4) = 1
    frminvoicemain.invoiceMainGrid.ColWidth(4) = 4500
    
    frminvoicemain.invoiceMainGrid.TextMatrix(0, 5) = "* Net Amount (Rs)"
    frminvoicemain.invoiceMainGrid.ColAlignment(5) = 1
    frminvoicemain.invoiceMainGrid.ColWidth(5) = 2500
End Sub
Sub openinvoicegridalign()

    frmopeninvoice.InvoiceGrid.ColWidth(0) = 600
    frmopeninvoice.InvoiceGrid.ColAlignment(0) = 3
    frmopeninvoice.InvoiceGrid.TextMatrix(0, 0) = " S.No "
    
    frmopeninvoice.InvoiceGrid.TextMatrix(0, 1) = "  * Item Name     "
    frmopeninvoice.InvoiceGrid.ColWidth(1) = 2800
    
    frmopeninvoice.InvoiceGrid.TextMatrix(0, 2) = "  * Colour     "
    frmopeninvoice.InvoiceGrid.ColWidth(2) = 2000
    
    frmopeninvoice.InvoiceGrid.TextMatrix(0, 3) = "  * Size   "
    frmopeninvoice.InvoiceGrid.ColWidth(3) = 2100
    frmopeninvoice.InvoiceGrid.ColAlignment(3) = 1
    
    frmopeninvoice.InvoiceGrid.TextMatrix(0, 4) = " Inv.Qty "
    frmopeninvoice.InvoiceGrid.ColWidth(4) = 1600
    
    frmopeninvoice.InvoiceGrid.TextMatrix(0, 5) = "     UOM "
    frmopeninvoice.InvoiceGrid.ColWidth(5) = 1650
    
    frmopeninvoice.InvoiceGrid.TextMatrix(0, 6) = "  Rate (Rs) "
    frmopeninvoice.InvoiceGrid.ColWidth(6) = 1500
    
    frmopeninvoice.InvoiceGrid.TextMatrix(0, 7) = "  Tot.Amount (Rs) "
    frmopeninvoice.InvoiceGrid.ColWidth(7) = 2000
End Sub
Sub frmopeninvoicemaingridload()
     frmopeninvoicemain.invoiceMainGrid.ColWidth(0) = 900
     frmopeninvoicemain.invoiceMainGrid.TextMatrix(0, 0) = "S.No"
     frmopeninvoicemain.invoiceMainGrid.ColAlignment(0) = 3
    
     frmopeninvoicemain.invoiceMainGrid.TextMatrix(0, 1) = " # Open Invoice No"
     frmopeninvoicemain.invoiceMainGrid.ColAlignment(1) = 3
     frmopeninvoicemain.invoiceMainGrid.ColWidth(1) = 1900
    
     frmopeninvoicemain.invoiceMainGrid.TextMatrix(0, 2) = "* Open Invoice Date"
     frmopeninvoicemain.invoiceMainGrid.ColAlignment(2) = 3
     frmopeninvoicemain.invoiceMainGrid.ColWidth(2) = 2000
    
     frmopeninvoicemain.invoiceMainGrid.TextMatrix(0, 3) = "#Sup.Bill.No"
     frmopeninvoicemain.invoiceMainGrid.ColAlignment(3) = 3
     frmopeninvoicemain.invoiceMainGrid.ColWidth(3) = 1600
    
     frmopeninvoicemain.invoiceMainGrid.TextMatrix(0, 4) = "#  Supplier Name"
     frmopeninvoicemain.invoiceMainGrid.ColAlignment(4) = 1
     frmopeninvoicemain.invoiceMainGrid.ColWidth(4) = 3400
     
     frmopeninvoicemain.invoiceMainGrid.TextMatrix(0, 5) = "#  Department Name"
     frmopeninvoicemain.invoiceMainGrid.ColAlignment(5) = 1
     frmopeninvoicemain.invoiceMainGrid.ColWidth(5) = 2300
    
     frmopeninvoicemain.invoiceMainGrid.TextMatrix(0, 6) = "* Net Amount (Rs)"
     'frmopeninvoicemain.invoiceMainGrid.ColAlignment(6) = 1
     frmopeninvoicemain.invoiceMainGrid.ColWidth(6) = 1800
End Sub
Sub frmpaymentgridload()
    frmpayment.PaymentGrid.ColWidth(0) = 900
    frmpayment.PaymentGrid.TextMatrix(0, 0) = "S.No"
    frmpayment.PaymentGrid.ColAlignment(0) = 3
    
    frmpayment.PaymentGrid.ColWidth(1) = 2000
    frmpayment.PaymentGrid.TextMatrix(0, 1) = "* Invoice No"
    frmpayment.PaymentGrid.ColAlignment(1) = 3
    
    frmpayment.PaymentGrid.ColWidth(2) = 2000
    frmpayment.PaymentGrid.TextMatrix(0, 2) = "Invoice Date"
    frmpayment.PaymentGrid.ColAlignment(2) = 3
    
    frmpayment.PaymentGrid.ColWidth(3) = 2300
    frmpayment.PaymentGrid.TextMatrix(0, 3) = "* Sup.Bill.No"
    frmpayment.PaymentGrid.ColAlignment(3) = 3
    
    frmpayment.PaymentGrid.ColWidth(4) = 2100
    frmpayment.PaymentGrid.TextMatrix(0, 4) = "Bill Amount (Rs)"
    
    frmpayment.PaymentGrid.ColWidth(5) = 2200
    frmpayment.PaymentGrid.TextMatrix(0, 5) = "Pay Amount (Rs)"
    
    frmpayment.PaymentGrid.ColWidth(6) = 0
    frmpayment.PaymentGrid.TextMatrix(0, 6) = "Inv.id"
    
    frmpayment.PaymentGrid.ColWidth(7) = 0
    frmpayment.PaymentGrid.TextMatrix(0, 7) = "Total Pay"
    
    frmpayment.DebtGrid.ColWidth(0) = 900
    frmpayment.DebtGrid.TextMatrix(0, 0) = "S.No"
    frmpayment.DebtGrid.ColAlignment(0) = 3
    
    frmpayment.DebtGrid.ColWidth(1) = 2400
    frmpayment.DebtGrid.TextMatrix(0, 1) = "Debit Reason"
    
    frmpayment.DebtGrid.ColWidth(2) = 1800
    frmpayment.DebtGrid.TextMatrix(0, 2) = "Debit Amount(Rs)"
    
End Sub
Sub frmpaymentmaingridload()
    frmpaymentmain.PaymentMainGrid.ColWidth(0) = 1000
    frmpaymentmain.PaymentMainGrid.TextMatrix(0, 0) = "S.No"
    frmpaymentmain.PaymentMainGrid.ColAlignment(0) = 3
    
    frmpaymentmain.PaymentMainGrid.ColWidth(1) = 1600
    frmpaymentmain.PaymentMainGrid.TextMatrix(0, 1) = " # Payment No"
    frmpaymentmain.PaymentMainGrid.ColAlignment(1) = 3
    
    frmpaymentmain.PaymentMainGrid.ColWidth(2) = 1600
    frmpaymentmain.PaymentMainGrid.TextMatrix(0, 2) = "Payment Date "
    frmpaymentmain.PaymentMainGrid.ColAlignment(2) = 3
    
    frmpaymentmain.PaymentMainGrid.ColWidth(3) = 1600
    frmpaymentmain.PaymentMainGrid.TextMatrix(0, 3) = " # Invoice No"
     frmpaymentmain.PaymentMainGrid.ColAlignment(3) = 3
    
    frmpaymentmain.PaymentMainGrid.ColWidth(4) = 2000
    frmpaymentmain.PaymentMainGrid.TextMatrix(0, 4) = " # Sup.Bill.No"
    frmpaymentmain.PaymentMainGrid.ColAlignment(4) = 3
    
    frmpaymentmain.PaymentMainGrid.ColWidth(5) = 3500
    frmpaymentmain.PaymentMainGrid.TextMatrix(0, 5) = " # Supplier Name"
    
    frmpaymentmain.PaymentMainGrid.ColWidth(6) = 2000
    frmpaymentmain.PaymentMainGrid.TextMatrix(0, 6) = " Net Amount (Rs)"
    
End Sub
Sub frmusersgridload()
    frmusers.loginGrid.ColWidth(0) = 900
    frmusers.loginGrid.TextMatrix(0, 0) = "S.No"
    frmusers.loginGrid.ColAlignment(0) = 3
    
    frmusers.loginGrid.ColWidth(1) = 2000
    frmusers.loginGrid.TextMatrix(0, 1) = "* User ID"
    frmusers.loginGrid.ColAlignment(1) = 3
    
    frmusers.loginGrid.ColWidth(2) = 2500
    frmusers.loginGrid.TextMatrix(0, 2) = "* User Name"
    
    frmusers.loginGrid.ColWidth(3) = 2500
    frmusers.loginGrid.TextMatrix(0, 3) = " * Password "
        
    frmusers.loginGrid.ColWidth(4) = 2500
    frmusers.loginGrid.TextMatrix(0, 4) = " * Date of Create "
    frmusers.loginGrid.ColAlignment(4) = 3
End Sub
Sub frmporeportingsload()
    frmporeporting.ReportGrid.ColWidth(0) = 900
    frmporeporting.ReportGrid.TextMatrix(0, 0) = "S.No"
    frmporeporting.ReportGrid.ColAlignment(0) = 3
    
    frmporeporting.ReportGrid.ColWidth(1) = 1800
    frmporeporting.ReportGrid.TextMatrix(0, 1) = " # PO No"
    frmporeporting.ReportGrid.ColAlignment(1) = 3
    
    frmporeporting.ReportGrid.ColWidth(2) = 1800
    frmporeporting.ReportGrid.TextMatrix(0, 2) = "Po Date"
    frmporeporting.ReportGrid.ColAlignment(2) = 3
    
    frmporeporting.ReportGrid.ColWidth(3) = 4000
    frmporeporting.ReportGrid.TextMatrix(0, 3) = "# Supplier Name"
    
    frmporeporting.ReportGrid.ColWidth(4) = 2300
    frmporeporting.ReportGrid.TextMatrix(0, 4) = "    Total PO.Qty "
   
    frmporeporting.ReportGrid.ColWidth(5) = 2300
    frmporeporting.ReportGrid.TextMatrix(0, 5) = "    Total Rec.Qty"
    
End Sub
Sub frmpurchaserepots()
    frmpurchasereporting.ReportGrid.ColWidth(0) = 900
    frmpurchasereporting.ReportGrid.TextMatrix(0, 0) = "S.No"
    frmpurchasereporting.ReportGrid.ColAlignment(0) = 3
    
    frmpurchasereporting.ReportGrid.ColWidth(1) = 1800
    frmpurchasereporting.ReportGrid.TextMatrix(0, 1) = " # PO No"
    frmpurchasereporting.ReportGrid.ColAlignment(1) = 3
    
    frmpurchasereporting.ReportGrid.ColWidth(2) = 1800
    frmpurchasereporting.ReportGrid.TextMatrix(0, 2) = "PO Date"
    frmpurchasereporting.ReportGrid.ColAlignment(2) = 3
    
    frmpurchasereporting.ReportGrid.ColWidth(3) = 4000
    frmpurchasereporting.ReportGrid.TextMatrix(0, 3) = " # Supplier Name"
    
    frmpurchasereporting.ReportGrid.ColWidth(4) = 3000
    frmpurchasereporting.ReportGrid.TextMatrix(0, 4) = " # Department"
   
    frmpurchasereporting.ReportGrid.ColWidth(5) = 2300
    frmpurchasereporting.ReportGrid.TextMatrix(0, 5) = "Net Amount (Rs)"
    
End Sub
