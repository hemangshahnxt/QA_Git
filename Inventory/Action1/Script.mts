Dim TestDBName : TestDBName = Environment.Value("DBUnderTest")

Dim IsPresent
Dim StartUpMethod
StartUpMethod = Environment.Value("StartUpMethod")


RunAction "LogIn [GlobalPracticeActions]", oneIteration, TestDBName

wait(2)

Window("regexpwndtitle:=Nextech.*","index:=0").WinMenu("menuobjtype:=2").Select "Modules;Inventory"

wait(2)


Call ChangeTabs (Window("regexpwndtitle:=Nextech.*","index:=0").Window("regexpwndtitle:=Inventory").ActiveX("acx_name:=NxTab Control"), "Product")
Call VerTabOpened ("Product")
Call VerifyProduct()

Call ChangeTabs (Window("regexpwndtitle:=Nextech.*","index:=0").Window("regexpwndtitle:=Inventory").ActiveX("acx_name:=NxTab Control"), "Order")
Call VerTabOpened("Order")
Call VerifyOrder()

Window("regexpwndtitle:=Nextech.*","index:=0").Window("regexpwndtitle:=Inventory").ActiveX("acx_name:=NxTab Control").Type micAltDwn + "l" + micAltUp
Call VerTabOpened("Allocation")
Call VerifyAllocation()

Call ChangeTabs (Window("regexpwndtitle:=Nextech.*","index:=0").Window("regexpwndtitle:=Inventory").ActiveX("acx_name:=NxTab Control"), "Overview")
Call VerTabOpened("Overview")
Call VerifyOverview()

Call ChangeTabs (Window("regexpwndtitle:=Nextech.*","index:=0").Window("regexpwndtitle:=Inventory").ActiveX("acx_name:=NxTab Control"), "Reports")
Call VerTabOpened("Reports")
Call VerifyReports()

Call ChangeTabs (Window("regexpwndtitle:=Nextech.*","index:=0").Window("regexpwndtitle:=Inventory").ActiveX("acx_name:=NxTab Control"), "Frames")
Call VerTabOpened("Frames")
Call VerifyFrames()

Call ChangeTabs (Window("regexpwndtitle:=Nextech.*","index:=0").Window("regexpwndtitle:=Inventory").ActiveX("acx_name:=NxTab Control"), "Optical Orders")
Call VerTabOpened("Optical Orders")
Call VerifyOpticalOrders()

Sub VerifyProduct()
	
Window("regexpwndtitle:=Nextech.*","index:=0").Window("regexpwndtitle:=Inventory").WinButton("regexpwndtitle:=New Item").Click

'Clicking on New Product WinMenu inside New Item
Window("regexpwndtitle:=Nextech.*","index:=0").WinMenu("menuobjtype:=3").Select "New Product"
If Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=New Item","index:=1").Exist Then
	Reporter.ReportEvent micPass, "New Item Dialog box is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "New Item Dialog Box is not Present"
End if

Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=New Item","index:=1").WinButton("regexpwndtitle:=Select").Click
If Dialog("regexpwndtitle:=Service / Inventory Categories").Exist Then
	Reporter.ReportEvent micPass, "Service/Inventory Categories Dialog is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Service/Inventory Categories Dialog is Not Present"
End If

Dialog("regexpwndtitle:=Service / Inventory Categories").WinButton("regexpwndtitle:=Add &New").Click

If Dialog("regexpwndtitle:=Input").WinButton("regexpwndtitle:=&Cancel").Exist Then
	Reporter.ReportEvent micPass, "Input Dialog box is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Input Dialog box is not Present"
End If
Dialog("regexpwndtitle:=Input").WinButton("regexpwndtitle:=&Cancel").Click
Dialog("regexpwndtitle:=Service / Inventory Categories").WinButton("regexpwndtitle:=Cancel").Click
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=New Item","index:=1").WinButton("regexpwndtitle:=Cancel").Click

'Clicking on the New Frame WinMenu Inside of New Item
Window("regexpwndtitle:=Nextech.*","index:=0").Window("regexpwndtitle:=Inventory").WinButton("regexpwndtitle:=New Item").Click
Set bc = CreateObject("Mercury.DeviceReplay")

Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Inventory").WinButton("regexpwndtitle:=New Item").Click

bc.PressKey 208
bc.PressKey 208
bc.PressKey 28
Set bc = Nothing

If Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Frames Data").Exist Then
	Reporter.ReportEvent micPass, "Frames Data Dialog Box is Present",""
Else	
	Reporter.ReportEvent micFail, "Test Failed", "Frames Data Dialog Box is not Present"
End If

'Clicking on the Select Button in Frames Data Dialog
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Frames Data").WinButton("regexpwndtitle:=Select").Click
If Dialog("regexpwndtitle:=Service / Inventory Categories").Exist Then
	Reporter.ReportEvent micPass, "Service/Inventory Categories Dialog is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Service/Inventory Categories Dialog is Not Present"
End If
Dialog("regexpwndtitle:=Service / Inventory Categories").WinButton("regexpwndtitle:=Add &New").Click

If Dialog("regexpwndtitle:=Input").WinButton("regexpwndtitle:=&Cancel").Exist Then
	Reporter.ReportEvent micPass, "Input Dialog box is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Input Dialog box is not Present"
End If
Dialog("regexpwndtitle:=Input").WinButton("regexpwndtitle:=&Cancel").Click
Dialog("regexpwndtitle:=Service / Inventory Categories").WinButton("regexpwndtitle:=Cancel").Click

Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Frames Data").WinButton("regexpwndtitle:=Cancel").Click
'

'Clicking on New Contact lens WinMenu inside the New items
Window("regexpwndtitle:=Nextech.*","index:=0").Window("regexpwndtitle:=Inventory").WinButton("regexpwndtitle:=New Item").Click
Window("regexpwndtitle:=Nextech.*","index:=0").WinMenu("menuobjtype:=3").Select "New Contact Lens"
If Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Contact Lens").Exist Then
	Reporter.ReportEvent micPass, "Contact lens Dialog box is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Contact lens Dialog box is not Present"	
End If

'Clciking on the Select Button on the Contact Lens Dialog box
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Contact Lens").WinButton("regexpwndtitle:=Select").Click
IsPresent = Dialog("regexpwndtitle:=Service / Inventory Categories").WinButton("regexpwndtitle:=Cancel").Exist
Call RecordResults (IsPresent, "Service / Inventory Categories Dialog box is Present")
Dialog("regexpwndtitle:=Service / Inventory Categories").WinButton("regexpwndtitle:=Add &New").Click

If Dialog("regexpwndtitle:=Input").WinButton("regexpwndtitle:=&Cancel").Exist Then
	Reporter.ReportEvent micPass, "Input Dialog box is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Input Dialog box is not Present"
End If
Dialog("regexpwndtitle:=Input").WinButton("regexpwndtitle:=&Cancel").Click
Dialog("regexpwndtitle:=Service / Inventory Categories").WinButton("regexpwndtitle:=Cancel").Click


'Clicking on the ellipsis Button on the Contact lens Dialog box
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Contact Lens").WinButton("regexpwndtitle:=\.\.\.","index:=1").Click
If Dialog("regexpwndtitle:=Edit Contact Lens Type").Exist Then
	Reporter.ReportEvent micPass, "Edit Contact Lens Type Dialog Box is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Edit Contact Lens Type Dialog Box is not Present"
End If
'Clicking on the Add Button
Dialog("regexpwndtitle:=Edit Contact Lens Type").WinButton("regexpwndtitle:=Add").Click
Dialog("regexpwndtitle:=Edit Contact Lens Type").WinButton("regexpwndtitle:=Edit").Click
Dialog("regexpwndtitle:=Edit Contact Lens Type").WinButton("regexpwndtitle:=Delete").Click

If Dialog("regexpwndtitle:=Edit Contact Lens Type").Dialog("regexpwndtitle:=Delete\?").Exist Then
	Dialog("regexpwndtitle:=Edit Contact Lens Type").Dialog("regexpwndtitle:=Delete\?").WinButton("regexpwndtitle:=&Yes").Click
End If
Dialog("regexpwndtitle:=Edit Contact Lens Type").WinButton("regexpwndtitle:=Close").Click

'Clicking on the 2nd ellipsis Button on the Contact Lens Dialog Box
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Contact Lens").WinButton("regexpwndtitle:=\.\.\.","index:=0").Click
If Dialog("regexpwndtitle:=Edit Contact Lens Manufacturer").Exist Then
	Reporter.ReportEvent micPass, "Edit Contact Lens Manufacturer Dialog is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Edit Contact lens manufacturer is not Present"
End If

'Clicking on the Add Button in the Edit Contact lens Manufacturer Dialog box
Dialog("regexpwndtitle:=Edit Contact Lens Manufacturer").WinButton("regexpwndtitle:=Add").Click
Dialog("regexpwndtitle:=Edit Contact Lens Manufacturer").WinButton("regexpwndtitle:=Edit").Click
Dialog("regexpwndtitle:=Edit Contact Lens Manufacturer").WinButton("regexpwndtitle:=Delete").Click

If Dialog("regexpwndtitle:=Edit Contact Lens Manufacturer").Dialog("regexpwndtitle:=Delete\?").Exist Then
	Dialog("regexpwndtitle:=Edit Contact Lens Manufacturer").Dialog("regexpwndtitle:=Delete\?").WinButton("regexpwndtitle:=&Yes").Click
End If
Dialog("regexpwndtitle:=Edit Contact Lens Manufacturer").WinButton("regexpwndtitle:=Close").Click
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Contact Lens").WinButton("regexpwndtitle:=Cancel").Click

'Clicking on Mark Inactive Button
Window("regexpwndtitle:=Nextech.*","index:=0").Window("regexpwndtitle:=Inventory").WinButton("regexpwndtitle:=Mark Inactive").Click
IsPresent = Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Deactivate Inventory Item").Exist
Call RecordResults(IsPresent, "Deactivate Inventory item Dialog box is Present")
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Deactivate Inventory Item").WinButton("regexpwndtitle:=&No").Click

'Clicking on the Delete Item Button
Window("regexpwndtitle:=Nextech.*","index:=0").Window("regexpwndtitle:=Inventory").WinButton("regexpwndtitle:=Delete Item").Click
IsPresent = Window("regexpwndtitle:=Nextech.*","index:=1").Dialog("regexpwndtitle:=Practice").Exist
Call RecordResults(IsPresent, "Nextech Practice Dialog is Present")
Window("regexpwndtitle:=Nextech.*","index:=1").Dialog("regexpwndtitle:=Nextech Practice").WinButton("regexpwndtitle:=&No").Click

 'Clicking on the Select Button 
Window("regexpwndtitle:=Nextech.*","index:=0").Window("regexpwndtitle:=Inventory").WinButton("regexpwndtitle:=Select").Click
If Window("regexpwndtitle:=Nextech.*","index:=0").Dialog( "regexpwndtitle:=Service / Inventory Categories").Exist Then
	Reporter.ReportEvent micPass, "Service/Inventory Categories Dialog box is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Service/Inventory Categories Dialog box is not Present"
End If

Window("regexpwndtitle:=Nextech.*","index:=0").Dialog( "regexpwndtitle:=Service / Inventory Categories").WinButton("regexpwndtitle:=Add &New").Click

If Dialog("regexpwndtitle:=Input").WinButton("regexpwndtitle:=&Cancel").Exist Then
	Reporter.ReportEvent micPass, "Input Dialog box is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Input Dialog box is not Present"
End If
Dialog("regexpwndtitle:=Input").WinButton("regexpwndtitle:=&Cancel").Click
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog( "regexpwndtitle:=Service / Inventory Categories").WinButton("regexpwndtitle:=Cancel").Click

'Clicking on the Remove Button
Window("regexpwndtitle:=Nextech.*","index:=0").Window("regexpwndtitle:=Inventory").WinButton("regexpwndtitle:=Remove").Click
IsPresent = Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Practice").Exist
Call RecordResults(IsPresent, "Nextech Practice Dialog is Present")
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Practice").WinButton("regexpwndtitle:=&No").Click

'Clicking on Configure Charge Level Providers
Window("regexpwndtitle:=Nextech.*","index:=0").Window("regexpwndtitle:=Inventory").WinButton("regexpwndtitle:=Configure Charge Level Providers").Click
If Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Configure Charge Level Providers","index:=1").Exist Then
	Reporter.ReportEvent micPass, "Configure Charge Level Dialog Box is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Configure Charge level Dialog is not Present"
End If
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Configure Charge Level Providers","index:=1").WinButton("regexpwndtitle:=Close").Click

'Clicking on Add Supplier
Window("regexpwndtitle:=Nextech.*","index:=0").Window("regexpwndtitle:=Inventory").WinButton("regexpwndtitle:=Add Supplier").Click
IsPresent = Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Choose A Supplier").Exist
Call RecordResults(IsPresent, "Choose a Supplier Dialog box is Present")

Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Choose A Supplier").WinButton("regexpwndtitle:=Add New").Click
If Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Create New Contact").Exist Then
	Reporter.ReportEvent micPass, "Create New Contact Dialog box is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Create New Contract Dialog box is not Present"
End If
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Create New Contact").WinButton("regexpwndtitle:=Cancel").Click
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Choose A Supplier").WinButton("regexpwndtitle:=Cancel").Click

'Clicking on Adv. Revenue Code Setup
Window("regexpwndtitle:=Nextech.*","index:=0").Window("regexpwndtitle:=Inventory").WinButton("regexpwndtitle:=Adv\. Revenue Code Setup").Click
IsPresent = Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Advanced Revenue Code Setup").Exist
Call RecordResults(IsPresent, "Advanced Rev Code Setup is Present")
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Advanced Revenue Code Setup").WinButton("regexpwndtitle:=Close").Click

'Clicking on the ellipsis Button next to the shop fee
Window("regexpwndtitle:=Nextech.*","index:=0").Window("regexpwndtitle:=Inventory").WinButton("regexpwndtitle:=\.\.\.").Click
IsPresent = Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Edit Shop Fees").Exist
Call RecordResults(IsPresent, "Edit Shop Fees Dialog box is Present")
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Edit Shop Fees").WinButton("regexpwndtitle:=Close").Click

'Clicking on NDC info Button
Window("regexpwndtitle:=Nextech.*","index:=0").Window("regexpwndtitle:=Inventory").WinButton("regexpwndtitle:=NDC Info").Click
Window("regexpwndtitle:=Nextech.*","index:=0").WinMenu("menuobjtype:=3").Select "Edit Claim Note..."
IsPresent = Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Input").Exist
Call RecordResults(IsPresent, "Input Dialog box is Present")
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Input").WinButton("regexpwndtitle:=&Cancel").Click

Window("regexpwndtitle:=Nextech.*","index:=0").Window("regexpwndtitle:=Inventory").WinButton("regexpwndtitle:=NDC Info").Click
Window("regexpwndtitle:=Nextech.*","index:=0").WinMenu("menuobjtype:=3").Select "NDC Defaults..."
IsPresent = Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=NDC Defaults").Exist
Call RecordResults(IsPresent, "NDC Defaults Dialog is Present")
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=NDC Defaults").WinButton("regexpwndtitle:=Cancel").Click

'Clicking on the Pending cases Button
Window("regexpwndtitle:=Nextech.*","index:=0").Window("regexpwndtitle:=Inventory").WinButton("regexpwndtitle:=Pending Cases").Click
IsPresent = Window("regexpwndtitle:=Nextech.*","index:=1").Dialog("regexpwndtitle:=Pending Case Histories").Exist
Call RecordResults(IsPresent, "Pending Cases Histories Dialog is Present")
Window("regexpwndtitle:=Nextech.*","index:=1").Dialog("regexpwndtitle:=Pending Case Histories").WinButton("regexpwndtitle:=Close").Click

'Clicking on view History Button
Window("regexpwndtitle:=Nextech.*","index:=0").Window("regexpwndtitle:=Inventory").WinButton("regexpwndtitle:=View History").Click
IsPresent = Window("regexpwndtitle:=Nextech.*","index:=0").Window("regexpwndtitle:=Product History").Exist
Call RecordResults(IsPresent, "Preview of Product History is Present")
Window("regexpwndtitle:=Nextech.*","index:=0").Window("regexpwndtitle:=Product History").Close

'Clicking on the Adjust Button
Window("regexpwndtitle:=Nextech.*","index:=0").Window("regexpwndtitle:=Inventory").WinButton("regexpwndtitle:=Adjust").Click
If Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Item Adjustment").Exist Then
	Reporter.ReportEvent micPass, "Item Adjustment Dialog Box is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Item Adjustment Dialog Box is not Present"
End If
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Item Adjustment").WinButton("regexpwndtitle:=\.\.\.").Click
IsPresent = Dialog("regexpwndtitle:=Product Adjustment Categories").WinButton("regexpwndtitle:=Add").Exist
Call RecordResults(IsPresent, "Product Adj Categories Dialog is Present")
Dialog("regexpwndtitle:=Product Adjustment Categories").WinButton("regexpwndtitle:=Add").Click

If Dialog("regexpwndtitle:=Input").WinButton("regexpwndtitle:=&Cancel").Exist Then
	Reporter.ReportEvent micPass, "Input Dialog box is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Input Dialog box is not Present"
End If

Dialog("regexpwndtitle:=Input").WinButton("regexpwndtitle:=&Cancel").Click
Dialog("regexpwndtitle:=Product Adjustment Categories").WinButton("regexpwndtitle:=Close").Click
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Item Adjustment").WinButton("regexpwndtitle:=Cancel").Click

'Clicking on Transfer Button
Window("regexpwndtitle:=Nextech.*","index:=0").Window("regexpwndtitle:=Inventory").WinButton("regexpwndtitle:=Transfer").Click
If Dialog("regexpwndtitle:=Nextech Practice").Exist Then
	Reporter.ReportEvent micPass, "Nextech Practice Dialog is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Nextech Practice Dialog is not Present"
End If
Dialog("regexpwndtitle:=Nextech Practice").WinButton("regexpwndtitle:=OK").Click

'Clicking on Inactive Inventory Items
Window("regexpwndtitle:=Nextech.*","index:=0").Window("regexpwndtitle:=Inventory").WinButton("regexpwndtitle:=Inactive Inventory Items").Click
IsPresent = Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Inactive Inventory Items","index:=1").Exist
Call RecordResults(IsPresent, "Inactive Inventory items Dialog box is Present")
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Inactive Inventory Items","index:=1").WinButton("regexpwndtitle:=Close").Click

'Clicking on Link Products To Services
Window("regexpwndtitle:=Nextech.*","index:=0").Window("regexpwndtitle:=Inventory").WinButton("regexpwndtitle:=Link Products To Services").Click
IsPresent = Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Link Products To Services","index:=1").Exist
Call RecordResults(IsPresent, "Link Products to Services Dialog is Present")
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Link Products To Services","index:=1").WinButton("regexpwndtitle:=Close").Click
End Sub

Sub VerifyOrder

'Working on the Order Tab on the Inventory Module
'Clicking on New Order Button
Window("regexpwndtitle:=Nextech.*","index:=0").Window("regexpwndtitle:=Inventory").WinButton("regexpwndtitle:=New Order").Click
If Window("regexpwndtitle:=Nextech Practice").WinButton("regexpwndtitle:=OK").Exist Then
	Window("regexpwndtitle:=Nextech Practice").WinButton("regexpwndtitle:=OK").Click
End If

Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Order").WinButton("regexpwndtitle:=\.\.\.","index:=0").Click
If Dialog("regexpwndtitle:=Edit Order Methods").Exist Then
	Reporter.ReportEvent micPass, "Edit Order Methods Dialog box is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Edit Order Methods Dialog box is not Present"
End If
Dialog("regexpwndtitle:=Edit Order Methods").WinButton("regexpwndtitle:=Add").Click
Dialog("regexpwndtitle:=Edit Order Methods").WinButton("regexpwndtitle:=Edit").Click
Dialog("regexpwndtitle:=Edit Order Methods").WinButton("regexpwndtitle:=Delete").Click
If Dialog("regexpwndtitle:=Edit Order Methods").Dialog("regexpwndtitle:=Delete\?").Exist Then
	Dialog("regexpwndtitle:=Edit Order Methods").Dialog("regexpwndtitle:=Delete\?").WinButton("regexpwndtitle:=&Yes").Click
End If
Dialog("regexpwndtitle:=Edit Order Methods").WinButton("regexpwndtitle:=Close").Click

'Clicking on Frame to Product
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Order").WinButton("regexpwndtitle:=Frame to Product").Click
If Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Order").Dialog("regexpwndtitle:=Convert To Product").Exist Then
	Reporter.ReportEvent micPass, "Convert To Product Dialog box is Open",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Convert To Product Dialog box is Not Open"
End If
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Order").Dialog("regexpwndtitle:=Convert To Product").WinButton("regexpwndtitle:=\.\.\.").Click
IsPresent = Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Order").Dialog("regexpwndtitle:=Convert To Product").Dialog("regexpwndtitle:=Select Frame to Convert").Exist
Call RecordResults(IsPresent, "Select Frame to Convert Dialog is Present")
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Order").Dialog("regexpwndtitle:=Convert To Product").Dialog("regexpwndtitle:=Select Frame to Convert").WinButton("regexpwndtitle:=Cancel").Click
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Order").Dialog("regexpwndtitle:=Convert To Product").WinButton("regexpwndtitle:=Cancel").Click


'Clicking on track via FedEx
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Order").WinButton("regexpwndtitle:=Track Via FedEx").Click
IsPresent = Window("regexpwndtitle:=Nextech.*","index:=0").WinButton("regexpwndtitle:=OK").Exist
Call RecordResults(IsPresent, "Nextech Practice Dialog is Present")
Window("regexpwndtitle:=Nextech.*","index:=0").WinButton("regexpwndtitle:=OK").Click

'Clicking on the second ellipsis button on the Order Dialog box
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Order").WinButton("regexpwndtitle:=\.\.\.","index:=1").Click
If Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Order").Dialog("regexpwndtitle:=Edit Charge Cards").Exist Then
	Reporter.ReportEvent micPass, "Edit Charge Cards Dialog box is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Edit Charge Cards Dialog box is not Present"
End If
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Order").Dialog("regexpwndtitle:=Edit Charge Cards").WinButton("regexpwndtitle:=&Add").Click
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Order").Dialog("regexpwndtitle:=Edit Charge Cards").ActiveX("acx_name:=NexTech DataList Control 2\.0").WinEdit("regexpwndclass:=Edit").Type "Diner"
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Order").Dialog("regexpwndtitle:=Edit Charge Cards").WinButton("regexpwndtitle:=&Add").Click
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Order").Dialog("regexpwndtitle:=Edit Charge Cards").WinButton("regexpwndtitle:=&Remove").Click
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Order").Dialog("regexpwndtitle:=Edit Charge Cards").WinButton("regexpwndtitle:=Cancel").Click

'Clicking on the third ellipsis Button on the Order Dialog box
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Order").WinButton("regexpwndtitle:=\.\.\.","index:=2").Click
If Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Order").Dialog("regexpwndtitle:=Edit Shipping Methods").Exist Then
	Reporter.ReportEvent micPass, "Edit Shipping Methods Dialog box is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Edit Shipping Methods Dialog box is not Present"
End If
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Order").Dialog("regexpwndtitle:=Edit Shipping Methods").WinButton("regexpwndtitle:=Add").Click
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Order").Dialog("regexpwndtitle:=Edit Shipping Methods").WinButton("regexpwndtitle:=Edit").Click
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Order").Dialog("regexpwndtitle:=Edit Shipping Methods").WinButton("regexpwndtitle:=Delete").Click
If Dialog("regexpwndtitle:=Edit Shipping Methods").Dialog("regexpwndtitle:=Delete\?").Exist Then
	Dialog("regexpwndtitle:=Edit Shipping Methods").Dialog("regexpwndtitle:=Delete\?").WinButton("regexpwndtitle:=&Yes").Click
End If
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Order").Dialog("regexpwndtitle:=Edit Shipping Methods").WinButton("regexpwndtitle:=Close").Click

'Clicking on Apply Discounts To Whole Order
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Order").WinButton("regexpwndtitle:=Apply Discounts To Whole Order").Click
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Order").WinMenu("menuobjtype:=3").Select "Apply a Percent Off"
IsPresent = Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Order").Dialog("regexpwndtitle:=Input").Exist
Call RecordResults(IsPresent, "Input Dialog box is Present")
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Order").Dialog("regexpwndtitle:=Input").WinButton("regexpwndtitle:=&Cancel").Click

'Clicking on the Aplly Discounts to Whole Order to check the 2nd WinMenu
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Order").WinButton("regexpwndtitle:=Apply Discounts To Whole Order").Click
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Order").WinMenu("menuobjtype:=3").Select "Apply a Discount Amount"
IsPresent = Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Order").Dialog("regexpwndtitle:=Input").Exist
Call RecordResults(IsPresent, "Input Dialog box is Present")
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Order").Dialog("regexpwndtitle:=Input").WinButton("regexpwndtitle:=&Cancel").Click

'Clicking on Preview Button
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Order").WinButton("regexpwndtitle:=Preview").Click
IsPresent = Window("regexpwndtitle:=Nextech.*","index:=0").WinButton("regexpwndtitle:=Cancel").Exist
Call RecordResults(IsPresent, "Nextech Practice Dialog is Present")
Window("regexpwndtitle:=Nextech.*","index:=0").WinButton("regexpwndtitle:=Cancel").Click
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Order").WinButton("regexpwndtitle:=Cancel").Click

'Clicking on Products Below Reorder Pt.
Window("regexpwndtitle:=Nextech.*","index:=0").Window("regexpwndtitle:=Inventory").WinButton("regexpwndtitle:=Products Below Reorder Pt\.").Click
wait(2)
IsPresent = Window("regexpwndtitle:=Nextech.*","index:=0").Window("regexpwndtitle:=Inventory Items to be Ordered").Exist
Call RecordResults(IsPresent, "Inventory to be Ordered preview is Present")
Window("regexpwndtitle:=Nextech.*","index:=0").Window("regexpwndtitle:=Inventory Items to be Ordered").Close

'Clicking on Allocations to be Ordered...
Window("regexpwndtitle:=Nextech.*","index:=0").Window("regexpwndtitle:=Inventory").WinButton("regexpwndtitle:=Allocations to be Ordered\.\.\.").Click
IsPresent = Window("regexpwndtitle:=Nextech.*","index:=0").WinButton("regexpwndtitle:=OK").Exist
Call RecordResults(IsPresent, "Nextech Practice Dialog is Present")
Window("regexpwndtitle:=Nextech.*","index:=0").WinButton("regexpwndtitle:=OK").Click

'Clicking on receive Frames
Window("regexpwndtitle:=Nextech.*","index:=0").Window("regexpwndtitle:=Inventory").WinButton("regexpwndtitle:=Receive Frames").Click
If Window("regexpwndtitle:=Nextech Practice").WinButton("regexpwndtitle:=OK").Exist Then
	Window("regexpwndtitle:=Nextech Practice").WinButton("regexpwndtitle:=OK").Click
End If


Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Frame Order.*").WinButton("regexpwndtitle:=\.\.\.","index:=0").Click
If Dialog("regexpwndtitle:=Edit Order Methods").Exist Then
	Reporter.ReportEvent micPass, "Edit Order Methods Dialog box is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Edit Order Methods Dialog box is not Present"
End If
Dialog("regexpwndtitle:=Edit Order Methods").WinButton("regexpwndtitle:=Add").Click
Dialog("regexpwndtitle:=Edit Order Methods").WinButton("regexpwndtitle:=Edit").Click
Dialog("regexpwndtitle:=Edit Order Methods").WinButton("regexpwndtitle:=Delete").Click
If Dialog("regexpwndtitle:=Edit Order Methods").Dialog("regexpwndtitle:=Delete\?").Exist Then
	Dialog("regexpwndtitle:=Edit Order Methods").Dialog("regexpwndtitle:=Delete\?").WinButton("regexpwndtitle:=&Yes").Click
End If
Dialog("regexpwndtitle:=Edit Order Methods").WinButton("regexpwndtitle:=Close").Click

'Clicking on Frame to Product
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Frame Order.*").WinButton("regexpwndtitle:=Frame to Product").Click

If Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Frame Order.*").Dialog("regexpwndtitle:=Convert To Product").Exist Then
	Reporter.ReportEvent micPass, "Convert To Product Dialog box is Open",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Convert To Product Dialog box is Not Open"
End If
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Frame Order.*").Dialog("regexpwndtitle:=Convert To Product").WinButton("regexpwndtitle:=\.\.\.").Click
IsPresent = Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Frame Order.*").Dialog("regexpwndtitle:=Convert To Product").Dialog("regexpwndtitle:=Select Frame to Convert").Exist
Call RecordResults(IsPresent, "Select From to Convert Dialog is Present")
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Frame Order.*").Dialog("regexpwndtitle:=Convert To Product").Dialog("regexpwndtitle:=Select Frame to Convert").WinButton("regexpwndtitle:=Cancel").Click
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Frame Order.*").Dialog("regexpwndtitle:=Convert To Product").WinButton("regexpwndtitle:=Cancel").Click



'Clicking on track via FedEx
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Frame Order.*").WinButton("regexpwndtitle:=Track Via FedEx").Click
IsPresent = Window("regexpwndtitle:=Nextech.*","index:=0").WinButton("regexpwndtitle:=OK").Exist
Call RecordResults(IsPresent, "Nextech Practice Dialog is Present")
Window("regexpwndtitle:=Nextech.*","index:=0").WinButton("regexpwndtitle:=OK").Click

'Clicking on the second ellipsis button on the Order Dialog box
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Frame Order.*").WinButton("regexpwndtitle:=\.\.\.","index:=1").Click
If Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Frame Order.*").Dialog("regexpwndtitle:=Edit Charge Cards").Exist Then
	Reporter.ReportEvent micPass, "Edit Charge Cards Dialog box is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Edit Charge Cards Dialog box is not Present"
End If
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Frame Order.*").Dialog("regexpwndtitle:=Edit Charge Cards").WinButton("regexpwndtitle:=&Add").Click
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Frame Order.*").Dialog("regexpwndtitle:=Edit Charge Cards").ActiveX("acx_name:=NexTech DataList Control 2\.0").WinEdit("regexpwndclass:=Edit").Type "Diner"
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Frame Order.*").Dialog("regexpwndtitle:=Edit Charge Cards").WinButton("regexpwndtitle:=&Add").Click
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Frame Order.*").Dialog("regexpwndtitle:=Edit Charge Cards").WinButton("regexpwndtitle:=&Remove").Click
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Frame Order.*").Dialog("regexpwndtitle:=Edit Charge Cards").WinButton("regexpwndtitle:=Cancel").Click

'Clicking on the third ellipsis Button on the Order Dialog box
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Frame Order.*").WinButton("regexpwndtitle:=\.\.\.","index:=2").Click
If Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Frame Order.*").Dialog("regexpwndtitle:=Edit Shipping Methods").Exist Then
	Reporter.ReportEvent micPass, "Edit Shipping Methods Dialog box is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Edit Shipping Methods Dialog box is not Present"
End If
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Frame Order.*").Dialog("regexpwndtitle:=Edit Shipping Methods").WinButton("regexpwndtitle:=Add").Click
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Frame Order.*").Dialog("regexpwndtitle:=Edit Shipping Methods").WinButton("regexpwndtitle:=Edit").Click
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Frame Order.*").Dialog("regexpwndtitle:=Edit Shipping Methods").WinButton("regexpwndtitle:=Delete").Click

If Dialog("regexpwndtitle:=Edit Shipping Methods").Dialog("regexpwndtitle:=Delete\?").Exist Then
	Dialog("regexpwndtitle:=Edit Shipping Methods").Dialog("regexpwndtitle:=Delete\?").WinButton("regexpwndtitle:=&Yes").Click
End If

Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Frame Order.*").Dialog("regexpwndtitle:=Edit Shipping Methods").WinButton("regexpwndtitle:=Close").Click

'Clicking on Apply Discounts To Whole Order
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Frame Order.*").WinButton("regexpwndtitle:=Apply Discounts To Whole Order").Click
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Frame Order.*").WinMenu("menuobjtype:=3").Select "Apply a Percent Off"
IsPresent = Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Frame Order.*").Dialog("regexpwndtitle:=Input").Exist
Call RecordResults(IsPresent, "Input Dialog box is Present")
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Frame Order.*").Dialog("regexpwndtitle:=Input").WinButton("regexpwndtitle:=&Cancel").Click

'Clicking on the Aplly Discounts to Whole Order to check the 2nd WinMenu
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Frame Order.*").WinButton("regexpwndtitle:=Apply Discounts To Whole Order").Click
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Frame Order.*").WinMenu("menuobjtype:=3").Select "Apply a Discount Amount"
IsPresent = Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Frame Order.*").Dialog("regexpwndtitle:=Input").Exist
Call RecordResults(IsPresent, "Input Dialog box is Present")
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Frame Order.*").Dialog("regexpwndtitle:=Input").WinButton("regexpwndtitle:=&Cancel").Click

'Clicking on Preview Button
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Frame Order.*").WinButton("regexpwndtitle:=Preview").Click
IsPresent = Window("regexpwndtitle:=Nextech.*","index:=0").WinButton("regexpwndtitle:=Cancel").Exist
Call RecordResults(IsPresent, "Nextech Practice Dialog is Present")
Window("regexpwndtitle:=Nextech.*","index:=0").WinButton("regexpwndtitle:=Cancel").Click
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Frame Order.*").WinButton("regexpwndtitle:=Cancel").Click

'Clicking on New Return Button
Window("regexpwndtitle:=Nextech.*","index:=0").Window("regexpwndtitle:=Inventory").WinButton("regexpwndtitle:=New Return").Click
If Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Supplier Product Return").Exist Then
	Reporter.ReportEvent micPass, "Supplier Product Return Dialog is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Supplier Product Return Dialog is not Present"
End If
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Supplier Product Return").WinButton("regexpwndtitle:=Track Via FedEx").Click
IsPresent = Window("regexpwndtitle:=Nextech.*","index:=0").WinButton("regexpwndtitle:=OK").Exist
Call RecordResults(IsPresent, "Nextech Practice Dialog is Present")
Window("regexpwndtitle:=Nextech.*","index:=0").WinButton("regexpwndtitle:=OK").Click

'Clicking on Configure Return Reasons...
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Supplier Product Return").WinButton("regexpwndtitle:=Configure Return Reasons\.\.\.").Click
If Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Supplier Product Return").Dialog("regexpwndtitle:=Edit Return Reasons").Exist Then
	Reporter.ReportEvent micPass, "Edit Return Reasons Dialog box is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Edit Return Reasons Dialog box is not Present"
End If
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Supplier Product Return").Dialog("regexpwndtitle:=Edit Return Reasons").WinButton("regexpwndtitle:=Add").Click
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Supplier Product Return").Dialog("regexpwndtitle:=Edit Return Reasons").WinButton("regexpwndtitle:=Edit").Click
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Supplier Product Return").Dialog("regexpwndtitle:=Edit Return Reasons").WinButton("regexpwndtitle:=Delete").Click
If Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Supplier Product Return").Dialog("regexpwndtitle:=Edit Return Reasons").Dialog("regexpwndtitle:=Delete\?").Exist Then
	Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Supplier Product Return").Dialog("regexpwndtitle:=Edit Return Reasons").Dialog("regexpwndtitle:=Delete\?").WinButton("regexpwndtitle:=&Yes").Click
End If

Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Supplier Product Return").Dialog("regexpwndtitle:=Edit Return Reasons").WinButton("regexpwndtitle:=Close").Click

'Clicking on the Configure Return Mehtods...
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Supplier Product Return").WinButton("regexpwndtitle:=Configure Return Methods\.\.\.").Click
If Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Supplier Product Return").Dialog("regexpwndtitle:=Edit Return Methods").Exist Then
	Reporter.ReportEvent micPass, "Edit Return Reasons Dialog box is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Edit Return Reasons Dialog box is not Present"
End If
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Supplier Product Return").Dialog("regexpwndtitle:=Edit Return Methods").WinButton("regexpwndtitle:=Add").Click
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Supplier Product Return").Dialog("regexpwndtitle:=Edit Return Methods").WinButton("regexpwndtitle:=Edit").Click
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Supplier Product Return").Dialog("regexpwndtitle:=Edit Return Methods").WinButton("regexpwndtitle:=Delete").Click
If Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Supplier Product Return").Dialog("regexpwndtitle:=Edit Return Methods").Dialog("regexpwndtitle:=Delete\?").Exist Then
	Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Supplier Product Return").Dialog("regexpwndtitle:=Edit Return Methods").Dialog("regexpwndtitle:=Delete\?").WinButton("regexpwndtitle:=&Yes").Click
End If
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Supplier Product Return").Dialog("regexpwndtitle:=Edit Return Methods").WinButton("regexpwndtitle:=Close").Click

'Clicking on the Preview Button on the Supplier Product Return
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Supplier Product Return").WinButton("regexpwndtitle:=Print Preview").Click
IsPresent = Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Supplier Product Return").Dialog("regexpwndtitle:=Inventory Returns").WinButton("regexpwndtitle:=Cancel").Exist
Call RecordResults(IsPresent, "Inventory Return Dialog is Present")
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Supplier Product Return").Dialog("regexpwndtitle:=Inventory Returns").WinButton("regexpwndtitle:=Cancel").Click

Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Supplier Product Return").WinButton("regexpwndtitle:=Cancel").Click
IsPresent = Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Supplier Product Return").Dialog("regexpwndtitle:=Practice").Exist
Call RecordResults(IsPresent, "Practice Dialog is Present")
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Supplier Product Return").Dialog("regexpwndtitle:=Practice").WinButton("regexpwndtitle:=&Yes").Click

'Clicking on Products to be Returned on Order Tab
Window("regexpwndtitle:=Nextech.*","index:=0").Window("regexpwndtitle:=Inventory").WinButton("regexpwndtitle:=Products to be Returned.*").Click
wait(1)
IsPresent = Window("regexpwndtitle:=Nextech.*","index:=0").WinButton("regexpwndtitle:=OK").Exist
Call RecordResults(IsPresent, "Nextech Practice Dialog is present")
Window("regexpwndtitle:=Nextech.*","index:=0").WinButton("regexpwndtitle:=OK").Click
wait(1)

'Clicking on Reconcile Consignment
Window("regexpwndtitle:=Nextech.*","index:=0").Window("regexpwndtitle:=Inventory").WinButton("regexpwndtitle:=Reconcile &Consignment").Click
If Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Consignment Reconciliation").Exist Then
	Reporter.ReportEvent micPass, "Consignment Reconciliation Dialog box is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Consignment Reconciliation Dialog box is not Present"
End If
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Consignment Reconciliation").WinButton("regexpwndtitle:=&Cancel").Click

End Sub

Sub VerifyAllocation()
	
If Window("regexpwndtitle:=Nextech.*","index:=0").WinButton("regexpwndtitle:=OK").Exist Then
	Window("regexpwndtitle:=Nextech.*","index:=0").WinButton("regexpwndtitle:=OK").Click
End If
'Allocation Tab on Inventory Module
'Clicking on Create New patient Allocation
Window("regexpwndtitle:=Nextech.*","index:=0").Window("regexpwndtitle:=Inventory").WinButton("regexpwndtitle:=Create New Patient Allocation").Click
If Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Patient Inventory Allocation").Exist Then
	Reporter.ReportEvent micPass, "Patient Inventory Allocation is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed","Patient Inventory Allocaiton is not Present"
End If
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Patient Inventory Allocation").WinButton("regexpwndtitle:=Preview").Click
IsPresent = Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Patient Inventory Allocation").Dialog("regexpwndtitle:=Practice").Exist
Call RecordResults(IsPresent, "Practice Dialog box is Present")
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Patient Inventory Allocation").Dialog("regexpwndtitle:=Practice").WinButton("regexpwndtitle:=&No").Click
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Patient Inventory Allocation").WinButton("regexpwndtitle:=Cancel").Click

'Clicking on Complete Allocations Button
Window("regexpwndtitle:=Nextech.*","index:=0").Window("regexpwndtitle:=Inventory").WinButton("regexpwndtitle:=Complete Allocations").Click
If Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Complete Allocations","index:=1").Exist Then
	Reporter.ReportEvent micPass, "Complete Allocation Dialog box is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Complete Allocation Dialog box is not Present"
End If
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Complete Allocations","index:=1").WinButton("regexpwndtitle:=Cancel").Click

'Clicking on Configure Required Allocations
Window("regexpwndtitle:=Nextech.*","index:=0").Window("regexpwndtitle:=Inventory").WinButton("regexpwndtitle:=Configure Required Allocations\.\.\.").Click
If Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Appointments Requiring Allocations").Exist Then
	Reporter.ReportEvent micPass, "Appointments Requiring Allocations Dialog box is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Appointments Requiring Allocations Dialog box is not Present"
End If
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Appointments Requiring Allocations").WinButton("regexpwndtitle:=&Add").Click
IsPresent = Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Appointments Requiring Allocations").Dialog("regexpwndtitle:=Required Allocation Detail").Exist
Call RecordResults(IsPresent, "Required Allocation Detail Dialog is Present")
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Appointments Requiring Allocations").Dialog("regexpwndtitle:=Required Allocation Detail").WinButton("regexpwndtitle:=Cancel").Click
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Appointments Requiring Allocations").WinButton("regexpwndtitle:=&Close").Click

End Sub

'Overview Tab in Inventory Module
Sub VerifyOverview()
	
End Sub

'Reports Tab in Inventory Module
Sub VerifyReports()
	

If Window("regexpwndtitle:=Nextech.*","index:=0").WinButton("regexpwndtitle:=OK").Exist Then
	Window("regexpwndtitle:=Nextech.*","index:=0").WinButton("regexpwndtitle:=OK").Click
End If

'Clicking on Consignment History by Date
Window("regexpwndtitle:=Nextech.*","index:=0").Window("regexpwndtitle:=Inventory").WinButton("regexpwndtitle:=Consignment History By Date").Click
If Window("regexpwndtitle:=Nextech.*","index:=0").Window("regexpwndtitle:=Reports").WinButton("regexpwndtitle:=New").Exist Then
	Reporter.ReportEvent micPass, "Reports Module is Open",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Report Module is not Open"
End If
'Clicking on the New Button
Window("regexpwndtitle:=Nextech.*","index:=0").Window("regexpwndtitle:=Reports").WinButton("regexpwndtitle:=New").Click
If Window("regexpwndtitle:=Nextech.*","index:=0").WinButton("regexpwndtitle:=&No").Exist Then
	Window("regexpwndtitle:=Nextech.*","index:=0").WinButton("regexpwndtitle:=&No").Click
End If
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Input").WinButton("regexpwndtitle:=&Cancel").Click

'Clicking on Save As Button
Window("regexpwndtitle:=Nextech.*","index:=0").Window("regexpwndtitle:=Reports").WinButton("regexpwndtitle:=Save as").Click
IsPresent = Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Input").Exist
Call RecordResults(IsPresent, "Input Dialog is Present")
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Input").WinButton("regexpwndtitle:=&Cancel").Click

'Clicking on the Search Button
Window("regexpwndtitle:=Nextech.*","index:=0").Window("regexpwndtitle:=Reports").WinButton("regexpwndtitle:=&Search").Click
If Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Search Report Descriptions").Exist Then
	Reporter.ReportEvent micPass, "Search Report Descriptions Dialog box is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Search Report Descriptions Dialog box is not Present"
End If
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Search Report Descriptions").WinButton("regexpwndtitle:=&Close").Click

'Clicking on the Help Button
Window("regexpwndtitle:=Nextech.*","index:=0").Window("regexpwndtitle:=Reports").WinButton("regexpwndtitle:=&Help").Click
IsPresent = Window("regexpwndtitle:=NexTech Practice User Manual").Exist
Call RecordResults(IsPresent, "Nextech Practice Usual Manual is Open")
Window("regexpwndtitle:=NexTech Practice User Manual").Close

'Clicking on Apply Above Filters To All
Window("regexpwndtitle:=Nextech.*","index:=0").Window("regexpwndtitle:=Reports").WinButton("regexpwndtitle:=Apply Above Filters To All").Click
If Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Practice").WinButton("regexpwndtitle:=No").Exist Then
	Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Practice").WinButton("regexpwndtitle:=No").Click
End If

'Clicking on Edit Report Button
Window("regexpwndtitle:=Nextech.*","index:=0").Window("regexpwndtitle:=Reports").WinButton("regexpwndtitle:=Edit Report").Click
If Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").Exist Then
	Reporter.ReportEvent micPass, "Pick the Report Dialog Box is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Pick the Report Dialog Box is not Present"
End If
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").WinButton("regexpwndtitle:=New").Click
IsPresent = Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").Dialog("regexpwndtitle:=Consignment History.*").Exist
Call RecordResults(IsPresent, "Consignment History by Date is Present")
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").Dialog("regexpwndtitle:=Consignment History.*").Close
IsPresent = Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").Dialog("regexpwndtitle:=Consignment History.*").Dialog("regexpwndtitle:=Save Custom Report").Exist
Call RecordResults(IsPresent, "Save Custom Report Dialog is Present")
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").Dialog("regexpwndtitle:=Consignment History.*").Dialog("regexpwndtitle:=Save Custom Report").WinButton("regexpwndtitle:=No").Click
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").WinButton("regexpwndtitle:=Delete").Click
IsPresent = Window("regexpwndtitle:=Nextech.*","index:=0").WinButton("regexpwndtitle:=OK").Exist
Call RecordResults(IsPresent, "Nextech Practice Dialog is Present")
Window("regexpwndtitle:=Nextech.*","index:=0").WinButton("regexpwndtitle:=OK").Click
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").WinButton("regexpwndtitle:=Close").Click

'Navigating back to the Inventory Module from Reports
Window("regexpwndtitle:=Nextech.*","index:=0").WinMenu("menuobjtype:=2").Select "Modules;Inventory"
If Window("regexpwndtitle:=Nextech.*","index:=0").Window("regexpwndtitle:=Inventory").Static("regexpwndtitle:=Consignment Reports").Exist Then
	Reporter.ReportEvent micPass, "Tab change to Inventory Module is confirmed",""
Else 
	Reporter.ReportEvent micFail, "Test Failed", "Tab change to Inventory module failed"
End If

'Clicking on Consignment List
Window("regexpwndtitle:=Nextech.*","index:=0").Window("regexpwndtitle:=Inventory").WinButton("regexpwndtitle:=Consignment List").Click
If Window("regexpwndtitle:=Nextech.*","index:=0").Window("regexpwndtitle:=Reports").WinButton("regexpwndtitle:=New").Exist Then
	Reporter.ReportEvent micPass, "Reports Module is Open",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Report Module is not Open"
End If
'Clicking on the New Button
Window("regexpwndtitle:=Nextech.*","index:=0").Window("regexpwndtitle:=Reports").WinButton("regexpwndtitle:=New").Click
If Window("regexpwndtitle:=Nextech.*","index:=0").WinButton("regexpwndtitle:=&No").Exist Then
	Window("regexpwndtitle:=Nextech.*","index:=0").WinButton("regexpwndtitle:=&No").Click
End If
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Input").WinButton("regexpwndtitle:=&Cancel").Click

'Clicking on Save As Button
Window("regexpwndtitle:=Nextech.*","index:=0").Window("regexpwndtitle:=Reports").WinButton("regexpwndtitle:=Save as").Click
IsPresent = Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Input").Exist
Call RecordResults(IsPresent, "Input Dialog is Present")
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Input").WinButton("regexpwndtitle:=&Cancel").Click

'Clicking on the Search Button
Window("regexpwndtitle:=Nextech.*","index:=0").Window("regexpwndtitle:=Reports").WinButton("regexpwndtitle:=&Search").Click
If Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Search Report Descriptions").Exist Then
	Reporter.ReportEvent micPass, "Search Report Descriptions Dialog box is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Search Report Descriptions Dialog box is not Present"
End If
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Search Report Descriptions").WinButton("regexpwndtitle:=&Close").Click

'Clicking on the Help Button
Window("regexpwndtitle:=Nextech.*","index:=0").Window("regexpwndtitle:=Reports").WinButton("regexpwndtitle:=&Help").Click
IsPresent = Window("regexpwndtitle:=NexTech Practice User Manual").Exist
Call RecordResults(IsPresent, "Nextech Practice User Manual window is Present")
Window("regexpwndtitle:=NexTech Practice User Manual").Close

'Clicking on Apply Above Filters To All
Window("regexpwndtitle:=Nextech.*","index:=0").Window("regexpwndtitle:=Reports").WinButton("regexpwndtitle:=Apply Above Filters To All").Click
If Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Practice").WinButton("regexpwndtitle:=No").Exist Then
	Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Practice").WinButton("regexpwndtitle:=No").Click
End If

'Clicking on Edit Report Button
Window("regexpwndtitle:=Nextech.*","index:=0").Window("regexpwndtitle:=Reports").WinButton("regexpwndtitle:=Edit Report").Click
If Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").Exist Then
	Reporter.ReportEvent micPass, "Pick the Report Dialog Box is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Pick the Report Dialog Box is not Present"
End If
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").WinButton("regexpwndtitle:=New").Click
IsPresent = Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").Dialog("regexpwndtitle:=Consignment List.*").Exist
Call RecordResults(IsPresent, "Consignment List window is Present")
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").Dialog("regexpwndtitle:=Consignment List.*").Close
IsPresent = Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").Dialog("regexpwndtitle:=Consignment List.*").Dialog("regexpwndtitle:=Save Custom Report").Exist
Call RecordResults(IsPresent, "Save Custom Report is Present")
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").Dialog("regexpwndtitle:=Consignment List.*").Dialog("regexpwndtitle:=Save Custom Report").WinButton("regexpwndtitle:=No").Click
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").WinButton("regexpwndtitle:=Delete").Click
IsPresent = Window("regexpwndtitle:=Nextech.*","index:=0").WinButton("regexpwndtitle:=OK").Exist
Call RecordResults(IsPresent, "Nextech Practice Dialog is Present")
Window("regexpwndtitle:=Nextech.*","index:=0").WinButton("regexpwndtitle:=OK").Click
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").WinButton("regexpwndtitle:=Close").Click

'Navigating back to the Inventory Module from Reports
Window("regexpwndtitle:=Nextech.*","index:=0").WinMenu("menuobjtype:=2").Select "Modules;Inventory"
If Window("regexpwndtitle:=Nextech.*","index:=0").Window("regexpwndtitle:=Inventory").Static("regexpwndtitle:=Consignment Reports").Exist Then
	Reporter.ReportEvent micPass, "Tab change to Inventory Module is confirmed",""
Else 
	Reporter.ReportEvent micFail, "Test Failed", "Tab change to Inventory module failed"
End If

'Clicking on Consignment Turn Rate by Month Button
Window("regexpwndtitle:=Nextech.*","index:=0").Window("regexpwndtitle:=Inventory").WinButton("regexpwndtitle:=Consignment Turn.*").Click

If Window("regexpwndtitle:=Nextech.*","index:=0").Window("regexpwndtitle:=Reports").WinButton("regexpwndtitle:=New").Exist Then
	Reporter.ReportEvent micPass, "Reports Module is Open",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Report Module is not Open"
End If
'Clicking on the New Button
Window("regexpwndtitle:=Nextech.*","index:=0").Window("regexpwndtitle:=Reports").WinButton("regexpwndtitle:=New").Click
If Window("regexpwndtitle:=Nextech.*","index:=0").WinButton("regexpwndtitle:=&No").Exist Then
	Window("regexpwndtitle:=Nextech.*","index:=0").WinButton("regexpwndtitle:=&No").Click
End If
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Input").WinButton("regexpwndtitle:=&Cancel").Click

'Clicking on Save As Button
Window("regexpwndtitle:=Nextech.*","index:=0").Window("regexpwndtitle:=Reports").WinButton("regexpwndtitle:=Save as").Click
IsPresent = Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Input").Exist
Call RecordResults(IsPresent, "Input Dialog is Present")
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Input").WinButton("regexpwndtitle:=&Cancel").Click

'Clicking on the Search Button
Window("regexpwndtitle:=Nextech.*","index:=0").Window("regexpwndtitle:=Reports").WinButton("regexpwndtitle:=&Search").Click
If Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Search Report Descriptions").Exist Then
	Reporter.ReportEvent micPass, "Search Report Descriptions Dialog box is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Search Report Descriptions Dialog box is not Present"
End If
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Search Report Descriptions").WinButton("regexpwndtitle:=&Close").Click

'Clicking on the Help Button
Window("regexpwndtitle:=Nextech.*","index:=0").Window("regexpwndtitle:=Reports").WinButton("regexpwndtitle:=&Help").Click
IsPresent = Window("regexpwndtitle:=NexTech Practice User Manual").Exist
Call RecordResults(IsPresent, "Nextech Practice User Manual is Present")
Window("regexpwndtitle:=NexTech Practice User Manual").Close

'Clicking on Apply Above Filters To All
Window("regexpwndtitle:=Nextech.*","index:=0").Window("regexpwndtitle:=Reports").WinButton("regexpwndtitle:=Apply Above Filters To All").Click
If Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Practice").WinButton("regexpwndtitle:=No").Exist Then
	Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Practice").WinButton("regexpwndtitle:=No").Click
End If

'Clicking on Edit Report Button
Window("regexpwndtitle:=Nextech.*","index:=0").Window("regexpwndtitle:=Reports").WinButton("regexpwndtitle:=Edit Report").Click
If Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").Exist Then
	Reporter.ReportEvent micPass, "Pick the Report Dialog Box is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Pick the Report Dialog Box is not Present"
End If
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").WinButton("regexpwndtitle:=New").Click
IsPresent = Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").Dialog("regexpwndtitle:=Consignment Turn.*").Exist
Call RecordResults(IsPresent, "Consignment Turn Rate by Month window is Present")
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").Dialog("regexpwndtitle:=Consignment Turn.*").Close

IsPresent = Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").Dialog("regexpwndtitle:=Consignment Turn.*").Dialog("regexpwndtitle:=Save Custom Report").Exist
Call RecordResults(IsPresent, "Save Custom Report Dialog is Present")

Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").Dialog("regexpwndtitle:=Consignment Turn.*").Dialog("regexpwndtitle:=Save Custom Report").WinButton("regexpwndtitle:=No").Click
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").WinButton("regexpwndtitle:=Delete").Click
IsPresent = Window("regexpwndtitle:=Nextech.*","index:=0").WinButton("regexpwndtitle:=OK").Exist
Call RecordResults(IsPresent, "Nextech Practice Dialog is Present")

Window("regexpwndtitle:=Nextech.*","index:=0").WinButton("regexpwndtitle:=OK").Click
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").WinButton("regexpwndtitle:=Close").Click

'Navigating back to the Inventory Module from Reports
Window("regexpwndtitle:=Nextech.*","index:=0").WinMenu("menuobjtype:=2").Select "Modules;Inventory"
If Window("regexpwndtitle:=Nextech.*","index:=0").Window("regexpwndtitle:=Inventory").Static("regexpwndtitle:=Consignment Reports").Exist Then
	Reporter.ReportEvent micPass, "Tab change to Inventory Module is confirmed",""
Else 
	Reporter.ReportEvent micFail, "Test Failed", "Tab change to Inventory module failed"
End If

'Clicking on Serial Number/Expirable Products By patients
Window("regexpwndtitle:=Nextech.*").Window("regexpwndtitle:=Inventory").WinButton("regexpwndtitle:=Serial Number.*","index:=0").Click
If Window("regexpwndtitle:=Nextech.*","index:=0").Window("regexpwndtitle:=Reports").WinButton("regexpwndtitle:=New").Exist Then
	Reporter.ReportEvent micPass, "Reports Module is Open",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Report Module is not Open"
End If

'Clicking on the New Button
Window("regexpwndtitle:=Nextech.*","index:=0").Window("regexpwndtitle:=Reports").WinButton("regexpwndtitle:=New").Click
If Window("regexpwndtitle:=Nextech.*","index:=0").WinButton("regexpwndtitle:=&No").Exist Then
	Window("regexpwndtitle:=Nextech.*","index:=0").WinButton("regexpwndtitle:=&No").Click
End If
IsPresent = Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Input").Exist
Call RecordResults(IsPresent, "Input Dialog is Present")
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Input").WinButton("regexpwndtitle:=&Cancel").Click

'Clicking on Save As Button
Window("regexpwndtitle:=Nextech.*","index:=0").Window("regexpwndtitle:=Reports").WinButton("regexpwndtitle:=Save as").Click
IsPresent = Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Input").Exist
Call RecordResults(IsPresent, "Input Dialog is Present")
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Input").WinButton("regexpwndtitle:=&Cancel").Click

'Clicking on the Search Button
Window("regexpwndtitle:=Nextech.*","index:=0").Window("regexpwndtitle:=Reports").WinButton("regexpwndtitle:=&Search").Click
If Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Search Report Descriptions").Exist Then
	Reporter.ReportEvent micPass, "Search Report Descriptions Dialog box is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Search Report Descriptions Dialog box is not Present"
End If
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Search Report Descriptions").WinButton("regexpwndtitle:=&Close").Click

'Clicking on the Help Button
Window("regexpwndtitle:=Nextech.*","index:=0").Window("regexpwndtitle:=Reports").WinButton("regexpwndtitle:=&Help").Click
IsPresent = Window("regexpwndtitle:=NexTech Practice User Manual").Exist
Call RecordResults(IsPresent, "Nextech Practice User Manual Window is Open")
Window("regexpwndtitle:=NexTech Practice User Manual").Close

'Clicking on Apply Above Filters To All
Window("regexpwndtitle:=Nextech.*","index:=0").Window("regexpwndtitle:=Reports").WinButton("regexpwndtitle:=Apply Above Filters To All").Click
If Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Practice").WinButton("regexpwndtitle:=No").Exist Then
	Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Practice").WinButton("regexpwndtitle:=No").Click
End If

'Clicking on Edit Report Button
Window("regexpwndtitle:=Nextech.*","index:=0").Window("regexpwndtitle:=Reports").WinButton("regexpwndtitle:=Edit Report").Click
If Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").Exist Then
	Reporter.ReportEvent micPass, "Pick the Report Dialog Box is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Pick the Report Dialog Box is not Present"
End If

Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").WinButton("regexpwndtitle:=New").Click
IsPresent = Window("regexpwndtitle:=Nextech.*","index:=0").WinButton("regexpwndtitle:=OK").Exist
Call RecordResults(IsPresent, "Nextech Practice dialog box is Present")
Window("regexpwndtitle:=Nextech.*","index:=0").WinButton("regexpwndtitle:=OK").Click
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").WinButton("regexpwndtitle:=Delete").Click

IsPresent = Window("regexpwndtitle:=Nextech.*","index:=0").WinButton("regexpwndtitle:=OK").Exist
Call RecordResults(IsPresent, "Nextech Practice dialog box is Present")
Window("regexpwndtitle:=Nextech.*","index:=0").WinButton("regexpwndtitle:=OK").Click
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").WinButton("regexpwndtitle:=Close").Click

' Clicking on Create Merge Group
Window("regexpwndtitle:=Nextech.*","index:=0").Window("regexpwndtitle:=Reports").WinButton("regexpwndtitle:=Create Merge Group").Click
If Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=Serial Numbered.*","index:=1").Exist Then
	Reporter.ReportEvent micPass, "Serial Numbered/Patient is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Serial Number/Patient is Not Present"
End If
Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=Serial Numbered.*","index:=1").WinButton("regexpwndtitle:=Cancel").Click

'Navigating back to the Inventory Module from Reports
Window("regexpwndtitle:=Nextech.*","index:=0").WinMenu("menuobjtype:=2").Select "Modules;Inventory"
If Window("regexpwndtitle:=Nextech.*","index:=0").Window("regexpwndtitle:=Inventory").Static("regexpwndtitle:=Consignment Reports").Exist Then
	Reporter.ReportEvent micPass, "Tab change to Inventory Module is confirmed",""
Else 
	Reporter.ReportEvent micFail, "Test Failed", "Tab change to Inventory module failed"
End If

'Clicking on Serial Number/Expirable Products in Stock
Window("regexpwndtitle:=Nextech.*").Window("regexpwndtitle:=Inventory").WinButton("regexpwndtitle:=Serial Number.*","index:=1").Click

If Window("regexpwndtitle:=Nextech.*","index:=0").Window("regexpwndtitle:=Reports").WinButton("regexpwndtitle:=New").Exist Then
	Reporter.ReportEvent micPass, "Reports Module is Open",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Report Module is not Open"
End If

'Clicking on the New Button
Window("regexpwndtitle:=Nextech.*","index:=0").Window("regexpwndtitle:=Reports").WinButton("regexpwndtitle:=New").Click
If Window("regexpwndtitle:=Nextech.*","index:=0").WinButton("regexpwndtitle:=&No").Exist Then
	Window("regexpwndtitle:=Nextech.*","index:=0").WinButton("regexpwndtitle:=&No").Click
End If
IsPresent = Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Input").Exist
Call RecordResults(IsPresent, "Input Dialog box is Present")
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Input").WinButton("regexpwndtitle:=&Cancel").Click

'Clicking on Save As Button
Window("regexpwndtitle:=Nextech.*","index:=0").Window("regexpwndtitle:=Reports").WinButton("regexpwndtitle:=Save as").Click
IsPresent = Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Input").Exist
Call RecordResults(IsPresent, "Input Dialog box is Present")
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Input").WinButton("regexpwndtitle:=&Cancel").Click

'Clicking on the Search Button
Window("regexpwndtitle:=Nextech.*","index:=0").Window("regexpwndtitle:=Reports").WinButton("regexpwndtitle:=&Search").Click
If Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Search Report Descriptions").Exist Then
	Reporter.ReportEvent micPass, "Search Report Descriptions Dialog box is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Search Report Descriptions Dialog box is not Present"
End If
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Search Report Descriptions").WinButton("regexpwndtitle:=&Close").Click

'Clicking on the Help Button
Window("regexpwndtitle:=Nextech.*","index:=0").Window("regexpwndtitle:=Reports").WinButton("regexpwndtitle:=&Help").Click
IsPresent = Window("regexpwndtitle:=NexTech Practice User Manual").Exist
Call RecordResults(IsPresent, "Nextech Practice user Manual window is Present")
Window("regexpwndtitle:=NexTech Practice User Manual").Close

'Clicking on Apply Above Filters To All
Window("regexpwndtitle:=Nextech.*","index:=0").Window("regexpwndtitle:=Reports").WinButton("regexpwndtitle:=Apply Above Filters To All").Click
If Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Practice").WinButton("regexpwndtitle:=No").Exist Then
	Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Practice").WinButton("regexpwndtitle:=No").Click
End If

'Clicking on Edit Report Button
Window("regexpwndtitle:=Nextech.*","index:=0").Window("regexpwndtitle:=Reports").WinButton("regexpwndtitle:=Edit Report").Click
If Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").Exist Then
	Reporter.ReportEvent micPass, "Pick the Report Dialog Box is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Pick the Report Dialog Box is not Present"
End If
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").WinButton("regexpwndtitle:=New").Click
IsPresent = Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").Dialog("regexpwndtitle:=Serial Numbered.*").Exist
Call RecordResults(IsPresent, "Serial Numbered/Expirable Products In stock window is Present")

Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").Dialog("regexpwndtitle:=Serial Numbered.*").Close
IsPresent = Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").Dialog("regexpwndtitle:=Serial Numbered.*").Dialog("regexpwndtitle:=Save Custom Report").Exist
Call RecordResults(IsPresent, "Save Custom Report Dialog is Present")
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").Dialog("regexpwndtitle:=Serial Numbered.*").Dialog("regexpwndtitle:=Save Custom Report").WinButton("regexpwndtitle:=No").Click

Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").WinButton("regexpwndtitle:=Delete").Click
IsPresent = Window("regexpwndtitle:=Nextech.*","index:=0").WinButton("regexpwndtitle:=OK").Exist
Call RecordResults(IsPresent, "Nextech Practice Dialog is Present")
Window("regexpwndtitle:=Nextech.*","index:=0").WinButton("regexpwndtitle:=OK").Click
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").WinButton("regexpwndtitle:=Close").Click

'Navigating back to the Inventory Module from Reports
Window("regexpwndtitle:=Nextech.*","index:=0").WinMenu("menuobjtype:=2").Select "Modules;Inventory"
If Window("regexpwndtitle:=Nextech.*","index:=0").Window("regexpwndtitle:=Inventory").Static("regexpwndtitle:=Consignment Reports").Exist Then
	Reporter.ReportEvent micPass, "Tab change to Inventory Module is confirmed",""
Else 
	Reporter.ReportEvent micFail, "Test Failed", "Tab change to Inventory module failed"
End If

'Clicking on Physical Inventory-Serialized-Tally Sheet
Window("regexpwndtitle:=Nextech.*").Window("regexpwndtitle:=Inventory").WinButton("regexpwndtitle:= Tally Sheet").Click

If Window("regexpwndtitle:=Nextech.*","index:=0").Window("regexpwndtitle:=Reports").WinButton("regexpwndtitle:=New").Exist Then
	Reporter.ReportEvent micPass, "Reports Module is Open",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Report Module is not Open"
End If

'Clicking on the New Button
Window("regexpwndtitle:=Nextech.*","index:=0").Window("regexpwndtitle:=Reports").WinButton("regexpwndtitle:=New").Click
If Window("regexpwndtitle:=Nextech.*","index:=0").WinButton("regexpwndtitle:=&No").Exist Then
	Window("regexpwndtitle:=Nextech.*","index:=0").WinButton("regexpwndtitle:=&No").Click
End If
IsPresent = Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Input").Exist
Call RecordResults(IsPresent, "Input Dialog box is Present")
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Input").WinButton("regexpwndtitle:=&Cancel").Click

'Clicking on Save As Button
Window("regexpwndtitle:=Nextech.*","index:=0").Window("regexpwndtitle:=Reports").WinButton("regexpwndtitle:=Save as").Click
IsPresent = Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Input").Exist
Call RecordResults(IsPresent, "Input Dialog box is Present")
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Input").WinButton("regexpwndtitle:=&Cancel").Click

'Clicking on the Search Button
Window("regexpwndtitle:=Nextech.*","index:=0").Window("regexpwndtitle:=Reports").WinButton("regexpwndtitle:=&Search").Click
If Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Search Report Descriptions").Exist Then
	Reporter.ReportEvent micPass, "Search Report Descriptions Dialog box is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Search Report Descriptions Dialog box is not Present"
End If
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Search Report Descriptions").WinButton("regexpwndtitle:=&Close").Click

'Clicking on the Help Button
Window("regexpwndtitle:=Nextech.*","index:=0").Window("regexpwndtitle:=Reports").WinButton("regexpwndtitle:=&Help").Click
IsPresent = Window("regexpwndtitle:=NexTech Practice User Manual").Exist
Call RecordResults(IsPresent, "Nextech practice User Manual window is Present")
Window("regexpwndtitle:=NexTech Practice User Manual").Close

'Clicking on Apply Above Filters To All
Window("regexpwndtitle:=Nextech.*","index:=0").Window("regexpwndtitle:=Reports").WinButton("regexpwndtitle:=Apply Above Filters To All").Click
If Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Practice").WinButton("regexpwndtitle:=No").Exist Then
	Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Practice").WinButton("regexpwndtitle:=No").Click
End If

'Clicking on Edit Report Button
Window("regexpwndtitle:=Nextech.*","index:=0").Window("regexpwndtitle:=Reports").WinButton("regexpwndtitle:=Edit Report").Click
If Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").Exist Then
	Reporter.ReportEvent micPass, "Pick the Report Dialog Box is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Pick the Report Dialog Box is not Present"
End If
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").WinButton("regexpwndtitle:=New").Click

IsPresent = Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").Dialog("regexpwndtitle:=Tally Sheet.*").Exist
Call RecordResults(IsPresent, "Physical Inventory Serialized by Tally sheet window is Present")
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").Dialog("regexpwndtitle:=Tally Sheet.*").Close

IsPresent = Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").Dialog("regexpwndtitle:=Tally Sheet.*").Dialog("regexpwndtitle:=Save Custom Report").Exist
Call RecordResults(IsPresent, "Save Custom Report Dialog is Present")
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").Dialog("regexpwndtitle:=Tally Sheet.*").Dialog("regexpwndtitle:=Save Custom Report").WinButton("regexpwndtitle:=No").Click

Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").WinButton("regexpwndtitle:=Delete").Click
IsPresent = Window("regexpwndtitle:=Nextech.*","index:=0").WinButton("regexpwndtitle:=OK").Exist
Call RecordResults(IsPresent, "Nextech Practice dialog is Present")
Window("regexpwndtitle:=Nextech.*","index:=0").WinButton("regexpwndtitle:=OK").Click
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").WinButton("regexpwndtitle:=Close").Click

'Navigating back to the Inventory Module from Reports
Window("regexpwndtitle:=Nextech.*","index:=0").WinMenu("menuobjtype:=2").Select "Modules;Inventory"
If Window("regexpwndtitle:=Nextech.*","index:=0").Window("regexpwndtitle:=Inventory").Static("regexpwndtitle:=Consignment Reports").Exist Then
	Reporter.ReportEvent micPass, "Tab change to Inventory Module is confirmed",""
Else 
	Reporter.ReportEvent micFail, "Test Failed", "Tab change to Inventory module failed"
End If

'Clicking on Allocation List in Inventory Module in Reports Tab
Window("regexpwndtitle:=Nextech.*").Window("regexpwndtitle:=Inventory").WinButton("regexpwndtitle:=Allocation List").Click
If Window("regexpwndtitle:=Nextech.*","index:=0").Window("regexpwndtitle:=Reports").WinButton("regexpwndtitle:=New").Exist Then
	Reporter.ReportEvent micPass, "Reports Module is Open",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Report Module is not Open"
End If

'Clicking on the New Button
Window("regexpwndtitle:=Nextech.*","index:=0").Window("regexpwndtitle:=Reports").WinButton("regexpwndtitle:=New").Click
If Window("regexpwndtitle:=Nextech.*","index:=0").WinButton("regexpwndtitle:=&No").Exist Then
	Window("regexpwndtitle:=Nextech.*","index:=0").WinButton("regexpwndtitle:=&No").Click
End If
IsPresent = Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Input").Exist
Call RecordResults(IsPresent, "Input Dialog is Present")
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Input").WinButton("regexpwndtitle:=&Cancel").Click

'Clicking on Save As Button
Window("regexpwndtitle:=Nextech.*","index:=0").Window("regexpwndtitle:=Reports").WinButton("regexpwndtitle:=Save as").Click
IsPresent = Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Input").Exist
Call RecordResults(IsPresent, "Input Dialog is Present")
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Input").WinButton("regexpwndtitle:=&Cancel").Click

'Clicking on the Search Button
Window("regexpwndtitle:=Nextech.*","index:=0").Window("regexpwndtitle:=Reports").WinButton("regexpwndtitle:=&Search").Click
If Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Search Report Descriptions").Exist Then
	Reporter.ReportEvent micPass, "Search Report Descriptions Dialog box is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Search Report Descriptions Dialog box is not Present"
End If
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Search Report Descriptions").WinButton("regexpwndtitle:=&Close").Click

'Clicking on the Help Button
Window("regexpwndtitle:=Nextech.*","index:=0").Window("regexpwndtitle:=Reports").WinButton("regexpwndtitle:=&Help").Click
IsPresent = Window("regexpwndtitle:=NexTech Practice User Manual").Exist
Call RecordResults(IsPresent, "Nextech Practice User Manual Window is Present")
Window("regexpwndtitle:=NexTech Practice User Manual").Close

'Clicking on Apply Above Filters To All
Window("regexpwndtitle:=Nextech.*","index:=0").Window("regexpwndtitle:=Reports").WinButton("regexpwndtitle:=Apply Above Filters To All").Click
If Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Practice").WinButton("regexpwndtitle:=No").Exist Then
	Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Practice").WinButton("regexpwndtitle:=No").Click
End If

'Clicking on Edit Report Button
Window("regexpwndtitle:=Nextech.*","index:=0").Window("regexpwndtitle:=Reports").WinButton("regexpwndtitle:=Edit Report").Click
If Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").Exist Then
	Reporter.ReportEvent micPass, "Pick the Report Dialog Box is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Pick the Report Dialog Box is not Present"
End If
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").WinButton("regexpwndtitle:=New").Click
IsPresent = Window("regexpwndtitle:=Nextech.*","index:=0").WinButton("regexpwndtitle:=OK").Exist
Call RecordResults(IsPresent, "Nextech Practice Dialog is Present")
Window("regexpwndtitle:=Nextech.*","index:=0").WinButton("regexpwndtitle:=OK").Click

Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").WinButton("regexpwndtitle:=Delete").Click
IsPresent = Window("regexpwndtitle:=Nextech.*","index:=0").WinButton("regexpwndtitle:=OK").Exist
Call RecordResults(IsPresent, "Nextech Practice Dialog is Present")
Window("regexpwndtitle:=Nextech.*","index:=0").WinButton("regexpwndtitle:=OK").Click
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").WinButton("regexpwndtitle:=Close").Click

' Clicking on Create Merge Group
Window("regexpwndtitle:=Nextech.*","index:=0").Window("regexpwndtitle:=Reports").WinButton("regexpwndtitle:=Create Merge Group").Click
IsPresent = Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Practice").WinButton("regexpwndtitle:=Cancel").Exist
Call RecordResults(IsPresent, "Practice Dialog box is Present")
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Practice").WinButton("regexpwndtitle:=Cancel").Click


'Navigating back to the Inventory Module from Reports
Window("regexpwndtitle:=Nextech.*","index:=0").WinMenu("menuobjtype:=2").Select "Modules;Inventory"
If Window("regexpwndtitle:=Nextech.*","index:=0").Window("regexpwndtitle:=Inventory").Static("regexpwndtitle:=Consignment Reports").Exist Then
	Reporter.ReportEvent micPass, "Tab change to Inventory Module is confirmed",""
Else 
	Reporter.ReportEvent micFail, "Test Failed", "Tab change to Inventory module failed"
End If

'Clicking on Go to the Reports Module
Window("regexpwndtitle:=Nextech.*").Window("regexpwndtitle:=Inventory").WinButton("regexpwndtitle:=Go to the.*").Click
If Window("regexpwndtitle:=Nextech.*","index:=0").Window("regexpwndtitle:=Reports").WinButton("regexpwndtitle:=New").Exist Then
	Reporter.ReportEvent micPass, "Reports Module is Open",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Report Module is not Open"
End If
'Clicking on the New Button
Window("regexpwndtitle:=Nextech.*","index:=0").Window("regexpwndtitle:=Reports").WinButton("regexpwndtitle:=New").Click
IsPresent = Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Input").Exist
Call RecordResults(IsPresent, "Input Dialog box is Present")
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Input").WinButton("regexpwndtitle:=&Cancel").Click

'Clicking on Save As Button
Window("regexpwndtitle:=Nextech.*","index:=0").Window("regexpwndtitle:=Reports").WinButton("regexpwndtitle:=Save as").Click
IsPresent = Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Input").Exist
Call RecordResults(IsPresent, "Input Dialog box is Present")
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Input").WinButton("regexpwndtitle:=&Cancel").Click

'Clicking on the Search Button
Window("regexpwndtitle:=Nextech.*","index:=0").Window("regexpwndtitle:=Reports").WinButton("regexpwndtitle:=&Search").Click
If Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Search Report Descriptions").Exist Then
	Reporter.ReportEvent micPass, "Search Report Descriptions Dialog box is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Search Report Descriptions Dialog box is not Present"
End If
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Search Report Descriptions").WinButton("regexpwndtitle:=&Close").Click

'Clicking on the Help Button
Window("regexpwndtitle:=Nextech.*","index:=0").Window("regexpwndtitle:=Reports").WinButton("regexpwndtitle:=&Help").Click
IsPresent = Window("regexpwndtitle:=NexTech Practice User Manual").Exist
Call RecordResults(IsPresent, "Nextech Practice User Manual window is Present")
Window("regexpwndtitle:=NexTech Practice User Manual").Close

'Clicking on Apply Above Filters To All
Window("regexpwndtitle:=Nextech.*","index:=0").Window("regexpwndtitle:=Reports").WinButton("regexpwndtitle:=Apply Above Filters To All").Click
If Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Practice").WinButton("regexpwndtitle:=No").Exist Then
	Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Practice").WinButton("regexpwndtitle:=No").Click
End If

'Clicking on Edit Report Button
Window("regexpwndtitle:=Nextech.*","index:=0").Window("regexpwndtitle:=Reports").WinButton("regexpwndtitle:=Edit Report").Click
If Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").Exist Then
	Reporter.ReportEvent micPass, "Pick the Report Dialog Box is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Pick the Report Dialog Box is not Present"
End If
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").WinButton("regexpwndtitle:=New").Click
IsPresent = Window("regexpwndtitle:=Nextech.*","index:=0").WinButton("regexpwndtitle:=OK").Exist
Call RecordResults(IsPresent, "Nextech Practice Dialog is Present")
Window("regexpwndtitle:=Nextech.*","index:=0").WinButton("regexpwndtitle:=OK").Click

Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").WinButton("regexpwndtitle:=Delete").Click
IsPresent = Window("regexpwndtitle:=Nextech.*","index:=0").WinButton("regexpwndtitle:=OK").Exist
Call RecordResults(IsPresent, "Nextech Practice Dialog is Present")
Window("regexpwndtitle:=Nextech.*","index:=0").WinButton("regexpwndtitle:=OK").Click
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").WinButton("regexpwndtitle:=Close").Click

' Clicking on Create Merge Group
Window("regexpwndtitle:=Nextech.*","index:=0").Window("regexpwndtitle:=Reports").WinButton("regexpwndtitle:=Create Merge Group").Click
IsPresent = Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Practice").WinButton("regexpwndtitle:=Cancel").Exist
Call RecordResults(IsPresent, "Practice Dialog box is Present")
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Practice").WinButton("regexpwndtitle:=Cancel").Click


'Navigating back to the Inventory Module from Reports
Window("regexpwndtitle:=Nextech.*","index:=0").WinMenu("menuobjtype:=2").Select "Modules;Inventory"
If Window("regexpwndtitle:=Nextech.*","index:=0").Window("regexpwndtitle:=Inventory").Static("regexpwndtitle:=Consignment Reports").Exist Then
	Reporter.ReportEvent micPass, "Tab change to Inventory Module is confirmed",""
Else 
	Reporter.ReportEvent micFail, "Test Failed", "Tab change to Inventory module failed"
End If

End Sub

'Frames Tab on Inventory Module
'Clicking on Create Inventory Products For Selected on the Frames Tab
Sub VerifyFrames()
	
Window("regexpwndtitle:=Nextech.*","index:=0").Window("regexpwndtitle:=Inventory").WinButton("regexpwndtitle:=Create Inventory.*").Click
IsPresent = Window("regexpwndtitle:=Nextech.*","index:=0").WinButton("regexpwndtitle:=OK").Exist
Call RecordResults(IsPresent, "Nextech Practice Dialog is Present")
Window("regexpwndtitle:=Nextech.*","index:=0").WinButton("regexpwndtitle:=OK").Click

'Clicking on Apply Markup to All
Window("regexpwndtitle:=Nextech.*","index:=0").Window("regexpwndtitle:=Inventory").WinButton("regexpwndtitle:=Apply Markup to All").Click
IsPresent = Window("regexpwndtitle:=Nextech.*","index:=0").WinButton("regexpwndtitle:=OK").Exist
Call RecordResults(IsPresent, "Nextech Practice Dialog is Present")
Window("regexpwndtitle:=Nextech.*","index:=0").WinButton("regexpwndtitle:=OK").Click

'Clicking on Import Updated Frames Data
Window("regexpwndtitle:=Nextech.*","index:=0").Window("regexpwndtitle:=Inventory").WinButton("regexpwndtitle:=Import Updated Frames Data").Click
If Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Browse For Folder").Exist Then
	Reporter.ReportEvent micPass, "Browse for Folder Dialog is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Browse for Folder Dialog is not Present"
End If
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Browse For Folder").WinButton("regexpwndtitle:=Cancel").Click

'Clicking on Update All Existing Frames Products
Window("regexpwndtitle:=Nextech.*","index:=0").Window("regexpwndtitle:=Inventory").WinButton("regexpwndtitle:=Update All.*").Click
IsPresent = Window("regexpwndtitle:=Nextech.*","index:=0").WinButton("regexpwndtitle:=OK").Exist
Call RecordResults(IsPresent, "Nextech Practice Dialog is Present")
Window("regexpwndtitle:=Nextech.*","index:=0").WinButton("regexpwndtitle:=OK").Click

'Clicking on Run Price Change Report
Window("regexpwndtitle:=Nextech.*","index:=0").Window("regexpwndtitle:=Inventory").WinButton("regexpwndtitle:=Run Price.*").Click
IsPresent = Window("regexpwndtitle:=Nextech.*","index:=0").WinButton("regexpwndtitle:=OK").Exist
Call RecordResults(IsPresent, "Nextech Practice Dialog is Present")
Window("regexpwndtitle:=Nextech.*","index:=0").WinButton("regexpwndtitle:=OK").Click

'Clicking on the Options Button
Window("regexpwndtitle:=Nextech.*","index:=0").Window("regexpwndtitle:=Inventory").WinButton("regexpwndtitle:=Options").Click
If Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Frames Options").Exist Then
	Reporter.ReportEvent micPass, "Frames Options Dialog Box is Present", ""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Frames Options Dialog Box is not Present"
End If

Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Frames Options").WinButton("regexpwndtitle:=Cancel").Click

'Clicking on Edit Markups Button
Window("regexpwndtitle:=Nextech.*","index:=0").Window("regexpwndtitle:=Inventory").WinButton("regexpwndtitle:=Edit Markups").Click
If Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Edit Markup Formulas").Exist Then
	Reporter.ReportEvent micPass, "Edit Markup Formulas Dialog box is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Edit Markup Formulas Dialog box is not Present"
End If
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Edit Markup Formulas").WinButton("regexpwndtitle:=New").Click
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Edit Markup Formulas").WinButton("regexpwndtitle:=Delete").Click
IsPresent = Window("regexpwndtitle:=Nextech.*","index:=0").WinButton("regexpwndtitle:=&Yes").Exist
Call RecordResults(IsPresent, "Nextech Practice Dialog box is Present")
Window("regexpwndtitle:=Nextech.*","index:=0").WinButton("regexpwndtitle:=&Yes").Click

Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Edit Markup Formulas").WinButton("regexpwndtitle:=\?","index:=0").Click
IsPresent = Window("regexpwndtitle:=Nextech.*","index:=0").WinButton("regexpwndtitle:=OK").Exist
Call RecordResults(IsPresent, "Nextech Practice Dialog for markup Formula is Present")
Window("regexpwndtitle:=Nextech.*","index:=0").WinButton("regexpwndtitle:=OK").Click

Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Edit Markup Formulas").WinButton("regexpwndtitle:=\?","index:=1").Click
IsPresent = Window("regexpwndtitle:=Nextech.*","index:=0").WinButton("regexpwndtitle:=OK").Exist
Call RecordResults(IsPresent, "Nextech Practice Dialog for Round Sales Price is Present")
Window("regexpwndtitle:=Nextech.*","index:=0").WinButton("regexpwndtitle:=OK").Click
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Edit Markup Formulas").WinButton("regexpwndtitle:=Cancel").Click

End Sub

'Optical Orders Tab on Inventory Module
Sub VerifyOpticalOrders()
	
'Clicking on New Order Button on the Optical Orders Tab
Window("regexpwndtitle:=Nextech.*","index:=0").Window("regexpwndtitle:=Inventory").WinButton("regexpwndtitle:=New Order").Click
Window("regexpwndtitle:=Nextech.*","index:=0").WinMenu("menuobjtype:=3").Select "Glasses Order"
If Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Glasses Order").Exist Then
	Reporter.ReportEvent micPass, "Glasses Order Dialog Box is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Glasses Order Dialog box is not Present"
End If

'Clicking on Glasses Catalog Setup inside the Glasses Order Dialog box
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Glasses Order").WinButton("regexpwndtitle:=Glasses Catalog Setup\.\.\.").Click
If Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Glasses Order").Dialog("regexpwndtitle:=Glasses Catalog Setup").Exist Then
	Reporter.ReportEvent micPass, "Glasses Catalog Setup Dialog box is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Glasses Catalog Setup Dialog box is not Present"
End If
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Glasses Order").Dialog("regexpwndtitle:=Glasses Catalog Setup").WinButton("regexpwndtitle:=Billing Setup\.\.\.").Click
IsPresent = Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Glasses Order").Dialog("regexpwndtitle:=Glasses Catalog Setup").Dialog("regexpwndtitle:=Glasses Catalog Billing Configuration").Exist
Call RecordResults(IsPresent, "Glasses Catalog Billing Configuration Dialog is Present")
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Glasses Order").Dialog("regexpwndtitle:=Glasses Catalog Setup").Dialog("regexpwndtitle:=Glasses Catalog Billing Configuration").WinButton("regexpwndtitle:=Cancel").Click

'Clicking on Check VisionWeb Catalog Update inside Glasses Catalog setup
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Glasses Order").Dialog("regexpwndtitle:=Glasses Catalog Setup").WinButton("regexpwndtitle:=Check VisionWeb.*").Click
IsPresent = Window("regexpwndtitle:=Nextech.*","index:=0").WinButton("regexpwndtitle:=OK").Exist
Call RecordResults(IsPresent, "Nextech Practice dialog is Present")
Window("regexpwndtitle:=Nextech.*","index:=0").WinButton("regexpwndtitle:=OK").Click

IsPresent = Window("regexpwndtitle:=Nextech.*","index:=0").WinButton("regexpwndtitle:=OK").Exist
Call RecordResults(IsPresent, "Nextech Practice dialog is Present")
Window("regexpwndtitle:=Nextech.*","index:=0").WinButton("regexpwndtitle:=OK").Click

'Clicking on Copy from... inside the Glasses Catalog Setup
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Glasses Order").Dialog("regexpwndtitle:=Glasses Catalog Setup").WinButton("regexpwndtitle:=Copy From\.\.\.").Click
IsPresent = Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Glasses Order").Dialog("regexpwndtitle:=Glasses Catalog Setup").Dialog("regexpwndtitle:=Select a Supplier").Exist
Call RecordResults(IsPresent, "Select a Supplier Dialog is Present")
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Glasses Order").Dialog("regexpwndtitle:=Glasses Catalog Setup").Dialog("regexpwndtitle:=Select a Supplier").WinButton("regexpwndtitle:=Cancel").Click

'Clicking on Add Designs inside the Glasses Catalog Setup
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Glasses Order").Dialog("regexpwndtitle:=Glasses Catalog Setup").WinButton("regexpwndtitle:=Add Designs").Click
IsPresent = Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Glasses Order").Dialog("regexpwndtitle:=Glasses Catalog Setup").Dialog("regexpwndtitle:=Select one or.*","Index:=0").Exist
Call RecordResults(IsPresent, "Select one or more Design dialog is Present")

'Clicking on the Add Button
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Glasses Order").Dialog("regexpwndtitle:=Glasses Catalog Setup").Dialog("regexpwndtitle:=Select one or.*","index:=0").WinButton("regexpwndtitle:=Add").Click
IsPresent = Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Glasses Order").Dialog("regexpwndtitle:=Glasses Catalog Setup").Dialog("regexpwndtitle:=Select one or.*","index:=0").Dialog("regexpwndtitle:=Add Design").Exist
Call RecordResults(IsPresent, "Add design Dialog box is Present")
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Glasses Order").Dialog("regexpwndtitle:=Glasses Catalog Setup").Dialog("regexpwndtitle:=Select one or.*","index:=0").Dialog("regexpwndtitle:=Add Design").WinButton("regexpwndtitle:=Cancel").Click
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Glasses Order").Dialog("regexpwndtitle:=Glasses Catalog Setup").Dialog("regexpwndtitle:=Select one or.*","index:=0").WinButton("regexpwndtitle:=Cancel").Click

'Clicking on the Add Material Button in the Glasses Catalog setup
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Glasses Order").Dialog("regexpwndtitle:=Glasses Catalog Setup").WinButton("regexpwndtitle:=Add Materials").Click
IsPresent = Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Glasses Order").Dialog("regexpwndtitle:=Glasses Catalog Setup").Dialog("regexpwndtitle:=Select one or.*","Index:=1").Exist
Call RecordResults(IsPresent, "Select one or more materials Dialog is Present")
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Glasses Order").Dialog("regexpwndtitle:=Glasses Catalog Setup").Dialog("regexpwndtitle:=Select one or.*","Index:=1").WinButton("regexpwndtitle:=Add").Click
IsPresent = Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Glasses Order").Dialog("regexpwndtitle:=Glasses Catalog Setup").Dialog("regexpwndtitle:=Select one or.*","index:=1").Dialog("regexpwndtitle:=Add Material").Exist
Call RecordResults(IsPresent, "Add material Dialog box is Present")
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Glasses Order").Dialog("regexpwndtitle:=Glasses Catalog Setup").Dialog("regexpwndtitle:=Select one or.*","index:=1").Dialog("regexpwndtitle:=Add Material").WinButton("regexpwndtitle:=Cancel").Click
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Glasses Order").Dialog("regexpwndtitle:=Glasses Catalog Setup").Dialog("regexpwndtitle:=Select one or.*","index:=1").WinButton("regexpwndtitle:=Cancel").Click

'Clicking on Add Treatnments button in the Glasses Catalog Setup
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Glasses Order").Dialog("regexpwndtitle:=Glasses Catalog Setup").WinButton("regexpwndtitle:=Add Treatments").Click
IsPresent = Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Glasses Order").Dialog("regexpwndtitle:=Glasses Catalog Setup").Dialog("regexpwndtitle:=Select one or.*","Index:=2").Exist
Call RecordResults(IsPresent, "Select one or more materials Dialog is Present")
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Glasses Order").Dialog("regexpwndtitle:=Glasses Catalog Setup").Dialog("regexpwndtitle:=Select one or.*","Index:=2").WinButton("regexpwndtitle:=Add").Click
IsPresent = Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Glasses Order").Dialog("regexpwndtitle:=Glasses Catalog Setup").Dialog("regexpwndtitle:=Select one or.*","index:=2").Dialog("regexpwndtitle:=Add Treatment").Exist
Call RecordResults(IsPresent, "Add material Dialog box is Present")
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Glasses Order").Dialog("regexpwndtitle:=Glasses Catalog Setup").Dialog("regexpwndtitle:=Select one or.*","index:=2").Dialog("regexpwndtitle:=Add Treatment").WinButton("regexpwndtitle:=Cancel").Click
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Glasses Order").Dialog("regexpwndtitle:=Glasses Catalog Setup").Dialog("regexpwndtitle:=Select one or.*","index:=2").WinButton("regexpwndtitle:=Cancel").Click

'Clicking on Add Frame Types Button in the Glasses Catalog Setup
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Glasses Order").Dialog("regexpwndtitle:=Glasses Catalog Setup").WinButton("regexpwndtitle:=Add Frame Types").Click
IsPresent = Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Glasses Order").Dialog("regexpwndtitle:=Glasses Catalog Setup").Dialog("regexpwndtitle:=Select one or.*","Index:=3").Exist
Call RecordResults(IsPresent, "Select one or more materials Dialog is Present")
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Glasses Order").Dialog("regexpwndtitle:=Glasses Catalog Setup").Dialog("regexpwndtitle:=Select one or.*","Index:=3").WinButton("regexpwndtitle:=Add").Click
IsPresent = Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Glasses Order").Dialog("regexpwndtitle:=Glasses Catalog Setup").Dialog("regexpwndtitle:=Select one or.*","index:=3").Dialog("regexpwndtitle:=Add Frame Type").Exist
Call RecordResults(IsPresent, "Add material Dialog box is Present")
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Glasses Order").Dialog("regexpwndtitle:=Glasses Catalog Setup").Dialog("regexpwndtitle:=Select one or.*","index:=3").Dialog("regexpwndtitle:=Add Frame Type").WinButton("regexpwndtitle:=Cancel").Click
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Glasses Order").Dialog("regexpwndtitle:=Glasses Catalog Setup").Dialog("regexpwndtitle:=Select one or.*","index:=3").WinButton("regexpwndtitle:=Cancel").Click

Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Glasses Order").Dialog("regexpwndtitle:=Glasses Catalog Setup").WinButton("regexpwndtitle:=Close").Click

'Clicking on Save and Print Preview Order inside the Glasses Order Dialog
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Glasses Order").WinButton("regexpwndtitle:=Save And Print Preview Order").Click
IsPresent = Window("regexpwndtitle:=Nextech.*","index:=0").WinButton("regexpwndtitle:=OK").Exist
Call RecordResults(IsPresent, "Nextech Practice Dialog is Present")
Window("regexpwndtitle:=Nextech.*","index:=0").WinButton("regexpwndtitle:=OK").Click

'Clicking on Save and Print Preview Rx inside the Glasses order Dialog
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Glasses Order").WinButton("regexpwndtitle:=Save And Print Preview Rx").Click
IsPresent = Window("regexpwndtitle:=Nextech.*","index:=0").WinButton("regexpwndtitle:=OK").Exist
Call RecordResults(IsPresent, "Nextech Practice Dialog is Present")
Window("regexpwndtitle:=Nextech.*","index:=0").WinButton("regexpwndtitle:=OK").Click

Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Glasses Order").WinButton("regexpwndtitle:=Cancel").Click

'Clicking on New Order again to verify the Contact lens Order
Window("regexpwndtitle:=Nextech.*","index:=0").Window("regexpwndtitle:=Inventory").WinButton("regexpwndtitle:=New Order").Click
Window("regexpwndtitle:=Nextech.*","index:=0").WinMenu("menuobjtype:=3").Select "Contact Lens Order"
If Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Contact Lens Order").Exist Then
	Reporter.ReportEvent micPass, "Contact lens Order Dialog box is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Contact lens Order Dialog box is not Present"
End If
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Contact Lens Order").WinButton("regexpwndtitle:=Save And Print Preview Order").Click
IsPresent = Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Contact Lens Order").Dialog("regexpwndtitle:=Alert!").Exist
Call RecordResults(IsPresent, "Alert Message Dialog box is Present")
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Contact Lens Order").Dialog("regexpwndtitle:=Alert!").WinButton("regexpwndtitle:=OK").Click
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Contact Lens Order").WinButton("regexpwndtitle:=Cancel").Click

'Clicking on Glasses Catalog Setup...
Window("regexpwndtitle:=Nextech.*","index:=0").Window("regexpwndtitle:=Inventory").WinButton("regexpwndtitle:=Glasses Catalog Setup\.\.\.").Click
If Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Glasses Catalog Setup").Exist Then
	Reporter.ReportEvent micPass, "Glasses Catalog Setup Dialog box is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Glasses Catalog Setup Dialog box is not Present"
End If
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Glasses Catalog Setup").WinButton("regexpwndtitle:=Billing Setup\.\.\.").Click
IsPresent = Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Glasses Catalog Setup").Dialog("regexpwndtitle:=Glasses Catalog Billing.*").Exist
Call RecordResults(IsPresent, "Glasses Catalog Billing Configuration Dialog is Present")
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Glasses Catalog Setup").Dialog("regexpwndtitle:=Glasses Catalog Billing.*").WinButton("regexpwndtitle:=Cancel").Click

'Clicking on Check VisionWeb Catalog Update inside Glasses Catalog Setup
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Glasses Catalog Setup").WinButton("regexpwndtitle:=Check VisionWeb.*").Click
IsPresent = Window("regexpwndtitle:=Nextech.*","index:=0").WinButton("regexpwndtitle:=OK").Exist
Call RecordResults(IsPresent, "Nextech Practice for Vision web catalog is Present")
Window("regexpwndtitle:=Nextech.*","index:=0").WinButton("regexpwndtitle:=OK").Click

IsPresent = Window("regexpwndtitle:=Nextech.*","index:=0").WinButton("regexpwndtitle:=OK").Exist
Call RecordResults(IsPresent, "Nextech Practice for Failed update catalog is Present")
Window("regexpwndtitle:=Nextech.*","index:=0").WinButton("regexpwndtitle:=OK").Click

'Clicking on Copy from... inside the Glasses Catalog Setup
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Glasses Catalog Setup").WinButton("regexpwndtitle:=Copy From\.\.\.").Click
IsPresent = Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Glasses Catalog Setup").Dialog("regexpwndtitle:=Select a Supplier").Exist
Call RecordResults(IsPresent, "Select a Supplier Dialog is Present")
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Glasses Catalog Setup").Dialog("regexpwndtitle:=Select a Supplier").WinButton("regexpwndtitle:=Cancel").Click

'Clicking on Add Designs inside the Glasses Catalog Setup
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Glasses Catalog Setup").WinButton("regexpwndtitle:=Add Designs").Click
IsPresent = Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Glasses Catalog Setup").Dialog("regexpwndtitle:=Select one or.*","Index:=0").Exist
Call RecordResults(IsPresent, "Select one or more Design dialog is Present")

'Clicking on the Add Button
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Glasses Catalog Setup").Dialog("regexpwndtitle:=Select one or.*","index:=0").WinButton("regexpwndtitle:=Add").Click
IsPresent = Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Glasses Catalog Setup").Dialog("regexpwndtitle:=Select one or.*","index:=0").Dialog("regexpwndtitle:=Add Design").Exist
Call RecordResults(IsPresent, "Add design Dialog box is Present")
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Glasses Catalog Setup").Dialog("regexpwndtitle:=Select one or.*","index:=0").Dialog("regexpwndtitle:=Add Design").WinButton("regexpwndtitle:=Cancel").Click
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Glasses Catalog Setup").Dialog("regexpwndtitle:=Select one or.*","index:=0").WinButton("regexpwndtitle:=Cancel").Click

'Clicking on the Add Material Button in the Glasses Catalog setup
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Glasses Catalog Setup").WinButton("regexpwndtitle:=Add Materials").Click
IsPresent = Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Glasses Catalog Setup").Dialog("regexpwndtitle:=Select one or.*","Index:=1").Exist
Call RecordResults(IsPresent, "Select one or more materials Dialog is Present")
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Glasses Catalog Setup").Dialog("regexpwndtitle:=Select one or.*","Index:=1").WinButton("regexpwndtitle:=Add").Click
IsPresent = Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Glasses Catalog Setup").Dialog("regexpwndtitle:=Select one or.*","index:=1").Dialog("regexpwndtitle:=Add Material").Exist
Call RecordResults(IsPresent, "Add material Dialog box is Present")
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Glasses Catalog Setup").Dialog("regexpwndtitle:=Select one or.*","index:=1").Dialog("regexpwndtitle:=Add Material").WinButton("regexpwndtitle:=Cancel").Click
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Glasses Catalog Setup").Dialog("regexpwndtitle:=Select one or.*","index:=1").WinButton("regexpwndtitle:=Cancel").Click

'Clicking on Add Treatnments button in the Glasses Catalog Setup
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Glasses Catalog Setup").WinButton("regexpwndtitle:=Add Treatments").Click
IsPresent = Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Glasses Catalog Setup").Dialog("regexpwndtitle:=Select one or.*","Index:=2").Exist
Call RecordResults(IsPresent, "Select one or more materials Dialog is Present")
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Glasses Catalog Setup").Dialog("regexpwndtitle:=Select one or.*","Index:=2").WinButton("regexpwndtitle:=Add").Click
IsPresent = Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Glasses Catalog Setup").Dialog("regexpwndtitle:=Select one or.*","index:=2").Dialog("regexpwndtitle:=Add Treatment").Exist
Call RecordResults(IsPresent, "Add material Dialog box is Present")
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Glasses Catalog Setup").Dialog("regexpwndtitle:=Select one or.*","index:=2").Dialog("regexpwndtitle:=Add Treatment").WinButton("regexpwndtitle:=Cancel").Click
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Glasses Catalog Setup").Dialog("regexpwndtitle:=Select one or.*","index:=2").WinButton("regexpwndtitle:=Cancel").Click

'Clicking on Add Frame Types Button in the Glasses Catalog Setup
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Glasses Catalog Setup").WinButton("regexpwndtitle:=Add Frame Types").Click
IsPresent = Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Glasses Catalog Setup").Dialog("regexpwndtitle:=Select one or.*","Index:=3").Exist
Call RecordResults(IsPresent, "Select one or more materials Dialog is Present")
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Glasses Catalog Setup").Dialog("regexpwndtitle:=Select one or.*","Index:=3").WinButton("regexpwndtitle:=Add").Click
IsPresent = Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Glasses Catalog Setup").Dialog("regexpwndtitle:=Select one or.*","index:=3").Dialog("regexpwndtitle:=Add Frame Type").Exist
Call RecordResults(IsPresent, "Add material Dialog box is Present")
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Glasses Catalog Setup").Dialog("regexpwndtitle:=Select one or.*","index:=3").Dialog("regexpwndtitle:=Add Frame Type").WinButton("regexpwndtitle:=Cancel").Click
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Glasses Catalog Setup").Dialog("regexpwndtitle:=Select one or.*","index:=3").WinButton("regexpwndtitle:=Cancel").Click


Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Glasses Catalog Setup").WinButton("regexpwndtitle:=Close").Click

'Clicking on Vision Web Setup...
Window("regexpwndtitle:=Nextech.*","index:=0").Window("regexpwndtitle:=Inventory").WinButton("regexpwndclass:=Button","index:=2").Click
If Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=VisionWeb Setup").Exist Then
	Reporter.ReportEvent micPass, "VisionWeb Setup Dialog box is Present", ""
Else
	Reporter.ReportEvent micFail, "Test Failed", "VisionWeb Setup Dialog box is not Present"
End If
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=VisionWeb Setup").WinButton("regexpwndtitle:=Close").Click

End Sub


Sub VerTabOpened (TabName)
	Select Case TabName
		Case "Product"
		
		Case "Order"
		
		Case "Allocation"
		
		Case "Overview"
		
		Case "Reports"
		
		Case "Frames"
		
		Case "Optical Orders"
		
	End Select
End Sub

Window("regexpwndtitle:=Nextech.*","index:=0").WinMenu("menuobjtype:=2").Select "Modules;Patients"



RunAction "LogOut [GlobalPracticeActions]", oneIteration



