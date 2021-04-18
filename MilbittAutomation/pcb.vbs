Dim rootDir : rootDir = "D:\MilbittAutomation\"
Dim arch : arch = "x64"

Dim pcbApp
Set pcbApp = Application
Dim guid
Dim lastAppId
guid = pcbApp.InstanceGuid

Dim gui
Dim AutomationMenu
Dim pcbDoc 

''' COMMON FUNCTIONS DO NOT MODIFY! >>>
	Function AddMenu(Caption, TargetMenu)
		If TargetMenu Is Nothing Then
			Exit Function
		End If
		Set AddMenu = Nothing
		Set AddMenu = FindMenu(Caption,TargetMenu)
		If AddMenu Is Nothing Then
			Set AddMenu = TargetMenu.Controls.Add(cmdControlPopup,,, -1)
			AddMenu.Caption = Caption
		End If
	End Function

	Function AddMenuAfter(Caption, afterMenuEntry, TargetMenu)
		Dim entryNum
		entryNum = GetMenuNumber(afterMenuEntry, TargetMenu) + 1

		If entryNum > TargetMenu.Controls.Count Then
			entryNum = -1
		End If

		Set AddMenuAfter = Nothing
		Set AddMenuAfter = FindMenu(Caption,TargetMenu)
		If AddMenuAfter Is Nothing Then
			Set AddMenuAfter = TargetMenu.Controls.Add(cmdControlButton,,,entryNum)
			AddMenuAfter.Caption = Caption
		End If
	End Function

	Function GetMenuNumber(menuToFind, menuBar)
		Dim ctrls : Set ctrls = menuBar.Controls
		Dim ctrl

		Dim menuCnt : menuCnt = 1
		GetMenuNumber = -1

		For Each ctrl In ctrls
		   Dim capt: capt = ctrl.Caption
		   capt = Replace(capt, "&", "")
		   If capt = menuToFind Then
			   GetMenuNumber = menuCnt
			   Exit For
		   End If
		   menuCnt = menuCnt + 1
		Next
	End Function

	Function FindMenu(menuToFind, menuBar)
		Dim ctrls : Set ctrls = menuBar.Controls
		Dim ctrl

		Set FindMenu = Nothing

		For Each ctrl In ctrls
		   Dim capt: capt = ctrl.Caption
		   capt = Replace(capt, "&", "")
		   If capt = menuToFind Then
			   Set FindMenu = ctrl
			   Exit For
		   End If
		Next
	End Function

	Function AddButton(Caption,TargetMethod,TargetMenu, disabled)
		If TargetMenu Is Nothing Then 
			Exit Function
		End If
		
		Set btnToAdd = Nothing
		Set btnToAdd = FindMenu(Caption, TargetMenu)
		If btnToAdd Is Nothing Then 
			Set btnToAdd = TargetMenu.Controls.Add(cmdControlButton,,,-1)
			btnToAdd.Caption = Caption
			btnToAdd.Target = ScriptEngine
			btnToAdd.DescriptionText = Caption
			btnToAdd.ExecuteMethod = TargetMethod
			Scripting.DontExit = True
			If Disabled = True Then
				btnToAdd.Enabled = False
			End If
		End If		
	End Function

	Function AddLauncherButton(Caption,AppID,TargetMenu, disabled, shortcut)
		If TargetMenu Is Nothing Then 
			Exit Function
		End If
		
		Set btnToAdd = Nothing
		Set btnToAdd = FindMenu(Caption, TargetMenu)
		If btnToAdd Is Nothing Then 
			Set btnToAdd = TargetMenu.Controls.Add(cmdControlButton,,,-1)
			btnToAdd.Caption = Caption
			btnToAdd.Target = ScriptEngine
			btnToAdd.DescriptionText = Caption
			btnToAdd.ExecuteMethod = "LaunchApp"
			btnToAdd.TooltipText = AppId
			Scripting.DontExit = True
			If Disabled = True Then
				btnToAdd.Enabled = False
			End If
		End If	

		If shortcut <> "" Then
			Dim path : path = Replace(AutomationMenu.Caption + "->" + TargetMenu.Caption + "->" + Caption, "&", "")
			Call AddKeyBindingForMenuCmd(shortcut, path)
		End If
	End Function

	Function AddSeperator(TargetMenu)
		If TargetMenu Is Nothing Then 
			Exit Function
		End If
			
		Set btnToAdd = Nothing
		Set btnToAdd = TargetMenu.Controls.Add(cmdControlButtonSeparator,,,-1)
	End Function

	

	Function AddKeyBindingForKeyInCmd(Key, Cmd)
		gui.Bindings("Document").AddKeyBinding Key, Cmd, 1, 1
	End Function

	Function AddKeyBindingForMenuCmd(Key, Cmd)
		gui.Bindings("Document").AddKeyBinding Key, Cmd, 0, 1
	End Function

	Function GetLicensedDoc(app)
	  On Error Resume next
	  Dim key,licenseServer,licenseToken,docObj
	  Set GetLicensedDoc = Nothing
	  
	  Set docObj = app.ActiveDocument
	  If (Err) Then 
		 Call app.Gui.StatusBarText("No active document: " + _
			Err.Description,epcbStatusFieldError)
		 Exit Function
	  End If
	  
	  key = docObj.Validate(0)
	  
	  Set licenseServer = _
		 CreateObject("MGCPCBAutomationLicensing.Application")
	  licenseToken = licenseServer.GetToken(key)
	  Set licenseServer = Nothing
	  
	  Err.Clear 
	  docObj.Validate(licenseToken)
	  If (Err) Then 
		 Call app.Gui.StatusBarText("No active document license: " + _
			Err.Description,epcbStatusFieldError)
		 Exit Function
	  End If
	  
	  Set GetLicensedDoc = docObj 
	End Function

	Function GetToolTipOfButton(btn_id)
		Dim docMenuBar
		For Each docMenuBar In Gui.CommandBars	
			For i = 1 To docMenuBar.Controls.Count
				On Error Resume Next
				Set menu = docMenuBar.Controls.Item(i)
				Call WriteMenuIDs(menu, btn_id)
			Next
		Next
		GetToolTipOfButton = lastAppId
		lastAppId = ""
	End Function

	Function WriteMenuIDs(menu, btn_id)
		WriteMenuIDs = "Nothing"
		Set menuCtrls = menu.Controls
		For j = 1 To menuCtrls.Count
			cmdName = menuCtrls.Item(j).Caption
			On Error Resume Next
			id = menuCtrls.Item(j).Id
			
			If Err Then
				 Err.Clear
				 WriteMenuIDs = WriteMenuIDs(menuCtrls.Item(j), btn_id)
			Elseif id <> 0 Then
				If Trim(id) = Trim(btn_id) Then
					lastAppId = menuCtrls.Item(j).TooltipText
					Exit For
				End If
			End If
		 Next
	End Function

	Function LaunchApp(btnId)
		appPath = GetToolTipOfButton(btnId)

		appPath = Replace(appPath, "{arch}", arch)

		Dim path
		path = rootDir + appPath + " " + guid

		Dim exec
		Set exec = CreateObject("ViewLogic.Exec")
		Call exec.Run(path,1,false)
	End Function
''' <<< COMMON FUNCTIONS


Function Generate_Menu_Structure()
    If pcbApp Is Nothing Then
        MsgBox "Automation connection to PCB application failed."
        Exit Function
    End If
    Set gui = pcbApp.Gui
	Scripting.AttachEvents pcbApp, "pcbApp"	
    Dim docMenuBar
    Set docMenuBar = gui.CommandBars("Document Menu Bar")
    Set AutomationMenu = Nothing
    Set AutomationMenu = FindMenu("Milbitt Automation", docMenuBar)
    If AutomationMenu Is Nothing Then
        Set AutomationMenu = docMenuBar.Controls.Add(cmdControlPopup,,,-1)
        AutomationMenu.Caption = "&Milbitt Automation"   
    End If

    Call AddMenu("Tools", AutomationMenu)
	Call AddSeperator(AutomationMenu)
	Call AddMenu("Placement", AutomationMenu)
	Call AddSeperator(AutomationMenu)
	Call AddMenu("View", AutomationMenu)
	Call AddSeperator(AutomationMenu)
	Call AddMenu("Mounting Holes", AutomationMenu)
	Call AddSeperator(AutomationMenu)
	Call AddButton("Milbitt Engineering - www.milbitt.com", "", AutomationMenu, True) 'keep this line for credit!
End Function
Generate_Menu_Structure()

Function Define_Shortcuts()
	AddKeyBindingForKeyInCmd "Alt+F2", "r 45"
	AddKeyBindingForMenuCmd "Ctrl+B", "Route->Tune Routes->Manual Tune"
	AddKeyBindingForMenuCmd "Ctrl+Shift+B", "Route->Tune Routes->Manual Saw Tune"
	AddKeyBindingForMenuCmd "Ctrl+N", "Edit->Snap->Toggle Hover Snap"
End Function
Define_Shortcuts()


''' Automation -> Tools
Function Generate_Tools_Menu()
	If AutomationMenu Is Nothing Then
		Exit Function
	End If

	Dim TheMenu
	Set TheMenu = AddMenu("Tools", AutomationMenu)
	Call AddLauncherButton("Unit Converter", "UnitConverter\UnitConverter.exe", TheMenu, False, "")
	Call AddSeperator(TheMenu)
	Call AddLauncherButton("Outline Assigner", "xPCB_OutlineAssigner\xPCB_OutlineAssigner_{arch}.exe", TheMenu, False, "")
End Function
Generate_Tools_Menu()



''' Automation -> Placement
Function Generate_Placement_Menu()
	If AutomationMenu Is Nothing Then
		Exit Function
	End If

	Dim TheMenu
	Set TheMenu = AddMenu("Placement", AutomationMenu)
	Call AddLauncherButton("Arrange Reference Designators", "xPCB_RefDesArranger\xPCB_RefDesArranger_{arch}.exe", TheMenu, False, "")
	Call AddSeperator(TheMenu)
	Call AddLauncherButton("Colorize Pins by Symbols", "xPCB_BankPainter\xPCB_BankPainter_{arch}.exe", TheMenu, False, "")
End Function
Generate_Placement_Menu()


''' Select Parts -> Right Click Menu -> Automation
Function Generate_RB_SelectedParts()
	Dim RB_menu
	Set RB_menu = gui.CommandBars("SelectedParts")
	Dim TheMenu
	Set TheMenu = AddMenu("Milbitt Automation", RB_menu)
	
	Call AddLauncherButton("Arrange Reference Designators", "xPCB_RefDesArranger\xPCB_RefDesArranger_{arch}.exe", TheMenu, False, "")
	Call AddSeperator(TheMenu)
	Call AddLauncherButton("Colorize Pins by Symbols", "xPCB_BankPainter\xPCB_BankPainter_{arch}.exe", TheMenu, False, "")
	
End Function
Generate_RB_SelectedParts()


''' Automation -> Mounting Holes
Function Generate_MountingHoles_Menu()
	If AutomationMenu Is Nothing Then
		Exit Function
	End If

	Dim TheMenu
	Set TheMenu = AddMenu("Mounting Holes", AutomationMenu)
	Call AddLauncherButton("Shave Mounting Holes", "Mentor_MountingHole_Shaver\Mentor_MountingHole_Shaver_{arch}.exe", TheMenu, False, "")
End Function
Generate_MountingHoles_Menu()


Function Generate_View_Menu()
	If AutomationMenu Is Nothing Then
		Exit Function
	End If

	Dim TheMenu
	Set TheMenu = AddMenu("View", AutomationMenu)

	Call AddButton("Zoom to Cursor", "xPCB_ZoomToCursor", TheMenu, False)
	AddKeyBindingForMenuCmd "Ctrl+Shift+C", "Milbitt Automation->View->Zoom to Cursor"
End Function
Generate_View_Menu()

Function xPCB_ZoomToCursor(nId)
	Set pcbDoc = GetLicensedDoc(pcbApp)
	Dim offset
	offset = 500
	Call pcbDoc.ActiveView.SetExtents(pcbDoc.ActiveViewEx.MousePositionX(epcbUnitMils) - offset, pcbDoc.ActiveViewEx.MousePositionY(epcbUnitMils), pcbDoc.ActiveViewEx.MousePositionX(epcbUnitMils) + offset, pcbDoc.ActiveViewEx.MousePositionY(epcbUnitMils), epcbUnitMils)
End Function

