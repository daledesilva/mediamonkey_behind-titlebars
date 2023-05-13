' MediaMonkey Script

' NAME: BehindTitlebars 0.9
' Author: Dale de Silva
' Website: www.oiltinman.com
' Date first started: 4/02/2009
' Date last edited: 7/02/2009

' INSTALL: Copy to Scripts\Auto\

' FILES THAT SHOULD BE PRESENT UPON A FRESH INSTALL:
' BehindTitlebars.vbs

' Special thanks to...
' Teknojnky		- without Trixmoto's original "Browse By Art" script, the foundation code for caching the album art would have been unavailable and this script would have been a far too daunting task for me.




Option Explicit
Dim startupState, arrValueNames, sMnu, hMnu
startupState = "hide"




Sub OnStartup()
	Initialise()
End Sub


Sub Initialise()
	Dim Mnu, aMnu, strComputer, strKeyPath, oReg, arrValueTypes, i
	Const HKEY_LOCAL_MACHINE = &H80000002
	Const HKEY_CURRENT_USER = &H80000001
	
	strComputer = "."
	 
	Set oReg = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & _ 
	    strComputer & "\root\default:StdRegProv")
	strKeyPath = "Software\MediaMonkey\Interface\Toolbars\Persistent"
	oReg.EnumValues HKEY_CURRENT_USER, strKeyPath,_
	    arrValueNames, arrValueTypes
	 
	
	' create submenu in view menu with seperators
	Set aMnu = SDB.UI.AddMenuItemSep(SDB.UI.Menu_View,-1,-1)
	
	Set Mnu = SDB.UI.AddMenuItemSub(SDB.UI.Menu_View,-1,-1)
	Mnu.Caption = "Show/Hide Panel Titlebars"
	
	'Set aMnu = SDB.UI.AddMenuItem(Mnu,-1,-1)
	'aMnu.Caption = "clear registry"
	'aMnu.Checked = False
	'Script.RegisterEvent aMnu, "OnClick", "ClearPanelsInRegistry"
	
	Set sMnu = SDB.UI.AddMenuItem(Mnu,-1,-1)
	sMnu.Caption = "Show All"
	sMnu.Checked = False
	Script.RegisterEvent sMnu, "OnClick", "ChangeAllTitlebars"
	sMnu.shortcut = "Ctrl+`"
	
	Set hMnu = SDB.UI.AddMenuItem(Mnu,-1,-1)
	hMnu.Caption = "Hide All"
	hMnu.Checked = False
	Script.RegisterEvent hMnu, "OnClick", "ChangeAllTitlebars"
	'hMnu.shortcut = "Ctrl+`"
	
	Set aMnu = SDB.UI.AddMenuItemSep(Mnu,-1,-1)
	 
	For i=0 To UBound(arrValueNames)
		CreateLinks Mnu, arrValueNames(i)
	Next

End Sub




Sub CreateLinks(Mnu, PnlName)
	Dim aMnu
	Dim Pnl : Set Pnl = SDB.UI.NewDockablePersistentPanel(PnlName)
	
	'add links to submenu within view menu
	Set aMnu = SDB.UI.AddMenuItem(Mnu,-1,-1)
	aMnu.Caption = PnlName
	
	If startupState = "hide" Then
		Pnl.showcaption = False
	ElseIf startupState = "show" Then
		Pnl.showcaption = True
	End If	
	
	'aMnu.Checked = Pnl.showcaption
	Script.RegisterEvent aMnu, "OnClick", "ChangeTitlebar"

End Sub



Sub ChangeTitlebar(Mnu)
	Dim Pnl : Set Pnl = SDB.UI.NewDockablePersistentPanel(Mnu.Caption)

	If Pnl.showcaption = False Then
		Pnl.showcaption = True
		'Mnu.Checked = True
	Else
		Pnl.showcaption = False
		'Mnu.Checked = False
	End If

End Sub



Sub ChangeAllTitlebars(Mnu)
	Dim Pnl, i
	
	For i=0 To UBound(arrValueNames)
		Set Pnl = SDB.UI.NewDockablePersistentPanel(arrValueNames(i))
		
		If Mnu.Caption = "Show All" Then
			Pnl.showcaption = True
			hMnu.shortcut = "Ctrl+`"
			sMnu.shortcut = ""
		Else
			Pnl.showcaption = False
			hMnu.shortcut = ""
			sMnu.shortcut = "Ctrl+`"
		End If
	Next

End Sub



Sub clearPanelsInRegistry(Mnu)
	Dim strComputer, strKeyPath, oReg, arrValueTypes
	Const HKEY_CURRENT_USER = &H80000001

	'arrValueNames is already initialised, but it needs to be initialised again after because the registry will be different
	strComputer = "."
	Set oReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & _ 
	    strComputer & "\root\default:StdRegProv")
		 	
	strKeyPath = "Software\MediaMonkey\Interface\Toolbars\Persistent"
	
	
	For i=0 To UBound(arrValueNames)
		oReg.DeleteValue HKEY_CURRENT_USER, strKeyPath, arrValueNames(i)
	Next
	
	oReg.EnumValues HKEY_CURRENT_USER, strKeyPath,_
	    arrValueNames, arrValueTypes
	
End Sub


Sub FirstRun()

  If SDB.VersionString < "3.1.0" Then
  	Uninstall()
  	Exit Sub
  Else
  	On Error Resume Next 'Suppress error messages
	  If SDB.VersionBuild > 1208 Then
	  
	  	Initialise()
	    Exit Sub
	    
	  End If
	  On Error Goto 0 'Unsuppress error messages
  End If
  
  Uninstall()

End Sub


Sub Uninstall()

	SDB.MessageBox "Behind Titlebars is only compatible with MediaMonkey 3.1.0.1209 and later."& VbNewLine &_
                    "Please Update your version of MediaMonkey. You might need to download from the Beta Forum to get the required version."& VbNewLine &_
                    "Mediamonkey will tell you the script is installed, however, it will be uninstalled for you automatically.", mtWarning, Array(mbOk)

	Dim path : path = SDB.ApplicationPath&"Scripts\Auto\"
	Dim fso : Set fso = CreateObject("Scripting.FileSystemObject")

	If fso.FileExists(path&"BehindTitlebars.vbs") Then
		Call fso.DeleteFile(path&"BehindTitlebars.vbs")
	End If

End Sub