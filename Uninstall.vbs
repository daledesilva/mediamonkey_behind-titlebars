myName = "Behind Titlebars"    'Put script name here (shown to user)

Dim path : path = SDB.ApplicationPath&"Scripts\Auto\"
Dim fso : Set fso = CreateObject("Scripting.FileSystemObject")

If fso.FileExists(path&"BehindTitlebars.vbs") Then
	Call fso.DeleteFile(path&"BehindTitlebars.vbs")
End If

MsgBox("I hope your experiences with Behind Titlebars were not all bad." & vbNewLine & "Please restart MediaMonkey for the uninstall to have full effect.")




