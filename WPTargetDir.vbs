Set WshShell = WScript.CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")
strMsg = "Completed!" & Chr(10) & Chr(10) & "WPTargetDir.vbs - © 2010 Ramesh Srinivasan" & Chr(10) & Chr(10) & "Visit us at http://www.winhelponline.com/blog"
strCurWP =""

On Error Resume Next
strCurWP = WshShell.RegRead("HKCU\Software\Microsoft\Internet Explorer\Desktop\General\WallpaperSource")
On Error Goto 0

If Trim(strCurWP) = "" Then
	MsgBox "No Wallpaper selected for Desktop Slideshow"
Else
	If fso.FileExists(strCurWP) Then
		Set obj = CreateObject("Scripting.FileSystemObject")
		obj.DeleteFile(strCurWP)
	Else
		MsgBox "Cannot find target for " & strCurWP
	End If
End If
