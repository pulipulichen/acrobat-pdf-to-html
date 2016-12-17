#include <File.au3>
#include <MsgBoxConstants.au3>
#include <Array.au3>

Local $acrobat= IniRead ( @ScriptDir & "\config.ini", "config", "acrobat", "C:\Program Files (x86)\Adobe\Acrobat DC\Acrobat\Acrobat.exe" )
Local $force_delete= IniRead ( @ScriptDir & "\config.ini", "config", "force_delete", "0" )
Local $open_html= IniRead ( @ScriptDir & "\config.ini", "config", "open_html", "0" )
Local $copy_to_clipboard= IniRead ( @ScriptDir & "\config.ini", "config", "copy_to_clipboard", "0" )
Local $remove_after_copy = IniRead ( @ScriptDir & "\config.ini", "config", "remove_after_copy", "0" )


Local $CommandLine = $CmdLine


;If $CommandLine[0] = 0 Then
   ;$CommandLine[0] = 1
   ;_ArrayAdd($CommandLine, "D:\PortableApps\acrobat-pdf-to-html\example\example.pdf")
;EndIf

For $i = 1 To $CommandLine[0]
   Local $file = $CommandLine[$i]
   ;MsgBox($MB_SYSTEMMODAL, "", $file)
   IF FileExists($file) = False Then
	  ContinueLoop
   EndIf

   Local $html = $file & ".html"


   If FileExists($html) and $force_delete = 1 Then
	  FileRecycle($html)
   EndIf

   If FileExists($html) = 1 Then
	  Exit
   EndIf



   Run($acrobat & " " & $file )
   WinWait("[CLASS:AcrobatSDIWindow]")
   WinActivate("[CLASS:AcrobatSDIWindow]")
   Send("{ALTDOWN}fth{ALTUP}")
   Local $hWnd = WinWait("Save As PDF")
   WinActivate("Save As PDF")

   ;Send("{ALTDOWN}n{ALTUP}")

   Local $dirPath = _PathFull($file & "\.")

   ;ControlSetText($hWnd, "", "ToolbarWindow323", $dirPath )
   ControlSetText($hWnd, "", "Edit1", $html )

   ;Send("{ALTDOWN}v{ALTUP}")
   ControlCommand($hWnd, "", "Button2", "UnCheck") ; 讓勾選的項目取消
   ;MsgBox($MB_SYSTEMMODAL, "", ControlGetText($hWnd, "", "Button2"))
   ;MsgBox($MB_SYSTEMMODAL, "", GUICtrlRead("[CLASS:Button; INSTANCE:2]"))
   ;Exit

   Send("{ENTER}")

   While 1
	   Sleep(100)
	   If FileExists($html) = 1 Then
		   Sleep(1000)
		   ExitLoop
	   EndIf
   WEnd

   WinActivate("[CLASS:AcrobatSDIWindow]")
   Send("{CTRLDOWN}w{CTRLUP}")

   ;MsgBox($MB_SYSTEMMODAL, "", $open_html & " " & $html)

   If $open_html <> "0" Then
	  Run($open_html & " " & $html)
   EndIf
   If $copy_to_clipboard = 1 Then
	  ClipPut(FileRead($html))
	  If $remove_after_copy = 1 Then
		 FileRecycle($html)
	  EndIf
   EndIf
Next