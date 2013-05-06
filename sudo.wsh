Option Explicit

function checkInput(inputParams)

if inputParams.Count = 0 then
   checkInput = false
end if

if StrComp(inputParams(0), "--help") = 0 then
  showHelp
  checkInput = false
end if

checkInput = true

end function

sub showHelp

WScript.Echo "sudo - this is handy command-line util for launching programs with administrator's privilegies." & vbCrLf & vbCrLf & _
             "Example:" & vbCrLf & vbCrLf & _
			 "sudo <launching_program>"

end sub

dim argc
dim argv
dim i

argc = WScript.Arguments.Count



if Not checkInput(WScript.Arguments) then
	WScript.quit
end if

set argv = WScript.Arguments
if argc < 1 then 
MsgBox "Usage: sudo <arg1 arg2 .. argN>"
WScript.quit
end if
dim str
for i = 1 to argc-1
MsgBox "Param[" & i & "] = " & argv(i)
str = str + " " + argv(i)
next

dim objShell
set objShell = CreateObject("Shell.Application") 
objShell.ShellExecute argv(0), str, "", "runas", 1