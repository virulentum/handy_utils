Option Explicit

' Checks entered parameters
function checkInput(inputParams)

   dim retValue

   retValue = True

   if inputParams.Count = 0 then
      showHelp
      retValue = False
   else
   if StrComp(inputParams(0), "--help") = 0 then
      showHelp
      retValue = False
   end if
end if

checkInput = retValue

end function

' Displays help message
sub showHelp

WScript.Echo "sudo" & vbCrLf & _
		     vbTab & "command-line util for launching programs" & vbCrLf & _
			 vbTab & "with administrator's privilegies." & vbCrLf & vbCrLf & _
             "Examples:" & vbCrLf & _
			 vbTab & "sudo   <program_for_launching>" & vbCrLf & _
			 vbTab & "sudo   --help"

end sub

dim argv

if Not checkInput(WScript.Arguments) then
	WScript.quit
end if

set argv = WScript.Arguments

dim objShell
set objShell = CreateObject("Shell.Application") 
objShell.ShellExecute argv(0), "", "", "runas", 1