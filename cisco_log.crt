# $language = "VBScript"
# $interface = "1.0"
'--------------------------------------------------------------
'Cisco IOS
'
'Created by David Gomez, davidgomez.255@gmail.com
'---------------------------------------------------------------

Sub Main
  ' turn on synchronous mode so we don't miss any data
  crt.Screen.Synchronous = True
  Dim bStartLog, bAppendToLog
  logfile = crt.Dialog.Prompt("Logfile name", "Capture Completion Notice Info", "C:\Running_Capture\Script_Completion_Log.log", False)
  crt.Session.LogFileName = logfile
  'crt.Session.LogFileName = "C:\Running_Capture\%Y-%M-%D--%h-%m-%s.%t-%S(%H).txt"
  crt.session.log True, True
 
  'crt.Session.Log True
  crt.Screen.Send "!-----Begin---" & VbCr
  crt.Screen.WaitForString "#"
  crt.Screen.Send "term length 0" & VbCr
  crt.Screen.WaitForString "#"
  crt.Screen.Send "!" & VbCr
  crt.Screen.WaitForString "#"
  crt.Screen.Send "!############################################################" & VbCr
  crt.Screen.Send "!This command will Log the Running Config on a Cisco Device" & VbCr
  crt.Screen.Send "!Revision 1.0" & VbCr
  crt.Screen.Send "!############################################################" & VbCr
  crt.Screen.WaitForString "#"
  crt.Screen.Send "!############################################################" & VbCr
  crt.Screen.Send "!STARTING EXECUTION" & VbCr
  crt.Screen.Send "!############################################################" & VbCr
  crt.Screen.WaitForString "#"
  crt.Screen.Send "SHOW RUN" & VbCr
  crt.Screen.WaitForString "#"
  crt.Screen.Send "!" & VbCr
  crt.Screen.WaitForString "#"
 
  ' turn off synchronous mode to restore normal input processing
  crt.Screen.Synchronous = False
  crt.Sleep 3500
  crt.session.log false
end Sub

