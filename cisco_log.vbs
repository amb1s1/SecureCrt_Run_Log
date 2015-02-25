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
  crt.session.log True, True
 
  crt.Screen.Send "!-----Begin---" & VbCr
  crt.Screen.WaitForString "#"
  
  'This will grab the hostname and use it to save the log file
  Set objTab = crt.GetScriptTab()
  nRow = objTab.Screen.CurrentRow
  hostname = objTab.screen.Get(nRow, 0, nRow, objTab.Screen.CurrentColumn - 2)
  hostname = Trim(hostname)
  
  crt.Session.LogFileName = ("C:\Running_Capture\"& hostname & ".txt")
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

