REM The following script was written to log into the SAP server automatically.
REM To view historical information and credit for this script please see
REM the following thread on the SAP Community Network:
REM http://scn.sap.com/thread/3763970

REM This script was last updated by Paul Street on 7/1/15

REM Directives
Option Explicit

REM Variables!  Must declare before using because of Option Explicit
Dim WSHShell, SAPGUIPath, SID, InstanceNo, WinTitle, SapGuiAuto, application, connection, session, W_Transaction, wks, Table, rows, cols, columns, dados, ie

REM Main
Set WSHShell = WScript.CreateObject("WScript.Shell")
If IsObject(WSHShell) Then

  REM Set the path to the SAP GUI directory
    SAPGUIPath = "C:\Program Files (x86)\SAP\FrontEnd\SapGui\"

  REM Set the SAP system ID
    SID = "172.19.19.10"

  REM Set the instance number of the SAP system
    InstanceNo = "02"

  REM Starts the SAP GUI
    WSHShell.Exec SAPGUIPath & "SAPgui.exe " & SID & " " & _
      InstanceNo

  REM Set the title of the SAP GUI window here
    WinTitle = "SAP"

  While Not WSHShell.AppActivate(WinTitle)
  Wend

  Set WSHShell = Nothing
End If

On Error Resume Next
W_Transaction = "ST22"
 
If Not IsObject(application) Then
  Set SapGuiAuto = GetObject("SAPGUI")
  Set application = SapGuiAuto.GetScriptingEngine
End If
If Not IsObject(connection) Then
  Set connection = application.Children(0)
End If
If Not IsObject(session) Then
  Set session = connection.Children(0)
End If
If IsObject(WScript) Then
  WScript.ConnectObject session, "on"
  WScript.ConnectObject application, "on"
End If

session.findById("wnd[0]").maximize

session.findById("wnd[0]/usr/txtRSYST-BNAME").text = "ddic"

'Coletando Erro de Login
If Err.number <> 0 Then
  function readFromRegistry(strRegistryKey, strDefault)
    Dim WSHShell, value

    On Error Resume Next
    Set WSHShell = CreateObject ("WScript.Shell")
    value = WSHShell.RegRead (strRegistryKey)

    if err.number <> 0 then
      readFromRegistry = strDefault
    else
      readFromRegistry = value
    end if

    set WSHShell = nothing
  end function

  function OpenWithChrome(strURL)
    Dim strChrome
    Dim WShellChrome

    strChrome = readFromRegistry ( "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\chrome.exe\Path", "") 
    if (strChrome = "") then
        strChrome = "chrome.exe"
    else
        strChrome = strChrome & "\chrome.exe"
    end if
    Set WShellChrome = CreateObject("WScript.Shell")
    strChrome = """" & strChrome & """" & " " & strURL
    WShellChrome.Run strChrome, 1, false
  end function

  OpenWithChrome "http://localhost:11500/error"

  WScript.CreateObject("WScript.Shell").Run "taskkill /f /im Cscript.exe", , True 
  WScript.CreateObject("WScript.Shell").Run "taskkill /f /im wscript.exe", , True  
End If

session.findById("wnd[0]/usr/pwdRSYST-BCODE").text = "CorpFlex@2019"

session.findById("wnd[0]/usr/pwdRSYST-BCODE").setFocus

session.findById("wnd[0]/usr/pwdRSYST-BCODE").caretPosition = 10

session.findById("wnd[0]").sendVKey 0

session.findById("wnd[0]/tbar[0]/okcd").Text = W_Transaction

session.findById("wnd[0]").sendVKey 0

session.findById("wnd[0]").sendVKey 8

WScript.CreateObject("WScript.Shell").SendKeys "^{PGDN}"

WScript.CreateObject("WScript.Shell").SendKeys "^(c)"

function readFromRegistry(strRegistryKey, strDefault)
  Dim WSHShell, value

  On Error Resume Next
  Set WSHShell = CreateObject ("WScript.Shell")
  value = WSHShell.RegRead (strRegistryKey)

  if err.number <> 0 then
    readFromRegistry = strDefault
  else
    readFromRegistry = value
  end if

  set WSHShell = nothing
end function

function OpenWithChrome(strURL)
  Dim strChrome
  Dim WShellChrome

  strChrome = readFromRegistry ( "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\chrome.exe\Path", "") 
  if (strChrome = "") then
      strChrome = "chrome.exe"
  else
      strChrome = strChrome & "\chrome.exe"
  end if
  Set WShellChrome = CreateObject("WScript.Shell")
  strChrome = """" & strChrome & """" & " " & strURL
  WShellChrome.Run strChrome, 1, false
end function

OpenWithChrome "http://localhost:11500/applicationres"

WScript.Sleep 1000

WScript.CreateObject("WScript.Shell").SendKeys "{TAB}"

WScript.CreateObject("WScript.Shell").SendKeys "^(v)"

WScript.CreateObject("WScript.Shell").SendKeys "{ENTER}"

'Finish Session SAP GUI
session.findById("wnd[0]").sendVKey 15

session.findById("wnd[0]").sendVKey 15

session.findById("wnd[1]/usr/btnSPOP-OPTION1").press