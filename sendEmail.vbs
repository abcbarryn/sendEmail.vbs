'''''''''''''''''''''''''''''''''''''''''''''''
'
' c:\scripts\events\sendEmail.vbs
'
' @Author: Nestor Urquiza
'
'
' @Description: Sends an email
'
'
'
''''''''''''''''''''''''''''''''''''''''''''''''

'
' Functions
'
Sub usage
    Wscript.Echo Wscript.ScriptName & " <host> <port> <from> <to> <subject> <body> [/Attach:file1[,file2][,filen]] [/User:username]"
    Wscript.Echo "  [/Password:password] [/SSL:true] [/NTLM:true]"
    Wscript.Echo
    Wscript.Echo "  If /SSL:true is present SSL/TLS will be used."
    Wscript.Echo "  SSL/TLS will only work on ports 25 and 465."
    Wscript.Echo "  If /NTLM:true is present NTLM authentication will be used if a user name and password is specified."
    WScript.Quit
End Sub

Function IsBlank(Value)
  IsBlank = False
  If IsEmpty(Value) or IsNull(Value) Then
    IsBlank = True
  End If
End Function


'Constants
strComputer = "."

'System config
Set wshShell = WScript.CreateObject( "WScript.Shell" )
strComputerName = wshShell.ExpandEnvironmentStrings( "%COMPUTERNAME%" )

'Parameters
If ( Wscript.Arguments.Count < 6 or Wscript.Arguments.Count = 7 ) Then
  Call usage
End If

host = Wscript.Arguments(0)
port = Wscript.Arguments(1)
from = Wscript.Arguments(2)
strTo = Wscript.Arguments(3)
subject = Wscript.Arguments(4)
body = Wscript.Arguments(5)

If len(WScript.Arguments.Named("user")) > 1 then
  sslUser = WScript.Arguments.Named("user")
End If
If len(WScript.Arguments.Named("password")) > 1 then
  sslPassword = WScript.Arguments.Named("password")
End If

useSSL = false
if WScript.Arguments.Named("ssl") then
  useSSL = true
End If
authmode=1
if WScript.Arguments.Named("ntlm") then
  authmode = 2
End If

'Prepare email
Set objEmail = CreateObject("CDO.Message")
objEmail.From = from
objEmail.To = strTo
objEmail.Subject = "[" & strComputerName & "]" & " " & subject
objEmail.Textbody = body


Set emailCfg = objEmail.Configuration
' cdoSendUsingPickup (1)
' cdoSendUsingPort (2)
' cdoSendUsingExchange (3)
emailCfg.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2

emailCfg.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = host
emailCfg.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = port

If ( useSSL ) then
  Wscript.Echo "Using SSL"
End If
emailCfg.Fields("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = useSSL

If ( Not isBlank(sslUser) ) Then
' 0 cdoAnonymous Perform no authentication.
' 1 cdoBasic     Use the basic (clear text) authentication mechanism.
' 2 cdoNTLM      Use the NTLM authentication mechanism.
  emailCfg.Fields("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = authmode
  Wscript.Echo "Using user authentication"
  emailCfg.Fields("http://schemas.microsoft.com/cdo/configuration/sendusername") = sslUser
  emailCfg.Fields("http://schemas.microsoft.com/cdo/configuration/sendpassword") = sslPassword
End If

Set fso = CreateObject("Scripting.FileSystemObject")
If len(WScript.Arguments.Named("attach")) > 1 then
        Attachments = split(WScript.Arguments.Named("attach"), ",")
        If ubound(Attachments) >= 0 then
                for each item in Attachments
                        Set f = fso.GetFile(item)
                        Wscript.Echo "Attached " & f.Path
                        objEmail.AddAttachment f.Path
                next
        End If
End If

objEmail.Configuration.Fields.Update
objEmail.Send
