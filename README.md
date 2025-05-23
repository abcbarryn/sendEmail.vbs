# sendEmail.vbs
Flexible VBS script to send email.

# Usage
cscript sendEmail.vbs &lt;host&gt; &lt;port&gt; &lt;from&gt; &lt;to&gt; &lt;subject&gt; &lt;body&gt; [/Attach:file1[,file2][,filen]] [/User:username]
  [/Password:password] [/SSL:true] [/NTLM:true]

  If /SSL:true is present SSL/TLS will be used if the server supports .
  SSL/TLS will only work on ports 25 and 465, read below for details.
  If /NTLM:true is present NTLM authentication will be used if a user name and password is specified otherwise if a user name and password is given basic authentication will be used.

If you use /ssl:true when connecting to a mail server providing SMTPS services on port 465, it is a very simple situation. Due to the all-or-nothing approach of the TLS encryption of the TCP/IP connection on port 465, this will work right away, or it won’t if there are configuration issues outside of the CDO library. Please note that the host name must match one of the ones in the SSL certificate.

The only other port where SSL can be used with VB script is port 25. When this CDO VB script is configured to send via port 25 with /SSL:true, VB will issue a STARTTLS command. But using the very same configuration on the client and on the server and just changing the port to 587, which is the standard port for opportunistic TLS, will equally consistently fail with an Error 0x80040213 “The transport failed to connect to the server.”. The Windows CDO DLL contains hardcoded logic to behave differently when connecting to port 25 as opposed to other ports. CDO’s behavior in handling sending emails over an encrypted connection, is hardcoded into the DLL file and cannot be changed by any means of configuration. Neither directly via the Configuration object, nor indirectly via modifying the Windows Registry or any configuration file.

If /SSL:false is specified or the /SSL switch is not given, CDO will never try to use any encryption. If the server expects encryption, sending emails will fail; either immediately or after expiry of the connection timeout for servers expecting SMTPS like on port 465.

If the server does not enforce encryption, emails will be sent, but through an unencrypted connection.

If /SSL:true is specified, the situation is much less straightforward. When beginning to send an email, CDO will evaluate the target port used for the connection.

Special Behavior on Port 25
If this or any script using CDO is configured for sending on port 25, if the server is configured to support STARTTLS on port 25, either optionally or mandatory, CDO will initially connect unencrypted and then initiate the encryption of the connection by sending the STARTTLS command. – This works as opportunistic TLS is intended to.
