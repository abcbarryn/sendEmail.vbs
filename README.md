# sendEmail.vbs
VBS flexible VBS script to send email.

# Usage
cscript sendEmail.vbs <host> <port> <from> <to> <subject> <body> [/Attach:file1[,file2][,filen]] [/User:username]
  [/Password:password] [/SSL:true] [/NTLM:true]

  If /SSL:true is present SSL/TLS will be used.
  SSL/TLS will only work on ports 25 and 465.
  If /NTLM:true is present NTLM authentication will be used if a user name and password is specified.
