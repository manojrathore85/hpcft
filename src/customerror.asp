<%
dim ae
dim messagebody
set ae = server.GetLastError()

messagebody = "Asp Code = " & ae.aspcode & vbcrlf
messagebody = messagebody & "Asp error no = " & ae.number & vbcrlf
messagebody = messagebody & "Asp error source = " & ae.source & vbcrlf
messagebody = messagebody & "Asp error category = " & ae.category & vbcrlf
messagebody = messagebody & "Asp line = " & ae.line & vbcrlf
messagebody = messagebody & "Asp error column = " & ae.column & vbcrlf
messagebody = messagebody & "Error Description = " & ae.Description & vbcrlf
messagebody = messagebody & "Asp Description = " &ae.ASPDescription & vbcrlf
messagebody = messagebody & "Error Date = " & now & vbcrlf
messagebody = messagebody & "Error Path info = " &request.ServerVariables("PATH_TRANSLATED") & vbcrlf

'Set Mailer = Server.CreateObject("SMTPsvg.Mailer")
'Mailer.FromName   = "Asp Errors"
'Mailer.FromAddress = varSiteSpecFromAddress 
'Mailer.RemoteHost = varSiteSpecRemoteHost
'Mailer.AddRecipient "Errors", varSiteSpecFromAddress
'Mailer.AddRecipient "Errors", "saurabhz@gmail.com"
'Mailer.Subject    = "error ocurred"
'Mailer.BodyText   = messagebody
'Mailer.SendMail
response.write(messagebody)
%>
<html>
<head>
</head>
<body topmargin="0" leftmargin="0">
<table width="100%" height="100%">
<Tr><td align="center" valign="middle">
<table width="20%" height="20%" border="1" cellpadding="0" cellspacing="0" bordercolor="#000000">
<tr><td align="center" valign="middle">Sorry for inconvenience.<br>An internal server error has occured</td></tr>
</table>
</td></Tr>
</table>
</body>
</html>