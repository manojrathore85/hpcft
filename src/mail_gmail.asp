<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<%

mail

function mail()

		'allRecipients = setRecipients(Recipients,CC,BCC,assignto)
		FromName1 = "Atul Agrawal"
		ReplyTo = "atul@elensoft.com"
		Subject = "Test mail from gmail"
		Body= "Test mail body"
		dim receipients
		receipients = "agrawalatul@gmail.com;atul@elensoft.com;agrawalatul@live.com;saurabhz@gmail.com"
        pGmailEmail = "ims@elensoft.com"
        pGmailPassword = "B88735"
        Set myMail=CreateObject("CDO.Message")
        'Name or IP of remote SMTP server
        myMail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "smtp.gmail.com"
        'Server port
        myMail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 465
        myMail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing")= 2
        myMail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = "1" 'Use 0 for anonymous
        myMail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = pGmailEmail
        myMail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = pGmailPassword
        myMail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = "true"
        myMail.Configuration.Fields.Update
        myMail.Subject = Subject
        myMail.From = FromName1 & "<ims@elensoft.com>"
		'myMail.Bcc="agrawalatul@gmail.com"
		'myMail.Cc="atul007ag@yahoo.co.in"
'        if isarray(allRecipients) then
'			for i = 0 to ubound(allRecipients)
'				toaddr = chomp(allRecipients(i))
'				if i = 0 then
'	                receipients = toaddr
'                else
'                	receipients = receipients & ";"  & toaddr
'                end if
'			next
'		end if
		myMail.To = receipients
        myMail.TextBody = Body
        
		on error resume next
		myMail.Send
		if err.number <> 0 then
		response.write "Error number " & err.number & " Error Desc " & err.description
		err.clear
		response.End()
		end if
		on error goto 0
        set myMail=nothing

end function 


%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<title>Untitled Document</title>
</head>

<body>
</body>
</html>
