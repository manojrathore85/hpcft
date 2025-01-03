<%
 Set myMail=CreateObject("CDO.Message")
        'Name or IP of remote SMTP server
        myMail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "email-smtp.us-east-2.amazonaws.com"
		myMail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "smtpout.secureserver.net"
        'Server port
        myMail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 587
        myMail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing")= 2
        myMail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = "1" 'Use 0 for anonymous
        'myMail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = "AKIAYDNQMB3VIYPLVDU3"
		myMail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = "info@hpchft.ai"
		
        'myMail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = "BOWgmsYuRYmPimR4uhO08tj9id2uPj2NLD8EXwpJNOSg"
		myMail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = "j4d8%5QrzN_8ZMb"
		
        myMail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpusetls") = "true"
        myMail.Configuration.Fields.Update
		if err.number <> 0 then
			response.write "Error number -" & err.number & "| Error Desc -" & err.description & "| Error Line -" & err.line
			err.clear
		end if
        myMail.Subject = "Test Mail"
        myMail.From = "HpcHpft Inc.<info@hpchft.ai>"
		myMail.To = "agrawalatul@gmail.com"
		myMail.ReplyTo ="hpchft.info@hpchft.com"
        myMail.TextBody = "test mail body."
  		
		on error resume next
		myMail.Send
		if err.number <> 0 then
			response.write "Error number " & err.number & " Error Desc " & err.description
			err.clear
			response.End()
		end if

        set myMail=nothing
%>