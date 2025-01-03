<%

' removes duplicate entries in the arrays sent as input
' and return a array with unique entries ( email addressess )

function setRecipients(arr1,arr2,arr3,assignto)
	Dim a,i
	Set a = CreateObject("Scripting.Dictionary")
	if isarray(arr1) then
		for i = 0 to ubound(arr1)
			if not a.exists(arr1(i)) then
				'if receipent_is_active( arr1(i) ) then
					a.add arr1(i),""
				'end if
			end if
		next
	end if
	if isarray(arr2) then
		for i = 0 to ubound(arr2)
			if not a.exists(arr2(i)) then
				'if receipent_is_active( arr2(i) ) then
					a.add arr2(i),""
				'end if
				'a.add arr2(i),""
			end if
		next
	end if
	
	if isarray(arr3) then
		for i = 0 to ubound(arr3)
			if not a.exists(arr3(i)) then
				'if receipent_is_active( arr3(i) ) then
					a.add arr3(i),""
				'end if
				'a.add arr3(i),""
			end if
		next
	end if
	
	if not a.exists(assignto) then	
		a.add assignto,""
	end if
	
	setRecipients = a.keys
	
end function


' mail function using persist mailsender object
function persist_mail(FromName1,ReplyTo,Recipients,CC,BCC,Subject,Body)

		allRecipients = setRecipients(Recipients,CC,BCC,assignto)
		
		dim Mailer
		Set Mailer = Server.CreateObject("Persits.MailSender")
		Mailer.FromName   = FromName1
		Mailer.From = varSiteSpecFromAddress
		Mailer.AddReplyTo  ReplyTo
		Mailer.Host = varSiteSpecRemoteHost
	    Mailer.Username = varSiteSpecEmailFromUsername
	    Mailer.Password = varSiteSpecEmailFromPassword
		if isarray(allRecipients) then
			for i = 0 to ubound(allRecipients)
				toaddr = chomp(allRecipients(i))
				Mailer.AddAddress toaddr,""
			next
		end if
	

'		if CC <> "" then
'			Mailer.AddCC CC, ""
'		end if
'		
'		if BCC <> "" then
'			Mailer.AddBCC BCC, ""
'		end if
		
       Mailer.Subject =  Subject
	   Mailer.Body = Body
	   on error resume next
	   Mailer.Send
	   
end function 
' 
function mail(FromName1,ReplyTo,Recipients,CC,BCC,Subject,Body)
'on error resume next
''
'response.write( FromName1 & "<br />") 
'response.write( ReplyTo & "<br />" )
'response.write( Recipients & "<br />" )
for i = 0 to ubound(Recipients)
	'response.write(Recipients(i))
			
		next
'response.write( CC & "<br />" )
'response.write( BCC & "<br />" )
'response.write( assignto & "<br />" )
'response.write( Subject & "<br />" )
'response.write( Body )
'
'		if err.number <> 0 then
'			response.write "Error number -" & err.number & "| Error Desc -" & err.description & "| Error Line -" & err.line
'			err.clear
'		end if


		allRecipients = setRecipients(Recipients,CC,BCC,assignto)
		dim receipients
        pGmailEmail = "info@hpchft.ai"
        pGmailPassword = "j4d8%5QrzN_8ZMb"
        Set myMail=CreateObject("CDO.Message")
        'Name or IP of remote SMTP server
        'myMail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "smtp.gmail.com"
		myMail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "smtpout.secureserver.net"
		'"relay-hosting.secureserver.net
        'Server port
        'myMail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 465
		myMail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 587
        myMail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing")= 2
        myMail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = "1" 'Use 0 for anonymous
        myMail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = pGmailEmail
        myMail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = pGmailPassword
        'myMail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = "true"
		myMail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpusetls") = "true"
        myMail.Configuration.Fields.Update
		if err.number <> 0 then
			response.write "Error number -" & err.number & "| Error Desc -" & err.description & "| Error Line -" & err.line
			err.clear
		end if
        myMail.Subject = Subject
        myMail.From = FromName1 & "<" & varSiteSpecFromAddress & ">"
        
        if isarray(allRecipients) then
			for i = 0 to ubound(allRecipients)
				toaddr = chomp(allRecipients(i))
				if i = 0 then
	                receipients = toaddr
                else
                	receipients = receipients & ";"  & toaddr
                end if
			next
		end if
		'response.write(receipients & "<br />")
		if err.number <> 0 then
			response.write "Error number -" & err.number & "| Error Desc -" & err.description & "| Error Line -" & err.line
			err.clear
		end if
        myMail.To = receipients
		if IsHTMLContent(Body) then
			myMail.HtmlBody = Body
		else
			myMail.TextBody = Body
		end if 
		
		if ReplyTo <> "" then
			myMail.ReplyTo = ReplyTo & ";" & varSiteSpecFromAddress
		else
			myMail.ReplyTo = varSiteSpecFromAddress
		end if
  		
		on error resume next
		myMail.Send
		if err.number <> 0 then
		response.write "Error number " & err.number & " Error Desc " & err.description
		err.clear
		response.End()
		end if

        set myMail=nothing

end function


%>
