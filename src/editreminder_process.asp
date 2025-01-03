<%
' Program to process add/edit reminders
'Atul	200600414	
%>
<!-- #include file ="CheckSession.asp" --> 
<!-- #include file ="Connect.asp" -->
<!-- #include file="generalfunctions.asp" -->
<!-- #include file ="sendmail.asp" -->

<%

'******* procedure to send mail on reminder add or edit **************************
sub sendMail_old(issueid,typeCol,duedate,personResponsible,alertFreq,Notes,mode)

	dim rs
	dim i
	dim reportUsers ' users added in reportusers of a issue
	dim reportUsers1 ' users added in projectwatchlist
	dim reportUsers2 ' users added in the form 'One Time Email'
	dim rptuser ' this variable is used to keep the value in recordset, for it does not come after checking the recordset for eof
	dim projectwatchlist ' this variable is used to keep the value in recordset, for it is does not come after checking the recordset for eof
	Dim mailsTo
	dim temp
	dim subject
	dim bodyText
	set rs = server.CreateObject("adodb.recordset")
	rs.open "select p.projectname,p.projectwatchlist,i.summary,i.assignto,i.reportUsers from " &  varTblNameProjects & " p," &  varTblNameIssues & " i where i.projectid = p.projectid and i.issueid =" & issueid ,con
	projectwatchlist = rs("projectwatchlist")
	rptuser = rs("reportUsers")
	if not rs.eof then
		' getting all the mail receipients in array
		if not isnull(rptuser) and rptuser <> "" then
			reportUsers = split(rptuser,",")
		end if
		if not isnull(projectwatchlist) and projectwatchlist <> "" then
			reportUsers1 = split(projectwatchlist,",")
		end if
		
		mailsTo = setRecipients(reportUsers,reportUsers1,rs("assignto"))
		
		
		dim Mailer
		Set Mailer = Server.CreateObject("SMTPsvg.Mailer")
		Mailer.FromName   = session("name")
		Mailer.FromAddress = varSiteSpecFromAddress
		Mailer.ReplyTo = session("user")
		Mailer.RemoteHost = varSiteSpecRemoteHost
		
		for i = 0 to ubound(mailsTo)
			'response.Write(mailsTo(i) & vbTab)
			Mailer.AddRecipient "",replace(replace(replace(mailsTo(i),chr(13),""),chr(32),""),chr(10),"")
		next
			'response.Write("<br>")
			Mailer.AddRecipient "",personResponsible
		if varSiteSpecAddCC <> "" then
			Mailer.AddCC "", varSiteSpecAddCC
		end if
		subject 		  =  varSiteSpecMailSubjectPrefix & _
							 "Reminder<" & varSpecialclientid & ">" & _
							 " <" & rs("summary") & "><" & rs("projectname") & ">"
							 
		Mailer.Subject = subject
		
		bodyText   =  mode & ":Reminder for project :-" & rs("projectname") & vbcrlf & _
							 "Issue Summary:-" & rs("summary") & vbcrlf & varSiteSpecURL & _
						  			Mid(varServerVariablePath_info, 1, InStrRev(varServerVariablePath_info, "/")) & _
									"issuedetails.asp?issueid=" & issueid & vbcrlf & _
							  "Reminder Type :" & typeCol & vbcrlf & _
							  "Due Date :" & duedate & vbcrlf & _
							  "Person Responsible :" & personResponsible & vbcrlf & _
							  "Alert Frequency :" & alertFreq 

							  if request("alertFreq") = "daily" then
									bodyText = bodyText & vbcrlf & "Days of Week:" & request("chk_daily_weekly") & vbcrlf
							  elseif request("alertFreq") = "monthly" then
									bodyText = bodyText & vbcrlf & "Months of Year:" & request("chk_monthly_yearly") & vbcrlf
							  end if

		bodyText = bodyText & vbcrlf & "Summary :" & Notes & vbcrlf & _
							  "Description :" & request("Description")
		
		Mailer.BodyText = bodyText
		
							  
		Mailer.SendMail
		
		'response.Write("Reminder for issue summary:- " & rs("summary"))

	end if
	rs.close
	set rs = nothing
end sub

sub sendMail(issueid,typeCol,duedate,personResponsible,alertFreq,Notes,mode)

	dim rs
	dim i
	dim reportUsers ' users added in reportusers of a issue
	dim reportUsers1 ' users added in projectwatchlist
	dim reportUsers2 ' users added in the form 'One Time Email'
	dim rptuser ' this variable is used to keep the value in recordset, for it does not come after checking the recordset for eof
	dim projectwatchlist ' this variable is used to keep the value in recordset, for it is does not come after checking the recordset for eof
	Dim mailsTo
	dim temp
	dim subject
	dim bodyText
	set rs = server.CreateObject("adodb.recordset")
	rs.open "select p.projectname,p.projectwatchlist,i.summary,i.assignto,i.reportUsers from " &  varTblNameProjects & " p," &  varTblNameIssues & " i where i.projectid = p.projectid and i.issueid =" & issueid ,con
	projectwatchlist = rs("projectwatchlist")
	rptuser = rs("reportUsers")
	if not rs.eof then
		' getting all the mail receipients in array

		if not isnull(rptuser) and rptuser <> "" then
			reportUsers = split(rptuser,",")
		end if
		if not isnull(projectwatchlist) and projectwatchlist <> "" then
			reportUsers1 = split(projectwatchlist,",")
		end if
		dim Receipent()
		redim Receipent(1)  
		Receipent(0) = rs("assignto")
		Receipent(1) = personResponsible

		if isarray(reportUsers) then
			for i = 0 to ubound(reportUsers)
				redim preserve Receipent(ubound(Receipent)+1)    
				Receipent(ubound(Receipent)) = reportUsers(i)
			next
		end if

		if isarray(reportUsers1) then
			for i = 0 to ubound(reportUsers1)
				redim preserve Receipent(ubound(Receipent)+1)    
				Receipent(ubound(Receipent)) = reportUsers1(i)
			next
		end if
		
		
		
		FromName   = session("name")
		ReplyTo = session("user")
		
		subject 		  =  varSiteSpecMailSubjectPrefix & _
							 "Reminder<" & varSpecialclientid & ">" & _
							 " <" & rs("summary") & "><" & rs("projectname") & ">"
							 
		
		body   =  mode & ":Reminder for project :-" & rs("projectname") & vbcrlf & _
							 "Issue Summary:-" & rs("summary") & vbcrlf & varSiteSpecURL & _
						  			Mid(varServerVariablePath_info, 1, InStrRev(varServerVariablePath_info, "/")) & _
									"issuedetails.asp?issueid=" & issueid & vbcrlf & _
							  "Reminder Type :" & typeCol & vbcrlf & _
							  "Due Date :" & duedate & vbcrlf & _
							  "Person Responsible :" & personResponsible & vbcrlf & _
							  "Alert Frequency :" & alertFreq 

							  if request("alertFreq") = "daily" then
									bodyText = bodyText & vbcrlf & "Days of Week:" & request("chk_daily_weekly") & vbcrlf
							  elseif request("alertFreq") = "monthly" then
									bodyText = bodyText & vbcrlf & "Months of Year:" & request("chk_monthly_yearly") & vbcrlf
							  end if

		body = body & vbcrlf & "Summary :" & Notes & vbcrlf & _
							  "Description :" & request("Description")
		
		
		BCC = ""

		mail FromName,ReplyTo,Receipent,varSiteSpecAddCC,BCC,Subject,Body


	end if
	rs.close
	set rs = nothing
end sub


'******* function to check the permission ************************
function checkPermission
	dim rsperm
	set rsperm = server.CreateObject("adodb.recordset")
	rsperm.open "select perm.* from " &  varTblNamePermissions & " perm, " &  varTblNameUsers & " u, " &  varTblNameissues & _
	 " i where i.projectid = u.projectid and u.permissionid = perm.permissionid and i.issueid = " &  request("issueid") & _
	 " and u.email ='" &  session("user") & "'",con
	if not rsperm.eof then
		if rsperm("pwrite") = "T" then
			checkPermission = true
		else
			checkPermission = false
		end if
	else
		checkPermission = false
	end if
	rsperm.close
	set rsperm = nothing
end function

function saveData()

			dim rs
			dim dueDate
			dim projectid
			dim bodyText
			dueDate = request("year") & "-" & request("month") + 1 & "-" & request("day") & " " & request("hour") & ":" & request("min") & ":01"
			set rs =server.CreateObject("adodb.recordset")
			rs.CursorLocation=3
			rs.Open "select * from " &  varTblNameReminder_table & " where 1=2",con,3,3
			rs.AddNew
			rs("issueidCol") = request("issueid")
			rs("TypeCol") = request("lstType")
			rs("due_dateTimeCol") = dueDate
			rs("TimeZoneCol") = request("timeZone")
			rs("date_enteredCol") = now
			rs("person_responsibleCol") = request("personResponsible")
			rs("AlertFrequencyCol") = request("alertFreq")
			rs("stateCol") = request("state")
			rs("NotesCol") = request("notes")
			rs("DescCol") = request("Description")
			if request("alertFreq") = "daily" then
				rs("DaysOfWeekCol") = request("chk_daily_weekly")
			elseif request("alertFreq") = "monthly" then
				rs("MonthsOfYearCol") = request("chk_monthly_yearly")
			end if
			rs.Update
			rs.close
			
			dueDate = dueDate & "(" & request("timezone") & ")"
			'********     retreiving project id *********************
			rs.open "select projectid from issues where issueid = " & request("issueid"),con
			if not rs.eof then
				projectid = rs("projectid")
			end if
			rs.close
			
			bodyText   =  "Reminder Created :-" & vbcrlf & _
						  "Reminder Type :" & request("lstType") & vbcrlf & _
						  "Due Date :" & duedate
						  if request("alertFreq") = "daily" then
								bodyText = bodyText & vbcrlf & "Days of Week:" & request("chk_daily_weekly") & vbcrlf
						  elseif request("alertFreq") = "monthly" then
								bodyText = bodyText & vbcrlf & "Months of Year:" & request("chk_monthly_yearly") & vbcrlf
						  end if
			bodyText   = bodyText & "Person Responsible :" & request("personResponsible") & vbcrlf & _
						  "Alert Frequency :" & request("alertFreq") & vbcrlf & _
						  "Summary :" & request("notes") & vbcrlf & _
						  "Description :" & request("Description")
			
			'****************   logging  the reminder update in issuechangescomments  *********************
				rs.CursorLocation=3
				rs.Open "select * from " &  varTblNameIssueChangesComments & " where 1=2",con,3,3
				rs.AddNew
				rs("projectid")=projectid
				rs("issueid")= request("issueid")
				rs("comments")= bodyText
				rs("updateDate")=now()
				rs("changesBy")=session("user")
				rs.Update
				rs.Close 
			set rs = nothing
			
			con.execute "update " &  varTblNameIssues & " set updateDate= now() where issueId=" & request("issueid")
			
			sendmail request("issueid"),request("lstType"),duedate,request("personResponsible"),request("alertFreq"),request("Notes"),"Added"
			
end function

function updateData()
	dim rs
	dim dueDate
	dim projectid
	dim bodyText
	dim varState
	set rs =server.CreateObject("adodb.recordset")

	dueDate = request("year") & "-" & request("month") + 1 & "-" & request("day") & " " & request("hour") & ":" & request("min") & ":01"

	if request("alertFreq") = "daily" then
		con.execute "update " &  varTblNameReminder_table & " set typeCol='" & request("lstType") & "',due_dateTimeCol='" & duedate & "',TimezoneCol='" & request("timezone") & "',person_responsibleCol='" & request("personResponsible") & "',AlertFrequencyCol='" & request("alertFreq") & "',stateCol='" & request("state") & "',NotesCol='" & request("notes") & "',DescCol='" & request("Description") & "',DaysOfWeekCol='" & request("chk_daily_weekly") & "',MonthsOfYearCol='' where rem_idCol=" & request("remid")
	elseif request("alertFreq") = "monthly" then
		con.execute "update " &  varTblNameReminder_table & " set typeCol='" & request("lstType") & "',due_dateTimeCol='" & duedate & "',TimezoneCol='" & request("timezone") & "',person_responsibleCol='" & request("personResponsible") & "',AlertFrequencyCol='" & request("alertFreq") & "',stateCol='" & request("state") & "',NotesCol='" & request("notes") & "',DescCol='" & request("Description") & "',DaysOfWeekCol='',MonthsOfYearCol='" & request("chk_monthly_yearly") & "' where rem_idCol=" & request("remid")
	else
		con.execute "update " &  varTblNameReminder_table & " set typeCol='" & request("lstType") & "',due_dateTimeCol='" & duedate & "',TimezoneCol='" & request("timezone") & "',person_responsibleCol='" & request("personResponsible") & "',AlertFrequencyCol='" & request("alertFreq") & "',stateCol='" & request("state") & "',NotesCol='" & request("notes") & "',DescCol='" & request("Description") & "',DaysOfWeekCol='',MonthsOfYearCol='' where rem_idCol=" & request("remid")
	end if

	'con.execute "update " &  varTblNameReminder_table & " set typeCol='" & request("lstType") & "',due_dateTimeCol='" & duedate & "',TimezoneCol='" & request("timezone") & "',person_responsibleCol='" & request("personResponsible") & "',AlertFrequencyCol='" & request("alertFreq") & "',stateCol='" & request("state") & "',NotesCol='" & request("notes") & "',DescCol='" & request("Description") & "',DaysOfWeekCol='" & request("chk_daily_weekly") & "',MonthsOfYearCol='" & request("chk_monthly_yearly") & "' where rem_idCol=" & request("remid")
	dueDate = dueDate & "(" & request("timezone") & ")"
	'********     retreiving project id *********************
			rs.open "select projectid from issues where issueid = " & request("issueid"),con
			if not rs.eof then
				projectid = rs("projectid")
			end if
			rs.close
			
				bodyText   =  "Reminder Updated :-" & vbcrlf & _
							  "Reminder Type :" & request("lstType") & vbcrlf & _
							  "Due Date :" & duedate
  							  
							  if request("alertFreq") = "daily" then
							  		bodyText = bodyText & vbcrlf & "Days of Week:" & request("chk_daily_weekly") & vbcrlf
							  elseif request("alertFreq") = "monthly" then
								  	bodyText = bodyText & vbcrlf & "Months of Year:" & request("chk_monthly_yearly") & vbcrlf
							  end if
							  
				bodyText   = bodyText & "Person Responsible :" & request("personResponsible") & vbcrlf & _
							  "Alert Frequency :" & request("alertFreq") & vbcrlf & _
							  "Summary :" & request("notes") & vbcrlf & _
							  "Description :" & request("Description")

							  
			
			'****************   logging  the reminder update in issuechangescomments  *********************
				rs.CursorLocation=3
				rs.Open "select * from " &  varTblNameIssueChangesComments & " where 1=2",con,3,3
				rs.AddNew
				rs("projectid")=projectid
				rs("issueid")= request("issueid")
				rs("comments")= bodyText
				rs("updateDate")=now()
				rs("changesBy")=session("user")
				rs.Update
				rs.Close 
			set rs = nothing
	if request("state") = "close" then
		varState = "Closed"
	else
		varState = "Updated"
	end if
	
	con.execute "update " &  varTblNameIssues & " set updateDate= now() where issueId=" & request("issueid")
	
	sendmail request("issueid"),request("lstType"),duedate,request("personResponsible"),request("alertFreq"),request("Notes"),varState
end function

'******* procedures and function end here  **************************

'****** actual programming goes here *********************************
dim varServerVariablePath_info
varServerVariablePath_info = request.ServerVariables("PATH_INFO")

'****** checking permission ******************************************
if checkPermission = false then
	server.Execute("permissiondenied.asp")
	response.End()
end if





' ************  code to add or edit reminder  ************************************
if request("cmdSave") = "Save Reminder" then
	if request("remid") <> "" then
		call updateData
	else
		call saveData
	end if
end if

' ************  code to close a reminder
if request("remid") <> "" and request("close") = "true" then
	dim duedate
	dueDate = request("year") & "-" & request("month") + 1 & "-" & request("day") & " " & request("hour") & ":" & request("min") & ":01"
	con.execute "update " &  varTblNameReminder_table & " set stateCol='close' where rem_idCol=" & request("remid") 
	sendmail request("issueid"),request("lstType"),duedate,request("personResponsible"),request("alertFreq"),request("Notes"),"Closed"
end if
response.Redirect("editreminder.asp?issueid=" & request("issueid") & "&reload_opener=true")
%>