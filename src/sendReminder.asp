<%
' *******************  COMMENTS  ***************************************
'Program to check the reminders in reminder_table of IMS and send mail for the purpose
'This program is called by a scheduler
' 


'atul	20060414		created
'atul	20061113		modified	checking day and month added

'***********************************************************************
%>

<!-- #include file ="Connect.asp" -->
<!-- #include file="generalfunctions.asp" -->
<!-- #include file="sendmail.asp" -->
<%


dim aspCallInterval

if request("inTime") <> "" then
	aspCallInterval= request("inTime")
else
	aspCallInterval= 30
end if

' ************** start of functions and procedures *************************************
' *************   returns all the mail recipients of the issue in an array
' function setRecipients(arr1,arr2,assignto)
	' Dim a,i
	' Set a = CreateObject("Scripting.Dictionary")
	' if isarray(arr1) then
		' for i = 0 to ubound(arr1)
			' if not a.exists(arr1(i)) then
					' a.add arr1(i),""
			' end if
		' next
	' end if
	' if isarray(arr2) then
		' for i = 0 to ubound(arr2)
			' if not a.exists(arr2(i)) then
				' a.add arr2(i),""
			' end if
		' next
	' end if
	
	' if not a.exists(assignto) then	
		' a.add assignto,""
	' end if
	
	' setRecipients = a.keys
	
' end function

' **************************  sends mail to the users of issue ************************************************* 
' sub sendMail(issueid,typeCol,duedate,personResponsible,alertFreq,Notes,descCol)
	' debug = "------coming<br>"
	' response.Write(debug)
	' hndl_file.writeline(debug)
	' dim rs
	' dim i
	' dim reportUsers ' users added in reportusers of a issue
	' dim reportUsers1 ' users added in projectwatchlist
	' dim reportUsers2 ' users added in the form 'One Time Email'
	' dim rptuser ' this variable is used to keep the value in recordset, for it does not come after checking the recordset for eof
	' dim projectwatchlist ' this variable is used to keep the value in recordset, for it is does not come after checking the recordset for eof
	' Dim mailsTo
	' dim temp
	' dim subject
	' dim bodyText
	' set rs = server.CreateObject("adodb.recordset")
	' rs.open "select p.projectname,p.projectwatchlist,i.summary,i.assignto,i.reportUsers from " &  varTblNameProjects & " p," &  varTblNameIssues & " i where i.projectid = p.projectid and i.issueid =" & issueid ,con
	' projectwatchlist = rs("projectwatchlist")
	' rptuser = rs("reportUsers")
	' if not rs.eof then
		' ' getting all the mail receipients in array
		' if not isnull(rptuser) and rptuser <> "" then
			' reportUsers = split(rptuser,",")
		' end if
		' if not isnull(projectwatchlist) and projectwatchlist <> "" then
			' reportUsers1 = split(projectwatchlist,",")
		' end if
		
		' mailsTo = setRecipients(reportUsers,reportUsers1,rs("assignto"))
		
		
		' dim Mailer
		' Set Mailer = Server.CreateObject("SMTPsvg.Mailer")
		' Mailer.FromName   = session("name")
		' Mailer.FromAddress = varSiteSpecFromAddress
		' Mailer.ReplyTo = session("user")
		' Mailer.RemoteHost = varSiteSpecRemoteHost
		
		' for i = 0 to ubound(mailsTo)
			' response.Write(mailsTo(i) & vbTab)
			' Mailer.AddRecipient "",replace(replace(replace(mailsTo(i),chr(13),""),chr(32),""),chr(10),"")
		' next
			' response.Write("<br>")
			' Mailer.AddRecipient "",personResponsible
		' if varSiteSpecAddCC <> "" then
			' Mailer.AddCC "", varSiteSpecAddCC
		' end if
		' subject 		  =  varSiteSpecMailSubjectPrefix & _
							 ' "Reminder<" & varSpecialclientid & ">" & _
							 ' " <" & Notes & ">" & _
							 ' " <" & rs("summary") & "><" & rs("projectname") & ">"
							 
		' Mailer.Subject = subject
		
		' bodyText   =  "Reminder for project :-" & rs("projectname") & vbcrlf & _
							 ' "Issue Summary:-" & rs("summary") & vbcrlf & varSiteSpecURL & _
						  			' Mid(varServerVariablePath_info, 1, InStrRev(varServerVariablePath_info, "/")) & _
									' "issuedetails.asp?issueid=" & issueid & vbcrlf & _
							  ' "Reminder Type :" & typeCol & vbcrlf & _
							  ' "Due Date :" & duedate & vbcrlf & _
							  ' "Person Responsible :" & personResponsible & vbcrlf & _
							  ' "Alert Frequency :" & alertFreq & vbcrlf & _
							  ' "Summary :" & Notes & vbcrlf & _
							  ' "Description :" & descCol
		
		' Mailer.BodyText = bodyText
		
							  
		' Mailer.SendMail
		
		' 'response.Write("Reminder for issue summary:- " & rs("summary"))

	' end if
	' rs.close
	' set rs = nothing
' end sub

sub sendMail(issueid,typeCol,duedate,personResponsible,alertFreq,Notes,descCol)
	debug = "in sendmail procedure<br>"
	response.Write(debug)
	hndl_file.writeline(debug)
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
				'Mailer.AddRecipient "",replace(replace(replace(reportUsers1(i),chr(13),""),chr(32),""),chr(10),"")
			next
		end if
		
		
		FromName   = session("name")
		ReplyTo = session("user")
		
		subject 		  =  varSiteSpecMailSubjectPrefix & _
							 "Reminder<" & varSpecialclientid & ">" & _
							 " <" & rs("summary") & "><" & rs("projectname") & ">"
							 
		
		body   =  "Reminder for project :-" & rs("projectname") & vbcrlf & _
							 "Issue Summary:-" & rs("summary") & vbcrlf & varSiteSpecURL & _
						  			Mid(varServerVariablePath_info, 1, InStrRev(varServerVariablePath_info, "/")) & _
									"issuedetails.asp?issueid=" & issueid & vbcrlf & _
							  "Reminder Type :" & typeCol & vbcrlf & _
							  "Due Date :" & duedate & vbcrlf & _
							  "Person Responsible :" & personResponsible & vbcrlf & _
							  "Alert Frequency :" & alertFreq & vbcrlf & _
							  "Summary :" & Notes & vbcrlf & _
							  "Description :" & descCol
		
		
							  
		BCC = ""

		mail FromName,ReplyTo,Receipent,varSiteSpecAddCC,BCC,Subject,Body


	end if
	rs.close
	set rs = nothing
end sub

function checkMin(min_remain) ' 20061113 frequency,days_or_months parameter added
	checkMin = false
	
	'if min_remain >= 30 and min_remain <= (aspCallInterval + 30) then
		
	'end if
	
	select case typeCol
	case "Major_After"
		
		' if last_call for due date is negative, then email the reminder
		last_call = min_remain - aspCallInterval
		
		if last_call <= 0 then
			checkMin = true
		end if

	case "Major_Before"
		' if last_call for due date is positive, then email the reminder
		last_call = min_remain + aspCallInterval
		
		if last_call >= 0 then
			checkMin = true
		end if
		
	end select
	
end function 

function check_type_condition
	dim last_call
	select case typeCol 
	case "Major_After"
		'means difference in min should be positive
		if diff_in_min => 0 then
			typeCondition = true
			'diff_in_min = abs(diff_in_min) + 60  ' + 60 is done so that we can send the mail earlier before the duedate
		end if
	case "Major_Before"
		'means difference in min should be negative
		if diff_in_min <= 0 then
			typeCondition = true
			'diff_in_min = abs(diff_in_min) + 30  ' + 30 is done so that we can send the mail earlier before the duedate
		end if		
	case "On_DueDate"
		' if last_call for due date is positive, means we have to mail the reminder
		last_call = diff_in_min + aspCallInterval
		
		' This tells that if diff_in_min is negative and (diff_in_min + aspCallInterval) is positive then we should send the mail.
		' This will be called last call to mail.
		if diff_in_min <= 0 and last_call >= 0 then
			flag_mail = true
		end if

	end select
end function

function check_alert_freq_condition
	if alertFreq = "daily" and daysofweekcol <> "" then
		if instr(1,daysofweekcol,weekdayname(weekday(date_time),true)) <> 0  then
			alertFreqCondition = true
		end if
		'end if
	elseif frequency = "monthly" and monthsofyearcol <> "" then
		if instr(1,monthsofyearcol,monthname(month(date_time),true)) <> 0  then
			alertFreqCondition = true
		end if		
	else
		alertFreqCondition = true
	end if
end function

function calc_min_remain
	
	select case alertFreq
	case "hourly"
		min_remain = diff_in_min  mod 60 
		flag_mail = checkMin(min_remain)
	case "daily"  ' 60 * 24
		min_remain = diff_in_min mod 1440 
		flag_mail = checkMin(min_remain)
	case "weekly"' 60 * 24 * 7
		min_remain = diff_in_min mod 10080 
		flag_mail = checkMin(min_remain)
	case "monthly" ' 60 * 24 * 30
		min_remain = diff_in_min mod 43200 
		flag_mail = checkMin(min_remain)
	case "yearly" ' 60 * 24 * 365
		min_remain = diff_in_min mod 525600 
		flag_mail = checkMin(min_remain)
	end select

end function

function day_light_saving_figure
	Dim dst_starts_2007
	Dim dst_ends_2007
	Dim dst_starts_2008
	Dim dst_ends_2008
	Dim dst_starts_2009
	Dim dst_ends_2009
	Dim diff_in_min
	
	dst_starts_2007 = "3/11/2007 1:59:59 AM"
	dst_ends_2007 = "11/4/2007 1:59:59 AM"
	dst_starts_2008 = "3/9/2008 1:59:59 AM"
	dst_ends_2008 = "11/2/2008 1:59:59 AM"
	dst_starts_2009 = "3/8/2009 1:59:59 AM"
	dst_ends_2009 = "11/1/2009 1:59:59 AM"
	
	
	if (DateDiff("s", dst_starts_2007, Now) > 0 and DateDiff("s", dst_ends_2007, Now) <= 0) or (DateDiff("s", dst_starts_2008, Now) > 0 and DateDiff("s", dst_ends_2008, Now) <= 0) or ( DateDiff("s", dst_starts_2009, Now) > 0 and DateDiff("s", dst_ends_2009, Now) <= 0 ) then
		day_light_saving_figure = -240
	else
		day_light_saving_figure = -300
	end if

end function
' ********************  end of functions and procedures		************************************



' **************    programs starts here  **************************
' *****************************************************************************************************************************
dim rs
dim date_time  ' current date time converted into IST,EST,GMT , depending upon the time zone selected with due date
dim typeCondition ' flag to keep whether current date falls in major_after,major_before,or on_duedate clause.
dim alertFreqCondition ' flag to keep whether duedate pass the alertFreq condition
dim min_remain  ' difference of duedate and current date in minutes
dim flag_mail ' flag to keep whether mail is to be sent or not
dim varServerVariablePath_info
dim issueid
dim typeCol
dim duedateTime
dim perResp
dim alertFreq
dim notes
dim timezone
dim remid
dim descCol
dim daysofweekcol
dim monthsofyearcol
dim debug

' #############   START SETTINGS TO DEBUG REMINDERS ############################
dim fso
dim hndl_file
dim path
path = "\upload\log_reminders\reminders" & Replace(Date, "/", "_") & ".htm"
set fso = server.CreateObject("scripting.filesystemobject")
if fso.fileexists(server.MapPath(".") & path ) then
	Set hndl_file = fso.OpenTextFile(server.MapPath(".") & path, 8)  ' 8 is for appending
else
	Set hndl_file = fso.CreateTextFile(server.MapPath(".") & path)
end if
' #############   END SETTINGS TO DEBUG REMINDERS ############################
varServerVariablePath_info = request.ServerVariables("PATH_INFO")
set rs = server.CreateObject("adodb.recordset")
rs.open "select * from reminder_table where stateCol='open'",con
while not rs.eof

	typeCondition = false
	flag_mail = false
	alertFreqCondition = false
	
	issueid = rs("issueidCol")
	typeCol = rs("TypeCol")
	duedateTime = rs("due_datetimecol") 
	timezone =  rs("TimeZoneCol")
	perResp = rs("person_ResponsibleCol")
	alertFreq = rs("AlertFrequencyCol")
	notes = rs("NotesCol")
	descCol = rs("DescCol")
	remid = rs("rem_idCol")
	daysofweekcol = rs("daysofweekcol")
	monthsofyearcol = rs("monthsofyearcol")
	
'	select case rs("TimeZoneCol") ' as shivankz server uses GMT time
	' select case timezone ' as shivankz server uses GMT time
	' case "IST"
		' date_time = DateAdd("n", 330, now())
	' case "EST"
		' 'date_time = DateAdd("n", -300, now())
		' date_time = DateAdd("n", day_light_saving_figure, now())
	' case "GMT"
		' date_time = now()
	' end select

	select case timezone ' as local server uses IST time
	case "IST"
		date_time = now()
	case "EST"
		date_time = DateAdd("n", -630, now())
	case "GMT"
		date_time = DateAdd("n", -330, now())
	end select
	


	diff_in_min = datediff("n",duedateTime,date_time) ' will give positive integer if date_time variable is bigger else negative
	debug = diff_in_min & " type-" & typeCol & " dd:" & duedatetime & " ct:" & now() & " date_time:" & date_time & " id-" & remid & "<br>"
	response.Write(debug)
	hndl_file.writeline(debug)
	
	check_type_condition


	if typeCondition = true then
		'diff_in_min = abs(diff_in_min) + 60  ' + 60 is done so that we can send the mail earlier before the duedate
		debug = "------type condition is true<br>"
		response.Write(debug)
		hndl_file.writeline(debug)

		check_alert_freq_condition
		
	end if
	
	if alertFreqCondition = true then
		
		debug = "------ DOW:(" & daysofweekcol & ")- MOY:(" & monthsofyearcol & ")- frequency condition is true<br>"
		response.write(debug)

		hndl_file.writeline(debug)
		calc_min_remain
		
		debug = "------ AF:" & alertFreq & " - min_remain:" & min_remain & "<br>"
		response.Write(debug)
		hndl_file.writeline(debug)
		
	end if

	if flag_mail = true then

		'sendMail_test issueid,typeCol,duedateTime & " (" & timezone & ")" ,perResp,alertFreq,notes,descCol
		sendMail issueid,typeCol,duedateTime & " (" & timezone & ")" ,perResp,alertFreq,notes,descCol

	end if
	
	rs.movenext
wend
hndl_file.writeline("----------------------------^ " & now & " ^------------------------------------<BR>")
hndl_file.close

%>