<%@LANGUAGE="VBSCRIPT"%>
<!-- #include file="checksession.asp" -->
<!-- #include file ="Connect.asp" -->
<!-- #include file="generalfunctions.asp" -->
<!-- #include file ="sendmail.asp" -->
<%

' ********************* CHECKING PERMISSION ***************************

dim rsperm
set rsperm = server.CreateObject("adodb.recordset")
rsperm.open "select perm.* from " &  varTblNamePermissions & " perm, " &  varTblNameUsers & " u, " &  varTblNameissues & _
 " i where i.projectid = u.projectid and u.permissionid = perm.permissionid and i.issueid = " &  request("issueid") & _
 " and u.email ='" &  session("user") & "'",con
if not rsperm.eof then
	if rsperm("pwrite") = "T" then
		rsperm.close
		set rsperm = nothing
	else
		server.Execute("permissiondenied.asp")
		rsperm.close
		set rsperm = nothing
		response.End()
	end if
else
		server.Execute("permissiondenied.asp")
		rsperm.close
		set rsperm = nothing
		response.End()
end if

' ********************* CHECKING PERMISSION ENDS***************************
%>

<%
on error resume next
if request("issueid") <> "" and request("status") <> "" then
	dim strstatus
	dim prevstatus
	dim rs
	set rs= server.CreateObject("adodb.recordset")
	con.begintrans
	rs.open "select status from " &  varTblNameIssues & " where issueid=" & request("issueid"),con,3,3
	if not rs.eof then
		prevstatus = rs(0)
	end if
	rs.close
	set rs = nothing

	strstatus = request("status")
	
	select case strstatus
	case "startprogress"
		con.execute "update " &  varTblNameIssues & " set status ='In Progress',updatedate=now() where issueid=" & request("issueid")
		con.execute "insert into " &  varTblNameIssueChangesComments & "(ProjectId,IssueId,Field,ChangesBy,UpdateDate,NewValue,Comments,LastValue) values(" & session("projectid") & "," & request("issueid") & ",'Status','" & session("user") & "',now(),'In Progress',null,'" & prevstatus & "')"
	case "stopprogress"
		con.execute "update " &  varTblNameIssues & " set status ='Stop Progress',updatedate=now() where issueid=" & request("issueid")
		con.execute "insert into " &  varTblNameIssueChangesComments & "(ProjectId,IssueId,Field,ChangesBy,UpdateDate,NewValue,Comments,LastValue)  values(" & session("projectid") & "," & request("issueid") & ",'Status','" & session("user") & "',now(),'Stop Progress',null,'" & prevstatus & "')"
	case "resolve"
		con.execute "update " &  varTblNameIssues & " set status ='Resolved',updatedate=now() where issueid=" & request("issueid")
		con.execute "insert into " &  varTblNameIssueChangesComments & "(ProjectId,IssueId,Field,ChangesBy,UpdateDate,NewValue,Comments,LastValue)  values(" & session("projectid") & "," & request("issueid") & ",'Status','" & session("user") & "',now(),'Resolved',null,'" & prevstatus & "')"
	case "close"
		con.execute "update " &  varTblNameIssues & " set status ='Closed',updatedate=now() where issueid=" & request("issueid")
		con.execute "insert into " &  varTblNameIssueChangesComments & "(ProjectId,IssueId,Field,ChangesBy,UpdateDate,NewValue,Comments,LastValue)  values(" & session("projectid") & "," & request("issueid") & ",'Status','" & session("user") & "',now(),'Closed',null,'" & prevstatus & "')"
	case "reopen"
		con.execute "update " &  varTblNameIssues & " set status ='Reopened',updatedate=now() where issueid=" & request("issueid")
		con.execute "insert into " &  varTblNameIssueChangesComments & "(ProjectId,IssueId,Field,ChangesBy,UpdateDate,NewValue,Comments,LastValue)  values(" & session("projectid") & "," & request("issueid") & ",'Status','" & session("user") & "',now(),'Reopened',null,'" & prevstatus & "')"
	end select
	con.committrans
	If Err.Number<>0 then
		response.Write(err.description & "<br>")
		response.End()
	end if

	call sendmail
	con.close
	set con = nothing
	response.Redirect("issuedetails.asp?issueid=" & request("issueid"))
end if

private sub sendmail_old()
	dim rstemp
	dim i
	dim reportUsers ' users added in reportusers of a issue
	dim reportUsers1 ' users added in projectwatchlist
	dim rptuser ' this variable is used to keep the value in recordset, for it is does not come after checking the recordset for eof
	dim projectwatchlist ' this variable is used to keep the value in recordset, for it is does not come after checking the recordset for eof
	set rstemp = server.CreateObject("adodb.recordset")
	rstemp.open "select p.projectname,p.projectwatchlist,i.summary,i.assignto,i.reportUsers from " &  varTblNameProjects & " p," &  varTblNameIssues & " i where i.projectid = p.projectid and i.issueid =" & request("issueid") ,con
	projectwatchlist = rstemp("projectwatchlist")
	rptuser = rstemp("reportUsers")
	if not rstemp.eof then
		if not isnull(rptuser) and rptuser <> "" then
			reportUsers = split(rptuser,",")
		end if
		if not isnull(projectwatchlist) and projectwatchlist <> "" then
			reportUsers1 = split(projectwatchlist,",")
		end if
		dim Mailer
		Set Mailer = Server.CreateObject("SMTPsvg.Mailer")
		Mailer.FromName   = session("name")
		Mailer.FromAddress = varSiteSpecFromAddress 
		Mailer.RemoteHost = varSiteSpecRemoteHost
		Mailer.AddRecipient "",rstemp("assignto")
		if isarray(reportUsers) then
			for i = 0 to ubound(reportUsers)
				Mailer.AddCC "",reportUsers(i)
			next
		end if
		if isarray(reportUsers1) then
			for i = 0 to ubound(reportUsers1)
				Mailer.AddCC "",reportUsers1(i)
			next
		end if
		if varSiteSpecAddCC <> "" then
			mailer.addCC "", varSiteSpecAddCC
		end if
		Mailer.Subject    = varSiteSpecMailSubjectPrefix & "<" & varSpecialclientid & ">" & " ( " & session("user") & " ) ( " & rstemp("summary") & " )"
		'Mailer.BodyText   = "Project :-" & rstemp("projectname") & vbcrlf & "Issue Summary:-" & rstemp("summary") & vbcrlf & varSiteSpecURL & "/issuedetails.asp?issueid=" & request("issueid") & vbcrlf & "Status : " & request("status")
		Mailer.BodyText   = "Project :-" & rstemp("projectname") & vbcrlf & "Issue Summary:-" & rstemp("summary") & vbcrlf & varSiteSpecURL & Mid(request.ServerVariables("PATH_INFO"), 1, InStrRev(request.ServerVariables("PATH_INFO"), "/")) & "issuedetails.asp?issueid=" & request("issueid") & vbcrlf & "Status : " & request("status")
		Mailer.SendMail
	end if
	rstemp.close
	set rstemp = nothing
end sub

private sub sendmail()
	dim rstemp
	dim i
	dim reportUsers ' users added in reportusers of a issue
	dim reportUsers1 ' users added in projectwatchlist
	dim rptuser ' this variable is used to keep the value in recordset, for it is does not come after checking the recordset for eof
	dim projectwatchlist ' this variable is used to keep the value in recordset, for it is does not come after checking the recordset for eof
	set rstemp = server.CreateObject("adodb.recordset")
	rstemp.open "select p.projectname,p.projectwatchlist,i.summary,i.assignto,i.reportUsers from " &  varTblNameProjects & " p," &  varTblNameIssues & " i where i.projectid = p.projectid and i.issueid =" & request("issueid") ,con
	projectwatchlist = rstemp("projectwatchlist")
	rptuser = rstemp("reportUsers")
	if not rstemp.eof then
		if not isnull(rptuser) and rptuser <> "" then
			reportUsers = split(rptuser,",")
		end if
		if not isnull(projectwatchlist) and projectwatchlist <> "" then
			reportUsers1 = split(projectwatchlist,",")
		end if
		
		FromName   = session("name")
		dim Receipent()
		redim Receipent(0)  
		Receipent(0) = rstemp("assignto")

	
		if isarray(reportUsers) then
			for i = 0 to ubound(reportUsers)
				redim preserve Receipent(ubound(Receipent)+1)    
				Receipent(ubound(Receipent)) = reportUsers(i)
				'Mailer.AddRecipient "",replace(replace(replace(reportUsers(i),chr(13),""),chr(32),""),chr(10),"")
			next
		end if
		if isarray(reportUsers1) then
			for i = 0 to ubound(reportUsers1)
				redim preserve Receipent(ubound(Receipent)+1)    
				Receipent(ubound(Receipent)) = reportUsers1(i)
				'Mailer.AddRecipient "",replace(replace(replace(reportUsers1(i),chr(13),""),chr(32),""),chr(10),"")
			next
		end if
		
		Subject    = varSiteSpecMailSubjectPrefix & "<" & varSpecialclientid & ">" & " ( " & session("user") & " ) ( " & rstemp("summary") & " )"
		'Mailer.BodyText   = "Project :-" & rstemp("projectname") & vbcrlf & "Issue Summary:-" & rstemp("summary") & vbcrlf & varSiteSpecURL & "/issuedetails.asp?issueid=" & request("issueid") & vbcrlf & "Status : " & request("status")
		Body   = "Project :-" & rstemp("projectname") & vbcrlf & "Issue Summary:-" & rstemp("summary") & vbcrlf & varSiteSpecURL & Mid(request.ServerVariables("PATH_INFO"), 1, InStrRev(request.ServerVariables("PATH_INFO"), "/")) & "issuedetails.asp?issueid=" & request("issueid") & vbcrlf & "Status : " & request("status")
		BCC = ""

		mail FromName,ReplyTo,Receipent,varSiteSpecAddCC,BCC,Subject,Body

	end if
	rstemp.close
	set rstemp = nothing
end sub
%>
