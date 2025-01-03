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
	if rsperm("pdelete") = "T" then
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

private sub sendmail_old(comments)
	dim rstemp
	dim i
	dim reportUsers ' users added in reportusers of a issue
	dim reportUsers1 ' users added in projectwatchlist
	dim rptuser ' this variable is used to keep the value in recordset, for it is does not come after checking the recordset for eof
	dim projectwatchlist ' this variable is used to keep the value in recordset, for it is does not come after checking the recordset for eof
	dim temp
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
		Mailer.ReplyTo = session("user")
		Mailer.RemoteHost = varSiteSpecRemoteHost
		Mailer.AddRecipient "",rstemp("assignto")
		if isarray(reportUsers) then
			for i = 0 to ubound(reportUsers)
				Mailer.AddRecipient "",replace(replace(replace(reportUsers(i),chr(13),""),chr(32),""),chr(10),"")
			next
		end if
		if isarray(reportUsers1) then
			for i = 0 to ubound(reportUsers1)
				Mailer.AddRecipient "",replace(replace(replace(reportUsers1(i),chr(13),""),chr(32),""),chr(10),"")
			next
		end if
		if isarray(reportUsers2) then
			for i = 0 to ubound(reportUsers2)
				Mailer.AddRecipient "",replace(replace(replace(reportUsers2(i),chr(13),""),chr(32),""),chr(10),"")
			next
		end if
		if varSiteSpecAddCC <> "" then
			Mailer.AddCC "", varSiteSpecAddCC
		end if
		'Mailer.Subject    =  varSiteSpecMailSubjectPrefix & "<" & varSpecialclientid & ">" & "A comment deleted <" & rstemp("summary") & "><" & rstemp("projectname") & ">"
		Mailer.Subject    =  "(" & varSiteSpecMailSubjectPrefix & ")" & varSpecialclientid & " | " & "A comment deleted - " & rstemp("summary") & " | " & rstemp("projectname") 'CR20060821Atul
		Mailer.BodyText   = "Project :-" & rstemp("projectname") & vbcrlf & "Issue Summary:-" & rstemp("summary") & vbcrlf & varSiteSpecURL & Mid(varServerVariablePath_info, 1, InStrRev(varServerVariablePath_info, "/")) & "issuedetails.asp?issueid=" & request("issueid") & vbcrlf & "Deleted by : " & session("Name") & vbcrlf & "Comments :- " & vbcrlf & comments
		Mailer.SendMail
	end if
	rstemp.close
	set rstemp = nothing
end sub

private sub sendmail(comments)
	dim rstemp
	dim i
	dim reportUsers ' users added in reportusers of a issue
	dim reportUsers1 ' users added in projectwatchlist
	dim rptuser ' this variable is used to keep the value in recordset, for it is does not come after checking the recordset for eof
	dim projectwatchlist ' this variable is used to keep the value in recordset, for it is does not come after checking the recordset for eof
	dim temp
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
		ReplyTo = session("user")

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
		if isarray(reportUsers2) then
			for i = 0 to ubound(reportUsers2)
				redim preserve Receipent(ubound(Receipent)+1)    
				Receipent(ubound(Receipent)) = reportUsers2(i)
				'Mailer.AddRecipient "",replace(replace(replace(reportUsers2(i),chr(13),""),chr(32),""),chr(10),"")
			next
		end if
		Subject    =  "(" & varSiteSpecMailSubjectPrefix & ")" & varSpecialclientid & " | " & "A comment deleted - " & rstemp("summary") & " | " & rstemp("projectname") 'CR20060821Atul
		Body   = "Project :-" & rstemp("projectname") & vbcrlf & "Issue Summary:-" & rstemp("summary") & vbcrlf & varSiteSpecURL & Mid(varServerVariablePath_info, 1, InStrRev(varServerVariablePath_info, "/")) & "issuedetails.asp?issueid=" & request("issueid") & vbcrlf & "Deleted by : " & session("Name") & vbcrlf & "Comments :- " & vbcrlf & comments
		BCC = ""

		mail FromName,ReplyTo,Receipent,varSiteSpecAddCC,BCC,Subject,Body
	end if
	rstemp.close
	set rstemp = nothing
end sub

' function to return the comments
function getComments
	dim rs
	set rs = server.CreateObject("adodb.recordset")
	rs.open "select * from " &  varTblNameIssueChangesComments & " where Rowid=" & clng(request("id")),con
	if not rs.eof then
		getComments  = rs("Comments")
	end if
	rs.close
	set rs = nothing
end function
%>

<%
if request("id") <> "" then
	dim comments
	dim rs
	comments = getComments
	con.execute "delete from " &  varTblNameIssueChangesComments & " where Rowid=" & clng(request("id"))
	call sendmail(comments)
end if
con.close
set con = nothing
response.Redirect("issuedetails.asp?issueid=" & request("issueid"))
%>
