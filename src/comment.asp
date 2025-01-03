<%@Language="VBScript" CodePage=65001 %>
<!-- #include file="checksession.asp" -->
<!-- #include file="generalfunctions.asp" -->
<!-- #include file ="Connect.asp" -->
<!-- #include file ="sendmail.asp" -->

<%
	
	if request("submit_comment") = "Submit Comment" and trim(request("txtcomments")) <> "" then
	
			dim objcomment
			set objcomment =server.CreateObject("adodb.recordset")
				objcomment.CursorLocation=3
				objcomment.Open "select * from " &  varTblNameIssueChangesComments & " where 1=2",con,3,3
				objcomment.AddNew
				objcomment("projectid")=session("projectid")
				objcomment("issueid")= request("issueid")
				strComment = trim(request("txtcomments"))			'24May06Atul
				objcomment("comments")= strComment 			'24May06Atul 
				objcomment("updateDate")=now()
				objcomment("changesBy")=session("user")
				objcomment.Update
				objcomment.Close 
				set objcomment = nothing
				dim update_query
				update_query = "update " & varTblNameIssues & " set UpdateDate = now() where projectid=" & clng(session("projectid")) & " and issueid = " & clng(request("issueid"))
				'response.write(update_query)
				'response.end()
				con.execute update_query
				
				
				
				call sendmail
				call send_response
				response.redirect("issuedetails.asp?issueid=" & request("issueid") & "&prjid=" & session("projectid"))
	end if
private sub send_response()
%>
	<table border=1 width="100%">
					   <tr bgcolor="#CCCCCC" class="regularTextBold"> 	
						<td colspan="3"> Comments By : &nbsp; <%=changesby%> &nbsp; &nbsp; [ <%= DateAdd("h", +1, now())%> EST ] </td>
						</tr>
						<tr>
						<td colspan="3" id="comnt<%'=counter%>" style="overflow:scroll;"><%=formatData(trim(request("txtcomments")))%></td>	
						</tr></table><br>
<%
end sub
private sub sendmail()
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

		Subject    =  "(" & varSiteSpecMailSubjectPrefix & ")" & varSpecialclientid & " | " & "Comment Added - " & rstemp("summary") & " | " & rstemp("projectname") 'CR20060821Atul
		Body   = "Project :-" & rstemp("projectname") & vbcrlf & "Issue Summary:-" & rstemp("summary") & vbcrlf & varSiteSpecURL & Mid(varServerVariablePath_info, 1, InStrRev(varServerVariablePath_info, "/")) & "/issuedetails.asp?issueid=" & request("issueid") & vbcrlf & "Comments :" & vbcrlf & request("txtcomments")	
		BCC = ""

		mail FromName,ReplyTo,Receipent,varSiteSpecAddCC,BCC,Subject,Body
	end if
	rstemp.close
	set rstemp = nothing
end sub
%>
