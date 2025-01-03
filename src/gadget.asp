<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!-- #include file="connect.asp" -->
<!-- #include file="generalFunctions.asp" -->
<!-- #include file="sendmail.asp" -->
<!-- #include file="IMSAdminEmail.asp" -->
<%
dim Username

if request("option") = "help" then
	do_help	
	response.End()
end if

if request("user") <> "" and request("pass") <> "" then
	uname = replace(request("user"),"'","") 
	password= replace(request("pass"),"'","") 
	if check_username(uname,password) then
		select case request("option")
		case ""
			print_issues(uname)
		case "outstanding"
			print_issues(uname)
		case "addcomments"
			add_comments(uname)
		case "show_details"
			show_details request("projectid"), request("issueid")
		end select
	else
		response.write("<?xml version='1.0' encoding='UTF-8' ?>")
		response.write("<message>Invalid Username/Password.&lt;br&gt;Use Edit Setting to enter correct username/password.</message>")
	end if
end if

function check_username(uname,password)
	flag = false
	sqlValidUser= "select * from " &  varTblNameUserProfile & " where email='" & uname & "' and password='" & password & "'"
	set rsValidUser = server.CreateObject("adodb.recordset")
	rsValidUser.Open sqlValidUser,con
	if not rsValidUser.eof then
		flag = true
		UserName = rsValidUser("firstname") & " " & rsValidUser("lastname")
		session("name") = UserName
		session("user") = uname
	end if
	check_username = flag
end function

sub print_issues(uname)
	
	number_of_issues = request("number_of_issues")
	summmary_length_in_tab = request("summmary_length_in_tab")
	ims_project_id = request("ims_project_id")

	
	sqlIssues ="select i.issueid,i.description,i.summary,p.projectname,p.projectid from " &  varTblNameIssues & " i," &  varTblNameUsers & " u," &  varTblNameProjects & " p where i.projectid = u.projectid and i.projectid = p.projectid and u.email ='" & uname & "' " 
	if trim(ims_project_id) <> "" then
		sqlIssues = sqlIssues & " and projectid =" & ims_project_id 
	end if
	sqlIssues = sqlIssues & " order by updatedate desc "
	
	if trim(number_of_issues) <> "" then
		sqlIssues = sqlIssues & " limit " & number_of_issues
	else
		sqlIssues = sqlIssues & " limit 3" 
	end if
	
	set rs = server.CreateObject("adodb.recordset")
	rs.open sqlIssues,con
	
  response.write("<?xml version='1.0' encoding='UTF-8' ?>")
  response.write("<issues title='Outstanding Issues'>")
  response.write("<response></response>")
	counter = 0
	while not rs.eof
	projectname = rs("projectname")
	summary = rs("summary")
	issueid = rs("issueid")
	desc = rs("description")
	projectid = rs("projectid")
	if trim(summmary_length_in_tab) <> "" then
		tab_title = mid(summary,1,summmary_length_in_tab)
	else 
		tab_title = mid(summary,1,10)
	end if
	%>
  <issue>
  	 <projectName><%=projectname%></projectName>
  	 <tabText><%=tab_title%>..</tabText>
     <summary><%'=summary%></summary> 
     <description><![CDATA[<br /><div id="data<%=counter%>" class="details"><b>Summary: <%=summary%></b><br /><br />
	 <b>Description:</b><br /><%=desc%></div><br />
<div style="padding-left:5px;">
<input type="hidden" name="issueid" id="issueid" value="<%=issueid%>" />
<input type="hidden" name="projectid" id="projectid" value="<%=projectid%>" />
<textarea cols="28" rows="5" name="txtComments<%=counter%>" id="txtComments<%=counter%>"></textarea>
<br />
<input name="cmdSubmit" type="button" value="Submit" onclick="ims_add_comments(this.form,'txtComments<%=counter%>');" />&nbsp;<input name="cmdDetails" type="button" value="Details" onclick="showDetails(this.form);" />&nbsp;<input type="reset" value="Reset" />
</div>
]]>
</description> 
  </issue>
<%
	counter = counter + 1
	rs.movenext
	wend
	response.write("</issues>")
	rs.close
	con.close
end sub

function add_comments(uname)
	on error resume next
	if request("issueid") <> "" and trim(request("comments")) <> "" then
		if check_permission(request("issueid")) = false then
			response.write("Write permission denied on the issue.")
		else
			dim strComment		'24May06Atul
			set objcomment =server.CreateObject("adodb.recordset")
				objcomment.CursorLocation=3
				objcomment.Open "select * from " &  varTblNameIssueChangesComments & " where 1=2",con,3,3
				objcomment.AddNew
				objcomment("projectid")=request("projectid")
				objcomment("issueid")= request("issueid")
				strComment = trim(request("comments"))			'24May06Atul
				objcomment("comments")= strComment 			'24May06Atul 
				objcomment("updateDate")=now()
				objcomment("changesBy")=uname
				objcomment.Update
				objcomment.Close 
				set objcomment = nothing
				sendmail request("projectid"), request("issueid"), strComment 
				con.close
	
			if err.number <> 0 then 
				response.write("There was a error while entering comments" & vbcrlf & err.description )
				err.clear()
			else 
				response.write("Comments added successfully")
				'response.write("Comments added successfully!Projectid:" & request("projectid") & "|Issue:" & request("issueid"))
			end if
		end if
	end if
end function

function show_details(projectid,issueid)
	'on error goto HandlerE
		   dim objissue ' issue change comments 
		   set objissue =server.CreateObject("adodb.recordset") 'setting the record oobhject for issue chages goes here 
			   'objissue.open "select * from " &  varTblNameIssueChangesComments & " where projectid =" & projectid & " and issueid=" & issueid & " order by updateDate",con
			   dim sql 
			   sql = "select *,if(comments is null,'',comments) as newcomments  from " &  varTblNameIssueChangesComments & " where issueid=" & issueid & " order by updateDate"
			   objissue.open sql,con
		   dim dattemp
		   dim flag
		   	dattemp = ""
			counter = 1
			flag = false
			if objissue.eof =false then 
				while not objissue.eof
					if  dattemp <> objissue("updatedate") and objissue("newcomments") = "" then
					dattemp =objissue("updatedate")	 
         	 
html = html & "<table border=1 width='100%' cellpadding='0' cellspacing='0'><tr class='regularTextBold'><td colspan='3'> Change By : &nbsp;" & objissue("changesby") & " &nbsp; &nbsp; [ " & objissue("updatedate") & " ] &nbsp;</td></tr><tr  bgcolor='#CCCCCC' class='regularTextBold'><td width='20%'>Field</td><td width='40%'> Old Value </td><td width='40%'> New Value</td>					</tr>"
					do
						if objissue.eof = true then
							exit do
						end if
						if (objissue("newcomments") = "") and (dattemp = objissue("updatedate")) then
							dattemp = objissue("updatedate")
						   html = html & "<tr><td>" & objissue("field") & "</td><td>" & formatData(objissue("lastvalue")) & "</td><td>" & formatData(objissue("newvalue")) & "</td></tr>"
					   else
							exit do
					   end if
					   objissue.movenext
				   loop
				   'while (not )
				   flag = true
				   html =html & "</table><br>"
				elseif objissue("newcomments") <> "" then
					flag = false
					   html =html & "<table border=1 width='100%'><tr bgcolor='#CCCCCC' class='regularTextBold'><td colspan='3'> Comments By : &nbsp; " & objissue("changesby") & " &nbsp; &nbsp; [ " & objissue("updatedate") & " ]</td></tr><tr><td colspan='3' id='' style='overflow:scroll;'>" & formatData(objissue("newcomments")) & "</td></tr></table><br>"
					counter = counter + 1
					end if 
				if objissue.eof = false and flag <> true  then
		         	 objissue.movenext
				end if
				
			  wend
		end if
			objissue.close
			set objissue = nothing	  
			con.close
			html = html & "</td></tr></table>"
			response.write(html)
end function

private sub sendmail(projectid,issueid,strComment)
	dim rstemp
	dim i
	dim reportUsers ' users added in reportusers of a issue
	dim reportUsers1 ' users added in projectwatchlist
	dim reportUsers2 ' users added in the form 'One Time Email'
	dim rptuser ' this variable is used to keep the value in recordset, for it is does not come after checking the recordset for eof
	dim projectwatchlist ' this variable is used to keep the value in recordset, for it is does not come after checking the recordset for eof
	dim temp
	set rstemp = server.CreateObject("adodb.recordset")
	rstemp.open "select p.projectname,p.projectwatchlist,i.summary,i.assignto,i.reportUsers from " &  varTblNameProjects & " p," &  varTblNameIssues & " i where i.projectid = p.projectid and i.issueid =" & issueid & " and i.projectid = " & projectid ,con
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
			next
		end if
		if isarray(reportUsers1) then
			for i = 0 to ubound(reportUsers1)
				redim preserve Receipent(ubound(Receipent)+1)    
				Receipent(ubound(Receipent)) = reportUsers1(i)
			next
		end if
		if isarray(reportUsers2) then
			for i = 0 to ubound(reportUsers2)
				redim preserve Receipent(ubound(Receipent)+1)    
				Receipent(ubound(Receipent)) = reportUsers2(i)
			next
		end if

		Subject    = "(" & varSiteSpecMailSubjectPrefix & ")" & varSpecialclientid & " | " &  rstemp("summary") & " | " & rstemp("projectname") 'CR20060821Atul
		
			Body = "Project :-" & rstemp("projectname") & vbcrlf & "Issue Summary:-" & summary & vbcrlf & "URL: -" &  varSiteSpecURL & Mid(varServerVariablePath_info, 1, InStrRev(varServerVariablePath_info, "/")) & "issuedetails.asp?issueid=" & issueid & "&prjid=" & projectid & vbcrlf & "Comments :" & strComment	' 24May06Atul
		
		BCC = ""

		mail FromName,ReplyTo,Receipent,varSiteSpecAddCC,BCC,Subject,Body
		
	end if
	rstemp.close
	set rstemp = nothing
	
end sub

function check_permission(issueid)
	dim rsperm
	dim sstr
	set rsperm = server.CreateObject("adodb.recordset")
	check_permission = false
	sstr = "select perm.* from " &  varTblNamePermissions & " perm, " &  varTblNameUsers & " u, " &  varTblNameissues & _
	 " i where i.projectid = u.projectid and u.permissionid = perm.permissionid and i.issueid = " &  issueid & _
	 " and u.email ='" &  session("user") & "'"
	
	 rsperm.open sstr,con
	 
	if not rsperm.eof then
		if rsperm("pwrite") = "T" then
			rsperm.close
			set rsperm = nothing
			check_permission = true
		end if
	end if

end function

sub do_help
	
	dim action
	
	select case request("help_action")
	case "request_username"
		action = "New username requested."
	case "change_password"
		action = "Change password requested."
	case "report_issue"
		action = "Reporting a issue."
	end select 
	
	html = "Action: " & action & vbcrlf
	html = html & "Username: " & request("request_username") & vbcrlf
	html = html & "Current Password: " & request("current_password") & vbcrlf
	html = html & "New Password: " & request("new_password") & vbcrlf
	html = html & "Confirm Password: " & request("confirm_password") & vbcrlf
	html = html & "Notes: " & request("notes") & vbcrlf
	FromName   = session("name") 
	ReplyTo = session("user")
	dim Receipent(0)
	Receipent(0) = IMSAdminEmail
	
	mail session("name"),session("user"),Receipent,"","","Help requested from IMS google gadget",html
	
	response.write("Help request has been sent to the Administrator.")
	
end sub

%>
