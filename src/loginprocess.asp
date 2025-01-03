<!--     Description : The file is basically for login & password for user if use give correct id and password the it move to ims_index.asp -->
<!--     Date        : 22/march/2k5      -->

<!-- #include file="connect.asp" -->
<!-- #include file="generalfunctions.asp" -->
<!-- #include file="sendmail.asp" -->
<%
   'if page is post back by submit button then this event fire
 if request("cmdSubmit")<> "" then
 	' remove all the session objects if exist currently
	call manage_cookies
	
	session.Contents.Remove("projectid")
	session.Contents.Remove("permission")
	session.Contents.Remove("user")
	session.Contents.Remove("name")
	session.Contents.Remove("client")
   dim objrsUserLogin  'varible for recordset
   dim sqlstr ' for query string
   dim flag ' for storing the result
   dim UserName
   dim i
   flag = false
	 set objrsUserLogin =server.CreateObject("adodb.recordset") ' creating the object for recordset
  		
		sqlstr= "select * from " &  varTblNameUserProfile & " where email='" & replace(request("txtUserName"),"'","") & "' and password='" & replace(request("txtPassword"),"'","") & "'"   
			objrsUserLogin.Open sqlstr,con
			if not objrsUserLogin.eof then
				flag = true
				UserName = objrsUserLogin("firstname") & " " & objrsUserLogin("lastname")
			end if
			'temp =cint(objrsUserLogin(0))

		'Destroying the un used object  
		if flag = false then
			  objrsUserLogin.Close() 
			  set objrsUserLogin  =nothing
			  con.close
			  set con = nothing	
			  dim Receipent()
			  redim Receipent(1)
			  Receipent(0) = varSiteSpecFromAddress
			  'Receipent(1) = "saurabhz@gmail.com"
			  'BCC = "saurabz@gmail.com"
			  Subject = "Invalid userlogin and password attempt"
		      Body = "Invalid userlogin and Password Attempt" & vbcrlf & "Username: " & request("txtusername") & vbcrlf & "Password: " & request("txtpassword")
			  From = "IMS"
			  replyto = ""
			  
			  mail From ,replyto ,Receipent, varSiteSpecAddCC, BCC, Subject, Body
			  
			  if request("txtpath")="website_index" then 'finding the repost direction
					Response.redirect("website_index.asp?message=Invalid userlogin and Password&issueid=" & request("issueid"))
			  end if
			  Response.redirect("login_index.asp?message=Invalid userlogin and Password&issueid=" & request("issueid"))
        else
		         'session.Timeout = 300
				 session.Timeout = 1440 ' for a day
		 		 session("user") =request("txtUserName") 'creating the session
				 session("name") = username
				 session("client") = varSpecialclientid
				 dim rsIms_Usage
				set rsIms_Usage = server.CreateObject("adodb.recordset")
				rsIms_Usage.cursorlocation = 3
				rsIms_Usage.open "select * from " &  varTblNameIms_Usage & " where 1=2",con,3,3
				rsIms_Usage.AddNew 
				rsIms_Usage("email") = session("user")
				rsIms_Usage("u_datetime") = now()
				rsIms_Usage("action")= "login"
				rsIms_Usage("subaction") = request.ServerVariables("QUERY_STRING")
				rsIms_Usage.update
				rsIms_Usage.close
				set rsIms_Usage = nothing
				 
				 if request("issueid") <> ""  then
				 	if objrsUserLogin.state = 1 then objrsUserLogin.close
					'objrsUserLogin.open "select permissionid,projectid from users where email = '" & request("txtUserName") & "' and projectid = (select projectid from issues where issueid = " & request("issueid") & " )",con
					if request("prjid") <> "" then
						sqlstr = "select u.permissionid,u.projectid from " &  varTblNameUsers & " u," &  varTblNameIssues & " i where i.projectid = u.projectid and i.issueid =" & request("issueid") & " and i.projectid=" & request("prjid") & " and u.email = '" & request("txtUserName") & "'"
					else
						sqlstr = "select u.permissionid,u.projectid from " &  varTblNameUsers & " u," &  varTblNameIssues & " i where i.projectid = u.projectid and i.issueid =" & request("issueid") & " and u.email = '" & request("txtUserName") & "'"
					end if
					
					objrsUserLogin.open sqlstr,con
					if not objrsUserLogin.eof then
						session("projectid") = objrsUserLogin("projectid")
						session("permission") = objrsUserLogin("permissionid")
						objrsUserLogin.Close() 
						set objrsUserLogin  =nothing
						con.close
						set con = nothing	
						response.Redirect("issuedetails.asp?issueid=" & request("issueid") & "&prjid=" & request("prjid"))
					end if
				 end if
				  if objrsUserLogin.state = 1 then objrsUserLogin.Close
				  set objrsUserLogin  =nothing
				  con.close
				  set con = nothing	
				 response.Redirect("selectproject.asp")
		 end if
  end if
sub manage_cookies
	if request("RemEmail") = "true" then
		response.Cookies("user") = request("txtUserName")
		response.Cookies("user").expires = date + 15
	else
		response.Cookies("user").expires = now - 2
	end if
	if request("RemPassword") = "true" then
		response.Cookies("password") = request("txtPassword")
		'response.Cookies("user").expires = date + 15
	else
		response.Cookies("password").expires = now - 2
	end if
end sub

%>

