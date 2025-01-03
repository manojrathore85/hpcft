<%@LANGUAGE="VBSCRIPT"%>
<!-- checking the session -->
<%  if session("user")="" then  %>
<HTML>
<HEAD>
<LINK href="include/style.css" type="text/css" rel="stylesheet">
</HEAD>
<BODY> 
 <table cellpadding="0" cellspacing="0" align="center">
	<tr>
		<td class="redbold" align="center"> Session Expired</td>
	</tr>
	<tr>
	     <td> Please Re-login <a href="login_index.asp"> Login Here</a></td>
	</tr>
	
	</table>
</BODY>
</HTML>

 <%
		 Response.End ()
 
	end if
  %>
 <!-- it checks the session whether the valid user or not -->
<!-- #include file="Connect.asp" --> <!-- including the connection file -->
<!-- #include file="generalFunctions.asp" --> <!-- including the functions file -->
<!-- ******************************************************** script language  Vb script coding goes here**************** -->
<%
	
	if request("cmdSubmit") = "Next >>"  then
		dim rsChkPrj
		set rsChkPrj = server.CreateObject("adodb.recordset")
		rsChkPrj.open "select permissionid from " &  varTblNameUsers & " where email='" & session("user") & "' and projectid =" & request("lstProjects"),con
		if not rsChkPrj.eof then
			session("projectid") = request("lstProjects")
			session("permission") = rsChkPrj("permissionid")
		'else
		end if
		rsChkPrj.close
		set rsChkPrj = nothing

		'if session("projects") <> "" then
		'	dim projects
		'	dim i
		'	projects = split(session("projects"),",")
		'	for i= 0  to ubound(projects)
		'		if projects(i) = request("lstProjects") then
		'			session("projectid") = request("lstProjects")
		'			i = ubound(projects) + 1
		'		end if
		'	next
			
			'if session("projectid") = "" then
		if session("projectid") <> request("lstProjects") then
			response.Write("<script>alert('You have no permission on this project');window.history.go(-1);</script>")
			con.close
			set con = nothing
			response.End()
		end if
	end if
		'elseif session("projectid") = "" then
		'	con.close
		'	set con = nothing
		'	response.End()
		'end if
	if session("projectid") = "" then
		response.Write("<script>alert('Please select the project');window.location='selectproject.asp';</script>") ' CR20070101
		con.close
		set con = nothing		
		response.End()
	end if
	
	if request("visitlink") <> "" then
		con.close
		set con = nothing
		response.redirect("issuenavigator.asp?search=alloutstanding")
		response.End()
	end if
	
	'session.Contents.Remove("projects")  '''''''''''''this can't be done as session("projects") 
										  '''''''''''''will be required every time users select the project
	'dim rsPerms
	'set rsPerms = server.CreateObject("adodb.recordset")
	'rsPerms.open "select permissionid from users where email = '" & session("user") & "' and projectid=" & session("projectid") ,con
	'if not rsPerms.eof then
	'	session("permission") = rsPerms("permissionid")
	'end if
	'rsPerms.close
	'set rsPerms = nothing
	'rsPerms.open "select u.permissionid,read,write,add,delete from users u,permissions p where u.permissionid = p.permission and email = '" & session("user") & "' and projectid = " & session("projectid"),con
	
     dim objview 'for viewing the projects
	 dim str
	 dim filename
	 dim counter
	 set objview =server.CreateObject("adodb.recordset") 
	 objview.open "select * from " &  varTblNameProjects & " where projectid=" & session("projectid") ,con ' recodset of projects 
	 str = objview("designdocpath")
	   if isnull(objview("designdocpath")) = false then
			 filename = split(str,",")
	   end if
%>
<!-- ***************************************  html coding goes here ************************************ -->

<html>
<head>
<title><%=setTitle%></title>
<LINK 
href="include/style.css" type=text/css rel=stylesheet>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1"></head>
<body bgcolor="#CCCCCC" leftmargin="0" topmargin="0">
<table width="100%" border="0" cellspacing="0" cellpadding="0" height="100%">
  <tr>
    <td valign="top" height="100%">
	  <table width="100%" height="100%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="#FFFFFF">
        <tr bgcolor="#006699"> 
          <td colspan="2"> <p><strong>&nbsp;<font color="#FFFFFF">Company Banner 
              and Information</font></strong></p>
            <p>&nbsp;</p></td>
        </tr>
        <tr align="left" valign="top"> 
          <td colspan="2"> 
            <div align="center">
			  <table width="100%" border="0" cellspacing="0" cellpadding="0" height="100%">
                <tr> 
                  <td valign="top"> <table width="98%" border="0" align="center" cellpadding="0" cellspacing="0">
                      <tr> 
                        <td width="66%">&nbsp;</td>
                        <td width="34%" align="left"><!-- #include file="include/TopRightNavBar.asp" --></td>
                      </tr>
                      <tr>
                        <td>&nbsp;</td>
                        <td>&nbsp;</td>
                      </tr>
                      <tr> 
                        <td colspan="2"> <!-- #include file="include/GeneralTopNavBar.asp" --></td>
                      </tr>
                    </table>
                    <font color="#FFFFFF">-</font><br>
                    <table width="80%" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#F3F3F3">
                      <tr>
                        <td><p class="redbold">&nbsp;&nbsp;<%=objview("projectName")%></p>
                          <table width="80%" border="0" align="center" cellpadding="0" cellspacing="0">
                            <tr> 
                              <td colspan="2"><strong>Project Lead :<%=objview("LeadDeveloper")%></strong></td>
                            </tr>
                            <tr> 
                              <td colspan="2">&nbsp;</td>
                            </tr>
                            <tr> 
                              <td colspan="2"><span class="regularTextBold">Description :-</span><br>
                                <%=replace(objview("description"),chr(13),"<br>")%></td>
                            </tr>
							<tr> 
                              <td colspan="2"><span class="regularTextBold">Design Docs :-</span>
                                <%
				if isArray(filename) then
					if ubound(filename) >= 0 then
						for counter=0 to ubound(filename)
						if counter > 0 then
							response.Write(", ")
						end if
						%>
						<a href="upload/<%=filename(counter)%>" target="_blank"><%=filename(counter)%></a>
						<%
						'response.Write(filename(counter))	
						next
					else
						'response.Write(filename(-1))
					end if
				else
					'response.Write("isArray not working here")
				end if
				'response.Write(objview("designdocpath"))
			%></td>
                            </tr>
                            <tr> 
                              <td colspan="2">&nbsp;</td>
                            </tr>
                            <tr> 
                              <td colspan="2"><font color="#000066"><strong>Filter 
                                Issues</strong></font></td>
                            </tr>
                            <tr> 
                              <td>&nbsp;</td>
                              <td>&nbsp;</td>
                            </tr>
                            <tr valign="bottom"> 
                              <td height="23">- <a href="issuenavigator.asp?search=all">All</a></td>
                              <td height="23">- <a href="issuenavigator.asp?search=resolved">Resolved recently</a></td>
                            </tr>
                            <tr valign="bottom"> 
                              <td height="23">- <a href="issuenavigator.asp?search=outstanding">Outstanding</a></td>
                              <td height="23">- <a href="issuenavigator.asp?search=addrecent">Added recently</a></td>
                            </tr>
                            <tr valign="bottom"> 
                              <td height="23">- <a href="issuenavigator.asp?search=asgndtome">Assigned to me</a></td>
                              <td height="23">- <a href="issuenavigator.asp?search=updaterecent">Updated recently</a></td>
                            </tr>
                            <tr valign="bottom"> 
                              <td height="23">- <a href="issuenavigator.asp?search=reportbyme">Reported by me</a></td>
                              <td height="23">- <a href="issuenavigator.asp?search=important">Most important</a> 
                              </td>
                            </tr>
                            <tr> 
                              <td height="23">&nbsp;</td>
                              <td height="23">&nbsp;</td>
                            </tr>
                          </table></td>
                      </tr>
                    </table>
                    
                  </td>
                </tr>
              </table>
              
            </div></td>
        </tr>
        <tr> 
          <td width="45%">&nbsp;</td>
          <td width="55%" align="center" valign="middle">&nbsp;</td>
        </tr>
        <tr> 
          <td>&nbsp;</td>
          <td align="center" valign="middle">&nbsp;</td>
        </tr>
        <tr bgcolor="#003366"> 
          <td colspan="2" align="center"><font color="#FFFFFF">Email Links etc 
            </font></td>
        </tr>
      </table>
   </td>
  </tr>
</table>
</body>
</html>
<%
objview.close
set objview = nothing
con.close
set con = nothing
%>
