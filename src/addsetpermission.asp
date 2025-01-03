<!-- Description :- This file is to add user --->
<!-- Date 22 march 2k5 -->
<!-- #include file ="checksession.asp" -->
<!-- #include file ="connect.asp" -->
<%

' ********************* CHECKING PERMISSION ***************************

dim rsperm
set rsperm = server.CreateObject("adodb.recordset")
rsperm.open "select * from " &  varTblNamePermissions & " where permissionid = '" & session("permission") & "'",con
if not rsperm.eof then
	if rsperm("permissionid") = "all" and rsperm("pread") = "T" and rsperm("pwrite") = "T" and rsperm("padd") = "T" and rsperm("pdelete") = "T" then
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
dim message
function check_user_project(userid,projectid)
	dim rs
	dim sql
	set rs = server.CreateObject("adodb.recordset")
	sql = "select email from " &  varTblNameUsers & " where projectid=" & projectid & " and email = '" & userid & "'"
	rs.open sql , con
	if rs.eof then
		check_user_project = "not present"		' set status
	else
		sql = "update " & varTblNameUsers & " set permissionid = '" & request("optPermission") & "' where projectid=" & projectid & " and email = '" & userid & "'"
		'response.Write(sql & "<br>")
		con.execute sql
	end if
	rs.close
	set rs = nothing
end function
sub set_permission(userid)
	dim sql
	dim i
	For i = 1 To Request("chkProjects").Count
		if check_user_project(userid,Request("chkProjects")(i)) = "not present" then
			sql = "insert into " & varTblNameUsers & " values ('" & userid & "'," & Request("chkProjects")(i) & ",'" & Request("optPermission") & "')"
		'	response.Write(sql & "<br>")
			con.execute sql
		end if
	next
end sub
dim i
if request("cmdSubmit") = "Set Permission" then 
   if (request("chkUsers") <> "" and request("chkProjects") <> "" and request("optPermission") <> "")  then
   		For i = 1 To Request("chkUsers").Count
			set_permission(Request("chkUsers")(i))
		next
		message = "Permission set/updated as given."
	else
		message = "User/Project/Permission not provided to set the permission."
	end if
	'con.close
	'set con = nothing
	'response.Redirect("viewprojects.asp?mainlink=setpermissionlist")
end if

function print_users
	dim rs
	set rs = server.CreateObject("adodb.recordset")
	rs.open "select firstname,lastname,email from " &  varTblNameUserProfile & "",con
	if not rs.eof then %><table><%
		while not rs.eof %><tr><td><input type="checkbox" name="chkUsers" value="<%=rs("email")%>"></td><td><%=rs("firstname") & " " & rs("lastname") %></td></tr><%
			rs.movenext
		wend%></table><%
	end if
end function
sub print_projects
	dim rs
	set rs = server.CreateObject("adodb.recordset")
	rs.open "select projectid,projectname from " &  varTblNameProjects & "",con
	if not rs.eof then %><table><%
		while not rs.eof %><tr><td><input type="checkbox" name="chkProjects" value="<%=rs("projectid")%>"></td><td><%=rs("projectname")%></td></tr><%
			rs.movenext
		wend%></table><%
	end if
end sub
sub print_permissions
	dim rs
	set rs = server.CreateObject("adodb.recordset")
	rs.open "select permissionid from " &  varTblNamePermissions & "",con
	if not rs.eof then %><table><%
		while not rs.eof %><tr><td><input type="radio" name="optPermission" value="<%=rs("permissionid")%>"></td><td><%=rs("permissionid")%></td></tr><%
			rs.movenext
		wend%></table><%
	end if
end sub		 
	
%>
<script language ="Javascript">

 function checkvalidation()
  {
  		var flag=0;
		for(i=0;i<frmaddsetpermission.chkUsers.length;i++)
		{
			if(frmaddsetpermission.chkUsers(i).checked == true)
			{
				flag=1;
			}
		}
		if(flag != 1) { alert('Please select a user'); return false; } //returning if user is not selected
		 
		for(i=0;i<frmaddsetpermission.chkProjects.length;i++)
		{
			if(frmaddsetpermission.chkProjects(i).checked == true)
			{
				flag=2;
			}
		}

		if(flag != 2) { alert('Please select a project'); return false; } //returning if Project is not selected

		for(i=0;i<frmaddsetpermission.optPermission.length;i++)
		{
			if(frmaddsetpermission.optPermission(i).checked == true)
			{
				flag=3;
			}
		}		
		
		if(flag != 3) { alert('Please select a permission'); return false; } //returning if permission is not selected
  } 
</script>
<!-- <form name="frmaddpermission" method="post" action="addpermission.asp" onsubmit="return checkvalidation();">  --><!-- form start from here -->
<form name="frmaddsetpermission" method="post" action="viewprojects.asp" onsubmit="return checkvalidation();"> <!-- form start from here -->
<input type="hidden" value="addsetpermission" name="mainlink">
<table width="95%" border="0" cellspacing="0" cellpadding="0" height="100%">
		<tr> 
		  <td valign="top"> <font color="#FFFFFF">-</font><br> 
		  <%
		  if message <> "" then
		  %>
		  <p align="center"><%=message%></p>
		  
        <%
		  end if
		  %>
        <table width="90%" border="0" align="center" cellpadding="0" cellspacing="0">
          <tr bgcolor="#F3F3F3"> 
            <td height="43" class="redbold"><div align="center" class="redbold">Give 
                Permission to User</div></td>
          </tr>
          <tr>
            <td height="35"><br /><table width="100%" border="0">
              <tr bgcolor="#FFFFFF" class=head>
                <td>&nbsp;Users</td>
                <td>&nbsp;Projects</td>
                <td>&nbsp;Permission</td>
              </tr>
              <tr>
                <td valign="top">&nbsp;<%print_users%></td>
                <td valign="top">&nbsp;<%print_projects%></td>
                <td valign="top">&nbsp;<%print_permissions%></td>
              </tr>
            </table></td>
          </tr>
          <tr> 
            <td height="35" align="left" valign="top"><br />&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
              <input type="submit" value="Set Permission" name="cmdSubmit" class="formbutton">
&nbsp;&nbsp;<input  type="reset" value="Cancel" name="cmdCancel" class="formbutton"></td>
          </tr>
        </table>
      </td>
    </tr></table>
 </form>
</body>
</html>
<%
con.close
set con = nothing 
%>