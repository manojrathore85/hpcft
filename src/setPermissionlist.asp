<!-- ********************Description : This the file which is basically for showing the all the project list which are avail for us -->
<!-- ****************************** Date :--  29 march 2k5 ***************************************** -->
<!-- #include file="checksession.asp" -->
<!-- #include file="connect.asp" --> <!-- includeing the connection file -->

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
dim rs
set rs = server.CreateObject("adodb.recordset")
rs.open "select up.firstname,up.lastname,p.projectname,u.* from " &  varTblNameUserProfile & " up," &  varTblNameProjects & " p," &  varTblNameUsers & " u where u.email = up.email and u.projectid = p.projectid",con
%>
<script>
function checkDelete()
{
	if (confirm('Are you sure, you want to delete permission assigned to user.'))
		return true;
	else
		return false;
}
</script>

<table width="95%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td class="SectionHead">Permission Provided<br>
      &nbsp; </td>
  </tr>
  <tr> 
    <td>Below is the list of all the projects on the site.<br>
      <br>
      ----<a class="bluebold" href="viewprojects.asp?mainlink=addsetpermission"> 
      Set Permission</a> </td>
  </tr>
</table>
<BR>
<table width="95%" border="0" cellpadding="0" cellspacing="0" bgcolor="#CCCCCC">
<tr>
  <td><table width="100%" border="0" cellpadding="1" cellspacing="1">
        <tr bgcolor="#FFFFFF" class=head> 
          <td id=head>&nbsp;User Name</td>
          <td id=head>Email-id</td>
          <td id=head>Project Name</td>
          <td id=head>Permission</td>
          <td>&nbsp;Operations</td>
        </tr>
        <% 
    if rs.bof=false and rs.eof =false then ' checking whether the record is exist or not
	   while not rs.eof ' checkout all the user in database
%>
        <tr bgcolor="#FFFFFF"> 
          <td>&nbsp;<%=rs("firstname") & " " & rs("lastname") %> </td>
          <td>&nbsp;<%=rs("email")%></td>
          <td>&nbsp;<%=rs("projectname")%></td>
          <td>&nbsp;<%=rs("permissionid")%></td>
          <td>&nbsp;<a href="viewprojects.asp?mainlink=editsetpermission&projectid=<%=rs("projectid")%>&emailid=<%=rs("email")%>">Edit</a> 
            | <a href="deleteuserpermission.asp?projectid=<%=rs("projectid")%>&emailid=<%=rs("email")%>" onClick="return checkDelete();">Delete</a></td>
        </tr>
        <%
	    rs.movenext 
		wend
	else
	%>
        <tr bgcolor="#FFFFFF"> 
          <td colspan="5" align="center">NO PERMISSIONS GIVEN TO USERS</td>
        </tr>
        <%

    end if		
   %>
      </table></td>
</tr>
</table>
<%
 ' nullifying all the objects
 rs.close
 set rs = nothing
  con.close
  set con = nothing
 %> 
