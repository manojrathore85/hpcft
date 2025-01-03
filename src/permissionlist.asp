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
     dim i ' as incremental variable
   dim objview
   set objview =server.CreateObject("adodb.recordset") ' creating the recordset 
   objview.open "select * from " &  varTblNamePermissions & "",con
%>
<script>
function checkDelete()
{
	if (confirm('Are you sure you want to delete this permission.'))
		return true;
	else
		return false;
}
</script>

<table width="95%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td class="SectionHead">Permission<br>
      &nbsp; </td>
  </tr>
  <tr> 
    <td>Below is the list of all the projects on the site.<br>
      <br>
      ----<a class="bluebold" href="viewprojects.asp?mainlink=addPermission"> Add 
      Permission</a> </td>
  </tr>
</table>
<BR>
<table width="95%" border="0" cellpadding="0" cellspacing="0" bgcolor="#CCCCCC">
<tr>
  <td><table width="100%" border="0" cellpadding="1" cellspacing="1">
	  <tr bgcolor="#FFFFFF" class=head> 
		  <td id=head>&nbsp;Permission Name</td>
		  <td id=head>&nbsp;Read</td>
		  <td id=head>&nbsp;Write</td>
		  <td id=head>&nbsp;Add</td>
		  <td id=head>&nbsp;Delete</td>
		  <td>&nbsp;Details</td>
		<td>&nbsp;Operations</td>
	  </tr>
 <% 
    if objview.bof=false and objview.eof =false then ' checking whether the record is exist or not
	   while not objview.eof ' checkout all the user in database
%>	 
	  <tr bgcolor="#FFFFFF"> 
		  <td>&nbsp;<%=objview("permissionid")%></a> 
          </td>
		  <td>&nbsp;<%=objview("pread")%></td>
		  <td>&nbsp;<%=objview("pwrite")%></td>
		  <td>&nbsp;<%=objview("padd")%></td>
		  <td>&nbsp;<%=objview("pdelete")%></td>
		  <td>&nbsp;<%=objview("description")%></td>
		  <td>&nbsp;<a href="viewprojects.asp?mainlink=editPermission&id=<%=objview("permissionid")%>">Edit</a> 
            | <a href="deletepermission.asp?id=<%=objview("permissionid")%>" onClick="return checkDelete();">Delete</a></td>
	  </tr>
   <%
	    objview.movenext 
		wend
	else
	%>
	<tr bgcolor="#FFFFFF"> 
    <td colspan="7" align="center">NO PERMISSIONS CURRENTLY</td></tr>
	<%

    end if		
   %>

	</table></td>
</tr>
</table>
<%
 ' nullifying all the objects
  objview.close
  set objview= nothing
  con.close
  set con = nothing
 %> 
