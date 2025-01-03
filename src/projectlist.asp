<!-- ********************Description : This the file which is basically for showing the all the project list which are avail for us -->
<!-- ****************************** Date :--  29 march 2k5 ***************************************** -->
<!-- #include file="checksession.asp" -->
<!-- #include file="connect.asp" --> <!-- including the connection file -->

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
if request("order") = "" then
	order = "desc"
else
	order = ""
end if

function get_orderby
	select case request("orderBy")
	case ""
		get_orderby = "createdate " & request("order")
	case "pname"
		get_orderby = "projectname " & request("order")
	case "lead"
		get_orderby = "leaddeveloper " & request("order")
	case "createdate"
		get_orderby = "createdate " & request("order")
	end select
end function

     dim i ' as incremental variable
   dim objview
   dim sql
   set objview =server.CreateObject("adodb.recordset") ' creating the recordset 
   sql =  "select * from " &  varTblNameProjects & " order by " & get_orderby
   objview.open sql , con 'order by createdate desc ",con
%>
<script>
function checkDelete()
{
	if (confirm('Do u really want to delete this project.\nThis will delete all the data related to this project from the database.'))
		return true;
	else
		return false;
}
function orderby(field)
{
	frm_plist.orderBy.value=field;
	frm_plist.submit();	
}
</script>
<form action="viewprojects.asp" method="post" name="frm_plist">
<input type="hidden" name="mainlink" value="">
<input type="hidden" name="orderBy">
<input type="hidden" name="order" value="<%=order%>">
</form>
<table width="95%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td class="SectionHead">Administration<br>
      &nbsp; </td>
  </tr>
  <tr> 
    <td>Below is the list of all the projects on the site.<br>
      <br>
      ----<a class="bluebold" href="viewprojects.asp?mainlink=addProject"> Add Project</a> </td>
  </tr>
</table>
<BR>
<table width="95%" border="0" cellpadding="0" cellspacing="0" bgcolor="#CCCCCC">
<tr>
  <td><table width="100%" border="0" cellpadding="1" cellspacing="1">
	  <tr bgcolor="#FFFFFF" class=head> 
		<td id=head>&nbsp;<a href="javascript:orderby('pname')">Name</a></td>
		<td id=head>&nbsp;<a href="javascript:orderby('createdate')">Create Date</a></td>
		<td>&nbsp;<a href="javascript:orderby('lead')">Lead</a></td>
		<td>&nbsp;Operations</td>
	  </tr>
 <% 
    if objview.bof=false and objview.eof =false then ' checking whether the record is exist or not
	   while not objview.eof ' checkout all the user in database
%>	 
	  <tr bgcolor="#FFFFFF"> 
		  <td>&nbsp;<a href="viewprojects.asp?mainlink=viewproject&id=<%=objview("projectId")%>"><%=objview("projectName")%></a> 
          </td>
		  <td>&nbsp;<%=DateAdd("h", -5, objview("createDate"))%></td>
		  <td>&nbsp;<%=objview("leadDeveloper")%></td>
		<td>&nbsp;<a href="viewprojects.asp?mainlink=viewproject&id=<%=objview("projectId")%>">View</a> | <a href="viewprojects.asp?mainlink=editproject&id=<%=objview("projectId")%>">Edit</a> 
		  | <a href="deleteproject.asp?id=<%=objview("projectId")%>" onClick="return checkDelete();">Delete</a></td>
	  </tr>
   <%
	    objview.movenext 
		wend
	else
	%>
	<tr bgcolor="#FFFFFF"> 
    <td colspan="4" align="center">NO PROJECTS CURRENTLY</td></tr>
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
