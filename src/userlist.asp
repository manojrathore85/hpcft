<!-- ********************Description : This the file which is basically for showing the all the user list which are avail for us -->
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

if request("order") = "" then
	order = "desc"
else
	order = ""
end if

function get_orderby
	select case request("orderBy")
	case ""
		get_orderby = "email " & request("order")
	case "address"
		get_orderby = "address " & request("order")
	case "name"
		get_orderby = "mname " & request("order")
	case "email"
		get_orderby = "email " & request("order")
	end select
end function
%>


<% 
   dim i ' as incremental variable
   dim sql
   dim objview
   set objview =server.CreateObject("adodb.recordset") ' creating the recordset 
   sql = "select email,concat(firstname,' ',lastname) as mname,address from  " &  varTblNameUserProfile & " order by " & get_orderby
   objview.open sql,con
%>
<script>

function orderby(field)
{
	frm_ulist.orderBy.value=field;
	frm_ulist.submit();	
}
</script>
<form action="viewprojects.asp" method="post" name="frm_ulist">
<input type="hidden" name="mainlink" value="users">
<input type="hidden" name="orderBy">
<input type="hidden" name="order" value="<%=order%>">
</form>
<table width="95%" border="0" cellspacing="0" cellpadding="0">
<tr>
  <td class="SectionHead">Users<br>&nbsp;
  </td>
</tr>
<tr>
  <td>Below is the list of all the users on the site.<br>
	<br>
	----<a class="bluebold" href="viewprojects.asp?mainlink=adduser"> Add User</a></td>
</tr>
</table><br>
<table width="95%" border="0" cellspacing="1" cellpadding="1">
  <tr bgcolor="#FFFFFF" class="head"> 
    <td>&nbsp;<a href="javascript:orderby('email')">Email</a></td>
    <td>&nbsp;<a href="javascript:orderby('name')">Fullname</a></td>
    <td>&nbsp;<a href="javascript:orderby('address')">Address</a></td>
  </tr>
 <% 
    if objview.bof=false and objview.eof =false then ' checking whether the record is exist or not
	   while not objview.eof ' checkout all the user in database
%>	 
  <tr bgcolor="#FFFFFF"> 
    <td>&nbsp;<a href="viewprojects.asp?mainlink=viewuser&id=<%=objview("email")%>"><%=objview("email")%></a></td>
    <!--<td>&nbsp;<%'=objview("FirstName") & " " & objview("lastname")%></td>-->
	<td>&nbsp;<%=objview("mname")%></td>
    <td>&nbsp;<%=objview("Address")%></td>
  </tr>
   <%
	    objview.movenext 
		wend
	else
	%>
	<tr bgcolor="#FFFFFF"> 
    <td colspan="3" align="center">NO USERS CURRENTLY</td></tr>
	<%
    end if		
   %>

</table>							
<%
 ' nullifying all the objects
  objview.close
  set objview= nothing
  con.close
  set con = nothing
 %> 
