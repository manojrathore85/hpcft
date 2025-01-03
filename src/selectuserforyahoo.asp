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
		con.close
		set con = nothing
		response.End()
	end if
else
		server.Execute("permissiondenied.asp")
		rsperm.close
		set rsperm = nothing
		con.close
		set con = nothing
		response.End()
end if

' ********************* CHECKING PERMISSION ENDS***************************
%>


<% 
   dim i ' as incremental variable
   dim objview
   set objview =server.CreateObject("adodb.recordset") ' creating the recordset 
   objview.open "select * from  " &  varTblNameUserProfile & " order by Firstname",con
%>
<table width="95%" border="0" cellspacing="0" cellpadding="0">
<tr>
  <td class="SectionHead">Create Yahoo User-ID<br>&nbsp;
  </td>
</tr>
<tr>
    <td>Below is the list of all the users on the site.<br>
      Select the user for whom you want to create Yahoo-id.<br>
	<br>
    </td>
</tr>
</table><br>
<table width="95%" border="0" cellspacing="1" cellpadding="1">
  <tr bgcolor="#FFFFFF" class="head"> 
    <td>&nbsp;Email</td>
    <td>&nbsp;Fullname</td>
    <td>&nbsp;Address</td>
	<td>&nbsp;Yahoo-Id</td>
  </tr>
 <% 
    if objview.bof=false and objview.eof =false then ' checking whether the record is exist or not
	   while not objview.eof ' checkout all the user in database
%>	 
  <tr bgcolor="#FFFFFF"> 
    <td>&nbsp;<a href="createyahooid.asp?id=<%=objview("email")%>" target="_blank"><%=objview("email")%></a></td>
    <td>&nbsp;<%=objview("FirstName") & " " & objview("lastname")%></td>
    <td>&nbsp;<%=objview("Address")%></td>
	<td>&nbsp;<%=objview("yahooid")%></td>
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
