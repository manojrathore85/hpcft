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

     dim i ' as incremental variable
   dim objview
   set objview =server.CreateObject("adodb.recordset") ' creating the recordset 
   objview.open "select * from " &  varTblNameProjects & " order by projectname " & order,con
%>

<table width="95%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td class="SectionHead">Administration<br>
      &nbsp; </td>
  </tr>
  <tr> 
    <td>Below is the list of all the projects with the emails to be reported on.<br>
    </td>
  </tr>
</table>
<BR>
<table width="95%" border="0" cellpadding="0" cellspacing="0" bgcolor="#CCCCCC">
<tr>
  <td><table width="100%" border="0" cellpadding="1" cellspacing="1">
	  <tr bgcolor="#FFFFFF" class=head> 
		<td width="29%" id=head>&nbsp;<a class="bluebold" href="viewprojects.asp?mainlink=addwatcher&order=<%=order%>">Name</a></td>
		  <td width="51%" id=head>&nbsp;Watcher List</td>
		<td width="20%">&nbsp;Operations</td>
	  </tr>
 <% 
    if objview.bof=false and objview.eof =false then ' checking whether the record is exist or not
	   while not objview.eof ' checkout all the user in database
%>	 
	  <tr bgcolor="#FFFFFF"> 
		  <td width="29%">&nbsp;<a href="viewprojects.asp?mainlink=projectwatch&id=<%=objview("projectId")%>&project=<%=objview("projectName")%>"><%=objview("projectName")%></a></td>
		  <td width="51%">&nbsp;<%=objview("projectwatchlist")%></td>
		  <td width="20%">&nbsp;<a href="viewprojects.asp?mainlink=projectwatch&id=<%=objview("projectId")%>&project=<%=objview("projectName")%>">Edit</a> 
          </td>
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
