<%@LANGUAGE="VBSCRIPT"%>
<!--  ************ ********** Description : this file is basically for giving the result of the find issue *******   -->
<!-- #include file="checksession.asp" -->
<!-- #include file="connect.asp" -->
<!-- #include file="generalFunctions.asp" --> <!-- including the functions file -->

<%
'******* procedures and function start here  **************************
'******* function to check the permission ************************
function checkPermission
	dim rsperm
	set rsperm = server.CreateObject("adodb.recordset")
	rsperm.open "select * from " &  varTblNamePermissions & " where permissionid = '" & session("permission") & "'",con
	if not rsperm.eof then
		if rsperm("pread") = "T" then
			checkPermission = true
		else
			checkPermission = false
		end if
	else
		checkPermission = false
	end if
	rsperm.close
	set rsperm = nothing
end function
'******* procedures and function end here  **************************

'****** checking permission ******************************************
if checkPermission = false then
	server.Execute("permissiondenied.asp")
	response.End()
end if

function print_rem_desc()
	dim rs
	set rs = server.CreateObject("adodb.recordset")
	rs.open "select * from " & varTblNameReminder_table & " where rem_idCol=" & request("id"),con
	if not rs.eof then
	%>
        <tr> 
		  <td>Type</td>
          <td>&nbsp;<%=rs("TypeCol")%></td>
        </tr>
		<tr>
		  <td>Due Date Time </td>
		  <td>&nbsp;<%=rs("Due_DateTimeCol")%> ( <%=rs("TimeZoneCol")%> )</td>
		</tr>
		<tr>
		  <td>Date Entered</td>
			<td>&nbsp;<%=rs("Date_enteredCol")%> ( <%=rs("TimeZoneCol")%> )</td>
		</tr>
		<tr>
		  <td>Person Responsible </td>
			<td>&nbsp;<%=rs("Person_ResponsibleCol")%> </td>
		</tr>
		<tr>
		  <td>Alert Freq </td>
		<td>&nbsp;<%=rs("AlertFrequencyCol")%> </td>
		</tr>
		<tr>
		  <td>State</td>
			<td>&nbsp;<%=rs("StateCol")%> </td>
		</tr>
		<tr>
		  <td>Notes</td>
			<td>&nbsp;<%=rs("NotesCol")%> </td>
		</tr>
		<tr>
		  <td>Description</td>
		  <td>&nbsp;<%=rs("DescCol")%> </td>
	    </tr>
	<%
		'print_rem_desc = rs("DescCol")
	end if
	rs.close
	con.close
end function 

%>
<html>
<head>
<title>Home</title>
<LINK 
href="include/style.css" type=text/css rel=stylesheet>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1"></head>
<body bgcolor="#CCCCCC" leftmargin="0" topmargin="0">
<table width="100%" border="0" cellspacing="0" cellpadding="0" height="100%">
	<tr><td>&nbsp;</td></tr>
  <tr>
    <td valign="top" height="100%">
	  <table width="100%" height="65%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="#FFFFFF">
        <tr align="left"> 
          <td width="21%" height="35">Project&nbsp;</td>
		  <td width="79%">&nbsp;<%=getProjectName%></td>
        </tr>
        <tr> 
          <td height="38">Issue&nbsp;</td>
          <td>&nbsp;</td>
        </tr>
		<% print_rem_desc %>
      </table>
   </td>
  </tr>
</table>
</body>
</html>


