<%

'if request("issueid") <> "" and session("projectid")="" then  
'	response.Redirect("login_index.asp?issueid=" & request("issueid"))
'end if
%>
<% if session("user")="" then %>
<!--
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
	     <td> Please Re-login <a href="login_index.asp?issueid=<%=request("issueid")%>"> Login Here</a></td>
	</tr>
	
	</table>
</BODY>
</HTML>
-->
 <%
	response.write("<script>window.location= 'login_index.asp?issueid=" & request("issueid") & "&prjid=" & request("prjid") & "';</script>")
	'response.Redirect("login_index.asp?issueid=" & request("issueid") & "&prjid=" & request("prjid"))
 
end if
%>
