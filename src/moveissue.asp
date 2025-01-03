<%@LANGUAGE="VBSCRIPT"%>
<!-- #include file="checksession.asp" -->
<!-- #include file="connect.asp" -->

<%
' ********************* CHECKING PERMISSION ***************************

dim rsperm
set rsperm = server.CreateObject("adodb.recordset")
rsperm.open "select perm.* from " &  varTblNamePermissions & " perm, " &  varTblNameUsers & " u, " &  varTblNameissues & _
 " i where i.projectid = u.projectid and u.permissionid = perm.permissionid and i.issueid = " &  request("issueid") & _
 " and u.email ='" &  session("user") & "'",con
if not rsperm.eof then
	if rsperm("pdelete") = "T" then
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
if request("issueid") <> "" then
	con.begintrans
		con.execute "update " &  varTblNameIssues & " set projectid =" & request("selProject") & " where issueid=" & request("issueid")
		con.execute "update " &  varTblNameIssueChangesComments &  " set projectid =" & request("selProject") & " where issueid=" & request("issueid")
	con.committrans
end if
con.close
set con = nothing
response.Redirect("issuenavigator.asp?search=all")

%>
<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body>

</body>
</html>
