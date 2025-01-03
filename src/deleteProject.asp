<%@LANGUAGE="VBSCRIPT"%>
<!-- #include file="checksession.asp" -->
<!-- #include file="connect.asp" -->

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
if request("id") <> "" then
	con.begintrans
		con.execute "delete from " &  varTblNameProjects & " where projectid=" & request("id")
		con.execute "delete from " &  varTblNameIssues & " where projectid=" & request("id")
		'COULD NOT BE EXECUTED BECAUSE COMMENTS OF ANY ISSUE UNDER REQUESTED PROJECTID(FOR DELETE) MAY BE THERE FOR 
		'THE SAME ISSUE IN ANOTHER PROJECT 20070203
		'con.execute "delete from " &  varTblNameIssueChangesComments & " where projectid=" & request("id")
		con.execute "delete from " &  varTblNameUsers & " where projectid=" & request("id")
	con.committrans
end if
con.close
set con = nothing
response.Redirect("viewprojects.asp")

%>
<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body>

</body>
</html>
