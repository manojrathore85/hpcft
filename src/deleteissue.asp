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
on error resume next
if request("issueid") <> "" then
	if trim(request("prjid")) = "" then
		andclause = ""
	else
		andclause = " and projectid = " & request("prjid")
	end if
	con.begintrans
		con.execute "delete from " &  varTblNameIssues & " where issueid=" & request("issueid") & andclause
		con.execute "insert into " &  varTblNameIssueChangesComments & "(ProjectId,IssueId,Field,ChangesBy,UpdateDate,NewValue,Comments,LastValue)  values(" & session("projectid") & "," & request("issueid") & ",'deleted','" & session("user") & "',now(),'deleted',null,'deleted')"  
	con.committrans
	If Err.Number<>0 then
		response.Write(err.description & "<br>")
		response.End()
	end if

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
