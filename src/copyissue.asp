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
function check_it
	check_it = false
	dim rs
	set rs = server.CreateObject("adodb.recordset")
	rs.open "select 1 from " &  varTblNameIssues & " where projectid = " & request("selProject_to_copy") & " and issueid = " & request("issueid"),con
	if rs.eof then
		check_it = true
	end if
	rs.close
	set rs = nothing
end function
%>


<%
if request("issueid") <> "" then
	if check_it = true then
		con.execute "insert into " &  varTblNameIssues & " select " & request("selProject_to_copy") & ",issueid,createdate,updatedate,issuetype,severity,summary,assignto, reporter,description,status,attachedfilepath,reportUsers from " &  varTblNameIssues & " where issueid = " & request("issueid")
		response.Write("<script>alert('Issue copied successfully.');window.location = 'issuenavigator.asp?search=all';</script>")
	else
		response.Write("<script>alert('Issue cannot be copied in the same project.');history.go(-1);</script>")
	end if
end if
con.close
set con = nothing

'response.Redirect("issuenavigator.asp?search=all")
%>
<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body>

</body>
</html>
