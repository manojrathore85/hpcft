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
	dim rstemp
	dim sstr
	set rstemp = server.CreateObject("adodb.recordset")
	rstemp.open "select * from " &  varTblNameUsers & " where permissionid ='" & request("id") & "' order by email",con
	if not rstemp.eof then
		rstemp.close
		con.close
		response.Write("<script>alert('The permission you want to delete is given to one/more user-project.\nPlease change the permissions for them to delete this permission');window.history.go(-1);</script>")
		response.End()
	else
		con.execute "delete from " &  varTblNamePermissions & " where permissionid='" & request("id") & "'"
	end if
end if
con.close
set con = nothing
response.Redirect("viewprojects.asp?mainlink=permission")

%>
<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body>

</body>
</html>
