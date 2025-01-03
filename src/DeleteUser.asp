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
	rstemp.open "select * from " &  varTblNameProjects & " where leaddeveloper ='" & request("id") & "'",con
	if not rstemp.eof then
		while not rstemp.eof
			sstr = sstr & rstemp("projectname") & "\n"
			rstemp.movenext
		wend
		rstemp.close
		con.close
		response.Write("<script>alert('The user you want to delete is lead developer in this project/projects \n" & sstr & "Please change the developers of the projects to delete the user.');window.history.go(-1);</script>")
		response.End()
	else
		rstemp.close
		rstemp.open "select * from " &  varTblNameIssues & " where assignto ='" & request("id") & "' order by projectid",con
		if not rstemp.eof then
			sstr = "There are one or more issues assigned to this user.\nChange the user assign to for the isssue/issues to delete the user."
			rstemp.close
			con.close
			response.Write("<script>alert('" & sstr & "');window.history.go(-1);</script>")
			response.End()
		else
			con.begintrans
			con.execute "delete from " &  varTblNameUserProfile & " where email='" & request("id") & "'"
			con.execute "delete from " &  varTblNameUsers & " where email='" & request("id") & "'"
			con.committrans
		end if
	end if
end if
con.close
set con = nothing
response.Redirect("viewprojects.asp?mainlink=users")

%>
