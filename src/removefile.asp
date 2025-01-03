<%@LANGUAGE="VBSCRIPT"%>
<%
' Modified Date		Modified By			Comments
' 20060826			Atul Agrawal		CR20060826 
%>
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
	if rsperm("padd") = "T" then
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
dim rs
dim str
dim filenames
dim counter
dim pos
dim newfilenames
if request("issueid") <> "" then
	'set rs = server.CreateObject("adodb.recordset")
	'rs.open "select attachedfilepath from " &  varTblNameIssues & " where issueid="& request("issueid"),con
	'if not rs.eof then
	'	str = rs("attachedfilepath")
	'	if not isnull(rs("attachedfilepath")) then
	'		 filenames = split(str,",")
	'	 end if
	'	 rs.close
	'	 if ubound(filenames) >= 0 then
	'		for counter = 0 to ubound(filenames)
	'			if filenames(counter) = request("filename") then
	'				' do nothing
	'			else
	'				if counter = ubound(filenames) then
	'					newfilenames = newfilenames & filenames(counter)
	'				else
	'					newfilenames = newfilenames & filenames(counter) & ","
	'				end if
	'			end if
	'		next
	'	 end if	
	'	 if newfilenames <> "" then
	'		 If Mid(newfilenames, Len(newfilenames), Len(newfilenames)) = "," Then
	'			newfilenames = Mid(newfilenames, 1, Len(newfilenames) - 1)
	'		 End If
	'	 end if
	'	con.execute "update " &  varTblNameIssues & " set attachedfilepath = '" & newfilenames & "' where issueid =" & request("issueid")
	'	con.close
	'end if 
	'if rs.state = 1 then rs.close
	'set rs = nothing
	'set con = nothing
	call deletefile
	response.Redirect("editissue.asp?issueid=" & request("issueid"))
end if
str = ""
if request("projectid") <> "" then
	'set rs = server.CreateObject("adodb.recordset")
	'rs.open "select designdocpath from " &  varTblNameProjects & " where projectid="& request("projectid"),con
	'str = rs("designdocpath")
	'if not rs.eof then
	'	if not isnull(rs("designdocpath")) then
	'		 filenames = split(str,",")
	'	 end if
	'	 rs.close
	'	 if ubound(filenames) >= 0 then
	'		for counter = 0 to ubound(filenames)
	'			if filenames(counter) = request("filename") then
	'				' do nothing
	'			else
	'				if counter = ubound(filenames) then
	'					newfilenames = newfilenames & filenames(counter)
	'				else
	'					newfilenames = newfilenames & filenames(counter) & ","
	'				end if
	'			end if
	'		next
	'	 end if	
	'	 if newfilenames <> "" then
	'		 If Mid(newfilenames, Len(newfilenames), Len(newfilenames)) = "," Then
	'			newfilenames = Mid(newfilenames, 1, Len(newfilenames) - 1)
	'		 End If
	'	 end if
	'	con.execute "update " &  varTblNameProjects & " set designdocpath = '" & newfilenames & "' where projectid =" & request("projectid")
	'	con.close
	'end if 
	'if rs.state = 1 then rs.close
	'set rs = nothing
	'set con = nothing
	call deletefile
	response.Redirect("viewprojects.asp?mainlink=editproject&id=" & request("projectid"))
end if

sub deletefile
	dim fso
	set fso = server.CreateObject("scripting.filesystemobject")
	If fso.FileExists(server.MapPath(".") & "\" & "upload" & "\" & request("filename")) Then
		fso.deletefile(server.MapPath(".") & "\" & "upload" & "\" & request("filename"))
	end if
	set fso = nothing
end sub
%>