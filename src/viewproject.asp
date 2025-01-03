<!-- #include file="checksession.asp" -->
<!-- #include file="connect.asp" --> <!-- includeing the connection file -->

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
   dim objview
   dim filename
   dim str
   set objview =server.CreateObject("adodb.recordset") ' creating the recordset 
   objview.open "select p.*,u.firstname,u.lastname from " &  varTblNameProjects & " p," &  varTblNameUserProfile & " u where p.leaddeveloper = u.email and p.projectid =" & request("id"),con
   str = objview("designdocpath")
   if isnull(objview("designdocpath")) = false then
		 filename = split(str,",")
   end if
%>
<table width="95%" border="0" cellspacing="0" cellpadding="0" height="100%">
  <tr>
    <td valign="center" align="center">
	  <table height="80%" width="80%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="#FFFFFF">
        <tr bgcolor="#F3F3F3"> 
          <td height="38" class="SectionHead" align="center">&nbsp;&nbsp;<%=objview("projectname")%></td>
        </tr>
        <tr align="left" valign="top"> 
          <td><div align="left">&nbsp;&nbsp;<span class="redbold">Description</span><br>&nbsp;&nbsp;<%=objview("description")%><br>
              <br>
              &nbsp;</div>
            &nbsp;&nbsp;<span class="regularTextBold">Lead Developer :-</span>&nbsp;&nbsp;<a href="viewprojects.asp?mainlink=viewuser&id=<%=objview("leaddeveloper")%>"><%=objview("firstname") & " " & objview("lastname")%></a> <br>
			&nbsp;&nbsp;<span class="regularTextBold">Create Date :- </span>&nbsp;&nbsp;<%=DateAdd("h", -5, objview("createDate"))%><%'for showing eastern standard time as it is 5 hrs less then GMT%><br>
			&nbsp;&nbsp;<span class="regularTextBold">Design Docs :- </span>&nbsp;&nbsp;
			<%
				if isArray(filename) then
					if ubound(filename) >= 0 then
						for counter=0 to ubound(filename)
						if counter > 0 then
							response.Write(", ")
						end if
						%>
						<a href="upload/<%=filename(counter)%>" target="_blank"><%=filename(counter)%></a>
						<%
						'response.Write(filename(counter))	
						next
					else
						'response.Write(filename(-1))
					end if
				else
					'response.Write("isArray not working here")
				end if
				'response.Write(objview("designdocpath"))
			%>
			</td>
        </tr>
		<tr><td>&nbsp;</td></tr>
		<tr><td>&nbsp;</td></tr>
        <tr> 
          <td><p align="center"><a href="viewprojects.asp?mainlink=editproject&id=<%=objview("projectId")%>">Edit Project</a> | <a href="deleteproject.asp?id=<%=objview("projectId")%>">Delete 
              Project</a></p></td>
        </tr>
      </table>
   </td>
  </tr>
</table>
<%
objview.close
set objview = nothing

%>
