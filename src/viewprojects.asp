<!-- #include file="checksession.asp" -->
<!-- #include file="generalFunctions.asp" --> <!-- including the functions file -->
<!-- #include file="connect.asp" --> <!-- including the connection file -->
<%
if session("user") <> "sanjayz@gmail.com" then
	call check_permission
else
	session("projectid") = 1
	session("permission") = "all"
end if
' ********************* CHECKING PERMISSION ***************************
sub check_permission
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
end sub
' ********************* CHECKING PERMISSION ENDS***************************
%>
<html>
<head>
<title><%=setTitle%></title>
<LINK 
href="include/style.css" type=text/css rel=stylesheet>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1"></head>
<body bgcolor="#CCCCCC" leftmargin="0" topmargin="0">
<table width="100%" border="0" cellspacing="0" cellpadding="0" height="100%">
  <tr>
    <td valign="top" height="100%">
	  <table width="100%" height="100%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="#FFFFFF">
        <tr bgcolor="#006699"> 
          <td colspan="2"> <p><strong>&nbsp;<font color="#FFFFFF">Company Banner 
              and Information</font></strong></p>
            <p>&nbsp;</p></td>
        </tr>
        <tr align="left" valign="top"> 
          <td colspan="2"> 
            <div align="center">
<table width="100%" border="0" cellspacing="0" cellpadding="0" height="100%">
                <tr> 
                  <td valign="top"> <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
                      <tr> 
                        <td width="65%">&nbsp;</td>
                        <td width="35%" align="left"><!-- #include file="include/TopRightNavBar.asp" --></td>
                      </tr>
                      <tr>
                        <td>&nbsp;</td>
                        <td>&nbsp;</td>
                      </tr>
                      <tr> 
						<td colspan="2"> <!-- #include file="include/GeneralTopNavBar.asp" --></td>
                      </tr>
                    </table>
                    <font color="#FFFFFF">-</font><br>
                    <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
                      <tr> 
                        <td width="18%" valign="top">
<table width="90%" border="0" align="right" cellpadding="0" cellspacing="0">
                            <tr> 
                              <td>&nbsp;</td>
                            </tr>
                            <tr> 
                              <td><a  class="redbold" href="viewprojects.asp">Projects</a></td>
                            </tr>
                            <tr> 
                              <td><a class="redbold" href="viewprojects.asp?mainlink=users">Users</a></td>
                            </tr>
							<!--
                            <tr> 
                              <td><a class="redbold"href="#">Back-Up Database</a></td>
                            </tr>
                            <tr> 
                              <td><a class="redbold"href="#">Restore Database</a></td>
                            </tr>
							-->
                            <tr>
                              <td><a class="redbold"href="viewprojects.asp?mainlink=permission">Permissions</a></td>
                            </tr>
							<tr>
                              <td><a class="redbold"href="viewprojects.asp?mainlink=setpermissionlist">Set Permissions</a></td>
                            </tr>
							<tr>
                              <td><a class="redbold"href="viewprojects.asp?mainlink=addwatcher">Add Watchers</a></td>
                            </tr>
							
							<tr>
                              <td><a class="redbold"href="viewprojects.asp?mainlink=seeusage">See Usage</a></td>
                            </tr>
							<tr>
                              <td><a class="redbold"href="viewprojects.asp?mainlink=cleanusage">Clean Old Usage</a></td>
                            </tr>
							<tr>
                              <td><a class="redbold"href="viewprojects.asp?mainlink=imsStats">Ims Statistics</a></td>
                            </tr>

                          </table>
						</td>
						<td width="82%" valign="top" align="center"><br>
						<%
						If request("mainlink") = "" then
							server.Execute("projectlist.asp")
						elseif request("mainlink") = "users" then
							server.Execute("userlist.asp")
						elseif request("mainlink") = "addProject" then
							server.Execute("addProject.asp")
						elseif request("mainlink") = "viewproject" then
							server.Execute("viewproject.asp")
						elseif request("mainlink") = "editproject" then
							server.Execute("editproject.asp")
						elseif request("mainlink") = "adduser" then
							server.Execute("adduser.asp")
						elseif request("mainlink") = "viewuser" then
							server.Execute("viewuseradmin.asp")	
						elseif request("mainlink") = "edituser" then
							server.Execute("edituser.asp")
						elseif request("mainlink") = "userchangepass" then
							server.Execute("changepassall.asp")		
						elseif request("mainlink") = "permission" then
							server.Execute("permissionlist.asp")		
						elseif request("mainlink") = "addPermission" then
							server.Execute("addPermission.asp")		
						elseif request("mainlink") = "editPermission" then
							server.Execute("editpermission.asp")	
						elseif request("mainlink") = "setpermissionlist" then
							server.Execute("setPermissionlist.asp")
						elseif request("mainlink") = "addsetpermission" then
							server.Execute("addsetpermission.asp")
						elseif request("mainlink") = "editsetpermission" then
							server.Execute("editsetpermission.asp")
						elseif request("mainlink") = "yahoouser" and request("id") <> "" then
							server.Execute("selectyahooid.asp")
						elseif request("mainlink") = "yahoouser" then
							server.Execute("selectuserforyahoo.asp")
						elseif request("mainlink") = "addwatcher" then
							server.Execute("projectwatchlist.asp")
						elseif request("mainlink") = "projectwatch" then
							server.Execute("addwatcher.asp")
						elseif request("mainlink") = "seeusage" then
							server.Execute("seeusage.asp")
						elseif request("mainlink") = "cleanusage" then
							server.Execute("seeusage.asp")
						elseif request("mainlink") = "addClient" then
							server.Execute("addClient.asp")
						elseif request("mainlink") = "updateClient" then
							server.Execute("updateclient.asp")
						elseif request("mainlink") = "restoreClient" then
							server.Execute("restoreclient.asp")
						elseif request("mainlink") = "imsStats" then
							server.Execute("ims_stats.asp")
						end if
						%>
                        </td>
                      </tr>
					  </table>
                  </td>
                </tr>
              </table>
              
            </div></td>
        </tr>
        <tr> 
          <td width="45%">&nbsp;</td>
          <td width="55%" align="center" valign="middle">&nbsp;</td>
        </tr>
        <tr> 
          <td>&nbsp;</td>
          <td align="center" valign="middle">&nbsp;</td>
        </tr>
        <tr bgcolor="#003366"> 
          <td colspan="2" align="center"><font color="#FFFFFF">Email Links etc 
            </font></td>
        </tr>
      </table>
   </td>
  </tr>
</table>
</body>
</html>
