<!-- ************************ Description : showing the particular user's all the detail which you want to see *************** date 29 march 2005 Rajat Jaiswal -->
<!-- #include file="checksession.asp" --> <!-- includeing the  connection file -->
<!-- #include file="connect.asp" --> <!-- includeing the  connection file -->

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
  set objview =server.CreateObject("adodb.recordset") ' createing the object of the recordset 
  objview.open "select * from " &  varTblNameUserProfile & " where email='" & request("id") &"'",con
%>

<!--  ***************************** html Coding goes here ***************************** -->
<html>
<head>
<title><%=varSiteSpecTitle%></title>
<LINK href="include/style.css" type=text/css rel=stylesheet>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1"></head>
<script>
function checkDelete()
{
	if (confirm('Are you sure, you want to delete this user.'))
		return true;
	else
		return false;
}
</script>

<body bgcolor="#FFFFFF" leftmargin="0" topmargin="0">
<table width="100%" border="0" cellspacing="0" cellpadding="0" height="100%">
  <tr> 
    <td valign="top"> <font color="#FFFFFF">-</font><br> 
      <table width="90%" border="0" align="center" cellpadding="0" cellspacing="0">
        <tr bgcolor="#F3F3F3"> 
          <td height="43" colspan="2"><div align="left" class="redbold">&nbsp;&nbsp;User 
              :<%=objview("FirstName") & " " & objview("lastName")%></div></td>
        </tr>
        <tr align="left"> 
          <td height="35">&nbsp;&nbsp;<span class="regularTextBold">Address</span></td>
        	
          <td height="35" >:&nbsp;&nbsp;<%=objview("address")%></td>
		</tr>
        <tr align="left"> 
          <td height="35" >&nbsp;&nbsp;<span class="regularTextBold">Email </span></td>
          <td height="35" >:&nbsp;&nbsp;<%=objview("email")%></td>
        </tr>
        <tr>
          <td height="35" >&nbsp;&nbsp;<span class="regularTextBold">Mobile </span></td>
          <td height="35">:&nbsp;&nbsp;<%=objview("mobile")%></td>
        </tr>
        <tr> 
          <td height="35" >&nbsp;&nbsp;<span class="regularTextBold">House Phone 
            </span></td>
          <td height="35">:&nbsp;&nbsp;<%=objview("Hphone")%></td>
        </tr>
        <tr> 
          <td width="22%" height="35" >&nbsp;&nbsp;<span class="regularTextBold">Office 
            Phone </span> </td>
          <td width="78%" height="35">:&nbsp;&nbsp;<%=objview("Ophone")%> </td>
        </tr>
        <tr align="center" valign="middle"> 
          <td height="35" colspan="2"><a href="viewprojects.asp?mainlink=edituser&id=<%=objview("email")%>">Edit 
            Details</a> | <a href="viewprojects.asp?mainlink=userchangepass&id=<%=objview("email")%>">Change Password</a> | <a href="deleteuser.asp?id=<%=objview("email")%>" onClick="return checkDelete();">Delete 
            User</a> &nbsp;&nbsp;</td>
        </tr>
      </table></td>
  </tr>
</table>
</body>
</html>

