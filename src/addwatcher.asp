<!-- #include file="checksession.asp " -->
<!-- #include file="Connect.asp" -->

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
if request("cmdSubmit") = "Update" then
	con.execute "update " &  varTblNameProjects & " set projectwatchlist = '" & request("txtwatchers") & "' where projectid = " & request("pid")
	con.close
	set con = nothing
	response.Redirect("viewprojects.asp?mainlink=addwatcher")
end if
dim rs
set rs = server.CreateObject("adodb.recordset")
rs.open "select * from " &  varTblNameProjects & " where projectid=" & request("id"),con

%>
<!-- ********************************* Java Script ************************** -->

<form name="frmproject" method="post" action="addwatcher.asp">
<input type="hidden" name="pid" value="<%=request("id")%>">
  <table width="90%" border="0" align="center" cellpadding="0" cellspacing="0">
    <tr bgcolor="#F3F3F3"> 
      <td height="43" colspan="2"><div align="center" class="redbold">Add Watcher 
          for <%=request("project")%></div></td>
    </tr>
    <tr> 
      <td height="10" align="right">&nbsp;</td>
      <td height="10"><font color="#FF0000">Entries will be comma delimited</font></td>
    </tr>
    <tr> 
      <td width="23%" height="35" align="right" valign="top">Watcher List:<br>
        (e.g:-email@domain.com,<br>
        imsadmin@gtsims.com)&nbsp;&nbsp;</td>
      <td width="77%" height="35">&nbsp;&nbsp; 
        <textarea class="formTextbox" name="txtWatchers" cols="80" rows="20" wrap="VIRTUAL"><%=rs("projectwatchlist")%></textarea></td>
    </tr>
    <tr> 
      <td width="23%" height="35" align="right" valign="top">&nbsp;&nbsp;</td>
      <td width="77%" height="35" align="center"><input type="submit" value="Update" name="cmdSubmit" > 
        &nbsp;&nbsp; <input  type="Reset" value="Cancel" name="cmdCancel"></td>
    </tr>
  </table>
</form>
<%
rs.close
con.close
set rs= nothing
set con = nothing
%>