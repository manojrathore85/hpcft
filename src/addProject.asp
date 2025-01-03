<!-- #include file="checksession.asp" -->
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

' ********************* CHECKING PERMISSION ENDS ***************************
%>

<%
   dim strsql 
   strsql = "select email,firstname,lastname from " &  varTblNameUserProfile & " order by firstname"
   dim rs
   set rs = server.CreateObject("adodb.recordset")
   rs.open strsql,con
%>
<!-- ********************************* Java Script ************************** -->
<script language="JavaScript" src="include/validate.js"> </script> <!-- includeing the validation file -->

<script language="JavaScript">

	function checkvalidate()
	 {
	  if (frmproject.txtProject.value=="")
	    {
	     alert('Please Enter the Project');
	      frmproject.txtProject.focus();
	        return false;   
	     }
	     //Determine for leader
	     
	     if (frmproject.txtLeadDeveloper.value=="")      
	      {
	         alert('Please Enter the Leader Developer');
	         frmproject.txtLeadDeveloper.focus();
	         return false;  
	      }
	     //*********************** 
	     return true;
	 }
</script>

<form name="frmproject" method="post" action="Addproject_process.asp" onsubmit="return checkvalidate();" enctype="multipart/form-data">
  <table width="90%" border="0" align="center" cellpadding="0" cellspacing="0">
    <tr bgcolor="#F3F3F3"> 
      <td height="43" colspan="2"><div align="center" class="redbold">Add Project 
        </div></td>
    </tr>
    <tr> 
      <td height="35" align="right">* Name:&nbsp;&nbsp;</td>
      <td height="35">&nbsp;&nbsp; <input type="text" name="txtProject"></td>
    </tr>
    <tr> 
      <td width="23%" height="35" align="right">* Lead Developer:&nbsp;&nbsp;</td>
      <td width="77%" height="35">&nbsp;&nbsp;<select class="formTextbox" name="txtLeadDeveloper">
                              <%
									   		 if rs.eof =false and rs.bof =false then
											  while not rs.eof
									   %>
                              <option value="<%=rs("email")%>"><%=rs("firstname") & " " & rs("lastname")%></option>
                              <%
											rs.movenext
									   		wend
											end if
											rs.close
											set rs = nothing
											con.close
											set con = nothing
										%>
                            </select></td>
    </tr>
    <tr> 
      <td height="10" align="right">&nbsp;</td>
      <td height="10">&nbsp;</td>
    </tr>
    <tr> 
      <td width="23%" height="35" align="right" valign="top">Description:&nbsp;&nbsp;</td>
      <td width="77%" height="35">&nbsp;&nbsp;<textarea class="formTextbox" name="txtIssDesc" cols="80" rows="20" wrap="VIRTUAL"></textarea></td>
    </tr>
    <tr> 
      <td height="35" align="right" valign="top">Attach File:&nbsp;&nbsp;</td>
      <td height="35" align="left" valign="top">&nbsp;&nbsp;
        <INPUT class="formTextbox" type=file name=txtfile></td>
    </tr>
    <tr> 
      <td width="23%" height="35" align="right" valign="top">&nbsp;&nbsp;</td>
      <td width="77%" height="35" align="center"><input type="submit" value="Create" name="cmdSubmit" > 
        &nbsp;&nbsp; <input  type="Reset" value="Cancel" name="cmdCancel"></td>
    </tr>
  </table>
</form>
