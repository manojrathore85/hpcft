<!-- #include file ="checksession.asp" -->
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
  dim message
   if request("cmdSubmit")<>"" then
    if request("txtnpassword")<>"" then
		 con.execute "update " &  varTblNameUserProfile & " set password='" & request("txtnpassword") &"' where email='" & request("id" ) &"'"
		 message="Your Password is Updated"
		 con.close 
		 set con =nothing
		 end if
  end if
%>
<!--- *************************** script language -********  -->
<script language="Javascript">
function checkvalidation()
{
/* if (frmuserlogin.txtopassword.value="" )
  {
     alert('Please Enter the old  password');
     frmuserlogin.txtopassword.focus();
     return false;  
     
  } */
  //****************
  if (frmuserlogin.txtnpassword.value=="")
   {
     alert('Please Enter the new password');
     frmuserlogin.txtnpassword.focus();
     return false;  
   } 
  //*****************
  if (frmuserlogin.txtnpassword.value!=frmuserlogin.txtrpassword.value)
  {
     alert('Please Enter the same password in Re- Type password field');
     frmuserlogin.txtrpassword.value ="";
     frmuserlogin.txtrpassword.focus();
     return false;      
}
 return true;
}
</script>

<form name ="frmuserlogin" action="viewprojects.asp" method="post" onsubmit="return checkvalidation();">
<input type="hidden" value="userchangepass" name="mainlink">
<input type="hidden" value="<%=request("id")%>" name="id">
  <table width="375" cellspacing="1" cellpadding="1" border="1" bordercolor="#000000" 
					style="WIDTH: 375px; HEIGHT: 266px">
    <tr> 
      <td colspan="2" align="middle" valign="center"><%=message%></td>
    </tr>
    <tr> 
      <td><table width="99%" cellspacing="1" cellpadding="1">
          <tr> 
            <td height="22">&nbsp;</td>
            <td>&nbsp;</td>
          </tr>
          <tr> 
            <td align="right">Email&nbsp;&nbsp;&nbsp;</td>
            <td> :&nbsp;&nbsp; <%=request("id")%></td>
          </tr>
		  <!--
          <tr> 
            <td align="right">Old- Password</td>
            <td> :&nbsp;&nbsp;<input type="password" name="txtopassword" maxlength="16"  class="formTextbox"> </td>
          </tr> -->
          <tr> 
            <td align="right"> New-Password</td>
            <td>:&nbsp;&nbsp;<input type="password" name="txtnpassword" maxlength="16"  class="formTextbox">
            </td>
          </tr>
          <tr> 
            <td align="right"> Re-type Password</td>
            <td>:&nbsp;&nbsp;<input type="password" name="txtrpassword" maxlength="16"  class="formTextbox"></td>
          </tr>
          <tr> 
            <td colspan="2" align="center"><input type="submit" value="Save" class="formbutton" name="cmdSubmit"></td>
          </tr>
          <tr> 
            <td>&nbsp;</td>
            <td>&nbsp;</td>
          </tr>
          <tr align="middle" bgcolor="#f3f3f3"> 
            <td colspan="2">Email Links etc.</td>
          </tr>
        </table></td>
    </tr>
  </table>
</form>
