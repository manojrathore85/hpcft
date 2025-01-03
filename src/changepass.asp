<!-- #include file ="checksession.asp" -->
<!-- #include file="connect.asp" -->

<%
  dim message
   if request("cmdSubmit")<>"" then
    if request("txtnpassword")<>"" then
		 con.execute "update " &  varTblNameUserProfile & " set password='" & request("txtnpassword") &"' where email='" & session("user") &"'"
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
<html>
<title>Change Password</title>
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
                  <td valign="top"> <table width="98%" border="0" align="center" cellpadding="0" cellspacing="0">
                      <tr> 
                        <td width="66%">&nbsp;</td>
                        <td width="34%" align="left"><!-- #include file="include/TopRightNavBar.asp" --></td>
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
<form name ="frmuserlogin" action="changepass.asp" method="post" onSubmit="return checkvalidation();">
<!-- <input type="hidden" value="userchangepass" name="mainlink">
<input type="hidden" value="<%'=request("id")%>" name="id"> -->
<center>
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
            <td> :&nbsp;&nbsp; <%=session("user")%></td>
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
</td>
                </tr>
              </table>
 </center>
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
</body></html>