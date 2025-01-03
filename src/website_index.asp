<!--- description : this file is basically for the login of the admin to  the adminarea  -->
<!-- ***************************** ##### JAVA script ##### ****************** -->
<%
dim checkCookieemail
dim checkCookiePass
if request.Cookies("user") = "" then
	checkCookieemail = ""
else
	checkCookieemail = "checked"
end if
if request.Cookies("password") = "" then
	checkCookiePass = ""
else
	checkCookiePass = "checked"
end if
%>
<script language ="javascript">
	 //************* 
function checkvalidate()
{
	 if (frmlogin.txtUserName.value =="")
	  {
		 alert('Please Enter the UserName');
		 frmlogin.txtUserName.focus();
		 return false;
    	}    	  
   //************** 
     if (frmlogin.txtPassword.value =="")
      {
        alert('Please Enter the password');
        frmlogin.txtPassword.focus();
        return false;  
        
      }  	
      return true;
      
  }
</script>
 
<!-- &******************************* End of Java scripting *******************-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title>HOME</title>
<LINK 
href="include/style.css" type=text/css rel=stylesheet>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1"></head>
<body bgcolor="#CCCCCC" leftmargin="0" topmargin="0">
<table width="100%" border="0" cellspacing="0" cellpadding="0" height="100%">
  <tr>
    <td valign="top" height="100%">
	  <table width="100%" height="100%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="#FFFFFF">
        <tr bgcolor="#006699"> 
          <td colspan="2"> 
            <p><strong>&nbsp;<font color="#FFFFFF">Company Banner and Information</font></strong></p>
            <p>&nbsp;</p></td>
        </tr>
        <tr> 
          <td width="45%" align="center" valign="middle"> 
            <div align="center">
<table width="80%" border="0" align="center" cellpadding="0" cellspacing="0">
                <tr>
                  <td height="151" bgcolor="#F3F3F3"> 
                    <div align="center">
                      <p>&nbsp;</p>
                      <p>Company Information &amp; Links</p>
                      <p>&nbsp;</p>
                    </div></td>
                </tr>
              </table>
            </div></td>
          <td width="55%" align="center" valign="middle"> <table width="69%" cellspacing="1" cellpadding="1" border="1" bordercolor="#000000">
              <tr>
                <td>
                 <!-- **********************) statring the form from here (****** -->
                  <form name ="frmlogin" method="post" action="loginprocess.asp" onSubmit="return checkvalidate();">
                  <input type="hidden" name="txtpath" value="website_index"> <!-- for login process.asp use that where to post back if password is wrong -->
					<table width="99%" cellspacing="1" cellpadding="1">
					
					     <tr>
							<td colspan="2" align="center"><%=request("message")%> </td><!-- 22 march 2k5  if login faild then message goes here-->
						</tr>	
								    <tr> 
								      <td>&nbsp;</td>
								      <td>&nbsp;</td>
								    </tr>
								    <tr> 
								      <td align="right">UserName&nbsp;&nbsp;&nbsp;</td>
								      <td><input  class="formTextbox"type="text" name="txtUserName" value="<%=request.Cookies("user")%>"></td>
								    </tr>
								    <tr> 
								      <td align="right">Password&nbsp;&nbsp;&nbsp;</td>
								      <td><input  class="formTextbox"  name="txtPassword" type="password" value="<%=request.Cookies("password")%>"></td>
								    </tr>
								    <tr>
								      <td>&nbsp;</td>
								      <td>&nbsp;</td>
				      </tr>
                                <tr>
                                  <td align="right">&nbsp;
                                    <INPUT name=RemEmail type=checkbox value=true  <%=checkCookieemail%>>
                                  </td>
                                  <td>&nbsp;Remember Email</td>
                                </tr>
                                <tr>
                                  <td align="right"><font face="Addled" size="2">
                                    <INPUT name=RemPassword type=checkbox value=true <%=checkCookiepass%>>
                                  </font></td>
                                  <td>&nbsp;Remember Password</td>
                                </tr>
								    <tr> 
								      <td>&nbsp;</td>
								      <td>&nbsp;</td>
								    </tr>
								    <tr> 
								      <td>&nbsp;</td>
								      <td><input type="submit" value="Log In" class="formbutton" name="cmdSubmit">								      </td>
								    </tr>
								    <tr> 
								      <td>&nbsp;</td>
								      <td>&nbsp;</td>
								    </tr>
								    <tr> 
								      <td colspan="2">Forgot Password?<a href="forgetpassword.asp"> Click Here.</a></td>
								    </tr>
								    <tr> 
								      <td>&nbsp;</td>
								      <td>&nbsp;</td>
								    </tr>
								    <tr align="center" bgcolor="#F3F3F3"> 
								      <td colspan="2">Email Links etc.</td>
								    </tr>
						</table>
				  </form>
						<!---*************************** ) end of form (************************ -->
			 </td>
              </tr>
            </table></td>
        </tr>
        <tr> 
          <td>&nbsp;</td>
          <td align="center" valign="middle">&nbsp;</td>
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
