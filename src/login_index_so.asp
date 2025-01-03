<%@LANGUAGE="VBSCRIPT"%>
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
<!-- ************************ JAVA SCRIPT PART STARTED FROM HERE  ************ -->
<script language="JavaScript"  src="include/Validate.js"></script> <!-- including the Validate file -->
<script language ="JavaScript" >
function checkvalidation()
{
  var str; // for temp storage
  var flag // for temp return value storage
       
  //******* checking the name 
  
		if (frmuserlogin.txtUserName.value=="")
		 {
		   alert('Please Enter the UserId');
		   frmuserlogin.txtUserName.focus();  
		   return false;
		  }    
       
        if (frmuserlogin.txtUserName.value!="")
         {
            str=frmuserlogin.txtUserName.value;
           
            flag=checkEmail(str); //****** calling from validation.js
				if (flag==false)
				 {
				    alert('Please Enter the UserId in correct mail Format');
				    frmuserlogin.txtUserName.value ="";
				    frmuserlogin.txtUserName.focus();
				    return false;
				 }      
                    
         }   
         //************** checking password **********
         if (frmuserlogin.txtPassword.value =="")
           {
             alert('Please Enter the Password');
             frmuserlogin.txtPassword.focus();  
              return false; 
           }   
               		  
 } // End of function
</script>


<!-- ************************ HTML PART  STARTED FROM HERE ****************** -->

<html>
<head>
<title>Home</title>
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
                  <td valign="top"> <font color="#FFFFFF">-</font><br>
<div align="center"> 
            <Form name ="frmuserlogin" action="loginprocess_so.asp" method="post" onSubmit="return checkvalidation();"> 
			<input type="hidden" name="issueid" value="<%=request("issueid")%>">
			<input type="hidden" name="prjid" value="<%=request("prjid")%>">
					  <table width="375" cellspacing="1" cellpadding="1" border="1" bordercolor="#000000" 
					style="WIDTH: 375px; HEIGHT: 266px">
					<tr>  <td colspan="2" align="middle" valign="center"><%=request("message")%></td></tr>
					    <tr> 
					      <td><table width="99%" cellspacing="1" cellpadding="1">
                                <tr> 
                                  <td>&nbsp;</td>
                                  <td>&nbsp;</td>
                                </tr>
                                <tr> 
                                  <td align="right">Email&nbsp;&nbsp;&nbsp;</td>
                                  <td><input  class="formTextbox" name="txtUserName" readonly="true" value="<%=request.form("user")%>">
								  <input type="hidden" name="user" value="<%=request.form("user")%>"/>
								  </td>
                                </tr>
                                <tr> 
                                  <td align="right">Password&nbsp;&nbsp;&nbsp;</td>
                                  <td><input  class="formTextbox" name="txtPassword" type="password" value="<%=request.Cookies("password")%>"></td>
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
                                  <td><input type="submit" value="Log In" class="formbutton" name="cmdSubmit">                                  </td>
                                </tr>
                                <tr> 
                                  <td>&nbsp;</td>
                                  <td>&nbsp;</td>
                                </tr>
                                <tr> 
                                  <td colspan="2" align="center">Forgot Password?<A href="forgetpassword.asp"> 
                                    Click Here.</a></td>
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
					  
				</form> <!-- end of form -->	  
					  
           </div>                  </td>
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


</body>
</html>
