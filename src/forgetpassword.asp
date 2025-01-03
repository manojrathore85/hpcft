<!--  ********************************** Description  :  this the form if user lost his password then we can give him password from herer -->
<!-- #include file="connect.asp" --> <!-- includeing the database file -->
<!-- #include file="generalfunctions.asp" -->
<!-- #include file ="sendmail.asp" -->

<!-- **************************************************** scrript coding ************************-->
 <%
  if request("cmdsubmit")<>"" then
  dim objrs
  dim message ' for showing the message to user 
  set objrs = server.CreateObject("adodb.recordset")
  objrs.open "select Password from " &  varTblNameUserProfile & "  where email='" & request("txtuserid") & "'",con
  if objrs.eof =false then
        FromName = "IMS"
		ReplyTo = request("txtuserid")
		dim Receipent(0)
		Receipent(0) = request("txtuserid")
	    Subject = "Forgot password."
        Body = "Your password for login is :---> " &     objrs("password")
		BCC = ""

		mail FromName,ReplyTo,Receipent,varSiteSpecAddCC,BCC,Subject,Body
		'response.Write(objrs("password"))
		message = "Password sent to your login address."
	else
		message = "You are not the user of this site."
   end if
   objrs.close
   con.close
   set objrs = nothing
   end if
%>

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
                  <td valign="top"> <br/><br/><br/><br/><br/><br/><br/><br/><br/><br/><br/>
				  <table align="center" width="60%" cellpadding="0" cellspacing="0">
   <form name="frm" method="post" action="forgetpassword.asp" -->
   <tr>
        <td class="bluebold" align="center" colspan="2"> <%=message%></td>
   </tr>
   <tr>
		  <td>&nbsp;</td>
		  <td>&nbsp;</td>
		 </tr>
	  <tr>
		 <td class="regularText"> Enter your E- Mail </td>
		 <td ><input type="text" name="txtuserid"  class="formTextbox"> </td>
		 <tr>
		  <td>&nbsp;</td>
		  <td>&nbsp;</td>
		 </tr>
	  <tr>
	   <td colspan="2" align="center"><input type="submit" name="cmdsubmit">
	   <br/><br/><br/><br/><br/><br/><br/><br/><br/><br/><br/><br/><br/><br/><br/><br/><br/><br/>
	   </td>

	  </tr>
  </form>
  <tr><td></td></tr>
 </table>
				  
				  
				  
				  
				  
				  
				  
		   </td>
                </tr>
              </table>
              
            </div></td>
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
