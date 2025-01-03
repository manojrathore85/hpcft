<!-- Description :- This file is to Edit user and his password user  **************************************--->
<!-- ************************************************ Date 29 march 2k5 ********************************* -->
<!-- #include file ="checksession.asp" -->
<!-- #include file ="connect.asp" --> <!-- including the connection -->

<%
  ' ******************************************** storing the variable  which comes from the view user like '
	     dim message
		 dim objview  'for viewingthe record
		 dim email,password,address,hphone,ophone,mobile,fname
		  set objview =server.CreateObject("adodb.recordset") ' creating the object
			 objview.open "select * from " &  varTblNameUserProfile & " where email='" &  session("user") &"'",con
			   if objview.eof =false and objview.bof =false then 'storing the values in the variable
					address=objview("address")
					hphone =objview("Hphone")
					ophone=objview("Ophone")
					firstname= objview("firstname")
					lastname= objview("lastname")
					mobile=objview("mobile")
			   end if
		 objview.close ' nullify the objects
		 set objview = nothing 
		 
	'******************************** end of view records  ********************
   	
 ' ************************ after the post back of the form we are updateing the records ****************
   if request("cmdSubmit") <>"" then
		dim strsql ' string for sql
		 strsql= "update  " &  varTblNameUserProfile & " set address='" & trim(request("txtaddress")) & "',mobile='"& trim(request("txtmobile")) & "',Hphone='" & trim(request("txtphoneH")) & "',Ophone='" & trim(request("txtphoneO")) & "' where email='" & session("user") &"'"
		' response.Write(strsql)
		' response.End()
			con.execute strsql
			con.close
			set con = nothing 
			message = "Profile edited successfully."
end if		  
		 
    
%>
<!-- *****************    Script Java language ****************  --->
<html>
<head>
<title><%=varSiteSpecTitle%></title>
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
                        <td width="66%">&nbsp;</td>
                        <td width="34%" align="left"><!-- #include file="include/TopRightNavBar.asp" --></td>
                      </tr>
                      <tr>
                        <td>&nbsp;</td>
                        <td>&nbsp;</td>
                      </tr>
                      <tr><td colspan="2"> <!-- #include file="include/GeneralTopNavBar.asp" --></td>
                      </tr>
                    </table>
                    <font color="#FFFFFF">-</font><br>
                    <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
                      <tr> 
                        <td height="43" colspan="2">
						
<form name="frmadduser" method="post" action="editprofile.asp" onsubmit="return checkvalidation();"> <!-- form start from here -->
<table width="95%" border="0" cellspacing="0" cellpadding="0" height="100%">
	  
		<tr> 
		  <td valign="top"><center><span class="message"><%=message%></span></center> <font color="#FFFFFF">-</font><br> 
		  <table width="90%" border="0" align="center" cellpadding="0" cellspacing="0">
		      <tr bgcolor="#F3F3F3"> 
		        <td height="43" colspan="2"><div align="center" class="redbold">Edit 
                User</div></td>
		      </tr>
		      <tr> 
		        <td height="35" align="right">First Name:&nbsp;&nbsp;</td>
		        <td height="35">&nbsp;&nbsp; <input type="text" name="txtFirstName" size="80" class="formTextbox" value="<%=firstname%>" readonly="true"></td>
		      </tr>
			  <tr> 
		        <td height="35" align="right">Last Name:&nbsp;&nbsp;</td>
		        <td height="35">&nbsp;&nbsp; <input type="text" name="txtLastName" size="80" class="formTextbox" value="<%=lastname%>"  readonly="true"></td>
		      </tr>
		      <tr> 
		        <td height="35" align="right">Contact Address:&nbsp;&nbsp;</td>
		        <td height="35">&nbsp;&nbsp; <input type="text" name="txtAddress" size="80" class="formTextbox" value="<%=address%>"></td>
		      </tr>
		      <tr> 
		        <td width="21%" height="35" align="right">Mobile:&nbsp;&nbsp;</td>
		        <td width="79%" height="35">&nbsp;&nbsp; <input type="text" name="txtMobile" size="80" class="formTextbox" value="<%=mobile%>"></td>
		      </tr>
		      <tr> 
		        <td width="21%" height="35" align="right"> Phone(H):&nbsp;&nbsp;</td>
		        <td width="79%" height="35">&nbsp;&nbsp; <input type="text" name="txtPhoneH" size="80" class="formTextbox" value="<%=hphone%>"></td>
		      </tr>
		      <tr> 
		        <td width="21%" height="35" align="right">Phone(O):&nbsp;&nbsp;</td>
		        <td width="79%" height="35">&nbsp;&nbsp; <input type="text" name="txtPhoneO" size="80" class="formTextbox" value="<%=ophone%>"></td>
		      </tr>
		      <tr> 
		        <td height="35" align="right">&nbsp;</td>
		        <td height="35">&nbsp;&nbsp; </td>
		      </tr>
		      <tr> 
		        <td width="21%" height="35" align="right" valign="top">&nbsp;&nbsp;</td>
		        <td width="79%" height="35" align="center"><input type="submit" value="Update" name="cmdSubmit" > 
		          &nbsp;&nbsp; <input  type="reset" value="Cancel" name="cmdCancel"></td>
		      </tr>
 </table>
 </td>
  </tr></table></form>

						
						</td>
                      </tr>
                    </table>
                    <div align="center"><br>
                    </div></td>
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
