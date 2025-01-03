<!-- ************************ Description : showing the particular user's all the detail which you want to see *************** date 29 march 2005 Rajat Jaiswal -->
<!-- #include file="checksession.asp" --> <!-- includeing the  connection file -->
<!-- #include file="connect.asp" --> <!-- includeing the  connection file -->
<!--  ***************************   SCRIPTING GOES HERER *************************************** -->
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
<body bgcolor="#FFFFFF" leftmargin="0" topmargin="0">
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
          <td height="35" colspan="2"><a href="editprofile.asp">Edit 
            Details</a> | <a href="changepass.asp">Change Password</a></td>
        </tr>
      </table></td>
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

