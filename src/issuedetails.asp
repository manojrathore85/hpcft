<%@Language="VBScript" CodePage="65001" %>
<%
Session.CodePage = 65001
Response.CodePage = 65001
Response.CharSet = "utf-8"
'CR20060902
Session.LCID     = 1033 'en-US
%>
<!-- #include file="generalFunctions.asp" --> <!-- including the functions file -->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd"> 
<html>
<head>
<meta charset="utf-8">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title><%=setTitle%></title>
<LINK 
href="include/style.css" type=text/css rel=stylesheet>

<script language="JavaScript" src="include/encrypt.js" type="text/javascript">
</script>
<script>
var mycomntobj;
			function decodeIssue()
			{
				var msg;
				var varCmntText;
				// *********     when only one comment exists on page  ***********************
				//var varRegexp = new RegExp (/\s/, 'gi') ;

				if(typeof document.all("comnt").length == "undefined")	
				{
					//varCmntText = document.all("comnt").innerText;
					varCmntText = mySingleComment 	//CR20060706	mySinglecomment variable defined at the bottom in javascript
					msg = secureDecrypt(varCmntText,frm.deKey.value);
					if(msg != "")
						document.all("comnt").innerText = msg;
				}
				// *********     end 	*****************************************************

				
				// *********     when multiple comment exists on page  ***********************	
				for(i=0;i<document.all("comnt").length;i++)
				{
					varCmnt = document.all("comnt")[i].innerText;
					//varCmntText = myArrComments[i];	//CR20060706 myArrComments defined at the bottom in js
					msg = secureDecrypt(varCmnt,frm.deKey.value);
					if(msg != "")
						document.all("comnt")[i].innerText = msg;
				}
			}	// *********     end 	*****************************************************
			function Decrypt(obj){
				key = prompt("Please enter the key to Decrypt message");
				encrypted = obj.parentNode.parentNode.nextElementSibling.firstChild.nextSibling.innerText;
				decrypted = secureDecrypt(encrypted,key);
				obj.parentNode.parentNode.nextElementSibling.firstChild.nextSibling.innerText = decrypted;
				mycomntobj = obj;
			}
</script>
</head>
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
                      <tr> 
						<td colspan="2"> <!-- #include file="include/GeneralTopNavBar.asp" --></td>
                      </tr>
                    </table>
                    <font color="#FFFFFF">-</font><br>
					<!-- #include file="insideissuedetails.asp" --> <%'CR20060902%>
						<%
						'server.Execute("insideissuedetails.asp")		'CR20060902
						%>
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
<script>
function PopUp(theURL,winName,features) { //v2.0
  window.open(theURL,winName,features);
}
</script>

</body>
</html>
