<!-- #include file="checksession.asp" -->
<!-- #include file="connect.asp" -->
<!-- ********************************** vb script statrted from  here -->
 <%
   dim objproject ' for viewing the project
   set objproject =server.CreateObject("adodb.recordset")
   dim rs ' for viewing the project
   set rs =server.CreateObject("adodb.recordset")
   dim sstr
   rs.open "select yahooid from " &  varTblNameUserProfile & " where email='" & session("user") & "'",con
   if not (rs.eof or rs.bof) then
	   if not isnull(rs("yahooid")) and rs("yahooid") <> "" then
			sstr = "<a href=""yahooautologin.asp"" target=""_blank"">YahooCal Login</a>"
	   end if
   end if
   
   rs.close
   set rs = nothing
  ' objproject.Open "select * from projects order by projectName" ,con
   objproject.Open "select p.projectid,projectname from " &  varTblNameProjects & " p," &  varTblNameUsers & " u, issues i where p.projectid =u.projectid and u.email = '" & session("user") & "' AND i.projectid = p.Projectid GROUP BY i.Projectid ORDER BY MAX(i.updatedate) DESC, p.projectid DESC",con
   
   
    
 %>
<script language ="JavaScript" >
function checkvalidation()
{
       //*** checking the list box
        if (frmuserlogin.lstProjects.value =='none')
         {
            alert('Please Select the project First');
            frmuserlogin.lstProjects.focus(); 
            return false;
          }    
}             
function submitform(for_type)
{
	if(for_type == 'issues')
		 window.location = "issuenavigator.asp?search=alloutstanding";
	else if(for_type == 'allissues')
		 window.location = "issuenavigator.asp?search=all_iss";	     
	else if(for_type == 'diary')
		 window.location = "mydiary.asp";	
	else
		window.location = "viewreminder.asp?search=all";    
}
</script>
<html>
<head>
<title>Select Project</title>
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
          <td colspan="2" align="center"> 
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
                    </table>
                    <font color="#FFFFFF">-</font><br>
					<center>
					<Form name ="frmuserlogin" action="ims_index.asp" method="post" onSubmit="return checkvalidation();">
					<input type="hidden" name="visitlink">
					<!-- <input type="hidden" name="cmdSubmit" value="Next >>"> -->	
                    <table width="375" cellspacing="1" cellpadding="1" border="1" bordercolor="#000000" 
					style="WIDTH: 375px; HEIGHT: 200px">
                      <tr> 
                        <td colspan="2" align="center" valign="middle"><%=request("message")%></td>
                      </tr>
                      <tr> 
                        <td><table width="99%" cellspacing="1" cellpadding="1">
						<%
							if objproject.EOF=true and objproject.BOF =true then ' checking whether the records is exist or not
									session("projectid") = 1
									session("permission") = "all"
									response.write("<br><strong>Looks like there are no issues or Project in IMS</strong><br>")
									response.write("<a href='viewprojects.asp'>Create Project</a><br>")
									response.write("<a href='createissue.asp'>Create issue</a>")
									'response.redirect("viewprojects.asp")
								else
								%>	
                            <tr> 
                              <td>&nbsp;</td>
                              <td>&nbsp;</td>
                            </tr>
                            <tr> 
                              <td align="right">Project&nbsp;&nbsp;&nbsp;</td>
                              <td><select class="formTextbox" name="lstProjects">
                                  <option value="none" selected>Select Project</option>
                                  <% 
									'if objproject.EOF=false and objproject.BOF =false then ' checking whether the records is exist or not
										 while not objproject.EOF 
    							  %>
                                  <option value=<%=objproject("projectId")%>><%=objproject("projectName")%></option>
                                  <%
												     objproject.MoveNext 
												     wend 
												     
											 %>
                                </select>
                              &nbsp;&nbsp;</td>
                            </tr>
                            <tr> 
                              <td>&nbsp;</td>
                              <td>&nbsp;<li><a onClick="return submitform('issues');" href="#">All Outstanding Issues</a></li>&nbsp;<li><a onClick="return submitform('allissues');" href="#">All Issues</a>&nbsp;</li>&nbsp;<li><a onClick="return submitform('reminders');" href="#">All Reminders</a></li>&nbsp;<li><a onClick="return submitform('diary');" href="#">My Diary</a></li>&nbsp;<br/>&nbsp;</td>
                            </tr>
                            <tr> 
                              <td>&nbsp;</td>
                              <td><input type="submit" value="Next >>" class="formbutton" name="cmdSubmit"> 
                              </td>
                            </tr>
							<%
							end if
											     objproject.Close 
											     set objproject = nothing
											     con.close
											     set con = nothing
							%>
                           <tr> 
                              <td colspan="2" align="center" class="bluebold"></td>
                             </tr>
							 
                          </table></td>
                      </tr>
                    </table></form> </center></td>
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
<%
                                  if session("referrer") <> "" then
								 ' response.Write(varSiteSpecURL & Mid( request.ServerVariables("PATH_INFO"), 1, InStrRev(request.ServerVariables("PATH_INFO"), "/")) & "getfile.asp?" & session("referrer") )
                                           response.write("<script>window.open('" & varSiteSpecURL & Mid(request.ServerVariables("PATH_INFO"), 1, InStrRev(request.ServerVariables("PATH_INFO"), "/")) & "getfile.asp?" & session("referrer") & "')</script>")

				  end if
%>
