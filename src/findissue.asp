<!--  *********************** description : the file is basically for fiding the particular issue 13 march 2k5 Rajat --->
<%'CR20060902	Option buttons for project selection added %>
<!-- #include file="checksession.asp" -->
<!-- #include file="connect.asp" --> <!-- include connection file -->
<!-- #include file="generalFunctions.asp" --> <!-- including the functions file -->
<%

' ********************* CHECKING PERMISSION ***************************

'dim rsperm
'set rsperm = server.CreateObject("adodb.recordset")
'rsperm.open "select * from " &  varTblNamePermissions & " where permissionid = '" & session("permission") & "'",con
'if not rsperm.eof then
'	if rsperm("pread") = "T" then
'		rsperm.close
'		set rsperm = nothing
'	else
'		server.Execute("permissiondenied.asp")
'		rsperm.close
'		set rsperm = nothing
'		response.End()
'	end if
'else
'		server.Execute("permissiondenied.asp")
'		rsperm.close
'		set rsperm = nothing
'		response.End()
'end if

' ********************* CHECKING PERMISSION ENDS***************************
%>

 <%

   dim objuser 
   set objuser= server.CreateObject("adodb.recordset")   
   objuser.Open "select * from  " &  varTblNameUserProfile & " order by firstname",con
   
 %>

<!--       *********** Java Script ***************** 
<script language="javascript">
/*
        function checkvalidate()
         {
           if (frmfind.lstIssTypes.value=="none")
            {
              frmfind.lstIssTypes.focus();
              alert('Please Select issue Type');
              return false;  
            }
           //***************** for
           if (frmfind.lstAssignedTo.value =='none')
            {
              frmfind.lstAssignedTo.focus();
              alert('Please Select the Assign To');
              return false;  
            }  
          //***************** for user
           if (frmfind.lstReportedBy.value =='none')
            {
               frmfind.lstReportedBy.focus();
               alert('Please Select the Reported By');
               return false;  
            }   
            
            return true;  
         }
 */ 
</script>-->
<!-- **************************** HTML CODING **************************** -->


<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title><%= setTitle %></title>
<LINK href="include/style.css" type=text/css rel=stylesheet>
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
                      <tr><td colspan="2"> <!-- #include file="include/GeneralTopNavBar.asp" --></td>
                        
                      </tr>
                    </table>
                   <font color="#FFFFFF">-</font><br>

                    <table width="90%" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#F3F3F3">
                      <!--	<form name ="frmfind" action="issuenavigator.asp" method="post" onsubmit="return checkvalidate();" > -->
                      <form name ="frmfind" action="issuenavigator.asp" method="post">
                        <tr> 
                          <td height="43" colspan="2"><div align="center" class="redbold">Find 
                              Issue</div></td>
                        </tr>
                        <tr>
                          <td height="35" align="right">Project:&nbsp;&nbsp;</td>
                          <td height="35">&nbsp;&nbsp; 
                            <input type="radio" name="optProject" value="CURRENT_PROJECT" <%if session("projectid") = "" then response.Write("disabled='disabled'")%>> 
                          Current Project
                            <input type="radio" name="optProject" value="ALL_PROJECT"> 
                            All Project</td>
                        </tr>
                        <tr> 
                          <td height="35" align="right">Issue Type:&nbsp;&nbsp;</td>
                          <td height="35">&nbsp;&nbsp; <select class="formTextbox" name="lstIssTypes">
                              <option value="all" selected>All</option>
                              <option value="BUG">Bug</option>
                              <option value="NEW">New Feature</option>
                              <option value="TASK">Task</option>
                              <option value="IMPROVEMENT">Improvement</option>
                            </select></td>
                        </tr>
                        <tr> 
                          <td height="35" align="right">Severity:&nbsp;&nbsp;</td>
                          <td height="35">&nbsp;&nbsp; <select class="formTextbox" name="lstSeverity">
                              <option selected value="all">All</option>
                              <option value="MAJOR">Major</option>
                              <option value="BLOCKER">Blocker</option>
                              <option value="MINOR">Minor</option>
                              <option value="TRIVAL">Trivial</option>
                              <option value="CRITICAL">Critical</option>
                            </select></td>
                        </tr>
                        <tr> 
                          <td width="18%" height="35" align="right"> Keywords:&nbsp;&nbsp;</td>
                          <td width="82%" height="35">&nbsp;&nbsp; <input  class="formTextbox"type="text" name="txtIssSummary" size="100"></td>
                        </tr>
                        <tr>
                          <td height="35" align="right">Full Text Search:&nbsp;&nbsp;</td>
                          <td height="35">&nbsp;
                            <input type="checkbox" name="chkFullText" value="True"></td>
                        </tr>
                        <tr> 
                          <td width="18%" height="35" align="right"> Assign To:&nbsp;&nbsp;</td>
                          <td width="82%" height="35">&nbsp;&nbsp; <select class="formTextbox" name="lstAssignedTo">
                              <option value="all" selected>All</option>
                              <%
								  if objuser.eof=false and objuser.bof=false then 
							  	 while not objuser.eof
						    %>
                              <option value="<%=objuser("EMAIL")%>"><%=objuser("firstname") & " " & objuser("lastname")%></option>
                              <%
									  objuser.movenext
									  wend
									   objuser.movefirst
									  end if
								%>
                            </select></td>
                        </tr>
                        <tr> 
                          <td width="18%" height="35" align="right">Reported By:&nbsp;&nbsp;</td>
                          <td width="82%" height="35">&nbsp;&nbsp; <select class="formTextbox" name="lstReportedBy">
                              <option value="all" selected>All</option>
                              <%
								  if objuser.eof=false and objuser.bof=false then 
								 
							  	 while not objuser.eof
						    %>
                              <option value="<%=objuser("EMAIL")%>"><%=objuser("firstname") & " " & objuser("lastname") %></option>
                              <%
									  objuser.movenext
									  wend
									  end if
									  objuser.close
									  set objuser = nothing
								%>
                            </select></td>
                        </tr>
                        <tr> 
                          <td height="10" align="right">&nbsp;</td>
                          <td height="10">&nbsp;</td>
                        </tr>
                        <tr> 
                          <td width="18%" height="35" align="right" valign="top">&nbsp;&nbsp;</td>
                          <td width="82%" height="35" align="left"> <input type="submit" value="Find" name="cmdFind" > 
                            &nbsp;&nbsp; <input  type="reset" value="Cancel" name="cmdCancel"></td>
                        </tr>
                      </form >
                      <!--***************************  End of the form ******* -->
                    </table>
                    
                  </td>
                </tr>
			
              </table>
			  
		</div>           
            
			
			 
			
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
