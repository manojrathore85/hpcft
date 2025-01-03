<%
' Program to edit the issue
'Atul	20060307	Added encryption program,used hidden textarea 'txtcomments' to pass the comments
' CR20060826		GTSATUL
%>
<!-- #include file ="CheckSession.asp" --> 
<!-- #include file ="Connect.asp" -->
<!-- #include file ="generalFunctions.asp" -->

<%

' ********************* CHECKING PERMISSION ***************************
Session.CodePage  = 65001
dim rsperm
set rsperm = server.CreateObject("adodb.recordset")
rsperm.open "select perm.* from " &  varTblNamePermissions & " perm, " &  varTblNameUsers & " u, " &  varTblNameissues & _
 " i where i.projectid = u.projectid and u.permissionid = perm.permissionid and i.issueid = " &  request("issueid") & _
 " and u.email ='" &  session("user") & "'",con
if not rsperm.eof then
	if rsperm("pwrite") = "T" then
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


 <!-- checkinh yhr drddion -->
<%
 '****************   for user  assign to  ****************
    dim objuser  
   set objuser= server.CreateObject("adodb.recordset")   
   
'   objuser.Open "select * from  userprofile order by firstname",con
   
  ' objuser.Open "select up.* from " &  varTblNameUserProfile & " up, " &  varTblNameUsers & " u where up.email = u.email and u.projectid =" & session("projectid") ,con
     objuser.Open "select up.* from " &  varTblNameUserProfile & " up, " &  varTblNameUsers & " u, " &  varTblNameissues & _
				  " i where i.projectid = u.projectid and up.email = u.email and i.issueid =" & request("issueid") ,con

'**************************** determining the issueid details  *******************
'if request("issueId")<>"" then
    dim issuetype,severity,summary,assignto,description,projectId,filename 'varible use for storing the initial values 
    dim issueId,counter 
	dim str
    dim objview ' for record set
	issueid = request("issueid")
	'if request("issueid") <> "" then
    '   issueid=cdbl(request("issueId"))
	'else
    '   issueid= cdbl(request("txtissueId"))
	'end if
	
	
		set objview = server.CreateObject("adodb.recordset") ' createing the object
		dim strsql 
		if request("prjid") <> "" then
			strsql = "select * from " &  varTblNameIssues & " where issueId="& issueid & " and projectid=" & request("prjid")
		else
			strsql = "select * from " &  varTblNameIssues & " where issueId="& issueid
		end if
'   			objview.Open "select * from issues where issueId="& request("txtissueId")
			 objview.Open strsql,con
					 if objview.EOF  =false and objview.BOF  =false then
							 projectid=objview("projectId")
							 issuetype=objview("issuetype")
							 severity =objview("severity")
							 summary =objview("summary")
							 assignto= objview("Assignto")
							 description =objview("description")
							 str = objview("attachedfilepath")
							 if not isnull(objview("attachedfilepath")) then
								 filename = split(str,",")
							 end if
					  end if 
			  objview.Close 
	      set objview = nothing
'end if


'function check_file_ondisk(filename)
'	dim fso
'	set fso = server.CreateObject("scripting.filesystemobject")
'	If fso.FileExists(server.MapPath(".") & "\" & "upload" & "\" & filename) Then
'		check_file_ondisk =  filename 
'	else
'		check_file_ondisk = "<strike>" & filename & "</strike>"
'	end if
'	set fso = nothing
'end function  

%>

<!-- *********************** Java SCript ******************************* -->
<script language="Javascript">
self.moveTo(0,0);

 function checkvalidate()
  {
   
    if (frmissue.lstIssTypes.value =='none')
     {
       alert('Please Select the issue type');
       frmissue.lstIssTypes.focus();  
       return false;
      }
      //************** checking summary *************
      if (frmissue.txtIssSummary.value=="")
        {
          alert('Please Enter the summary of issue');
          frmissue.txtIssSummary.focus();  
          return false;   
        }
        //******************** Assign to 
        if (frmissue.lstUsers.value =='none')
         {
           alert('Please Select ther person');
           frmissue.lstUsers.focus();
           return false;  
                   
         }
		 if(frmissue.encodedmessage.value != "")
		 	frmissue.txtcomments.value = frmissue.encodedmessage.value;
		else
			frmissue.txtcomments.value = frmissue.comments.value;
		
        
  }

function checkDelete()
{
	if (confirm('Do u really want to delete this file'))
		return true;
	else
		return false;
}
var flageditdesc = false;
function confirmeditdesc()
{

if(flageditdesc == false)
	if (confirm('Do u really want to edit the description')){
		frmissue.txtIssDesc.removeAttribute('readonly','readonly');
		flageditdesc = true;
	}
	else{
		
	}
}

</script>
<!-- ******************************* html part from herer ********************* -->

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title><%=varSiteSpecTitle%></title>
<LINK href="include/style.css" type=text/css rel=stylesheet>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<script language="JavaScript" src="include/encrypt.js" type="text/javascript">
</script>
<script>
function encrypt()
{
if(frmissue.chkEncrypt.checked == true)
	if(frmissue.enKey.value == "" || frmissue.comments.value == "")
	{
		alert('Please enter encryption key and comments to encrypt'); 
		frmissue.chkEncrypt.checked = false;
	}
	else {
		//varMsg = secureEncrypt(frmissue.comments.value,frmissue.enKey.value);
		//alert(varMsg);
		//varMsgCopy="";
		//if(varMsg.length >= 100) {
		//	for(i=0;i<=varMsg.length;i+=100) {
		//		varMsgCopy += varMsg.substr(i+1,100) + "-";
				//alert(varMsgCopy);
		//	}
		//}
		frmissue.encodedmessage.value = secureEncrypt(frmissue.comments.value,frmissue.enKey.value);
		//frmissue.encodedmessage.value = varMsgCopy;
	}
else
{
	frmissue.enKey.value = "";
	frmissue.encodedmessage.value = "";
}

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
                  <td valign="top"> <table width="98%" border="0" align="center" cellpadding="0" cellspacing="0">
                      <tr> 
                        <td width="67%">&nbsp;</td>
                        <td width="33%" align="left"><!-- #include file="include/TopRightNavBar.asp" --></td>
                      </tr>
                      <tr>
                        <td>&nbsp;</td>
                        <td>&nbsp;</td>
                      </tr>
                      <tr> 
                        <td colspan="2"><!-- #include file="include/GeneralTopNavBar.asp"--></td>
                      </tr>
                    </table>
                    <font color="#FFFFFF">-</font><br>
            <!-- **************************************** form starting from here 22/march/2k5 ****************************** -->
                <form name="frmissue" method="post" action="editissue_process.asp" enctype="multipart/form-data" onSubmit="return checkvalidate();" >    
                     <input  type="hidden" name="issueId" value="<%=issueId%>">
					 <input type="hidden" name="txtprojectId" value="<%=projectId%>">
								
                      <table width="98%" border="0" align="center" cellpadding="0" cellspacing="0">
                        <tr bgcolor="#F3F3F3"> 
                          <td height="43" colspan="2"><div align="center" class="redbold">Edit 
                              Issue</div></td>
                        </tr>
                        <tr> 
                          <td height="35" align="right">* Issue Type:&nbsp;&nbsp;</td>
                          <td height="35">&nbsp;&nbsp; <select class="formTextbox" name="lstIssTypes">
                              <option value="none" selected>Select Issue Type</option>
                              <option value="BUG">Bug</option>
                              <option value="NEW">New Feature</option>
                              <option value="TASK">Task</option>
                              <option value="IMPROVEMENT">Improvement</option>
                            </select></td>
                        </tr>
                        <tr> 
                          <td height="35" align="right">Severity:&nbsp;&nbsp;</td>
                          <td height="35">&nbsp;&nbsp; <select class="formTextbox" name="lstSeverity">
                              <option selected value="MAJOR">Major</option>
                              <option value="BLOCKER">Blocker</option>
                              <option value="MINOR">Minor</option>
                              <option value="TRIVAL">Trivial</option>
                              <option value="CRITICAL">Critical</option>
                            </select></td>
                        </tr>
                        <tr> 
                          <td width="15%" height="35" align="right">* Summary:&nbsp;&nbsp;</td>
                          <td width="85%" height="35">&nbsp;&nbsp; <input  class="formTextbox"type="text" name="txtIssSummary" size="100" value="<%=summary%>"></td>
                        </tr>
                        <tr> 
                          <td width="15%" height="35" align="right">* Assign To:&nbsp;&nbsp;</td>
                          <td width="85%" height="35">&nbsp;&nbsp; <select class="formTextbox" name="lstUsers">
                              <option value="none" selected>Select User</option>
                              <%
									   		 if objuser.eof =false and objuser.bof =false then
											  while not objuser.eof
											   %>
                              <option value="<%=objuser("email")%>"><%=objuser("firstname") & " " & objuser("lastname") %></option>
                              <%
														objuser.movenext
														wend
														end if
														objuser.close
														set objuser = nothing
													%>
                            </select> </td>
                        </tr>
                        <tr> 
                          <td width="15%" height="35" align="right">Reporter:&nbsp;&nbsp;</td>
                          <td width="85%" height="35">&nbsp;&nbsp; <%=session("user")%></td>
                        </tr>
                        <tr> 
                          <td width="15%" height="35" align="right" valign="top">&nbsp;&nbsp;</td>
                          <td width="85%" height="35" align="left">&nbsp;&nbsp;&nbsp;&nbsp;<input type="submit" value="Update" name="cmdSubmit"> 
                            &nbsp;&nbsp; <input  type="button" value="Cancel" name="cmdCancel"></td>
                        </tr>
                        <tr> 
                          <td height="35" align="right">One Time Email:&nbsp;&nbsp;<br> 
                            <font color="red">Comma Delimited</font></td>
                          <td height="35"><table width="100%" border="0">
  <tr>
    <td width="37%">&nbsp;&nbsp;
      <textarea class="formTextbox" name="txtMailsTo" cols="50" rows="1" wrap="VIRTUAL"></textarea></td>
    <td width="21%" align="right">Send Full History:&nbsp;&nbsp;</td>
    <td width="42%">&nbsp;&nbsp; <input type="checkbox" name="chkSendHistory" value="1"></td>
  </tr>
</table>
</td>
                        </tr>
                        <tr> 
                          <td width="15%" height="35" align="right" valign="top">Description:&nbsp;&nbsp;</td>
                          <td width="85%" height="35">&nbsp;&nbsp; <textarea class="formTextbox" name="txtIssDesc" cols="100" rows="10" wrap="VIRTUAL" readonly="readonly" onClick="javascript:confirmeditdesc();"><%=description%></textarea></td>
                        </tr>
                        <!-- adding the commentss -->
                        <tr> 
                          <td height="35" align="right" valign="top">&nbsp;</td>
                          <td height="35"> Encrypt: 
                            <input type="checkbox" name="chkEncrypt" onClick="encrypt();" value="ON">
                            &nbsp;&nbsp;&nbsp;&nbsp;Encryption&nbsp;Key: 
                            <input type="password" name="enKey" size="50" ></td>
                        </tr>
                        <tr> 
                          <td width="15%" height="35" align="right" valign="top">Comments:&nbsp;&nbsp;</td>
                          <td width="85%" height="35">&nbsp;&nbsp; <textarea class="formTextbox" name="comments" cols="100" rows="10" wrap="VIRTUAL"></textarea></td>
                        </tr>
                        
                        <tr> 
                          <td height="35" align="right" valign="middle">Attach 
                            File:&nbsp;&nbsp;</td>
                          <td height="35" valign="top" align="left"> <table width="100%" height="100%">
                              <tr> 
                                <td width="43%" align="left" valign="middle">&nbsp;&nbsp;<INPUT class="formTextbox" type=file name=txtfile size="20"><br>
								  &nbsp;&nbsp;<INPUT class="formTextbox" type=file name=txtfile1 size="20"><br>
								  &nbsp;&nbsp;<INPUT class="formTextbox" type=file name=txtfile2 size="20"><br>
								  &nbsp;&nbsp;<INPUT class="formTextbox" type=file name=txtfile3 size="20"><br>
								  &nbsp;&nbsp;<INPUT class="formTextbox" type=file name=txtfile4 size="20">
								  </td>
                                <td width="57%"> 
                                  <%
							if isArray(filename) then
							if ubound(filename) >= 0 then
								response.Write("<strong>Files already attached with the issue.</strong><br>")
								for counter=0 to ubound(filename)
									if check_file_ondisk(filename(counter)) =  filename(counter) then ' CR20060826
										response.Write(">> ")	
									%>
									  <a href="upload/<%=filename(counter)%>" target="_blank"><%=filename(counter)%></a>&nbsp;&nbsp;&nbsp;<a href="removefile.asp?issueid=<%=issueid%>&filename=<%=filename(counter)%>" onClick="return checkDelete();">Delete 
									  this file</a><br> 
									  <%
									  end if	
								next
							end if
							end if
							%>                                </td>
                              </tr>
                            </table></td>
                        </tr>
                        <tr> 
                          <td width="15%" height="35" align="right" valign="top">&nbsp;&nbsp;</td>
                          <td width="85%" height="35" align="center"><input type="submit" value="Update" name="cmdSubmit"> 
                            &nbsp;&nbsp; <input  type="button" value="Cancel" name="cmdCancel"></td>
                        </tr>
						<tr>
                          <td height="35" align="right" valign="middle">Encoded 
                            Text:&nbsp;&nbsp;</td>
                          <td height="35" valign="top" align="left">&nbsp;&nbsp; 
                            <textarea class="formTextbox" name="encodedmessage" cols="100" rows="2" wrap="VIRTUAL" readonly></textarea>
							<input type="hidden" name="txtcomments">							</td>
                        </tr>
                      </table>
								
           <!-- ********************* end of form            ********** -->
           </form>         
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
<!-- ************************ script language Java script for highlighting assian to & issue type ans severity option ************** -->
<script language="JavaScript">
//****************************** for highlighting the issue type of the issue that which issue is selected 
 <% if issuetype <>""  then %>
   for (i=0;i<=frmissue.lstIssTypes.options.length-1;i++)
     if (frmissue.lstIssTypes.options[i].value=='<%= issuetype%>')
	    frmissue.lstIssTypes.selectedIndex=i;
		
 <% end if %>
 //********************************* for highlighting the severity that which severity is selected *********

  <% if severity <>""  then %>
   for (i=0;i<=frmissue.lstSeverity.options.length-1;i++)
     if (frmissue.lstSeverity.options[i].value=='<%=severity%>')
	    frmissue.lstSeverity.selectedIndex=i;
		
 <% end if %>
 //*****************  for highlighting the user  which is assign for the task  **************************
  <% if assignTo <>""  then %>
   for (i=0;i<=frmissue.lstUsers.options.length-1;i++)
     if (frmissue.lstUsers.options[i].value=='<%=assignto%>')
	    frmissue.lstUsers.selectedIndex=i;
		
 <% end if %>
 

</script>
<!-- ******************************************************* closing the connection object  *****************-->
<%
  if con.state then
  con.close
   set con = nothing
   end if

%>