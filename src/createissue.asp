<!DOCTYPE HTML>
<%@Language="VBScript" CodePage="65001" %>
<%
Session.CodePage = 65001
Response.CodePage = 65001
Response.CharSet = "utf-8"
'CR20060902
Session.LCID     = 1033 'en-US
%>
<!-- #include file ="checkSession.asp" --> <!-- checking the session -->
<%
' we want that only alloutstanding issues link can only work without selecting a project  
if session("projectid") = "" then
	response.Redirect("selectproject.asp")
end if 
%>
<!-- #include file ="Connect.asp" --> <!-- including the database file -->
<!-- #include file="generalFunctions.asp" --> <!-- including the functions file -->
<%

' ********************* CHECKING PERMISSION ***************************

dim rsperm
set rsperm = server.CreateObject("adodb.recordset")
rsperm.open "select * from " &  varTblNamePermissions & " where permissionid = '" & session("permission") & "'",con
if not rsperm.eof then
	if rsperm("padd") = "T" then
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



<!-- ********************************* SCRIPTING AREA FOLLOWS HERE ************************* -->
<%
dim objuser 'for  assigner
dim projectId,projectName ' to hold the project Id
' if request("lstprojects")<>"" then
	 projectId =session("projectid")
	 projectname=proname(projectId ) 'determining the project name for showing using proname function
 'end if
 '****************   for user  assign to  ****************
  
   set objuser= server.CreateObject("adodb.recordset")   
   
'   objuser.Open "select * from  userprofile order by firstname",con
   objuser.Open "select up.* from " &  varTblNameUserProfile & " up, " &  varTblNameUsers & " u where up.email = u.email and u.projectid =" & projectid ,con
   
  '************************ form post ************************************
  
 '*********************************************** function for finding the name of the project using the project id
  function proname (id)
  dim objrs ,temp
  set objrs =server.CreateObject("adodb.recordset")
  objrs.open "select ProjectName from " &  varTblNameProjects & " where projectID="&session("projectid"),con
   if objrs.eof =false then
   temp=objrs("projectName")
   end if
   objrs.close
   set objrs = nothing
   proname=temp
  
  end function
  

%>

<!-- *********************** Java SCript ******************************* -->
<script language="Javascript">
 function checkvalidate()
  {
  
	frmissue.cmdSubmit.disabled= true;
   
    if (frmissue.lstIssTypes.value =='none')
     {
       alert('Please Select the issue type');
       frmissue.lstIssTypes.focus();  
	   frmissue.cmdSubmit.disabled= false;
       return false;
      }
      //************** checking summary *************
      if (frmissue.txtIssSummary.value=="")
        {
          alert('Please Enter the summary of issue');
          frmissue.txtIssSummary.focus();  
		  frmissue.cmdSubmit.disabled= false;
          return false;   
        }
        //******************** Assign to 
        if (frmissue.lstUsers.value =='none')
         {
           alert('Please Select ther person');
           frmissue.lstUsers.focus();
		   frmissue.cmdSubmit.disabled= false;
           return false;  
                   
         }   
        return true;
  }
</script>
<!-- ******************************* html part from herer ********************* -->


<html>
<head>
<title><%=setTitle%></title>
<script src="https://cdn.tiny.cloud/1/jfhvmzju5ge1ljiv30evfvsh676206icpk2z7i430ny5ouoo/tinymce/6/tinymce.min.js" referrerpolicy="origin"></script>
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
            <!-- **************************************** form starting from here 22/march/2k5 ****************************** -->
                <form name="frmissue" method="post" action="createissue_process.asp" onSubmit="return checkvalidate();" enctype="multipart/form-data">    
								
                      <table width="90%" border="0" align="center" cellpadding="0" cellspacing="0">
                        <tr bgcolor="#F3F3F3"> 
                          <td height="43" colspan="2"><div align="center" class="redbold">Create 
                              Issue</div></td>
                        </tr>
                        <tr> 
                          <td  height="35" align="right"> Project:&nbsp;&nbsp;</td>
                          <td height="35" class="SectionHead">&nbsp;&nbsp;<%=projectname%> 
                            <input type="hidden" name="txtproject" value="<%=projectId%>"></td>
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
                          <td width="12%" height="35" align="right">* Summary:&nbsp;&nbsp;</td>
                          <td width="88%" height="35">&nbsp;&nbsp; <input  class="formTextbox"type="text" name="txtIssSummary" size="100"></td>
                        </tr>
                        <tr> 
                          <td width="12%" height="35" align="right">* Assign To:&nbsp;&nbsp;</td>
                          <td width="88%" height="35">&nbsp;&nbsp; <select class="formTextbox" name="lstUsers">
                              <option value="none" selected>Select User</option>
                              <%
									   		 if objuser.eof =false and objuser.bof =false then
											  while not objuser.eof
									   %>
                              <option value="<%=objuser("email")%>"><%=objuser("firstname") & " " & objuser("lastname")%></option>
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
                          <td width="12%" height="35" align="right">Reporter:&nbsp;&nbsp;</td>
                          <td width="88%" height="35">&nbsp;&nbsp; Current user 
                            <input type="text" class="formTextbox" border="0" style="border:1px white;"  name="txtreporter" value="<%=session("user")%>" readonly> 
                          </td>
                        </tr>
                        <tr> 
                          <td height="10" align="right">&nbsp;</td>
                          <td height="10">&nbsp;</td>
                        </tr>
                        <tr> 
                          <td width="12%" height="35" align="right" valign="top">Description:&nbsp;&nbsp;</td>
                          <td width="88%" height="35">&nbsp;&nbsp; <textarea class="formTextbox" id="txtIssDesc"  name="txtIssDesc" cols="100" rows="20" wrap="VIRTUAL"></textarea></td>
                        </tr>
                        <tr>
                          <td height="35" align="right" valign="top">Attach File:&nbsp;&nbsp;</td>
                          <td height="35" align="left" valign="top">&nbsp;&nbsp;
                            <INPUT class="formTextbox" type=file name=txtfile></td>
                        </tr>
                        <tr> 
                          <td width="12%" height="35" align="right" valign="top">&nbsp;&nbsp;</td>
                          <td width="88%" height="35" align="center"><input type="submit" value="Create" name="cmdSubmit"> 
                            &nbsp;&nbsp; <input  type="button" value="Cancel" name="cmdCancel"></td>
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
<script>


tinymce.init({
  selector: '#txtIssDesc',  // change this value according to your HTML
  plugins: 'a_tinymce_plugin autolink image lists table',
  images_upload_url: 'tinymceeditorupload.asp',
  a_plugin_option: true,
  toolbar: 'undo redo | styles | bold italic | alignleft aligncenter alignright alignjustify | ' +
      'bullist numlist outdent indent | link image | table | print preview media fullscreen | fontsize | ' +
      'forecolor backcolor emoticons | help',
  a_configuration_option: 400,
  menubar: '',
});

</script>
<!-- SCript  language java script for highlighting the selected issue type in lstissuetype *************** 24 march 2k5 Rajat Jaiswal -->

<script language ="javascript">
 <%' if ttype<>"" then %>
  //for(i=1;i<=frmissue.lstIssTypes.options.length-1;i++)
   // if(frmissue.lstIssTypes.options[i].value=='<%'=ttype%>')
    //   frmissue.lstIssTypes.selectedIndex=i;
   <%' end if %>         
</script>