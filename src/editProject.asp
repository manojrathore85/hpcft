<!-- Description : This file is basically for editing any selected project      date 29 march 2005 Rajat Jaiswal-->
<!-- #include file ="checksession.asp" -->
<!-- #include file ="connect.asp" -->

<%

' ********************* CHECKING PERMISSION ***************************

dim rsperm
set rsperm = server.CreateObject("adodb.recordset")
rsperm.open "select * from " &  varTblNamePermissions & " where permissionid = '" & session("permission") & "'",con
if not rsperm.eof then
	if rsperm("permissionid") = "all" and rsperm("pread") = "T" and rsperm("pwrite") = "T" and rsperm("padd") = "T" and rsperm("pdelete") = "T" then
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


<%
'******************************************************  viewing the project ********************************************************
  dim project,developer,description,filename
  if request("id")<>"" then 'checking the project id
   dim objview 
   dim rs
   dim str
   set rs =server.CreateObject("adodb.recordset")
   rs.open "select email,firstname,lastname from " &  varTblNameUserProfile & "",con
    set  objview =server.CreateObject("adodb.recordset")
	 objview.open "select * from  " &  varTblNameProjects & " where projectId=" & request("id"),con
	    if objview.eof =false or objview.bof =false then
	   		project =objview("projectName")
	   		developer=objview("leadDeveloper")
			description=objview("description")
			projectId=request("id")
			str = objview("designdocpath")
			if not isnull(objview("designdocpath")) then
				filename = split(str,",")
			end if
		end if
		
		objview.close  ' nullifying the object
		set objview = nothing	
	   	 
  end if 
  
%>

<!-- *********************************** Starting JAVA SCRIPT ************************* -->

<script language ="Javascript">
 function checkvalidate()
  {
    if (frmeditproject.txtProject.value =="")
      {
        alert('Please Enter the project first');
        frmeditproject.txtProject.focus();  
        return false;
      } 
     //***********
    if (frmeditproject.txtLeadDeveloper.value=="")
     {
       alert('Please Enter the lead Developer Name');
       frmeditproject.txtLeadDeveloper.focus();
       return false;
      }
  }
function checkDelete()
{
	if (confirm('Do u really want to delete this file'))
		return true;
	else
		return false;
}

</script>

<!-- ********************************** HTML CODING ******************* -->
	<table width="95%" border="0" cellspacing="0" cellpadding="0" height="100%">
	  <tr>
	    <td align="center" valign="center"> 
	    
	    <!-- *************** starting form ***************************** 22 march 2k5 ******************************* ---->
		<Form name ="frmeditproject" method="post" action="editproject_process.asp" onsubmit="return checkvalidate();" enctype="multipart/form-data">
            <input type="hidden" name="txtprojectId" value=<%=ProjectId%>> <!-- showing the project Id -->
		    
				
        <table width="90%" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
          <tr bgcolor="#F3F3F3"> 
            <td height="43" colspan="2"><div align="center" class="redbold">Edit 
                Project :&nbsp;<%=project%></div></td>
          </tr>
          <tr> 
            <td height="35" align="right">* Name:&nbsp;&nbsp;</td>
            <td height="35">&nbsp;&nbsp; <input type="text" name="txtProject" value="<%=project%>"></td>
          </tr>
          <tr> 
            <td width="20%" height="35" align="right">* Lead Developer:&nbsp;&nbsp;</td>
            <td width="80%" height="35">&nbsp;&nbsp; <select class="formTextbox" name="txtLeadDeveloper">
                              <%
									   		 if rs.eof =false and rs.bof =false then
											  while not rs.eof
									   %>
                              <option value="<%=rs("email")%>"><%=rs("firstname") & " " & rs("lastname")%></option>
                              <%
											rs.movenext
									   		wend
											end if
											rs.close
											set rs = nothing
											con.close
											set con = nothing
										%>
                            </select></td>
          </tr>
          <tr> 
            <td height="10" align="right">&nbsp;</td>
            <td height="10">&nbsp;</td>
          </tr>
          <tr> 
            <td width="20%" height="35" align="right" valign="top">Description:&nbsp;&nbsp;</td>
            <td width="80%" height="35">&nbsp;&nbsp; <textarea class="formTextbox" name="txtIssDesc" cols="80" rows="20" wrap="VIRTUAL"><%=description%> </textarea></td>
          </tr>
		  <tr>
		  <td height="35" align="right" valign="middle">Attach File:&nbsp;&nbsp;</td>
		  <td height="35" valign="top" align="left">
		  <table width="100%" height="100%">
		  <tr>
			<td width="43%" align="left" valign="middle">&nbsp;&nbsp;<INPUT class="formTextbox" type=file name=txtfile></td>
			<td width="57%">
			<%
			if isArray(filename) then
			if ubound(filename) >= 0 then
				response.Write("<strong>Files already attached with the project.</strong><br>")
				for counter=0 to ubound(filename)
				response.Write(counter + 1 & ". ")
				%>
				<a href="upload/<%=filename(counter)%>" target="_blank"><%=filename(counter)%></a>&nbsp;&nbsp;&nbsp;<a href="removefile.asp?projectid=<%=projectid%>&filename=<%=filename(counter)%>" onClick="return checkDelete();">Delete this file</a><br>
				<%	
				next
			end if
			end if
			%>
			</td>
		  </tr>
		  </table>
		  
		  </td>
		</tr>
          <tr> 
            <td width="20%" height="35" align="right" valign="top">&nbsp;&nbsp;</td>
            <td width="80%" height="35" align="center"><input type="submit" value="Update" name="cmdUpdate" > 
              &nbsp;&nbsp; <input  type="reset" value="Cancel" name="cmdCancel"></td>
          </tr>
        </table>
		  </form>
		  <!-- *********************************** ending the form ******************* -->
	  </td>
  </tr>
</table>
<script>
<% if developer <>""  then %>
   for (i=0;i<=frmeditproject.txtLeadDeveloper.options.length-1;i++)
     if (frmeditproject.txtLeadDeveloper.options[i].value=='<%=developer%>')
	    frmeditproject.txtLeadDeveloper.selectedIndex=i;
		
 <% end if %>
</script>