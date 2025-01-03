<!-- Description :- This file is to Edit user and his password user  **************************************--->
<!-- ************************************************ Date 29 march 2k5 ********************************* -->
<!-- #include file ="checksession.asp" -->
<!-- #include file ="connect.asp" --> <!-- including the connection -->
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
  ' ******************************************** stroring the variable  which comes from the view user like '
    if request("id")<>"" then ' the id comes from the  viewuser  
	     dim objview  'for viewingthe record
		 dim permissionid,read,writee,add,deletee,descriptionn
		  set objview =server.CreateObject("adodb.recordset") ' creating the object
			 objview.open "select * from " &  varTblNamePermissions & " where permissionid='" &  request("id") &"'",con
			   if objview.eof =false and objview.bof =false then 'storing the values in the varible
					permissionid = objview("permissionid")
					read = objview("pread")
					writee = objview("pwrite")
					add = objview("padd")
					deletee = objview("pdelete")
					descriptionn = objview("description")
			   end if
		 objview.close ' nullify the objects
		 set objview = nothing 
		 
	end if
	'******************************** end of view records  ********************
   	
 ' ************************ after the post back of the form we are updateing the records ****************
   if request("cmdSubmit") <>"" then
		dim strsql,message ' string for sql
		if request("chkRead") <> "" then
			strsql = "pread = 'T'"
		else 
			strsql = "pread = 'F'"
		end if
		if request("chkWrite") <> "" then
			strsql = strsql & ",pwrite = 'T'"
		else
			strsql = strsql & ",pwrite = 'F'" 
		end if
		if request("chkAdd") <> "" then
			strsql = strsql & ",padd = 'T'"
		else
			strsql = strsql & ",padd = 'F'" 
		end if
		if request("chkDelete") <> "" then
			strsql = strsql & ",pdelete = 'T'"
		else
			strsql = strsql & ",pdelete = 'F'" 
		end if
		dim temp
		temp ="update " &  varTblNamePermissions & " set " & strsql & ",description = '" & request("txtPermissionDetail") & "' where permissionid = '" & request("txtpermissionid") & "'" 
			con.execute temp
			message = "Permission edited successfully."
			con.close
			set con = nothing 
end if		  
%>
<!-- *****************    Script Java language ****************  --->
<script language ="Javascript">
 //checking the validation
 function checkvalidation()
  {
			 // checking the name 
			if(frmeditpermission.txtPermissionId.value=="")
			  {
				 alert('Please Enter the Permission Name');
				 frmeditpermission.txtPermissionId.focus();  
			     return false;  
			  }
			 // checking one of the permission is clicked or not
			 
			 if(frmeditpermission.chkRead.checked == false && frmeditpermission.chkWrite.checked == false && frmeditpermission.chkAdd.checked == false && frmeditpermission.chkWrite.checked == false)
			 {
			 	alert('Please select atleast one permission for the permission name');
				return false;
			 }
			 
			 // checking permission detail entered or not
		      if(frmeditpermission.txtPermissionDetail.value=="")
			  {
				 alert('Please Enter the Permission Details');
				 frmeditpermission.txtPermissionDetail.focus();  
			     return false;  
			  }
		    return true;

  } //end of function
</script>
<!-- <form name="frmadduser" method="post" action="viewprojects.asp" onsubmit="return checkvalidation();"> --> <!-- form start from here --> 
<form name="frmeditpermission" method="post" action="viewprojects.asp" onsubmit="return checkvalidation();"> <!-- form start from here -->
<input type="hidden" value="editPermission" name="mainlink">
<table width="95%" border="0" cellspacing="0" cellpadding="0" height="100%">
	  
		<tr> 
		  <td valign="top"> <font color="#FFFFFF">-</font><br> 
		  <table width="90%" border="0" align="center" cellpadding="0" cellspacing="0">
		  <% if message <> "" then %>
		  <tr> 
            <td height="35" align="center" colspan="2"><%=message%></td>
          </tr>
		  <% end if%>
          <tr bgcolor="#F3F3F3"> 
            <td height="43" colspan="2"><div align="center" class="redbold">Edit 
                Permission</div></td>
          </tr>
          <tr> 
            <td height="35" align="right">* PermissionId:&nbsp;&nbsp;</td>
            <td height="35">&nbsp;&nbsp; <input type="text" name="txtPermissionId" size="80" class="formTextbox" value="<%=permissionid%>"  readonly=""></td>
          </tr>
          <tr> 
            <td height="35" align="right">* Permissions:&nbsp;&nbsp;</td>
            <td height="35">&nbsp;&nbsp; 
			<%if read = "T" then %>
			<input type="checkbox" name="chkRead" value="r" checked>
			<% else %>
			<input type="checkbox" name="chkRead" value="r">
			<% end if %>
              Read&nbsp;&nbsp; 
			  <%if writee = "T" then %>
			<input type="checkbox" name="chkWrite" value="w" checked>
			<% else %>
			<input type="checkbox" name="chkWrite" value="w">
			<% end if %>
			  Write&nbsp;&nbsp;
			  <%if add = "T" then %>
			<input type="checkbox" name="chkAdd" value="a" checked>
			<% else %>
			<input type="checkbox" name="chkAdd" value="a">
			<% end if %>
              Add&nbsp;&nbsp;
			  <%if deletee = "T" then %>
			<input type="checkbox" name="chkDelete" value="d" checked>
			<% else %>
			<input type="checkbox" name="chkDelete" value="d">
			<% end if %>
              Delete</td>
          </tr>
          <tr> 
            <td height="35" align="right">* Permission Details:&nbsp;&nbsp;</td>
            <td height="35">&nbsp;&nbsp; <input type="text" name="txtPermissionDetail" size="80" class="formTextbox" value="<%=descriptionn%>"></td>
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

</body>
</html>
