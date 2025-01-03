<!-- Description :- This file is to add user --->
<!-- Date 22 march 2k5 -->
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
	dim message
   if request("cmdSubmit") <> "" then 
	   if (request("chkread") <> "" or request("chkwrite") <> "" or request("chkadd") <> "" or request("chkdelete") <> "") then
			dim objAdduser,objcheck ' object for record set
			dim strSql ' string for sql
			dim counter ' for counting
				set objcheck= server.CreateObject("adodb.recordset") ' setting object to check that whether the user already exist or not
				 strsql= "select count(*) from " &  varTblNamePermissions & " where permissionid='" & request("txtpermissionid") & "'"
					  objcheck.Open strsql,con
					  counter =cint(objcheck(0)) 'for storing the result
					  objcheck.Close 
					  set objcheck = nothing
					 if counter= 0 then
						 set  objAdduser =server.CreateObject("adodb.recordset")
						 objAdduser.CursorLocation = 3 'adUseClient
						 objAdduser.Open "select * from " &  varTblNamePermissions & " where 1=2",con, 3,3 ',3  'opening the empty record
							objAdduser.AddNew
							objAdduser("Permissionid") = request("txtPermissionid")
							if request("chkRead") <> "" then
								objAdduser("pread") = "T"
							else
								objAdduser("pread") = "F"
							end if
							if request("chkWrite") <> "" then
								objAdduser("pread") = "T"
								objAdduser("pwrite") = "T"
							else
								objAdduser("pwrite") = "F"
							end if
							if request("chkAdd") <> "" then
								objAdduser("pread") = "T"
								objAdduser("pwrite") = "T"
								objAdduser("padd") = "T"
							else
								objAdduser("padd") = "F"
							end if
							if request("chkDelete") <> "" then
								objAdduser("pdelete") = "T"
							else
								objAdduser("pdelete") = "F"
							end if
							objAdduser("description") = request("txtPermissionDetail")
						objAdduser.update
						message = "Record is saved successfully."
					  else
						   message = "The permissionid already exist please enter some other permissionid." 
					  end if    
				 con.close
				 set con = nothing 
			else
				message = "Provide atleast one permission to add the permission."
			end if		  
		end if		 
    
%>
<!-- *****************    Script Java language ****************  --->
<script language ="Javascript">
 //checking the validation
 function checkvalidation()
  {
	   
			 // checking the name 
			if(frmaddpermission.txtPermissionId.value=="")
			  {
				 alert('Please Enter the Permission Name');
				 frmaddpermission.txtPermissionId.focus();  
			     return false;  
			  }
			 // checking one of the permission is clicked or not
			 
			 if(frmaddpermission.chkRead.checked == false && frmaddpermission.chkWrite.checked == false && frmaddpermission.chkAdd.checked == false && frmaddpermission.chkWrite.checked == false)
			 {
			 	alert('Please select atleast one permission for the permission name');
				return false;
			 }
			 
			 // checking permission detail entered or not
		      if(frmaddpermission.txtPermissionDetail.value=="")
			  {
				 alert('Please Enter the Permission Details');
				 frmaddpermission.txtPermissionDetail.focus();  
			     return false;  
			  }
		    return true;
		 
		  
  } //end of function
</script>
<!-- <form name="frmaddpermission" method="post" action="addpermission.asp" onsubmit="return checkvalidation();">  --><!-- form start from here -->
<form name="frmaddpermission" method="post" action="viewprojects.asp" onsubmit="return checkvalidation();"> <!-- form start from here -->
<input type="hidden" value="addPermission" name="mainlink">
<table width="95%" border="0" cellspacing="0" cellpadding="0" height="100%">
		<tr> 
		  <td valign="top"> <font color="#FFFFFF">-</font><br> 
		  <%
		  if message <> "" then
		  %>
		  <p align="center"><%=message%></p>
		  
        <%
		  end if
		  %>
        <table width="90%" border="0" align="center" cellpadding="0" cellspacing="0">
          <tr bgcolor="#F3F3F3"> 
            <td height="43" colspan="2"><div align="center" class="redbold">Add 
                Permission</div></td>
          </tr>
          <tr> 
            <td height="35" align="right">* PermissionId:&nbsp;&nbsp;</td>
            <td height="35">&nbsp;&nbsp; <input type="text" name="txtPermissionId" size="60" class="formTextbox" value=""></td>
          </tr>
          <tr> 
            <td height="35" align="right">* Permissions:&nbsp;&nbsp;</td>
            <td height="35">&nbsp;&nbsp; <input type="checkbox" name="chkRead" value="r">
              Read&nbsp;&nbsp; <input type="checkbox" name="chkWrite" value="w">
              Write&nbsp;&nbsp; <input type="checkbox" name="chkAdd" value="a">
              Add&nbsp;&nbsp; <input type="checkbox" name="chkDelete" value="d">
              Delete</td>
          </tr>
          <tr> 
            <td height="35" align="right">* Permission Details:&nbsp;&nbsp;</td>
            <td height="35">&nbsp;&nbsp; <input type="text" name="txtPermissionDetail" size="60" class="formTextbox" value=""></td>
          </tr>
          <tr> 
            <td height="35" align="right">&nbsp;</td>
            <td height="35">&nbsp;&nbsp; </td>
          </tr>
          <tr> 
            <td width="26%" height="35" align="right" valign="top">&nbsp;&nbsp;</td>
            <td width="74%" height="35" align="center"><input type="submit" value="Create" name="cmdSubmit"> 
            &nbsp;&nbsp; <input  type="reset" value="Cancel" name="cmdCancel"></td>
          </tr>
        </table></td></tr></table>
 </form>
</body>
</html>
