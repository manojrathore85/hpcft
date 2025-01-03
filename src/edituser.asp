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
		 dim email,address,hphone,ophone,mobile,fname,yahooid
		  set objview =server.CreateObject("adodb.recordset") ' creating the object
			 objview.open "select * from " &  varTblNameUserProfile & " where email='" &  request("id") &"'",con
			   if objview.eof =false and objview.bof =false then 'storing the values in the varible
					email =objview("email")
					address=objview("address")
					hphone =objview("Hphone")
					ophone=objview("Ophone")
					firstname= objview("firstname")
					lastname= objview("lastname")
					mobile=objview("mobile")
					password=objview("password")
					yahooid = objview("yahooid")
			   end if
		 objview.close ' nullify the objects
		 set objview = nothing 
		 
	end if
	'******************************** end of view records  ********************
   	
 ' ************************ after the post back of the form we are updateing the records ****************
   if request("cmdSubmit") <>"" then
		dim strsql,message ' string for sql
		 strsql= "update  " &  varTblNameUserProfile & " set FirstName='" & trim(request("txtfirstname")) & "',lastname='" & trim(request("txtlastname"))  & "',address='" & trim(request("txtaddress")) & "',mobile='"& trim(request("txtmobile")) & "',Hphone='" & trim(request("txtphoneH")) & "',Ophone='" & trim(request("txtphoneO")) & "',yahooid='" & trim(request("txtYahooid")) & "' where email='" & request("txtemail") &"'"
		' response.Write(strsql)
		' response.End()
			con.execute strsql
			con.close
			set con = nothing 
			response.Redirect("viewprojects.asp?mainlink=viewuser&id=" & request("txtemail"))
			 
end if		  
		 
    
%>
<!-- *****************    Script Java language ****************  --->
<script language="JavaScript"  src="include/Validate.js"></script> <!-- including the Validate file -->
<script language ="Javascript">
 //checking the validation
 function checkvalidation()
  {
		var str,flag; // temp variable
    
			 // checking the name 
			if(frmadduser.txtFirstName.value=="")
			  {
				 alert('Please Enter the First name');
				 frmadduser.txtFirstName.focus();  
			     return false;  
			  }       
			  //
			   if (frmadduser.txtFirstName.value!="")
			     {
			       str=frmadduser.txtFirstName.value ; 
			         flag= isNumber(str);
			         if (flag==true)
			         {
			           alert('First name does not accept numbers');
			           frmadduser.txtFirstName.focus();
			           frmadduser.txtFirstName.value="";  
			           return false;
			         }    
			      }
			if(frmadduser.txtLastName.value=="")
			  {
				 alert('Please Enter the Last name');
				 frmadduser.txtLastName.focus();  
			     return false;  
			  }       
			  //
			   if (frmadduser.txtLastName.value!="")
			     {
			       str=frmadduser.txtLastName.value ; 
			         flag= isNumber(str);
			         if (flag==true)
			         {
			           alert('Last name does not accept numbers');
			           frmadduser.txtLastName.focus();
			           frmadduser.txtLastName.value="";  
			           return false;
			         }    
			      }   
  } //end of function
</script>
<form name="frmadduser" method="post" action="viewprojects.asp" onsubmit="return checkvalidation();"> <!-- form start from here -->
<input type="hidden" value="edituser" name="mainlink">
<input type="hidden" name="txtemail" value="<%=email%>">
<table width="95%" border="0" cellspacing="0" cellpadding="0" height="100%">
	  
		<tr> 
		  <td valign="top"> <font color="#FFFFFF">-</font><br> 
		  <table width="90%" border="0" align="center" cellpadding="0" cellspacing="0">
          <tr bgcolor="#F3F3F3"> 
            <td height="43" colspan="2"><div align="center" class="redbold">Edit 
                User</div></td>
          </tr>
          <tr> 
            <td height="35" align="right">* First Name:&nbsp;&nbsp;</td>
            <td height="35">&nbsp;&nbsp; <input type="text" name="txtFirstName" size="80" class="formTextbox" value="<%=firstname%>"  ></td>
          </tr>
          <tr> 
            <td height="35" align="right">* Last Name:&nbsp;&nbsp;</td>
            <td height="35">&nbsp;&nbsp; <input type="text" name="txtLastName" size="80" class="formTextbox" value="<%=lastname%>"></td>
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
            <td height="35" align="right">Yahoo-Email-Id:&nbsp;</td>
            <td height="35">&nbsp;&nbsp; <input type="text" name="txtYahooid" size="80" class="formTextbox" value="<%=yahooid%>"></td>
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
