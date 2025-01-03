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
   if request("cmdSubmit") <>"" then
		dim objAdduser,objcheck ' object for record set
		dim strsql ' string for sql
		dim counter ' for counting
	
	    	set objcheck= server.CreateObject("adodb.recordset") ' setting object to check that whether the user already exist or not
	    	'objcheck.Open "select * from userprofile where 1=2",con,adOpenDynamic,adLockOptimistic
	    	'if objcheck.EOF then
	    '		Response.Write("asdfj")
	    '		Response.End
	   ' 	end if
	    		
	    	
			 strsql= "select count(*) from " &  varTblNameUserProfile & " where email='" & request("txtemail") & "'"
              
                   objcheck.Open strsql,con
              
                  counter =cint(objcheck(0)) 'for storing the result
                  
                   objcheck.Close 
                   set objcheck = nothing
                if counter= 0 then
				  		  
				 '	 set  objAdduser =server.CreateObject("adodb.recordset")
			'		 objAdduser.Open "select * from userprofile where 1=2",con, 2,3 ',3  'opening the empty record
			'	 	objAdduser.AddNew  
			'	 	          	objAdduser("Email")= cstr(request("txtEmail"))
			'					objAdduser("password")=cstr(request("txtPassword"))
			'					objAdduser("Fullname")= request("txtName")
			'					objAdduser("Address")=  request("txtAddress")
			'					objAdduser("Mobile")= request("txtmobile")
			'					objAdduser("Hphone")= request("txtPhoneH")
			'					objAdduser("ophone")= requesT("txtPhoneO")
			'		objAdduser.update
			
			 strsql= "insert into " &  varTblNameUserProfile & "(Email,password,FirstName,LastName,address,mobile,Hphone,Ophone) values ( '" & trim(request("txtEmail")) & "','" & trim(request("txtpassword"))& "','" & trim(request("txtfirstname")) & "','" & trim(request("txtlastname")) & "','" & trim(request("address")) & "','" & trim(request("txtmobile")) & "','" & trim(request("txtphoneH")) & "','" & trim(request("txtphoneO")) & "')" 
			 
			 
			con.execute strsql
					message = "Record is saved successfully."
				
						' nullify unused object
					' objAdduser.Close()
					' set objAdduser=nothing
					 
				  else
				       message = "The user of this e-mail already exist" 
				   
				  end if    
				  	 
			 con.close
			 set con = nothing 
			 
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
			  //**************** 
			if (frmadduser.txtEmail.value =="")
				 {
				    alert('Please Ente the Email-id');
				     frmadduser.txtEmail.focus();
				     return false;   
				 } 
			
				 //******************  
			 if (frmadduser.txtEmail.value!="")
				 {
					   str=frmadduser.txtEmail.value;
					  		 flag=checkEmail(str); //****** calling from validation.js
							if (flag==false)
							 {
							    alert('Please Enter the Email-Id in correct mail Format');
							    frmadduser.txtEmail.value ="";
							    frmadduser.txtEmail.focus();
							    return false;
							 }      
					           
					} 
		 //checking the password  
			 if (frmadduser.txtPassword.value=="")
			   {
			      frmadduser.txtPassword.focus();
				  alert('Please Enter the password');
				 return false;  
			   }
		  // checking the confirm password 
		  
		  if (frmadduser.txtConfirmPass.value=="")
		    {
		       frmadduser.txtConfirmPass.focus();
		       alert('Please Enter the confirmation password');
		       return false;
		     }       
		   // checking the both passowrd 
		   if (frmadduser.txtPassword.value != frmadduser.txtConfirmPass.value)
		     {
		    
		       frmadduser.txtConfirmPass.value=="";  
		       frmadduser.txtConfirmPass.focus();  
		          alert('Please Enter the simillar password in both password fields');
		      return false;
		     }
		     
		    
		 
		  
  } //end of function
</script>
<form name="frmadduser" method="post" action="viewprojects.asp" onsubmit="return checkvalidation();"> <!-- form start from here -->
<input type="hidden" value="adduser" name="mainlink">
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
		        <td height="43" colspan="2"><div align="center" class="redbold">Create New User</div></td>
		      </tr>
		      <tr> 
		        <td height="35" align="right">* First Name:&nbsp;&nbsp;</td>
		        <td height="35">&nbsp;&nbsp; <input type="text" name="txtFirstName" size="60" class="formTextbox"></td>
		      </tr>
			  <tr> 
		        <td height="35" align="right">* Last Name:&nbsp;&nbsp;</td>
		        <td height="35">&nbsp;&nbsp; <input type="text" name="txtLastName" size="60" class="formTextbox"></td>
		      </tr>
		      <tr> 
		        <td height="35" align="right">Contact Address:&nbsp;&nbsp;</td>
		        <td height="35">&nbsp;&nbsp; <input type="text" name="txtAddress" size="60" class="formTextbox"></td>
		      </tr>
		      <tr> 
		        <td width="26%" height="35" align="right">Mobile:&nbsp;&nbsp;</td>
		        <td width="74%" height="35">&nbsp;&nbsp; <input type="text" name="txtMobile" size="60" class="formTextbox"></td>
		      </tr>
		      <tr> 
		        <td width="26%" height="35" align="right"> Phone(H):&nbsp;&nbsp;</td>
		        <td width="74%" height="35">&nbsp;&nbsp; <input type="text" name="txtPhoneH" size="60" class="formTextbox"></td>
		      </tr>
		      <tr> 
		        <td width="26%" height="35" align="right">Phone(O):&nbsp;&nbsp;</td>
		        <td width="74%" height="35">&nbsp;&nbsp; <input type="text" name="txtPhoneO" size="60" class="formTextbox"></td>
		      </tr>
		      <tr> 
		        <td height="35" align="right"> *Email-Id:&nbsp;&nbsp;</td>
		        <td height="35">&nbsp;&nbsp; <input type="text" name="txtEmail" size="60" class="formTextbox"></td>
		      </tr>
		      <tr> 
		        <td height="35" align="right">*Password:&nbsp;&nbsp;</td>
		        <td height="35">&nbsp;&nbsp; <input type="password" name="txtPassword" size="60" class="formTextbox"></td>
		      </tr>
		      <tr> 
		        <td height="35" align="right">*Confirm Password:&nbsp;&nbsp;</td>
		        <td height="35">&nbsp;&nbsp; <input type="password" name="txtConfirmPass" size="60" class="formTextbox"></td>
		      </tr>
		      <tr> 
		        <td height="35" align="right">&nbsp;</td>
		        <td height="35">&nbsp;&nbsp; </td>
		      </tr>
		      <tr> 
		        <td width="26%" height="35" align="right" valign="top">&nbsp;&nbsp;</td>
		        <td width="74%" height="35" align="center"><input type="submit" value="Create" name="cmdSubmit" > 
		          &nbsp;&nbsp; <input  type="reset" value="Cancel" name="cmdCancel"></td>
		      </tr>
 </table>
 </td></tr></table>
 </form>
</body>
</html>
