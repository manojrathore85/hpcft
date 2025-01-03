<!-- Description :- This file is to add user --->
<!-- Date 22 march 2k5 -->
<!-- #include file =" checksession.asp" -->
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
if request("cmdSubmit") = "Set Permission" then 
	   if (request("lstUsers") = "none" or request("lstProjects") = "none" or request("lstPermissions") = "none")  then
	   		message = "User/Project/Permission not provided to set the permission."
		else
			dim objAddUserPerms,objcheck ' object for record set
			dim strSql ' string for sql
			dim counter ' for counting
			set objcheck= server.CreateObject("adodb.recordset") ' setting object to check that whether the user already exist or not
			strsql= "select email from " &  varTblNameUsers & " where projectid=" & request("lstProjects") & " and email = '" & request("lstUsers") & "'"
			objcheck.Open strsql,con
			if objcheck.eof = true then
				objcheck.Close 
				set  objAddUserPerms =server.CreateObject("adodb.recordset")
				 objAddUserPerms.CursorLocation = 3 'adUseClient
				 objAddUserPerms.Open "select * from " &  varTblNameUsers & " where 1=2",con, 3,3 ',3  'opening the empty record
					objAddUserPerms.AddNew
					objAddUserPerms("email") = request("lstUsers")
					objAddUserPerms("projectid") = request("lstProjects")
					objAddUserPerms("Permissionid") = request("lstPermissions")
				objAddUserPerms.update
				objAddUserPerms.close
				set objAddUserPerms = nothing
				message = "Permission Given Successfully."
			  else
				   message = "A permission already set to the selected user with selected permission." 
			  end if
			  if objcheck.state = 1 then objcheck.close
			  set objcheck = nothing    
	end if		  
end if		 
    dim rs
	dim rsusers
	dim rsprojects
	dim rspermissions
	set rs = server.CreateObject("adodb.recordset")
	set rsusers = server.CreateObject("adodb.recordset")
	set rsprojects = server.CreateObject("adodb.recordset")
	set rspermissions = server.CreateObject("adodb.recordset")
	rs.open "select up.firstname,up.lastname,p.projectname,u.* from " &  varTblNameUserProfile & " up," &  varTblNameProjects & " p," &  varTblNameUsers & " u where u.email = up.email and u.projectid = p.projectid",con
	rsusers.open "select firstname,lastname,email from " &  varTblNameUserProfile & "",con
	rsprojects.open "select projectid,projectname from " &  varTblNameProjects & "",con
	rspermissions.open "select permissionid from " &  varTblNamePermissions & "",con
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
<!-- <form name="frmaddpermission" method="post" action="addpermission.asp" onsubmit="return checkvalidation();">  --><!-- form start from here -->
<form name="frmaddpermission" method="post" action="viewprojects.asp"> <!-- form start from here -->
<input type="hidden" value="setpermission" name="mainlink">
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
        <%if rsusers.eof = false and rsprojects.eof = false and rspermissions.eof = false then%>
        <table width="90%" border="0" align="center" cellpadding="0" cellspacing="0">
          <tr bgcolor="#F3F3F3"> 
            <td height="43" colspan="2"><div align="center" class="redbold">Give 
                Permission to User</div></td>
          </tr>
          <tr> 
            <td height="35" align="right">* User:&nbsp;&nbsp;</td>
            <td height="35">&nbsp;&nbsp; <select name="lstUsers" class="formTextbox">
			<option value="none" selected>Select User</option>
			<%while not rsusers.eof%>
			<option value="<%=rsusers("email")%>"><%=rsusers("firstname") & " " & rsusers("lastname")%></option>
			<%
			rsusers.movenext
			wend
			%>
              </select></td>
          </tr>
          <tr> 
            <td height="35" align="right">* Project:&nbsp;&nbsp;</td>
            <td height="35">&nbsp;&nbsp;
              <select name="lstProjects" class="formTextbox">
                <option value="none" selected>Select Project</option>
			<%while not rsprojects.eof%>
			<option value="<%=rsprojects("projectid")%>"><%=rsprojects("projectname")%></option>
			<%
			rsprojects.movenext
			wend
			%>
              </select></td>
          </tr>
          <tr> 
            <td height="35" align="right">* Permission:&nbsp;&nbsp;</td>
            <td height="35">&nbsp;&nbsp; <select name="lstPermissions" class="formTextbox">
                <option value="none" selected>Select Permission</option>
			<%while not rspermissions.eof%>
			<option value="<%=rspermissions("permissionid")%>"><%=rspermissions("permissionid")%></option>
			<%
			rspermissions.movenext
			wend
			%>
              </select></td>
          </tr>
          <tr> 
            <td height="35" align="right">&nbsp;</td>
            <td height="35">&nbsp;&nbsp; </td>
          </tr>
          <tr> 
            <td width="26%" height="35" align="right" valign="top">&nbsp;&nbsp;</td>
            <td width="74%" height="35" align="left">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input type="submit" value="Set Permission" name="cmdSubmit" class="formbutton">
            &nbsp;&nbsp; <input  type="reset" value="Cancel" name="cmdCancel" class="formbutton"></td>
          </tr>
        </table>
		<% end if %>
		<p align="center" class="redbold">Existing Premisions</p>
		<table width="100%" border="0" cellpadding="1" cellspacing="1">
          <tr bgcolor="#FFFFFF" class=head> 
            <td id=head>&nbsp;User Name</td>
            <td id=head>Email-id</td>
            <td id=head>Project Name</td>
            <td id=head>Permission</td>
            <td>&nbsp;Operations</td>
          </tr>
          <% 
    if rs.bof=false and rs.eof =false then ' checking whether the record is exist or not
	   while not rs.eof ' checkout all the user in database
%>
          <tr bgcolor="#FFFFFF"> 
            <td>&nbsp;<%=rs("firstname") & rs("lastname") %> </td>
            <td>&nbsp;<%=rs("email")%></td>
            <td>&nbsp;<%=rs("projectname")%></td>
            <td>&nbsp;<%=rs("permissionid")%></td>
            <td>&nbsp;<a href="viewprojects.asp?mainlink=setpermission&projectid=<%=rs("projectid")%>&emailid=<%=rs("email")%>">Edit</a> 
              | <a href="viewprojects.asp?mainlink=deleteuserpermission&projectid=<%=rs("projectid")%>&emailid=<%=rs("email")%>">Delete</a></td>
          </tr>
          <%
	    rs.movenext 
		wend
	else
	%>
          <tr bgcolor="#FFFFFF"> 
            <td colspan="5" align="center">NO PERMISSIONS GIVEN TO USERS</td>
          </tr>
          <%

    end if		
   %>
        </table>
		
		</td></tr></table>
 </form>
</body>
</html>
<%
con.close
set con = nothing 
%>