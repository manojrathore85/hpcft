<!-- Description :- This file is to add user --->
<!-- Date 22 march 2k5 -->
<!-- #include file ="checksession.asp" -->
<!-- #include file ="connect.asp" -->
<!-- #include file="generalfunctions.asp" -->
<!-- #include file ="sendmail.asp" -->

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
   if (request("emailid") = "" or request("projectid") = "" or request("lstPermissions") = "none")  then
		message = "User/Project/Permission not provided to set the permission."
	else
		dim objcheck ' object for record set
		dim strSql ' string for sql
		dim counter ' for counting
		dim projectname
		set objcheck= server.CreateObject("adodb.recordset") ' setting object to check that whether the user already exist or not
		strsql= "select email from " & varTblNameUsers & " where projectid=" & request("projectid") & " and email = '" & request("emailid") & "'"
		objcheck.Open strsql,con
		if objcheck.eof = false then
			objcheck.Close 
			objcheck.Open "select email from " & varTblNameUsers & " where projectid=" & request("projectid") & " and email = '" & request("emailid") & "' and permissionid ='" & request("lstPermissions") & "'",con
			if objcheck.eof = true then
				objcheck.close
				objcheck.open "select projectname from " &  varTblNameProjects & " where projectid=" & request("projectid"),con
				if not objcheck.eof then
					projectname = objcheck("projectname")
				end if
				objcheck.close
				if request("lstPermissions") = "noperm" then
					con.execute "delete from " &  varTblNameUsers & " where projectid=" & request("projectid") & " and email = '" & request("emailid") & "'"
				else
					con.execute "update " &  varTblNameUsers & " set permissionid = '" & request("lstPermissions") & "' where projectid=" & request("projectid") & " and email = '" & request("emailid") & "'"
				end  if

				call sendmail
				
				message = "Permission edited Successfully."
				if objcheck.state = 1 then objcheck.close
			    set objcheck = nothing    
				con.close
				set con = nothing
				response.Redirect("viewprojects.asp?mainlink=setpermissionlist")
		  	else
			   message = "A permission already set to the selected user with selected permission." 
		  	end if
		else
			message = "No records found to edit."
		end if
	    if objcheck.state = 1 then objcheck.close
	    set objcheck = nothing    
	end if
end if

private sub sendmail_old()
		dim Mailer
		Set Mailer = Server.CreateObject("SMTPsvg.Mailer")
		Mailer.FromName   = session("name")
		Mailer.FromAddress = varSiteSpecFromAddress 
		Mailer.RemoteHost = varSiteSpecRemoteHost
		Mailer.AddRecipient "",request("emailid")
		if varSiteSpecAddCC <> "" then
			mailer.addCC "", varSiteSpecAddCC
		end if
		Mailer.Subject    = varSiteSpecMailSubjectPrefix & "<" & varSpecialclientid & ">" & " Permission changed"
		Mailer.BodyText   = "Project :-" & Projectname & "<br>You are now assigned permission as :-" & request("lstPermissions") 
		Mailer.SendMail
end sub

private sub sendmail()

		FromName   = session("name")
		dim Receipent(0)
		Receipent(0) = request("emailid")
		Subject    = varSiteSpecMailSubjectPrefix & "<" & varSpecialclientid & ">" & " Permission changed"
		Body   = "Project :-" & Projectname & "<br>You are now assigned permission as :-" & request("lstPermissions") 
		BCC = ""
		mail FromName,ReplyTo,Receipent,varSiteSpecAddCC,BCC,Subject,Body
		
end sub


	dim rsp
	set rsp = server.CreateObject("adodb.recordset")
	dim rspermissions
	set rspermissions = server.CreateObject("adodb.recordset")
	rsp.open "select up.firstname,up.lastname,p.projectname,u.permissionid from " &  varTblNameUserProfile & " up," &  varTblNameProjects & " p," &  varTblNameUsers & " u where u.email = up.email and u.projectid = p.projectid and u.email ='" & request("emailid") &"' and u.projectid=" & request("projectid"),con
	rspermissions.open "select permissionid from " &  varTblNamePermissions & "",con
%>
<!-- *****************    Script Java language ****************  --->
<script language="JavaScript"  src="include/Validate.js"></script> <!-- including the Validate file -->
<script language ="Javascript">
 //checking the validation
 function checkvalidation()
  {
	
    
			if(frmeditsetpermission.lstPermissions.value=="none")
			  {
				 alert('Please select the  permission ');
				 frmeditsetpermission.lstPermissions.focus();  
			     return false;  
			  }       
			  //
			
		    
		 
		  
  } //end of function
</script>
<!-- <form name="frmaddpermission" method="post" action="addpermission.asp" onsubmit="return checkvalidation();">  --><!-- form start from here -->
<form name="frmeditsetpermission" method="post" action="viewprojects.asp" onsubmit="return checkvalidation();"> <!-- form start from here -->
<input type="hidden" value="editsetpermission" name="mainlink">
<input type="hidden" value="<%=request("emailid")%>" name="emailid">
<input type="hidden" value="<%=request("projectid")%>" name="projectid">
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
        <%if rsp.eof = false and rspermissions.eof = false then%>
        <table width="90%" border="0" align="center" cellpadding="0" cellspacing="0">
          <tr bgcolor="#F3F3F3"> 
            <td height="43" colspan="2" class="redbold"><div align="center" class="redbold">Edit 
                Permission for User</div></td>
          </tr>
          <tr> 
            <td height="35" align="right">* User:&nbsp;&nbsp;</td>
            <td height="35">&nbsp;&nbsp; <input type="text" name="txtUser" class="formTextbox" value="<%=rsp("firstname") & " " & rsp("lastname")%>" size="50" readonly=""></td>
          </tr>
          <tr> 
            <td height="35" align="right">* Project:&nbsp;&nbsp;</td>
            <td height="35">&nbsp;&nbsp; <input type="text" name="txtProject" class="formTextbox" value="<%=rsp("projectname")%>" size="50"></td>
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
			<option value="noperm">no perms</option>
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
		<%else%>
		Users,Projects OR Permissions Missing
		<% end if %>
      </td>

    </tr></table>
 </form>
</body>
</html>
<script>
<%if rsp.eof = false then%>
for(i=0;i<frmeditsetpermission.lstPermissions.options.length;i++)
	if(frmeditsetpermission.lstPermissions.options[i].value == '<%=rsp("permissionid")%>')
		frmeditsetpermission.lstPermissions.selectedIndex = i;
<%end if%>
</script>
<%
if rsp.state = 1 then rsp.close
if rspermissions.state = 1 then rspermissions.close
set rsp = nothing
set rspermissions = nothing
con.close
set con = nothing 
%>