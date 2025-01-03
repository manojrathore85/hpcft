<!---   Description : This page is basically for giving the comment releated to particular issue  23 march 2k5 rajat ---->
<%
   dim projectname
  if request("issueid")<>"" then ' checking what is the no coming form there --
		 dim objview ' set view object
		 set objview =server.CreateObject("adodb.recordset")
		 objview.open "select projectname from " &  varTblNameIssues & " where issueId=" & request("pid"),con
		 
		  if objview.eof =false and objview.bof =false then
		  projectname=objview("projectName") 'stroing the name of the project
		  end if
		  ' nullify all the object 
		  objview.close
		  set objview = nothing
  end if	 
 '******************************************************  end of showing then name of the project  
 if request("
 
%>

<!-- ********************************* client side validation useing java Script **************************** -->
   <script language="JavaScript">
   
     function checkvalidate()
      {
        if (frmcomment.txtcomment.value =="") 
        {
          alert('Please Enter the some comment');
          frmcomment.txtcomment.focus();
          return false; 
      }
   </script>
<!-- ******************************* HTML CODING STATRED HERE ---->
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<LINK href="include/style.css" type="text/css" rel="stylesheet">
</HEAD>
<BODY  leftmargin="0" topmargin="0">
<table  border="0" cellspacing="0" cellpadding="0"  align="center">
<tr>
   <td> Project Name </td>
   <td>&nbsp;<%=projectName%></td>
</tr>
<!-- Begin of form goes here -->

<form name="frmcomment" action="issuecomments.asp" method="get" onsubmit="return checkvalidate();">
			<tr>
			 <td class="regularText" > Comments</td>
			 <td><textarea name="txtcomment" class="formTextbox" cols="30"></textarea></td>
			
			</tr>
			<tr>
			  <td><input type="submit" name="cmdSubmit" value="Save" class="formbutton" ></td>
			  <td><input type="reset" name="cmdClear" value="Clear" class="formbutton"></td>
			</tr>
	<!-- End of Form goes Here -->		
</form>
</table>




</BODY>
</HTML>
