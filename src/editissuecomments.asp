<!DOCTYPE HTML>
<%@Language="VBScript" CodePage="65001" %>
<%
Session.CodePage = 65001
Response.CodePage = 65001
Response.CharSet = "utf-8"
'CR20060902
Session.LCID     = 1033 'en-US
%>
<!-- #include file="checksession.asp" -->
<!-- #include file="generalfunctions.asp" -->
<!-- #include file ="Connect.asp" -->
<!-- #include file ="sendmail.asp" -->
<%

' ********************* CHECKING PERMISSION ***************************

dim rsperm
dim prevdata
set rsperm = server.CreateObject("adodb.recordset")
rsperm.open "select perm.* from " &  varTblNamePermissions & " perm, " &  varTblNameUsers & " u, " &  varTblNameissues & _
 " i where i.projectid = u.projectid and u.permissionid = perm.permissionid and i.issueid = " &  request("issueid") & _
 " and u.email ='" &  session("user") & "'",con
if not rsperm.eof then
	if rsperm("pdelete") = "T" then
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



private sub sendmail()
	dim rstemp
	dim i
	dim reportUsers ' users added in reportusers of a issue
	dim reportUsers1 ' users added in projectwatchlist
	dim rptuser ' this variable is used to keep the value in recordset, for it is does not come after checking the recordset for eof
	dim projectwatchlist ' this variable is used to keep the value in recordset, for it is does not come after checking the recordset for eof
	dim temp
	set rstemp = server.CreateObject("adodb.recordset")
	rstemp.open "select p.projectname,p.projectwatchlist,i.summary,i.assignto,i.reportUsers from " &  varTblNameProjects & " p," &  varTblNameIssues & " i where i.projectid = p.projectid and i.issueid =" & request("issueid") ,con
	projectwatchlist = rstemp("projectwatchlist")
	rptuser = rstemp("reportUsers")
	if not rstemp.eof then
		if not isnull(rptuser) and rptuser <> "" then
			reportUsers = split(rptuser,",")
		end if
		if not isnull(projectwatchlist) and projectwatchlist <> "" then
			reportUsers1 = split(projectwatchlist,",")
		end if
		FromName   = session("name")
		ReplyTo = session("user")
		dim Receipent()
		redim Receipent(0)  
		Receipent(0) = rstemp("assignto")
		if isarray(reportUsers) then
			for i = 0 to ubound(reportUsers)
				redim preserve Receipent(ubound(Receipent)+1)    
				Receipent(ubound(Receipent)) = reportUsers(i)
				'Mailer.AddRecipient "",replace(replace(replace(reportUsers(i),chr(13),""),chr(32),""),chr(10),"")
			next
		end if
		if isarray(reportUsers1) then
			for i = 0 to ubound(reportUsers1)
				redim preserve Receipent(ubound(Receipent)+1)    
				Receipent(ubound(Receipent)) = reportUsers1(i)
				'Mailer.AddRecipient "",replace(replace(replace(reportUsers1(i),chr(13),""),chr(32),""),chr(10),"")
			next
		end if
		if isarray(reportUsers2) then
			for i = 0 to ubound(reportUsers2)
				redim preserve Receipent(ubound(Receipent)+1)    
				Receipent(ubound(Receipent)) = reportUsers2(i)
				'Mailer.AddRecipient "",replace(replace(replace(reportUsers2(i),chr(13),""),chr(32),""),chr(10),"")
			next
		end if

		Subject    =  "(" & varSiteSpecMailSubjectPrefix & ")" & varSpecialclientid & " | " & "A comment updated - " & rstemp("summary") & " | " & rstemp("projectname") 'CR20060821Atul
		Body   = "Project :-" & rstemp("projectname") & "<br>" & "Issue Summary:-" & rstemp("summary") & "<br>" & varSiteSpecURL & Mid(varServerVariablePath_info, 1, InStrRev(varServerVariablePath_info, "/")) & "issuedetails.asp?issueid=" & request("issueid") & "<br>" & "Prev Comments :" & "<br>" &  prevdata & "<br>" & "-------------------------------------------------" & "<br>" & "Updated Comments :" & "<br>" & request("txtcomments")	
		BCC = ""

		mail FromName,ReplyTo,Receipent,varSiteSpecAddCC,BCC,Subject,Body
	end if
	rstemp.close
	set rstemp = nothing
end sub


if request("id") <> "" then
	dim rs
	dim rowid
	
	set rs  = server.CreateObject("adodb.recordset")
	rs.open "select * from " &  varTblNameIssueChangesComments & " where Rowid=" & clng(request("id")) ,con
	if not rs.eof then
		rowid = rs("rowid")
		prevdata = rs("comments")
	end if
	rs.close
	set rs = nothing
end if

if request("id") <> "" and request("cmdSubmit") = "Update" then
	
	dim comm
	comm = replace(trim(request("txtcomments")),"'","\'")
	comm = replace(comm,chr(34),"\""")  'chr(34) is for double quote ( " )
	con.execute "update " & varTblNameIssueChangesComments & " set comments = '" & comm & "' where Rowid=" & clng(request("id"))
	sendmail
end if
con.close
set con = nothing
'response.Redirect("issuedetails.asp?issueid=" & request("issueid"))
%>

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<title>Untitled Document</title>
<!-- *********************** Java SCript ******************************* -->
<script language="JavaScript" src="include/encrypt.js" type="text/javascript"></script>
<script src="https://cdn.tiny.cloud/1/jfhvmzju5ge1ljiv30evfvsh676206icpk2z7i430ny5ouoo/tinymce/6/tinymce.min.js" referrerpolicy="origin"></script>
<script language="Javascript">
self.moveTo(0,0);

function encrypt()
{
	if(frmissue.chkEncrypt.checked == true)
		if(frmissue.enKey.value == "" || frmissue.comments.value == "")
		{
			alert('Please enter encryption key and comments to encrypt'); 
			frmissue.chkEncrypt.checked = false;
		}
		else {
			frmissue.encodedmessage.value = secureEncrypt(frmissue.comments.value,frmissue.enKey.value);
			//frmissue.encodedmessage.value = varMsgCopy;
		}
	else
	{
		frmissue.enKey.value = "";
		frmissue.encodedmessage.value = "";
	}
}

 function checkvalidate()
  {
	 console.log(tinymce.activeEditor.getContent());
	 if(frmissue.encodedmessage.value != ""){
		frmissue.txtcomments.value = frmissue.encodedmessage.value;
		}
	else{
		frmissue.txtcomments.value = tinymce.activeEditor.getContent();
		}
		
		
  }
 </script>
 

</head>

<body>

<form name="frmissue" method="post" onSubmit="return checkvalidate();" >
<input type="hidden" name="hdn_prevcomment" value="<%'=prevdata%>">
<table width="98%" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr bgcolor="#F3F3F3">
    <td height="43" colspan="2"><div align="center" class="redbold">Edit 
      Issue Comments </div></td>
  </tr>
  <!-- adding the commentss -->
  <tr>
    <td height="35" align="right" valign="top">&nbsp;</td>
    <td height="35"> Encrypt:
      <input type="checkbox" name="chkEncrypt" onClick="encrypt();" value="ON" />
      &nbsp;&nbsp;&nbsp;&nbsp;Encryption&nbsp;Key:
      <input type="password" name="enKey" size="50"></td>
  </tr>
  <tr>
    <td width="15%" height="35" align="right" valign="top">Comments:&nbsp;&nbsp;</td>
    <td width="85%" height="35">&nbsp;&nbsp;
        <textarea class="formTextbox" name="comments" id="txtcomments1" cols="70" rows="10" wrap="virtual"><%=prevdata%></textarea></td>
  </tr>
  <tr>
    <td width="15%" height="35" align="right" valign="top">&nbsp;&nbsp;</td>
    <td width="85%" height="35" align="center"><input type="submit" value="Update" name="cmdSubmit" />
      &nbsp;&nbsp;
      <input  type="button" value="Cancel" name="cmdCancel" /></td>
  </tr>
  <tr>
    <td height="35" align="right" valign="middle">Encoded 
      Text:&nbsp;&nbsp;</td>
    <td height="35" valign="top" align="left">&nbsp;&nbsp;
        <textarea class="formTextbox" name="encodedmessage" cols="70" rows="2" wrap="virtual" readonly="readonly"></textarea>
        <input type="hidden" name="txtcomments"/>    </td>
  </tr>
</table>
</form>
</body>
</html>
<script>


tinymce.init({
  selector: '#txtcomments1',  // change this value according to your HTML
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

<%if request("id") <> "" and request("cmdSubmit") = "Update" then%>
<script language="JavaScript">
self.close();
opener.location.reload();
</script>
<%end if%>