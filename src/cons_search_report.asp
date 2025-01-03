<!-- #include file="checksession.asp" -->
<!-- #include file="connect.asp" -->
<!-- #include file="generalFunctions.asp" --> <!-- including the functions file -->
<%
' ********************* CHECKING PERMISSION ***************************

'dim rsperm
'set rsperm = server.CreateObject("adodb.recordset")
'rsperm.open "select * from " &  varTblNamePermissions & " where permissionid = '" & session("permission") & "'",con
'if not rsperm.eof then
'	if rsperm("pread") = "T" then
'		rsperm.close
'		set rsperm = nothing
'	else
'		server.Execute("permissiondenied.asp")
'		rsperm.close
'		set rsperm = nothing
'		response.End()
'	end if
'else
'		server.Execute("permissiondenied.asp")
'		rsperm.close
'		set rsperm = nothing
'		response.End()
'end if

' ********************* CHECKING PERMISSION ENDS***************************

dim rs 
dim strsql
set rs= server.CreateObject("adodb.recordset") 

if request("chkFullText") = "True" then  ' if full text search requested then only
		strsql = "select i.*,ic.newvalue,ic.lastvalue,ic.comments,p.projectname from " &  varTblNameIssues & " i," &  varTblNameUsers & " u," &  varTblNameProjects & " p, " & varTblNameIssuechangescomments & " ic " & _
		"where (ic.newvalue like '%" & request("txtIssSummary") & "%' or ic.lastvalue like '%" & request("txtIssSummary") & "%' or ic.comments like '%" & request("txtIssSummary") & "%' " & _
		" or i.summary like '%" & request("txtIssSummary") & "%' or i.description like '%" & request("txtIssSummary") & "%')" 	

		if request("optProject") <> "ALL_PROJECT" then	
			strsql = strsql  & " and i.projectid =" & session("projectid")
		end if
		strsql = strsql  & " and  u.email = '" & session("user") & "' and " & _	
		"i.projectid = p.projectid and i.projectid = u.projectid and i.issueid = ic.issueid group by i.issueid "

		rs.open strsql ,con

end if
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<title><%=setTitle%></title>
<LINK 
href="include/style.css" type=text/css rel=stylesheet>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1"></head>
<body bgcolor="#CCCCCC" leftmargin="0" topmargin="0">
<br>
<%

if not rs.eof  then
	while not rs.eof 
	   projectname =rs("projectname")
	   createdate=DateAdd("h", -5, rs("createdate")) 'for showing eastern standard time as it is 5 hrs less then GMT
	   updatedate =DateAdd("h", -5, rs("updateDate"))'for showing eastern standard time as it is 5 hrs less then GMT
	   issuetype=rs("issuetype")
	   severity=rs("severity")
	   summary =rs("summary")
	   assignto =rs("assignto")
	   reporter=rs("reporter")
	   attachedfilepath = rs("attachedfilepath")
	   if isnull(rs("attachedfilepath")) = false then
			filename = split(attachedfilepath,",")
	   end if
	   description= rs("description") & ""
	   status =rs("status")
	   show_content(rs("issueid"))
	   response.Write("<br>")
	   rs.movenext
	wend
end if		

sub format_and_print(mydata)
	mydata = formatData(mydata) ' formatData function called in generalFunctions.asp
	mydata = highlight(mydata,request("txtIssSummary"))
	if len(mydata) >= 100 then
		if InStr(1, mydata, " ") = 0 and InStr(1, mydata, "-") = 0 then
			for int_i = 1 to len(mydata)
				copy_data = copy_title_data & mid(mydata ,int_i,100) & "-"
				int_i = int_i + 100
			next
			response.Write(copy_title_data)
			copy_title_data = ""
		else
			response.Write(mydata)
		end if
	else
		response.Write(mydata)
	end if
end sub

%>

<%
sub show_content(myissueid)
'on error resume next
'if err.number <> 0 then
'	response.Write(err.description)
'	err.clear
'end if
%>
<table width="100%" height="100%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="#FFFFFF">
<tr><td width="5%">&nbsp;</td><td width="95%">
<table	border="0" cellspacing="0" cellpadding="0" width="100%">
	 <tr>
		<td class="bluebold" colspan="2" align="LEFT"><font size="+1"><%=projectname%></font>
		<br>
		&nbsp;&nbsp;
		<span class="regularText"><%=summary%></span><br>
		<span class="regularTextSmall">Created:</span><span class="regularTextBlueSmall">&nbsp;<%=createdate%></span><span class="regularTextSmall">&nbsp;&nbsp;Updated:</span>&nbsp;<span class="regularTextBlueSmall"><%=updateDate%></span><br>&nbsp;
		</td> <!-- displaying the project name on which the issue is created  -->
	 </tr>
	 <tr>
		<td class="regularTextBold">&nbsp;&nbsp;Issue URL </td> <!-- issue type of the issue -->
		<td>:&nbsp;&nbsp;<%="<a href='" & varSiteSpecURL & "/issuedetails.asp?issueid=" & myissueid & "'>" & varSiteSpecURL & "/issuedetails.asp?issueid=" & myissueid & "</a>" %> </td> <!-- severity of the issue goes here -->
	 </tr>
	 <tr>
		<td class="regularTextBold">&nbsp;&nbsp;Issue Type </td> <!-- issue type of the issue -->
		<td>:&nbsp;&nbsp;<%=issuetype%> </td> <!-- severity of the issue goes here -->
	 </tr>
	 <tr>
	   	<td class="regularTextBold">&nbsp;&nbsp;Severity</td> <!-- assign to details goes here -->
		<td>:&nbsp;&nbsp;<%=severity%> </td>
	 </tr>
	 <tr>
        <td class="regularTextBold">&nbsp;&nbsp;Reporter</td>
		<td>:&nbsp;&nbsp;<%=reporter%></td>
	 </tr>
	 <tr>  
	    <td class="regularTextBold">&nbsp;&nbsp;Assign To</td>
		<td>:&nbsp;&nbsp;<%=Assignto%></td>
	 </tr>
	 <tr>  
	    <td class="regularTextBold">&nbsp;&nbsp;Status</td>
		<td>:&nbsp;&nbsp;<span class="redbold"><%=status%></span></td>
	 </tr>
	 <tr>  
        <td class="regularTextBold">&nbsp;&nbsp;File Attached</td>
		<td valign="middle">
			<table width="100%" cellpadding="0" cellspacing="0">
				 <tr>
					<td width="2%">:&nbsp;&nbsp;</td>
					<td width="98%"><%
				if isArray(filename) then
					if ubound(filename) >= 0 then
						for counter=0 to ubound(filename)
						if counter > 0 then
							response.Write(", ")
						end if
						%>
						<a href="upload/<%=filename(counter)%>" target="_blank"><%=check_file_ondisk(filename(counter))%><%'=filename(counter)%></a>
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
   		 <td class="regularTextBold" colspan="2">&nbsp;&nbsp;Description :- </td>
	</tr>
	<tr>
		 <td colspan="2">
			 <table width="100%" cellpadding="0" cellspacing="0">
				 <tr>
				 	<td width="3%">&nbsp;</td>
					<td width="97%"><%format_and_print(description & "")%></td>
				 </tr>
			 </table>
		 </td>
	 </tr>
</table>
<br>

	<!-- *************************** showing the all the comment of the issue if exist  order by date of the update which is goes here  ************* -->
	<%
		   dim objissue ' issue change comments 
		   set objissue =server.CreateObject("adodb.recordset") 'setting the record oobhject for issue chages goes here 
			   objissue.open "select * from " &  varTblNameIssueChangesComments & " where issueid=" & myissueid & " order by updateDate",con
		   dim dattemp
		   dim flag
		   	dattemp = ""
			counter = 1
			flag = false
			if objissue.eof =false then 
				while not objissue.eof 	   
					if  dattemp <> objissue("updatedate") and isnull(objissue("comments")) then
					dattemp =objissue("updatedate")	 
         	 %>
					<table border=1 width="100%" cellpadding="0" cellspacing="0">
					   <tr class="regularTextBold"> 	
						<td colspan="3"> Change By : &nbsp; <%=objissue("changesby")%> &nbsp; &nbsp; [ <%=objissue("updatedate")%> ] &nbsp;</td>
						</tr>
						<tr  bgcolor="#CCCCCC" class="regularTextBold">
								<td width="20%">Field</td>
								<td width="40%"> Old Value </td>
								<td width="40%"> New Value</td>
						</tr>
						<tr>
								<td><%=objissue("field")%></td> 
								<td><%format_and_print(objissue("lastvalue") & "")%></td>
								<td><%format_and_print(objissue("newvalue") & "")%></td>
						   </tr>
					
				   'while (not )
				   flag = true
				   response.Write("</table><br>")
				   %>
				<% elseif not isnull(objissue("comments")) then%>
					<%flag = false%>
					   <table border=1 width="100%">
					   <tr bgcolor="#CCCCCC" class="regularTextBold"> 	
						<td colspan="3"> Comments By : &nbsp; <%=objissue("changesby")%> &nbsp; &nbsp; [ <%=objissue("updatedate")%> ] &nbsp;<a href="deleteissuecomments.asp?id=<%=objissue("rowid")%>&issueid=<%=request("issueId")%>" onClick="return checkDelete();">Delete</a> | <a href="#" onClick="PopUp('editissuecomments.asp?id=<%=objissue("rowid")%>&issueid=<%=request("issueid")%>','EditReminders','status=yes,scrollbars=yes,width=700,height=400')"><b>Edit</b></a></td>
						</tr>
						<tr>
						<td colspan="3" id="comnt<%'=counter%>" style="overflow:scroll;"><%format_and_print(objissue("comments") & "")%></td>	
						</tr></table><br>	
				  	<% 
					counter = counter + 1
					end if 
					%>
				<%
				
				objissue.movenext
			  wend
			  %>
			  <% 
			end if
			objissue.close
			set objissue = nothing	  
    	  %>			
		</td>
	</tr>
</table>
</td></tr></table>
<%
end sub
%>
<script>	
//CR20060706
// The below javascript keep a copy of all the comments in array or single variable for further use.
//var ab= "asdflasjdf lkjas";
//alert(ab.indexOf(' '));
varCmnt = "";
varCmntCopy = "";
mySingleComment = ""
if(typeof document.all("comnt").length == "undefined")	{	
	mySingleComment = document.all("comnt").innerText;
	varCmnt = document.all("comnt").innerText;
	if(varCmnt.length >= 100) {
		if(varCmnt.indexOf(' ') == -1) {
			for(j=0;j<varCmnt.length;j+=100) {
				varCmntCopy += varCmnt.substr(j+1,100) + "-";
			}
			document.all("comnt").innerText = varCmntCopy ;
		}
	}
}
else {
	var myArrComments = new Array(document.all("comnt").length);
	for(i=0;i<document.all("comnt").length;i++)	{
		varCmnt = document.all("comnt")(i).innerText;
		myArrComments[i] = varCmnt; 
		varCmntCopy = "";
		if(varCmnt.length >= 100) {
			if(varCmnt.indexOf(' ') == -1) {
				for(j=0;j<varCmnt.length;j+=100) {
					varCmntCopy += varCmnt.substr(j,100) + "-";
				}
				document.all("comnt")(i).innerText = varCmntCopy ;
			}
		}
	}
}
</script>
</body>
</html>
