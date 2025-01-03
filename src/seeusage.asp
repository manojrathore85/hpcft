<!-- ******************** Description : This the file which is used to show the usage history of the selected users -->
<!-- ****************************** Date :--  29 march 2k5 ***************************************** -->
<!-- #include file="checksession.asp" -->
<!-- #include file="connect.asp" --> <!-- including the connection file -->

<%

' ********************* CHECKING PERMISSION ***************************

dim rsperm
dim rs
dim str
dim sql
set rs = server.CreateObject("adodb.recordset")
dim prj_iss
dim my_action
dim my_subaction
prj_iss = array(0,0,0,0)

dim findstr1
dim findstr2
findstr1 = "editissue.asp"
findstr2 = "issuedetails.asp"
dim REO
Set REO = New RegExp

With REO
	.Pattern = "Projectid=([0-9]{1,5}),issueid=([0-9]{1,5})([&A-Za-z0-9=]+)"
	.IgnoreCase = True
	.Global = True
End With


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
' REMOVING THE RECORDS ON REQUEST 
'******************************************************
if request("mainlink") = "cleanusage" then
	con.execute "delete from " &  varTblNameIms_Usage & " where u_datetime <= date_sub(current_timestamp,INTERVAL " & varSiteSpecNUM_USAGE_DAYS & " DAY)"
end if

'**************************************************
   dim i ' as incremental variable
   dim objview,rsUsage
   set objview =server.CreateObject("adodb.recordset") ' creating the recordset 
   set rsUsage =server.CreateObject("adodb.recordset") ' creating the recordset 
   rsUsage.open "select distinct iu.email,up.firstname,up.lastname from " &  varTblNameIms_Usage & " iu," &  varTblNameUserProfile & " up where iu.email = up.email",con
   if request("lstUsers") <> "" then
   		objview.open "select * from " &  varTblNameIms_Usage & " where email = '" & request("lstUsers") & "' order by IdKey desc",con
   else
   		objview.open "select * from " &  varTblNameIms_Usage & " where 1=2",con
   '		if not rsUsage.eof then		
	'		objview.open "select * from " &  varTblNameIms_Usage & " where email = '" & rsUsage("email") & "'",con
	'	else
	'		objview.open "select * from " &  varTblNameIms_Usage & " where email = '" & session("user") & "'",con
	'	end if
   end if
   
function get_details(action,subaction)
	subaction = subaction & "&"
	if instr(action,findstr1) <> 0 or instr(action,findstr2) <> 0 then
		if REO.test(subaction) = true then
			str = REO.replace(subaction, "$1,$2")
			'response.Write("str=" & str )
			prj_iss = split(str,",")
			'response.Write(prj_iss(0) & " , " & prj_iss(1))
			get_details = get_summary()
		else
			get_details = ""
		end if
	end if
end function
 
function get_summary()
	if rs.state = 1 then rs.close
	sql = "select summary from " & varTblNameIssues & " where projectid=" & prj_iss(0) & " and issueid=" & prj_iss(1)
	rs.open sql,con
	if not rs.eof then
		get_summary = rs(0)
	else
		get_summary = ""
	end if
	rs.close
end function
%>
<script>
function checkDelete()
{
	if (confirm('Do u really want to delete this project.\nThis will delete all the data related to this project from the database.'))
		return true;
	else
		return false;
}
</script>

<table width="95%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td class="SectionHead">Administration<br>
      &nbsp; </td>
  </tr>
  <tr> 
    <td>View the usage history of selected user<br>
    </td>
  </tr>
</table>
<BR>
<table width="95%" border="0" cellpadding="0" cellspacing="0" bgcolor="#CCCCCC">
<tr>
  <td><table width="100%" border="0" cellpadding="1" cellspacing="1"><form name="frmUsage" method="post" action="">
        <tr bgcolor="#FFFFFF" class=head> 
          <td colspan="4" id=head><div align="center">Select the user :- 
                <select name="lstUsers" onChange="javascript:if(this.value != 'none') frmUsage.submit();">
				<option value="none">Select User</option>
				<%
			   while not rsUsage.eof
				%>
				<option value="<%=rsUsage("email")%>"><%=rsUsage("firstname") & " " & rsUsage("lastname")%></option>
				<%
				rsUsage.movenext
				wend
				%>
                </select>
            </div></td>
        </tr> </form>
        <tr bgcolor="#FFFFFF"> 
          <td id=head>&nbsp;</td>
          <td id=head>&nbsp;</td>
          <td>&nbsp;</td>
		  <td>&nbsp;</td>
        </tr>
        <tr bgcolor="#FFFFFF" class=head> 
          <td id=head>&nbsp;DateTime</td>
          <td>&nbsp;Action</td>
          <td>&nbsp;SubAction</td>
		  <td>&nbsp;Issue Summary</td>
        </tr>
        <% 
    if  objview.bof=false and objview.eof =false then ' checking whether the record is exist or not
	   while not objview.eof ' checkout all the user in database
	   		my_action = objview("action")
			my_subaction = objview("subaction") 
%>
        <tr bgcolor="#FFFFFF"> 
          <td>&nbsp;<%=objview("u_datetime")%></td>
          <td>&nbsp;<%=my_action%></td>
          <td>&nbsp;<%=my_subaction%></td>
		  <td>&nbsp;<% hello = get_details(my_action,my_subaction)
		  response.Write("<a href='issuedetails.asp?issueid=" & prj_iss(1) & "&prjid=" & prj_iss(0) & "' target='_blank'>" & hello & "</a>")%></td>
        </tr>
        <%
	    objview.movenext 
		wend
	else
	%>
        <tr bgcolor="#FFFFFF"> 
          <td colspan="4" align="center">---------------------</td>
        </tr>
        <%

    end if		
   %>
      </table></td>
</tr>
</table>
<script>
<%
if request("lstUsers") <> "" then
%>
   for (i=0;i<=frmUsage.lstUsers.options.length-1;i++)
     if (frmUsage.lstUsers.options[i].value=='<%=request("lstUsers")%>')
			frmUsage.lstUsers.selectedIndex=i;
<%
end if
%>
</script>
<%
 ' nullifying all the objects
  objview.close
  set objview= nothing
  con.close
  set con = nothing
 %> 
