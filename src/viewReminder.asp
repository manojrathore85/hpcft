<!--  ************ ********** Description : this file is basically for giving the result of the find issue *******   -->
<!-- #include file="checksession.asp" -->
<%
' we want that only alloutstanding issues link can only work without selecting a project  
if session("projectid") = "" and request("search") = "" then
	response.Redirect("selectproject.asp")
end if 
%>
<!-- #include file="connect.asp" -->
<!-- #include file="generalFunctions.asp" --> <!-- including the functions file -->

<%
'******* procedures and function start here  **************************
'******* function to check the permission ************************
function checkPermission
	dim rsperm
	set rsperm = server.CreateObject("adodb.recordset")
	rsperm.open "select * from " &  varTblNamePermissions & " where permissionid = '" & session("permission") & "'",con
	if not rsperm.eof then
		if rsperm("pread") = "T" then
			checkPermission = true
		else
			checkPermission = false
		end if
	else
		checkPermission = false
	end if
	rsperm.close
	set rsperm = nothing
end function
'******* procedures and function end here  **************************

'****** checking permission ******************************************
'if checkPermission = false then
'	server.Execute("permissiondenied.asp")
'	response.End()
'end if

dim rsRem  
dim strsql 
set rsRem =server.CreateObject("adodb.recordset")

if request("search") <> "" then
	'strsql ="select i.*,p.projectname from " &  varTblNameIssues & " i," &  varTblNameUsers & " u," &  varTblNameProjects & _
	 '		" p where (status = 'opened' or status = 'Reopened' or status = 'In Progress') and i.projectid = u.projectid and" & _
	'		" i.projectid = p.projectid and u.email ='" & session("user")& "' order by " & orderby & " desc"
	'strsql = "select rt.*,i.issueid,i.summary from " & varTblNameReminder_table & " rt," & varTblNameIssues & " i where rt.issueidcol = i.issueid and i.projectid = " & session("projectid") & " order by i.issueid"
		strsql = "select rt.*,p.projectname,i.issueid,i.summary from " & varTblNameReminder_table & " rt," & varTblNameIssues & " i," & _ 
			 varTblNameUsers & " u," &  varTblNameProjects & " p where rt.statecol = 'open' and i.projectid = u.projectid and" & _
			 " i.projectid = p.projectid and rt.issueidcol = i.issueid and u.email ='" & session("user")& "' order by " & get_orderby 'order by due_datetimecol desc"
			'"where rt.issueidcol = i.issueid and i.projectid = " & session("projectid") & " order by i.issueid"
else
	strsql = "select rt.*,i.issueid,i.summary from " & varTblNameReminder_table & " rt," & varTblNameIssues & " i where rt.issueidcol = i.issueid and i.projectid = " & session("projectid") & "  order by " & get_orderby
end if


rsRem.Open strsql,con 

dim order

if request("order") = "" then
	order = "desc"
else
	order = ""
end if

function get_orderby
	select case request("orderBy")
	case ""
		get_orderby = "due_datetimecol " & request("order")
	case "pname"
		get_orderby = "p.projectname " & request("order")
	case "isummary"
		get_orderby = "i.summary " & request("order")
	case "rtype"
		get_orderby = "rt.typecol " & request("order")
	case "duedate"
		get_orderby = "due_datetimecol " & request("order")
	case "duedatetz"
		get_orderby = "timezonecol " & request("order")
	case "createdate"
		get_orderby = "date_enteredcol " & request("order")
	case "pr"
		get_orderby = "person_responsiblecol " & request("order")
	case "af"
		get_orderby = "alertfrequencycol " & request("order")
	case "state"
		get_orderby = "statecol " & request("order")
	case "rsummary"
		get_orderby = "notescol " & request("order")
	end select
end function
 
%>
<html>
<head>
<title><%=setTitle%></title>
<LINK href="include/style.css" type=text/css rel=stylesheet>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<script>

function orderby(field)
{
	frm_reminder.orderBy.value=field;
	frm_reminder.submit();	
}
</script>

</head>
<body bgcolor="#CCCCCC" leftmargin="0" topmargin="0">

<form action="viewreminder.asp" method="post" name="frm_reminder">
<input type="hidden" name="search" value="<%=request("search")%>">
<input type="hidden" name="orderBy">
<input type="hidden" name="order" value="<%=order%>">
</form> 
<table width="100%" border="0" cellspacing="0" cellpadding="0" height="100%">
  <tr>
    <td valign="top" height="100%">
	  <table width="100%" height="100%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="#FFFFFF">
        <tr bgcolor="#006699"> 
          <td colspan="2"> <p><strong>&nbsp;<font color="#FFFFFF">Company Banner 
              and Information</font></strong></p>
            <p>&nbsp;</p></td>
        </tr>
        <tr align="left" valign="top"> 
          <td colspan="2"> 
            <div align="center">
<table width="100%" border="0" cellspacing="0" cellpadding="0" height="100%">
                <tr> 
                  <td valign="top"> <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
                      <tr> 
                        <td width="66%">&nbsp;</td>
                        <td width="34%" align="left"><!-- #include file="include/TopRightNavBar.asp" --></td>
                      </tr>
                      <tr>
                        <td>&nbsp;</td>
                        <td>&nbsp;</td>
                      </tr>
                      <tr><td colspan="2"> <!-- #include file="include/GeneralTopNavBar.asp" --></td>
                      </tr>
                    </table>
                    <font color="#FFFFFF">-</font><br>
                    <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
                      <tr> 
                        <td height="43" colspan="2"><div align="center" class="SectionHead">List Reminders 
                            : 
							<%if request("search") = "" then%>
							<%=getProjectName%>
							<%else%>
								All
							<%end if%>
							</div>
                          <br>
						</td>
                      </tr>
					  </table>
                    <div align="center"><br>
                      <table width="95%" border="0" cellpadding="0" cellspacing="0" bgcolor="#CCCCCC">
                        <tr>
                          <td><table width="100%" border="0" align="center" cellpadding="1" cellspacing="1">
                             <tr bgcolor="#FFFFFF" class="head">
							   <%
								if request("search") = "all" then
								%>
								<td><a href="javascript:orderby('pname')">Project Name</a></td>
								<%
								end if
								%>
							 <td><a href="javascript:orderby('isummary')">Issue Summary</a></td>
                              <td><a href="javascript:orderby('rtype')">Type</a></td>
                              <td><a href="javascript:orderby('duedate')">Due Date Time</a> </td>
                              <td><a href="javascript:orderby('duedatetz')">T-Zn</a></td>
                              <td><a href="javascript:orderby('createdate')">Date Entered</a></td>
                              <td>T-Zn</td>
                              <td><a href="javascript:orderby('pr')">Person Responsible</a> </td>
                              <td><a href="javascript:orderby('af')">Alert Freq</a> </td>
                              <td><a href="javascript:orderby('state')">State</a></td>
                              <td><a href="javascript:orderby('rsummary')">Summary</a></td>
                            </tr>
							 <%
								 if rsRem.BOF=false and rsRem.EOF =false then  
								 while not rsRem.EOF 
								 
								  %> 
                           
                              <tr bgcolor="#FFFFFF"> 
							  <%
								if request("search") = "all" then
								%>
								<td><%=rsRem("Projectname")%></td>
								<%
								end if
								%>
							  <td><a href="issuedetails.asp?issueid=<%=rsRem("issueId")%>" ><%=rsRem("summary")%></a></td>
                                <td><%=rsRem("TypeCol")%></td>
                              <td><%=rsRem("Due_DateTimeCol")%></td>
                              <td><%=rsRem("TimeZoneCol")%></td>
                              <td><%=rsRem("Date_enteredCol")%></td>
                              <td><%="GMT"%></td>
                              <td><%=rsRem("Person_ResponsibleCol")%></td>
                              <td><%=rsRem("AlertFrequencyCol")%></td>
                              <td><%=rsRem("StateCol")%></td>
                              <td><a href="rem_desc.asp?is=<%=rsRem("summary")%>&id=<%=rsRem("rem_idCol")%>" target="_blank"><%=rsRem("NotesCol")%></a></td>
	                          </tr>
                              <%
                                 rsRem.MoveNext 
                                 wend
                                 else 'if not found the we get another result
                                %> 
									<tr bgcolor="#FFFFFF"> 
										<Td colspan="11" align="center"> The Particular criteria not matched </td>
									</tr>
                              
                                  
                               <%
										rsRem.Close
										set rsRem = nothing
										con.close
										set con= nothing
										end if 
                                 %>
                            </table></td>
                        </tr>
                      </table>
                    </div></td>
                </tr>
              </table>
              
            </div></td>
        </tr>
        <tr> 
          <td width="45%">&nbsp;</td>
          <td width="55%" align="center" valign="middle">&nbsp;</td>
        </tr>
        <tr> 
          <td>&nbsp;</td>
          <td align="center" valign="middle">&nbsp;</td>
        </tr>
        <tr bgcolor="#003366"> 
          <td colspan="2" align="center"><font color="#FFFFFF">Email Links etc 
            </font></td>
        </tr>
      </table>
   </td>
  </tr>
</table>
</body>
</html>
