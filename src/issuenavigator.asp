<% 'CR20060902	full text implemented plus functionality for selected project %>
<!-- #include file="checksession.asp" -->
<%
' we want that only alloutstanding issues link can only work without selecting a project  
' ADDED (and request("optProject") <> "ALL_PROJECT") CR20070101
if session("projectid") = "" and request("search") <> "all_iss" and request("search") <> "alloutstanding" and request("optProject") <> "ALL_PROJECT" then
	response.Redirect("selectproject.asp")
end if 
%>
<!-- #include file="connect.asp" -->
<!-- #include file="generalFunctions.asp" --> <!-- including the functions file -->
<%

' ********************* CHECKING PERMISSION ***************************


' ********************* CHECKING PERMISSION ENDS***************************
%>


<%
 dim objresult  'for determine the result
 dim strsql 'for sql
 dim search
 dim orderby
 set objresult =server.CreateObject("adodb.recordset")
 search = request("search")
 'setting order by string for the query
 select case request("orderBy")
 case ""
 	orderby = "updatedate"
	if search = "addrecent" then
		orderby = "createdate"
	end if
 case "prjName"
 	orderby = "projectname"
 case "summary"
 	orderby = "summary"
 case "assignee"
 	orderby = "assignto"
 case "reporter"
 	orderby = "reporter"
 case "status"
 	orderby = "status"
 case "created"
 	orderby = "createdate"
 case "updated"
 	orderby = "updatedate"
 end select
' strsql ="select * from issues where issuetype='"& trim(request("lstIssTypes"))& "' and severity ='" & trim(request("lstSeverity")) & "' and summary='" & trim(request("txtIssSummary"))& "' and assignTo='" & trim(request("lstAssignedTo")) &"'"
' CASES ARRIVING FROM PROJECT DESCRIPTION PAGE I.E FILTER ISSUE
Select Case search
case "all"
	strsql ="select * from " &  varTblNameIssues & " where projectid =" & session("projectid") & " order by " & orderby & " desc" 
case "resolved"
	strsql ="select * from " &  varTblNameIssues & "  where status = 'resolved' and projectid =" & session("projectid") & " order by " & orderby & " desc"
case "outstanding"
	strsql ="select * from " &  varTblNameIssues & "  where (status = 'opened' or status = 'Reopened' or status = 'In Progress') and projectid =" & session("projectid") & " order by " & orderby & " desc"
case "alloutstanding"
	if orderby = "projectname" then
		orderby = "p.projectname"
	else
		orderby = "i." & orderby
	end if
	strsql ="select i.*,p.projectname from " &  varTblNameIssues & " i," &  varTblNameUsers & " u," &  varTblNameProjects & " p where (status = 'opened' or status = 'Reopened' or status = 'In Progress') and i.projectid = u.projectid and i.projectid = p.projectid and u.email ='" & session("user")& "' order by " & orderby & " desc"
case "all_iss"
	if orderby = "projectname" then
		orderby = "p.projectname"
	else
		orderby = "i." & orderby
	end if
	strsql ="select i.*,p.projectname from " &  varTblNameIssues & " i," &  varTblNameUsers & " u," &  varTblNameProjects & " p where i.projectid = u.projectid and i.projectid = p.projectid and u.email ='" & session("user")& "' order by " & orderby & " desc"	
case "addrecent"
	strsql ="select * from " &  varTblNameIssues & " where projectid =" & session("projectid") & " order by " & orderby & " desc"
case "asgndtome"
	strsql ="select * from " &  varTblNameIssues & "  where assignto = '" & session("user") & "' and projectid =" & session("projectid") & " order by " & orderby & " desc"
case "updaterecent"
	strsql ="select * from " &  varTblNameIssues & " where projectid =" & session("projectid") & " order by " & orderby & " desc"	
case "reportbyme"
	strsql ="select * from " &  varTblNameIssues & "  where reporter = '" & session("user") & "' and projectid =" & session("projectid") & " order by " & orderby & " desc"
case "important"
	strsql ="select * from " &  varTblNameIssues & "  where (severity = 'major' or severity = 'blocker' or severity = 'critical') and projectid =" & session("projectid") & " order by " & orderby & " desc"
End Select

' CASES ARRIVING FROM FIND ISSUE PAGE
if strsql = "" and request("lstIssTypes") <> "" then  
	if request("chkFullText") = "True" then  ' if full text search requested
	
		strsql = "select i.*,p.projectname from " &  varTblNameIssues & " i," &  varTblNameUsers & " u," &  varTblNameProjects & " p, " & varTblNameIssuechangescomments & " ic " & _
		"where (ic.newvalue like '%" & request("txtIssSummary") & "%' or ic.lastvalue like '%" & request("txtIssSummary") & "%' or ic.comments like '%" & request("txtIssSummary") & "%' " & _
		" or i.summary like '%" & request("txtIssSummary") & "%' or i.description like '%" & request("txtIssSummary") & "%')" 	'CR20060902

		if request("optProject") <> "ALL_PROJECT" then	'CR20060902
			strsql = strsql  & " and i.projectid =" & session("projectid")
		end if
		
		'CR20060902
		strsql = strsql  & " and  u.email = '" & session("user") & "' and " & _	
		"i.projectid = p.projectid and i.projectid = u.projectid and i.issueid = ic.issueid group by i.issueid "	'CR20060902
		
		
		'strsql = "select i.*,p.projectname from " &  varTblNameIssues & " i," &  varTblNameProjects & " p," & _
		'"(SELECT distinct( issueid) FROM " &  varTblNameIssueChangesComments & _
		'" where MATCH (newvalue,lastvalue,comments) AGAINST ('" & request("txtIssSummary") & "')" & _
		'" union " & _
		'"SELECT distinct( issueid) FROM " &  varTblNameIssues & _ 
		'" where MATCH (summary,description) AGAINST ('" & request("txtIssSummary") & "')) a " & _
		'"where " & _
		'"i.projectid = p.projectid and i.issueid = a.issueid "
		
	else  ' if not full text search requested
		if request("lstIssTypes") <> "all" then
			strsql = "select i.*,p.projectname from " &  varTblNameIssues & " i," &  varTblNameUsers & " u," &  varTblNameProjects & " p where i.projectid = u.projectid and i.projectid = p.projectid and u.email ='" & session("user")& "' and i.issuetype = '" & request("lstIssTypes") & "'"   '  & "' and projectid =" & session("projectid") 
		else
			strsql = "select i.*,p.projectname from " &  varTblNameIssues & " i," &  varTblNameUsers & " u," &  varTblNameProjects & " p where i.projectid = u.projectid and i.projectid = p.projectid and u.email ='" & session("user")& "'"
		end if
		if request("optProject") <> "ALL_PROJECT" then		'CR20060902
			strsql = strsql  & " and i.projectid =" & session("projectid")
		end if
		if request("lstSeverity") <> "all" then
			strsql = strsql & " and severity = '" & request("lstSeverity") & "'"
		end if
		if request("lstAssignedTo") <> "all" then
			strsql = strsql & " and assignto = '" & request("lstAssignedTo") & "'"
		end if
		if request("lstReportedBy") <> "all" then
			strsql = strsql & " and reporter = '" & request("lstReportedBy") & "'"
		end if
		if request("txtIssSummary") <> "" then
			strsql = strsql & " and summary like '%" & request("txtIssSummary") & "%'"
		end if
	end if
		if orderby = "projectname" then
			strsql = strsql & " order by p." & orderby & " desc"
		else
			strsql = strsql & " order by i." & orderby & " desc"
		end if
end if

' CASE WHEN ALL ISSUES ARE REQUESTED
if strsql = "" then
	 strsql = "select * from " &  varTblNameIssues & " where projectid =" & session("projectid") & " order by " & orderby & " desc" 
end if

objresult.Open strsql,con 
 
%>
<html>
<head>
<title><%=setTitle%></title>
<LINK 
href="include/style.css" type=text/css rel=stylesheet>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<script>
function view_cons_rpt() {
	frmIssues.target = "_blank";
	frmIssues.action = "cons_search_report.asp";
	frmIssues.submit();
}
function issue_description() {
	frmIssues.target = "_blank";
	frmIssues.action = "cons_search_report.asp";
	frmIssues.submit();
}
function PopUp(theURL,winName,features) { //v2.0
  window.open(theURL,winName,features);
}
</script>
</head>
<body bgcolor="#CCCCCC" leftmargin="0" topmargin="0">
<form action="issuenavigator.asp" method="post" name="frmIssues">
<input type="hidden" name="search" value="<%=request("search")%>">
<input type="hidden" name="optProject" value="<%=request("optProject")%>">
<input type="hidden" name="lstIssTypes" value="<%=request("lstIssTypes")%>">
<input type="hidden" name="lstSeverity" value="<%=request("lstSeverity")%>">
<input type="hidden" name="txtIssSummary" value="<%=request("txtIssSummary")%>">
<input type="hidden" name="chkFullText" value="<%=request("chkFullText")%>">
<input type="hidden" name="lstAssignedTo" value="<%=request("lstAssignedTo")%>">
<input type="hidden" name="lstReportedBy" value="<%=request("lstReportedBy")%>">
<input type="hidden" name="cmdFind" value="<%=request("cmdFind")%>">
<input type="hidden" name="orderBy">


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
                        <td height="43" colspan="2"><div align="center" class="SectionHead">Issue 
                            Navigator : <%
							if request("search") = "alloutstanding" or request("search") = "all_iss" or session("projectid") = "" then  ' added ( or session("projectid") = "" ) CR20070101
								response.write("All Outstanding")
							else
								response.Write(getProjectName)
							end if
							%></div>
                          <p align="center" class="redbold">Displaying --- matching 
                            issue</p>
                          <br>
						</td>
                      </tr>
					  </table>
                    <div align="center"><br>
                      <table width="95%" border="0" cellpadding="0" cellspacing="0" bgcolor="#CCCCCC">
					  <%if request("chkFullText") = "True" then%>
					  <tr bgcolor="#FFFFFF"><td><div align="left"><a href="#" onClick="view_cons_rpt();">View consolidated search report</a></div><br></td></tr>
					  <% end if %>
 					  <%if request("search")  = "alloutstanding" or request("search")  = "all_iss" then%>
					  <tr bgcolor="#FFFFFF"><td><div align="left"><a href="#" onClick="PopUp('issuenav_desc.asp?search=<%=request("search")%>','IssueDesc','status=yes,scrollbars=yes,width=800,height=550')"><b>Display All Issue With Description</b></a></div><br></td></tr>
					  <% end if %>
                        <tr>
                          <td><table width="100%" border="0" align="center" cellpadding="1" cellspacing="1">
                              <tr bgcolor="#FFFFFF" class="head"> 
                                <!--<td>Key</td>-->
								<%
								if request("search") = "alloutstanding" or request("search") = "all_iss" or request("cmdFind") <> "" then
								%>
								<td><a href="#" onClick="frmIssues.orderBy.value='prjName';frmIssues.submit();">Project Name</a></td>
								<%
								end if
								%>
                                <td><a href="#" onClick="frmIssues.orderBy.value='summary';frmIssues.submit();">Summary</a></td>
                                <td><a href="#" onClick="frmIssues.orderBy.value='assignee';frmIssues.submit();">Assignee</a></td>
                                <td><a href="#" onClick="frmIssues.orderBy.value='reporter';frmIssues.submit();">Reporter</a></td>
                                <td><a href="#" onClick="frmIssues.orderBy.value='status';frmIssues.submit();">Status</a></td>
                                <td><a href="#" onClick="frmIssues.orderBy.value='created';frmIssues.submit();">Created ( EST )</a></td>
                                <td><a href="#" onClick="frmIssues.orderBy.value='updated';frmIssues.submit();">Updated ( EST )</a></td>
                              </tr>
							 <%
								 if objresult.BOF=false and objresult.EOF =false then  
								 while not objresult.EOF 
								 
								  %> 
                           
                              <tr bgcolor="#FFFFFF"> 
                              <!--<td><%'=objresult("projectId")%></td>-->
							  <%
								if request("search") = "alloutstanding" or request("search") = "all_iss" or request("cmdFind") <> "" then
								%>
								<td><%=objresult("Projectname")%></td>
								<%
								end if
								%>
							  <td><a href="issuedetails.asp?issueid=<%=objresult("issueId")%>&prjid=<%=objresult("projectid")%>" ><%=objresult("summary")%></a></td>
                              <td><%=objresult("AssignTo")%></td>
                              <td><%=objresult("Reporter")%></td>
                              <td><%=objresult("status")%></td>
                              <td><%=DateAdd("h", +1, objresult("createDate")) %><%'for showing eastern standard time as it is 1 hrs + then American/Chicago time ( GMT - 6 HRS ) 17012011atul%></td>
                              <td><%=DateAdd("h", +1, objresult("updateDate"))%><%'for showing eastern standard time as it is 1 hrs + then American/Chicago time ( GMT - 6 HRS ) 17012011atul%></td>
                              </tr>
                              <%
                                 objresult.MoveNext 
                                 wend
                                 else 'if not found the we get another result
                                %> 
									<tr bgcolor="#FFFFFF"> 
										<Td colspan="7" align="center"> The Particular criteria not matched </td>
									</tr>
                              
                                  
                               <%
										objresult.Close
										set objreult = nothing
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
