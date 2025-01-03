<% 'CR20060902	full text implemented plus functionality for selected project %>
<!-- #include file="checksession.asp" -->
<!-- #include file="generalFunctions.asp" -->
<%
' we want that only alloutstanding issues link can only work without selecting a project  
' ADDED (and request("optProject") <> "ALL_PROJECT") CR20070101
'if session("projectid") = "" and request("search") <> "alloutstanding" and request("optProject") <> "ALL_PROJECT" then
'	response.Redirect("selectproject.asp")
'end if 
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
 dim order
 set objresult =server.CreateObject("adodb.recordset")
 if request("order") = "" then
	order = "desc"
 else
	order = ""
 end if
 search = request("search")
 'setting order by string for the query
 select case request("orderBy")
 case ""
 	orderby = "updatedate"
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
case "alloutstanding"
	if orderby = "projectname" then
		orderby = "p.projectname"
	else
		orderby = "i." & orderby
	end if
	strsql ="select i.*,p.projectname from " &  varTblNameIssues & " i," &  varTblNameUsers & " u," &  varTblNameProjects & " p where (status = 'opened' or status = 'Reopened' or status = 'In Progress') and i.projectid = u.projectid and i.projectid = p.projectid and u.email ='" & session("user")& "' order by " & orderby & " " & order
case "all_iss"
	strsql ="select i.*,p.projectname from " &  varTblNameIssues & " i," &  varTblNameUsers & " u," &  varTblNameProjects & " p where i.projectid = u.projectid and i.projectid = p.projectid and u.email ='" & session("user")& "' order by " & orderby & " " & order
End Select

objresult.Open strsql,con 
 
%>
<html>
<head>
<title><%=setTitle%></title>
<LINK 
href="include/style.css" type=text/css rel=stylesheet>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<body bgcolor="#CCCCCC" leftmargin="0" topmargin="0">
<form action="issuenav_desc.asp" method="post" name="frmIssues">
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
                  <td valign="top">
                    <font color="#FFFFFF">-</font><br>
                    <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
                      <tr> 
                        <td height="43" colspan="2"><div align="center" class="SectionHead">Issue 
                            Navigator - Descriptive</div>
						</td>
                      </tr>
					  </table>
                    <div align="center"><br>
                      <table width="95%" border="0" cellpadding="0" cellspacing="0" bgcolor="#CCCCCC">
                        <tr>
                          <td><table width="100%" border="0" align="center" cellpadding="1" cellspacing="1">
                              <tr bgcolor="#FFFFFF" class="head"> 
                                <!--<td>Key</td>-->
								<td><a href="#" onClick="frmIssues.orderBy.value='prjName';frmIssues.submit();">Project Name</a></td>
                                <td><a href="#" onClick="frmIssues.orderBy.value='summary';frmIssues.submit();">Summary</a></td>
                                <td><a href="#" onClick="frmIssues.orderBy.value='assignee';frmIssues.submit();">Assignee</a></td>
                                <td><a href="#" onClick="frmIssues.orderBy.value='reporter';frmIssues.submit();">Reporter</a></td>
                                <td><a href="#" onClick="frmIssues.orderBy.value='status';frmIssues.submit();">Status</a></td>
                                <td><a href="#" onClick="frmIssues.orderBy.value='created';frmIssues.submit();">Created</a></td>
                                <td><a href="#" onClick="frmIssues.orderBy.value='updated';frmIssues.submit();">Updated</a></td>
                              </tr>
							 <%
								 if objresult.BOF=false and objresult.EOF =false then  
								 while not objresult.EOF 
								 
								  %> 
                           
                              <tr bgcolor="#FFFFFF"> 
                              <!--<td><%'=objresult("projectId")%></td>-->
								<td><%=objresult("Projectname")%></td>
							  <td><a href="issuedetails.asp?issueid=<%=objresult("issueId")%>&prjid=<%=objresult("projectid")%>" ><%=objresult("summary")%></a></td>
                              <td><%=objresult("AssignTo")%></td>
                              <td><%=objresult("Reporter")%></td>
                              <td><%=objresult("status")%></td>
                              <td><%=DateAdd("h", -5, objresult("createDate")) %><%'for showing eastern standard time as it is 5 hrs less then GMT%></td>
                              <td><%=DateAdd("h", -5, objresult("updateDate"))%><%'for showing eastern standard time as it is 5 hrs less then GMT%></td>
                              </tr>
							<tr bgcolor="#CCCCCC"> 
								<td colspan="7"><%=formatData(objresult("description"))%></td></tr>
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
