<% 
' ####################################################################################

' ADDED TO IMS ON 22/02/2007

' ####################################################################################

%>
<!-- #include file="checksession.asp" -->
<!-- #include file="connect.asp" -->
<!-- #include file="generalFunctions.asp" --> <!-- including the functions file -->
<%

sub mydiary()
	dim rs
	dim sql
	set rs = server.CreateObject("adodb.recordset")
	sql ="select md.*,p.projectname,i.summary,i.updatedate from " & varTblNameIms_MyDiary_table & " md, " &  varTblNameProjects & " p, " &  varTblNameIssues & " i where  md.projectid = p.projectid and md.projectid = i.projectid and md.issueid = i.issueid and md.email ='" & session("user")& "' order by " & get_orderby 'i.updatedate desc"'" & orderby & " desc"
	rs.open sql ,con
	while not rs.eof
%>
	  <tr bgcolor="#FFFFFF"> 
	  <td><%=rs("projectname")%></td>
	  <td><a href="issuedetails.asp?issueid=<%=rs("issueId")%>&prjid=<%=rs("projectid")%>" ><%=rs("summary")%></a></td>
	  <td><%=DateAdd("h", -5, rs("date_of_creation")) %> <%'for showing eastern standard time as it is 5 hrs less then GMT%> </td>
	  <td><%=DateAdd("h", -5, rs("updatedate")) %> <%'for showing eastern standard time as it is 5 hrs less then GMT%> </td>
	  <td><%=rs("comments")%></td>
	  <td><a href="addtodiary.asp?mode=remove&issueid=<%=rs("issueId")%>&prjid=<%=rs("projectid")%>" target="_blank">Remove</a></td>
	  </tr>
<%
		rs.movenext
	wend
	rs.close
	con.close
	set rs = nothing
	set con = nothing
end sub

dim order

if request("order") = "" then
	order = "desc"
else
	order = ""
end if

function get_orderby
	select case request("orderBy")
	case ""
		get_orderby = "i.updatedate " & request("order")
	case "pname"
		get_orderby = "p.projectname " & request("order")
	case "summary"
		get_orderby = "i.summary " & request("order")
	case "createdate"
		get_orderby = "md.date_of_creation " & request("order")
	case "updatedate"
		get_orderby = "i.updatedate " & request("order")
	case "comments"
		get_orderby = "md.comments " & request("order")
	end select
end function
%>

<html>
<head>
<title><%=setTitle%></title>
<LINK 
href="include/style.css" type=text/css rel=stylesheet>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<script>

function orderby(field)
{
	frm_diary.orderBy.value=field;
	frm_diary.submit();	
}
</script>

</head>
<body bgcolor="#CCCCCC" leftmargin="0" topmargin="0">
<form action="mydiary.asp" method="post" name="frm_diary">
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
                        <td height="43" colspan="2"><div align="center" class="redbold">My Diary</div>
                          <br>
						</td>
                      </tr>
					  </table>
                    <div align="center"><br>
                      <table width="95%" border="0" cellpadding="0" cellspacing="0" bgcolor="#CCCCCC">
                        <tr>
                          <td><table width="100%" border="0" align="center" cellpadding="1" cellspacing="1">
                              <tr bgcolor="#FFFFFF" class="head"> 
                                <!--<td>Key</td>-->
								<td><a href="javascript:orderby('pname')">Project Name</a></td>
                                <td><a href="javascript:orderby('summary')">Summary</a></td>
                                <td><a href="javascript:orderby('createdate')">Date Created</a></td>
								<td><a href="javascript:orderby('updatedate')">Issue Update Date</a></td>
                                <td><a href="javascript:orderby('comments')">Comments</a></td>
								<td>&nbsp;</td>
                              </tr>
							 <%
								 call mydiary
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
