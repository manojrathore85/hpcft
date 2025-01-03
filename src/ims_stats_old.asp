<!-- ******************** Description : This the file which is used to show the usage history of the selected users -->
<!-- ****************************** Date :--  29 march 2k5 ***************************************** -->
<!-- #include file="checksession.asp" -->
<!-- #include file="connect.asp" --> <!-- including the connection file -->

<%
' ********************* Global Variables ***************************
dim rs
set rs = server.CreateObject("adodb.recordset")

' ******************************************************************

' ********************* CHECKING PERMISSION ***************************
sub check_permission
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
end sub
' ********************* CHECKING PERMISSION ENDS***************************

sub fill_users
	dim rs
	dim sql
	set rs = server.CreateObject("adodb.recordset")
	sql = "select email,firstname,lastname from " & varTblNameUserProfile 
	rs.open sql , con
	if not rs.eof then
	   while not rs.eof
		%>
		<option value="<%=rs("email")%>"><%=rs("firstname") & " " & rs("lastname")%></option>
		<%
		rs.movenext
		wend
	end if
end sub

function get_create_count(between_expr)
	dim sql
	'if end_date = "" then
	'	sql = "select count(*) from issues where date_format(createdate,'%d/%c/%Y') = '" & start_date & "' and reporter = '" & request("lstUsers") & "'"
	'else
		sql = "select count(*) from issues where createdate " & between_expr & " and reporter = '" & request("lstUsers") & "'"
	'end if	
	rs.open sql,con
	if not rs.eof then
		get_create_count = rs(0)
	else
		get_create_count = 0
	end if
	rs.close
end function
function get_update_count(between_expr)
	dim sql
	'if end_date = "" then
	'	sql = "select count(*) from issues where date_format(createdate,'%d/%c/%Y') = '" & start_date & "' and reporter = '" & request("lstUsers") & "'"
	'else
		sql = "select count(*) from issuechangescomments where updatedate " & between_expr & " and changesby = '" & request("lstUsers") & "' and field <> 'deleted'"
	'end if	
	rs.open sql,con
	if not rs.eof then
		get_update_count = rs(0)
	else
		get_update_count = 0
	end if
	rs.close
end function
function get_delete_count(between_expr)
		dim sql
	'if end_date = "" then
	'	sql = "select count(*) from issues where date_format(createdate,'%d/%c/%Y') = '" & start_date & "' and reporter = '" & request("lstUsers") & "'"
	'else
		sql = "select count(*) from issuechangescomments where updatedate " & between_expr & " and changesby = '" & request("lstUsers") & "' and field = 'deleted'"
	'end if	
	rs.open sql,con
	if not rs.eof then
		get_delete_count = rs(0)
	else
		get_delete_count = 0
	end if
	rs.close
end function

sub create_report
	dim count_of
	select case request("opt_report_type")
	case "daily"
		count_of = datediff("d",start_date,end_date)
		'get_data_for_report(1)
		get_data_for_report count_of,"d","DAY"
	case "weekly"
		count_of = datediff("ww",start_date,end_date,vbMonday)
		'get_data_for_report(7)
		get_data_for_report count_of,"ww","WEEK"
	case "monthly"
		count_of = datediff("m",start_date,end_date)
		get_data_for_report count_of,"m","MONTH"
		'get_data_for_report(30)
	case "yearly"
		count_of = datediff("yyyy",start_date,end_date)
		get_data_for_report count_of,"yyyy","YEAR"
		'get_data_for_report(365)
	end select
end sub

sub get_data_for_report(count_of,vb_add_interval_type,mysql_add_interval_type)
	dim i
	dim date_from 
	dim date_to
	dim between_expr
'	response.Write(count_of & " " & vb_add_interval_type & " " &  mysql_add_interval_type & " hello <br>")
	count_of = count_of - 1
	for i = 0 to count_of
		date_from = DateAdd(vb_add_interval_type,i,start_date)
		date_from = year(date_from) & "/" & month(date_from) & "/" & day(date_from)
		date_to = DateAdd(vb_add_interval_type,i + 1,start_date)
		date_to = year(date_to) & "/" & month(date_to) & "/" & day(date_to)
		'response.Write(date_from & " - " & date_to & "<br>")
		if vb_add_interval_type = "d" then
'			date_from = DateAdd(vb_add_interval_type,i,start_date)
'			date_from = year(date_from) & "-" & month(date_from) & "-" & day(date_from)
			'duration = DateAdd(vb_add_interval_type,i,start_date)
			'duration = cdate(start_date) + i
			duration = date_from
		else
'			date_from = DateAdd(vb_add_interval_type,i,start_date)
'			date_from = year(date_from) & "-" & month(date_from) & "-" & day(date_from)
'			date_to = DateAdd(vb_add_interval_type,i + 1,start_date)
'			date_to = year(date_from) & "-" & month(date_from) & "-" & day(date_from)
			duration = date_from & " - " & date_to
					
			'duration = DateAdd(vb_add_interval_type,i,start_date) & " - " & DateAdd(vb_add_interval_type,i + 1,start_date)
			'duration = cdate(start_date) + i & " - " & cdate(start_date) + i + constant
		end if
		between_expr = "between '" & date_from & "' and '" & date_to & "'" 

		create_count =	get_create_count(between_expr)
		update_count =	get_update_count(between_expr)
		delete_count =	get_delete_count(between_expr)
		
		if create_count = 0 and update_count = 0 and delete_count = 0 then
			' don't print the row
		else
		%>
		  <tr bgcolor="#FFFFFF"> 
          <td>&nbsp;<%=duration%></td>
          <td>&nbsp;<%=create_count%></td>
          <td>&nbsp;<%=update_count%></td>
		  <td>&nbsp;<%=delete_count%></td>
        </tr>
		<%
		end if
	next
end sub
call check_permission



%>
<script>
</script>

<table width="95%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td class="SectionHead">Administration<br>
      &nbsp; </td>
  </tr>
  <tr> 
    <td>View the IMS Statistics of selected user<br>
    </td>
  </tr>
</table>
<BR>
<table width="95%" border="0" cellpadding="0" cellspacing="0" bgcolor="#CCCCCC">
<tr>
  <td><table width="100%" border="0" cellpadding="1" cellspacing="1"><form name="frmStats" method="post" action="">
        <tr bgcolor="#FFFFFF" class=head> 
          <td width="17%" id=head>
		  
		  <div align="center">Select the user :- 
                <select name="lstUsers">
				<% call fill_users %>
                </select>
            </div>			</td>
			<td width="18%" id=head><input type="radio" name="opt_report_type" value="daily" />Daily<br />
			  <input type="radio" name="opt_report_type" value="weekly" />Weekly<br />
			    <input type="radio" name="opt_report_type" value="monthly" />
			    Monthly&nbsp;<br />
		    <input type="radio" name="opt_report_type" value="yearly" />Yearly</td>
			<td width="49%" id=head>Start Date:&nbsp;
			  <select name="lst_start_year" class="listbox">
                <option value="2005">2005</option>
				<option value="2006">2006</option>
                <option value="2007">2007</option>
                <option value="2008">2008</option>
                <option value="2009">2009</option>
                <option value="2010">2010</option>
                <option value="2011">2011</option>
                <option value="2012">2012</option>
                <option value="2013">2013</option>
                <option value="2014">2014</option>
                <option value="2015">2015</option>
              </select>
              <select name="lst_start_month" class="listbox">
                <option value="0" selected="selected">January</option>
                <option value="1">February</option>
                <option value="2">March</option>
                <option value="3">April</option>
                <option value="4">May</option>
                <option value="5">June</option>
                <option value="6">July</option>
                <option value="7">August</option>
                <option value="8">September</option>
                <option value="9">October</option>
                <option value="10">November</option>
                <option value="11">December</option>
              </select>
              <label for="day"></label>
              <select name="lst_start_day" class="listbox">
                <option value="1" selected="selected">1</option>
                <option value="2">2</option>
                <option value="3">3</option>
                <option value="4">4</option>
                <option value="5">5</option>
                <option value="6">6</option>
                <option value="7">7</option>
                <option value="8">8</option>
                <option value="9">9</option>
                <option value="10">10</option>
                <option value="11">11</option>
                <option value="12">12</option>
                <option value="13">13</option>
                <option value="14">14</option>
                <option value="15">15</option>
                <option value="16">16</option>
                <option value="17">17</option>
                <option value="18">18</option>
                <option value="19">19</option>
                <option value="20">20</option>
                <option value="21">21</option>
                <option value="22">22</option>
                <option value="23">23</option>
                <option value="24">24</option>
                <option value="25">25</option>
                <option value="26">26</option>
                <option value="27">27</option>
                <option value="28">28</option>
                <option value="29">29</option>
                <option value="30">30</option>
                <option value="31">31</option>
              </select>
            <br />
            &nbsp;End Date:&nbsp;
            <select name="lst_end_year" class="listbox">
			  <option value="2005">2005</option>
              <option value="2006">2006</option>
              <option value="2007">2007</option>
              <option value="2008">2008</option>
              <option value="2009">2009</option>
              <option value="2010">2010</option>
              <option value="2011">2011</option>
              <option value="2012">2012</option>
              <option value="2013">2013</option>
              <option value="2014">2014</option>
              <option value="2015">2015</option>
            </select>
            <select name="lst_end_month" class="listbox">
              <option value="0" selected="selected">January</option>
              <option value="1">February</option>
              <option value="2">March</option>
              <option value="3">April</option>
              <option value="4">May</option>
              <option value="5">June</option>
              <option value="6">July</option>
              <option value="7">August</option>
              <option value="8">September</option>
              <option value="9">October</option>
              <option value="10">November</option>
              <option value="11">December</option>
            </select>
            <label for="label"></label>
            <select id="label" name="lst_end_day" class="listbox">
              <option value="1" selected="selected">1</option>
              <option value="2">2</option>
              <option value="3">3</option>
              <option value="4">4</option>
              <option value="5">5</option>
              <option value="6">6</option>
              <option value="7">7</option>
              <option value="8">8</option>
              <option value="9">9</option>
              <option value="10">10</option>
              <option value="11">11</option>
              <option value="12">12</option>
              <option value="13">13</option>
              <option value="14">14</option>
              <option value="15">15</option>
              <option value="16">16</option>
              <option value="17">17</option>
              <option value="18">18</option>
              <option value="19">19</option>
              <option value="20">20</option>
              <option value="21">21</option>
              <option value="22">22</option>
              <option value="23">23</option>
              <option value="24">24</option>
              <option value="25">25</option>
              <option value="26">26</option>
              <option value="27">27</option>
              <option value="28">28</option>
              <option value="29">29</option>
              <option value="30">30</option>
              <option value="31">31</option>
            </select></td>
			<td width="16%" id=head><input name="submit" type="submit" value="Generate Report" /></td>
        </tr> </form>
        <tr bgcolor="#FFFFFF"> 
          <td colspan="4">Note : Row having create,update, and delete count equal to 0 will not be shown</td>
        </tr>
        <tr bgcolor="#FFFFFF" class=head> 
          <td id=head>&nbsp;Duration</td>
          <td>&nbsp;Create Count </td>
          <td>&nbsp;Update Count </td>
		  <td>&nbsp;Delete Count </td>
        </tr>
		
		<%
		if request("submit") = "Generate Report" then
			dim start_date
			dim end_date
			start_date = request("lst_start_year")  & "/" &  request("lst_start_month") + 1 & "/" & request("lst_start_day")
			end_date = request("lst_end_year")  & "/" &  request("lst_end_month") + 1 & "/" & request("lst_end_day")
			call create_report
		end if
		%>
<!--        <tr bgcolor="#FFFFFF"> 
          <td colspan="4" align="center">NO DETAILS CURRENTLY</td>
        </tr>
-->      </table></td>
</tr>
</table>
<script>
<%
if request("lstUsers") <> "" then
%>
   for (i=0;i<=frmStats.lstUsers.options.length-1;i++)
     if (frmStats.lstUsers.options[i].value=='<%=request("lstUsers")%>')
			frmStats.lstUsers.selectedIndex=i;


frmStats.lst_start_year.value = '<%=request("lst_start_year")%>';
frmStats.lst_start_month.value = <%=request("lst_start_month")%>; 
frmStats.lst_start_day.value = <%=request("lst_start_day")%>;
frmStats.lst_end_year.value = '<%=request("lst_end_year")%>';
frmStats.lst_end_month.value = <%=request("lst_end_month")%>; 
frmStats.lst_end_day.value = <%=request("lst_end_day")%>;

for(i=0;i<frmStats.opt_report_type.length;i++)
	if(frmStats.opt_report_type(i).value == '<%=request("opt_report_type")%>')
		frmStats.opt_report_type(i).checked = true;
<%
end if
%>

</script>
<%
  set rs = nothing
  con.close
  set con = nothing
 %> 
