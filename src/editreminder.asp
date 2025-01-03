<%
' Program to add/edit reminders
'Atul	200600413	
%>
<!-- #include file ="CheckSession.asp" --> 
<!-- #include file ="Connect.asp" -->

<%
'******* procedures and function start here  **************************
'******* function to check the permission ************************
function checkPermission
	dim rsperm
	set rsperm = server.CreateObject("adodb.recordset")
	rsperm.open "select perm.* from " &  varTblNamePermissions & " perm, " &  varTblNameUsers & " u, " &  varTblNameissues & _
	 " i where i.projectid = u.projectid and u.permissionid = perm.permissionid and i.issueid = " &  request("issueid") & _
	 " and u.email ='" &  session("user") & "'",con
	if not rsperm.eof then
		if rsperm("pwrite") = "T" then
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

' *********  procedure to set the variables while in edit mode  *****************************
sub setVariables
	dim rs
	set rs = server.CreateObject("adodb.recordset")
	rs.open "select * from " & varTblNameReminder_table & " where rem_idCol =" & request("remid"),con
	if not rs.eof then
		remType = rs("TypeCol")
		remDueDate = rs("due_datetimeCol")
		remTimeZone = rs("timezoneCol")
		remPR  = rs("person_ResponsibleCol")
		remAlertFreq= rs("alertFrequencyCol")
		remState = rs("stateCol")
		remNotes = rs("notesCol")
		remDesc=rs("DescCol") '20061113
		remDaysofWeek=rs("DaysOfWeekCol") '20061113
		remMonthsofYear=rs("MonthsOfYearCol") '20061113
	end if
	rs.close
	set rs = nothing
end sub
'******* procedures and function end here  **************************

'****** checking permission ******************************************
if checkPermission = false then
	server.Execute("permissiondenied.asp")
	response.End()
end if

' ***** edit reminder code **************

if request("remid") <> "" then
	dim remType
	dim remDueDate
	dim remTimeZone
	dim remPR ' person responsible
	dim remAlertFreq
	dim remState
	dim remNotes
	dim remDesc		'20061113
	dim remDaysofWeek '20061113
	dim remMonthsofYear '20061113

	call setVariables
end if

' ************  code to show the existing reminders  ************************************

dim rsRem,rsUsers,summary
set rsRem = server.CreateObject("adodb.recordset")
set rsUsers = server.CreateObject("adodb.recordset")

rsRem.open "select summary from " & varTblNameissues & " where issueid=" & request("issueid"),con
summary = rsRem("summary")
rsRem.close

rsRem.open "select rt.*,i.summary from " & varTblNameReminder_table & " rt," & varTblNameissues & " i where" & _
			   " rt.issueidCol = i.issueid and  rt.issueidCol=" & request("issueid") ,con
			   
rsUsers.Open "select u.email from  " &  varTblNameUsers & " u, " &  varTblNameissues & _
			  " i where i.projectid = u.projectid and i.issueid =" & request("issueid") ,con			    

' ********************* CHECKING PERMISSION ENDS***************************
%>


<!-- *********************** Java SCript ******************************* -->
<script language="Javascript">
self.moveTo(0,0);
<% if request("reload_opener") = "true" then %>
	opener.location.reload();
<% end if%>
</script>


<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title><%=varSiteSpecTitle%></title>
<LINK href="include/style.css" type=text/css rel=stylesheet>
<style type="text/css">
<!--
#lyr_daily_weekly {
	position:absolute;
	visibility:hidden;
	width:100%;
	height:100%;
	z-index:1;
}
#lyr_monthly_yearly {
	position:absolute;
	visibility:hidden;
	width:100%;
	height:100%;
	z-index:1;
}

-->
</style>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<body bgcolor="#CCCCCC" leftmargin="0" topmargin="0">
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
                  <td valign="top"> <table width="98%" border="0" align="center" cellpadding="0" cellspacing="0">
                      <tr> 
                        <td width="67%">&nbsp;</td>
                        <td width="33%" align="left"><!-- #include file="include/TopRightNavBar.asp" --></td>
                      </tr>
                      <tr>
                        <td>&nbsp;</td>
                        <td>&nbsp;</td>
                      </tr>
                      <tr> 
                        <td colspan="2"><!-- #include file="include/GeneralTopNavBar.asp"--></td>
                      </tr>
                    </table>
                    <font color="#FFFFFF">-</font><br>
            <!-- **************************************** form starting from here 22/march/2k5 ****************************** -->
                <form name="frmRem" method="post" action="editreminder_process.asp" onSubmit="return checkvalidate();">    
                     <input  type="hidden" name="issueid" value="<%=request("issueid")%>">
					 <input  type="hidden" name="remid" value="<%=request("remid")%>">
						
                      <table width="98%" border="0" align="center" cellpadding="0" cellspacing="0">
                        <tr bgcolor="#F3F3F3"> 
                          <td height="43" colspan="2"><div align="center" class="redbold">Edit 
                              Reminders</div></td>
                        </tr>
                        <tr>
                          <td width="12%" height="35" align="left">Issue Summary:&nbsp;&nbsp;</td>
                          <td width="88%" height="35">&nbsp;&nbsp;<%=summary%></td>
                        </tr>
                        <tr>
                          <td height="35" colspan="2" align="left">
						  <table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#CCCCCC">
						  <tr><td>
						  <table width="100%" border="0" align="center" cellpadding="1" cellspacing="1">
                            <tr bgcolor="#FFFFFF" class="head"><td align="center" colspan="10" class="regularText">
                              Reminders till now
                              </td></tr>
                            <tr bgcolor="#FFFFFF">
                              <td>Type</td>
                              <td>Due Date Time </td>
                              <td>T-Zn</td>
                              <td>Date Entered</td>
                              <td>T-Zn</td>
                              <td>Person Responsible </td>
                              <td>Alert Freq </td>
                              <td>State</td>
                              <td>Summary</td>
                              <td>Edit</td>
                            </tr>
							<%
							while not rsRem.eof
							%>
                            <tr bgcolor="#FFFFFF">
                              <td><%=rsRem("TypeCol")%></td>
                              <td><%=rsRem("Due_DateTimeCol")%></td>
                              <td><%=rsRem("TimeZoneCol")%></td>
                              <td><%=rsRem("Date_enteredCol")%></td>
                              <td><%="GMT"%></td>
                              <td><%=rsRem("Person_ResponsibleCol")%></td>
                              <td><%=rsRem("AlertFrequencyCol")%></td>
                              <td><%=rsRem("StateCol")%></td>
                              <td><a href="rem_desc.asp?is=<%=summary%>&id=<%=rsRem("rem_idCol")%>" target="_blank"><%=rsRem("NotesCol")%></a></td>
							  <td><a href="editreminder.asp?issueid=<%=request("issueid")%>&remid=<%=rsRem("rem_idCol")%>">Edit</a>/<a href="editreminder_process.asp?issueid=<%=request("issueid")%>&remid=<%=rsRem("rem_idCol")%>&close=true">Close</a></td>
                            </tr>
							<%
							rsRem.movenext
							wend
							%>
							<!--<tr bgcolor="#FFFFFF"><td colspan="10">&nbsp;</td></tr> -->

                          </table></td></tr></table></td>
                        </tr>
						<tr><td colspan="2">&nbsp;</td></tr>
                        <tr>
                          <td height="35" colspan="2" align="left"><table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#CCCCCC">
						  <tr><td>
						  <table width="100%" border="0" align="center" cellpadding="1" cellspacing="1">
                            <tr bgcolor="#FFFFFF" class="head"><td align="center" colspan="7" class="regularText">
                              Add Reminders
                              </td></tr>
                            <tr bgcolor="#FFFFFF">
                              <td width="14%">Type</td>
                              <td width="34%">Due Date Time </td>
                              <td width="7%">T-Zn</td>
                              <td width="12%">Person Responsible </td>
                              <td width="9%">Alert Freq </td>
                              <td width="8%">State</td>
                              <td width="16%">Summary</td>
                              </tr>
                            <tr bgcolor="#FFFFFF">
                              <td><select name="lstType" class="listbox" onChange="if(frmRem.lstType.value =='On_DueDate')frmRem.alertFreq.value='none';else frmRem.alertFreq.selectedIndex=0;">
							  <option value="Major_After" selected="selected">Major_After</option>
							  <option value="Major_Before">Major_Before</option>
							  <option value="On_DueDate">On_DueDate</option>							  
                              </select>                              </td>
                              <td><select id="year" name="year" class="listbox">
                                <option value="2024" selected="selected">2024</option>
                                <option value="2025">2025</option>
                                <option value="2026">2026</option>
                                <option value="2027">2027</option>
                                <option value="2028">2028</option>
								<option value="2029">2029</option>
								<option value="2030">2030</option>
								<option value="2031">2031</option>
								<option value="2032">2032</option>
								<option value="2033">2033</option>
                              </select>
                                <select id="month" name="month" class="listbox">
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
                                <select id="day" name="day" class="listbox">
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
                                <select id="hour" name="hour" class="listbox">
                                  <option value="0" selected="selected">00 / 12 am</option>
                                  <option value="1">01 / 1 am</option>
                                  <option value="2">02 / 2 am</option>
                                  <option value="3">03 / 3 am</option>
                                  <option value="4">04 / 4 am</option>
                                  <option value="5">05 / 5 am</option>
                                  <option value="6">06 / 6 am</option>
                                  <option value="7">07 / 7 am</option>
                                  <option value="8">08 / 8 am</option>
                                  <option value="9">09 / 9 am</option>
                                  <option value="10">10 / 10 am</option>
                                  <option value="11">11 / 11 am</option>
                                  <option value="12">12 / 12 pm</option>
                                  <option value="13">13 / 1 pm</option>
                                  <option value="14">14 / 2 pm</option>
                                  <option value="15">15 / 3 pm</option>
                                  <option value="16">16 / 4 pm</option>
                                  <option value="17">17 / 5 pm</option>
                                  <option value="18">18 / 6 pm</option>
                                  <option value="19">19 / 7 pm</option>
                                  <option value="20">20 / 8 pm</option>
                                  <option value="21">21 / 9 pm</option>
                                  <option value="22">22 / 10 pm</option>
                                  <option value="23">23 / 11 pm</option>
                                </select>
                                <input type="text" id="min" name="min" maxlength="2" size="2" value="00" class="inputbox">
                                <input type="hidden" value="01" name="sec" /></td>
                              <td><select id="select" name="timeZone" class="listbox">
							  <option value="EST" selected="selected">EST</option>
							  <option value="IST">IST</option>
							  <option value="GMT">GMT</option>
                              </select></td>
                              <td><select name="personResponsible" class="listbox">
							  <%
							  while not rsUsers.eof
							  %>
							  	<option value="<%=rsUsers("email")%>"><%=rsUsers("email")%></option>
							  <%
							  rsUsers.movenext
							  wend 
							  rsUsers.close
							  con.close
							  %>
                              </select></td>
                              <td><select name="alertFreq" class="listbox" onChange="if(this.value=='daily') show_daily(); else if(this.value=='monthly') show_monthly(); else hide_layers();">
							  <option value="hourly" selected="selected">Hourly</option>
							  <option value="daily">Daily</option>
							  <option value="weekly">Weekly</option>
							  <option value="monthly">Monthly</option>
							  <option value="Yearly">Yearly</option>
							  <option value="none">None</option>
                              </select></td>
                              <td>
							  <select name="State" class="listbox">
							  <option value="open">Open</option>
							  <option value="close">Close</option>
							  </select>							  </td>
                              <td><input type="text" name="Notes" class="inputbox"></td>
                              </tr>
                            <tr bgcolor="#FFFFFF" height="90px">
							  <td> Description -&gt; </td>
							  <td id="checkme"><textarea name="Description" cols="30" rows="5" wrap="virtual"></textarea></td>
							  <td colspan="5" valign="top">							  
							  <div id="lyr_daily_weekly">
<table width="100%" border="0" bgcolor="#CCCCCC" >
  <tr>
    <td colspan="7" align="left" class="head"><strong>Daily and Weekly</strong></td>
  </tr>
  <tr>
    <td align="center">Sun</td>
    <td align="center">Mon</td>
    <td align="center">Tue</td>
    <td align="center">Wed</td>
    <td align="center">Thu</td>
    <td align="center">Fri</td>
    <td align="center">Sat</td>
    </tr>
  <tr>
    <td align="center"><input name="chk_daily_weekly" type="checkbox" value="Sun" checked /> </td>
    <td align="center"><input name="chk_daily_weekly" type="checkbox" value="Mon" checked /></td>
    <td align="center"><input name="chk_daily_weekly" type="checkbox" value="Tue" checked /></td>
    <td align="center"><input name="chk_daily_weekly" type="checkbox" value="Wed" checked /></td>
    <td align="center"><input name="chk_daily_weekly" type="checkbox" value="Thu" checked /></td>
    <td align="center"><input name="chk_daily_weekly" type="checkbox" value="Fri" checked /></td>
    <td align="center"><input name="chk_daily_weekly" type="checkbox" value="Sat" checked /></td>
    </tr>
  <tr>
    <td colspan="7" align="center">Please check the days when you will like event to run </td>
    </tr>
  <!--<tr>
    <td colspan="7" align="center">
	<select name="lst_week">
	<option value="Current Week">Current Week</option>
	<option value="Next Week">Next Week</option>
	<option value="Every Week">Every Week</option>
	<option value="Every 2nd Week">Every 2nd Week</option>
	<option value="Every 3rd Week">Every 3rd Week</option>
	<option value="Every 4th Week">Every 4th Week</option>
	</select>&nbsp;
	</td>
  </tr>-->
</table>
</div>
<div id="lyr_monthly_yearly">
<table width="100%" border="0" bgcolor="#CCCCCC">
  <tr>
    <td colspan="12"><strong>Monthly and Yearly</strong></td>
  </tr>
  <tr>
    <td align="center">Jan</td>
    <td align="center">Feb</td>
    <td align="center">Mar</td>
    <td align="center">Apr</td>
    <td align="center">May</td>
    <td align="center">Jun</td>
    <td align="center">Jul</td>
    <td align="center">Aug</td>
    <td align="center">Sep</td>
    <td align="center">Oct</td>
    <td align="center">Nov</td>
    <td align="center">Dec</td>
  </tr>
  <tr>
    <td align="center"><input name="chk_monthly_yearly" type="checkbox" value="Jan" checked /> </td>
    <td align="center"><input name="chk_monthly_yearly" type="checkbox" value="Feb" checked /></td>
    <td align="center"><input name="chk_monthly_yearly" type="checkbox" value="Mar" checked /></td>
    <td align="center"><input name="chk_monthly_yearly" type="checkbox" value="Apr" checked /></td>
    <td align="center"><input name="chk_monthly_yearly" type="checkbox" value="May" checked /></td>
    <td align="center"><input name="chk_monthly_yearly" type="checkbox" value="Jun" checked /></td>
    <td align="center"><input name="chk_monthly_yearly" type="checkbox" value="Jul" checked /></td>
    <td align="center"><input name="chk_monthly_yearly" type="checkbox" value="Aug" checked /></td>
    <td align="center"><input name="chk_monthly_yearly" type="checkbox" value="Sep" checked /></td>
    <td align="center"><input name="chk_monthly_yearly" type="checkbox" value="Oct" checked /></td>
    <td align="center"><input name="chk_monthly_yearly" type="checkbox" value="Nov" checked /></td>
    <td align="center"><input name="chk_monthly_yearly" type="checkbox" value="Dec" checked /></td>
  </tr>
  <tr>
    <td colspan="12" align="center">Please check the months for the event to run </td>
    </tr>
  <tr>
    <td colspan="12"><!--<table width="100%" border="0">
  <tr>
    <td width="35%"><input type="radio" name="opt_monthly_yearly" value="date" onClick="lyr_monthly_yearly_date.style.visibility='visible';lyr_monthly_yearly_dow.style.visibility='hidden';" />Date</td>
    <td width="65%" rowspan="2">
	<div id="lyr_monthly_yearly_date" style="visibility:hidden">On &nbsp;
	<select name="lst_day_number">
	<option value="1">01</option>
	<option value="2">02</option>
	<option value="3">03</option>
	<option value="4">04</option>
	<option value="5">05</option>
	<option value="6">06</option>
	<option value="7">07</option>
	<option value="8">08</option>
	<option value="9">09</option>
	<option value="10">10</option>
	<option value="11">11</option>
	<option value="12">12</option>
	<option value="13">13</option>
	<option value="14">14</option>
	<option value="15">15</option>
	<option value="16">16</option>
	<option value="18">17</option>
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
	</select>&nbsp;day of month
	</div>
	<div id="lyr_monthly_yearly_dow" style="visibility:hidden">every
	<select name="lst_week_number">
	<option value="1">1st</option>
	<option value="2">2nd</option>
	<option value="3">3rd</option>
	<option value="4">4th</option>
	<option value="5">5th</option>
	<option value="last">last</option>
	</select>&nbsp;
	<select name="lst_week_day">
	<option value="Sun">Sunday</option>
	<option value="Mon">Monday</option>
	<option value="Tue">Tuesday</option>
	<option value="Wed">Wednesday</option>
	<option value="Thu">Thursday</option>
	<option value="Fri">Friday</option>
	<option value="Sat">Saturday</option>
	</select>
	</div>
	
	</td>
  </tr>
  <tr>
    <td><input type="radio" name="opt_monthly_yearly" value="day_of_week" onClick="lyr_monthly_yearly_date.style.visibility='hidden';lyr_monthly_yearly_dow.style.visibility='visible';"/>Day of week</td>
  </tr>
</table>-->
</td>
    </tr>
</table>

</div>							  </td>
                            </tr>
                            <tr bgcolor="#FFFFFF">
							  <td colspan="7"><br><input type="submit" value="Save Reminder" name="cmdSave" class="buttonc"> &nbsp;<input type="button" value="Cancel" name="cmdCancel" class="buttonc" onClick="cancel()"></td>
                            </tr>
                          </table></td></tr></table></td>
                        </tr>
                        <!-- adding the commentss -->
                      </table>
								
           <!-- ********************* end of form            ********** -->
           </form>         
                  </td>
                </tr>
              </table>
              
            </div></td>
        </tr>
        <tr> 
          <td width="45%" id="hello23">&nbsp;</td>
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

<script language="JavaScript">
// layers positioning, visualizing etc
function checkvalidate() {
	//opener.location.reload();
	//return t
}
function cancel() {
	frmRem.remid.value = "";
	frmRem.submit();
}
function hide_layers(){
lyr_monthly_yearly.style.visibility = 'hidden';
lyr_daily_weekly.style.visibility = 'hidden';	
//lyr_monthly_yearly_date.style.visibility='hidden';
//lyr_monthly_yearly_dow.style.visibility='hidden';
}
function show_daily() {
lyr_monthly_yearly.style.visibility = 'hidden';
lyr_daily_weekly.style.visibility = 'visible';	
//lyr_monthly_yearly_date.style.visibility='hidden';
//lyr_monthly_yearly_dow.style.visibility='hidden';
}
function show_monthly() {
lyr_monthly_yearly.style.visibility = 'visible';
lyr_daily_weekly.style.visibility = 'hidden';	
}
function getPageCoors (element) {
	var coords = {x: 0, y: 0};
	while (element) {
	coords.x += element.offsetLeft;
	coords.y += element.offsetTop;
	element = element.offsetParent;
	}
	return coords;
}

var element;
if (document.getElementById) {
element = document.getElementById('lyr_daily_weekly');
var coords = getPageCoors(element);
//alert(coords.x + ':' + coords.y);
lyr_monthly_yearly.style.left = coords.x;
lyr_monthly_yearly.style.top = coords.y;
}

<%
if request("remid") <> "" then
%>
frmRem.lstType.value = '<%=remType%>';
frmRem.year.value = <%=year(remdueDate)%>; 
frmRem.month.value = <%=month(remdueDate) - 1%>;
frmRem.day.value = <%=day(remdueDate)%>;
frmRem.hour.value = <%=hour(remdueDate)%>;
frmRem.min.value = <%=minute(remdueDate)%>;
frmRem.timeZone.value = '<%=remtimeZone%>';
frmRem.personResponsible.value = '<%=remPR%>';
frmRem.alertFreq.value = '<%=remalertFreq%>';
//if(frmRem.alertFreq.value == 'daily') show_daily();
//if(frmRem.alertFreq.value == 'monthly') show_monthly();
frmRem.State.value = '<%=remState%>';
frmRem.Notes.value = '<%=remNotes%>';
frmRem.Description.value = '<%=remDesc%>';
var flag = false;
if(frmRem.alertFreq.value == 'daily')
{
	var daily_weekly = '<%=remDaysofWeek%>';
	var dw = new Array(); 
	dw = daily_weekly.split(", ");
	for(j=0;j<frmRem.chk_daily_weekly.length;j++)
	{
		for(i=0;i<dw.length;i++)
		{
			if(frmRem.chk_daily_weekly[j].value == dw[i])
			{
				flag = true;
				break;
			}
		}
		if(flag == false)  frmRem.chk_daily_weekly[j].checked = false;
		flag = false;
	}

	show_daily();
}
if(frmRem.alertFreq.value == 'monthly')
{
	var monthly_yearly = '<%=remMonthsofYear%>';
	var my = new Array(); 
	my = monthly_yearly.split(", ");
	for(j=0;j<frmRem.chk_monthly_yearly.length;j++)
	{
		for(i=0;i<my.length;i++)
		{
			if(frmRem.chk_monthly_yearly[j].value == my[i])
			{
				flag = true;
				break;
			}
		}
		if(flag == false)  frmRem.chk_monthly_yearly[j].checked = false;
		flag = false;
	}
	show_monthly();
}

frmRem.lstType.focus();
<%
end if
%>

</script>
<!-- ******************************************************* closing the connection object  *****************-->
<%
  if con.state then
  con.close
  set con = nothing
  end if

%>