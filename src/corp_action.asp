<%
if session("user") = "" then
	response.Write("Please check the session")
end if 
%>
<!-- #include file ="connect.asp" -->
<%
if session("user") <> "shivankz@gmail.com" and session("user") <> "saurabhz@gmail.com" then
	response.Write("You are not a valid user for this page.")
	con.close
	set con = nothing
	response.End()
end if

dim rs
dim message
dim sortby
dim order
dim viewType
dim	symbol 
dim	subsymbol 
dim	operation 
dim	dateoperation 
dim	expirationdate 
dim	affect
dim strclose
dim notes

set rs = server.CreateObject("adodb.recordset")
if request("cmd") = "close" then
	con.execute "update " &  varTblNameCorp_Act_table & " set close = 'T' where id =" & request("id")
	message = "Entry Closed"
end if

if request("cmd") = "delete" then
	con.execute "delete from " &  varTblNameCorp_Act_table & " where id =" & request("id")
	message = "Entry deleted"
end if

if request("cmd") = "TransferToIMS" then
	rs.open "select * from " &  varTblNameCorp_Act_table & " where id =" & request("id") ,con
	if not rs.eof then
		symbol = rs("symbol")
		subsymbol = rs("subsymbol")
		operation = rs("operation")
		dateoperation = rs("dateoperation")
		expirationdate = rs("expirationdate")
		affect = rs("affect")
		strclose = rs("close")
		notes = rs("notes")
		rs.close
		rs.open "select * from " &  varTblNameIms_Compliance_Table & " where symbol ='" & symbol & "' and subsymbol = '" & subsymbol & "' and operation = '" & operation & "' and dateoperation = '" & dateoperation & "' and expirationdate = '" & expirationdate & "' and affect = '" & affect & "'",con
		if rs.eof then
			rs.close
			con.execute "insert into " &  varTblNameIms_Compliance_Table & " values('" & symbol & "','" & subsymbol & "','" & operation & "','" & dateformat(dateoperation) & "','" & dateformat(expirationdate) & "','','" & affect & "','" & strclose & "','" & notes & "')"
			message = "Entry inserted in ims compliance table"
		else
			rs.close
			message = "Entry already present in ims compliance table"
		end if
	end if
	if rs.state = 1 then rs.close
end if

if request("cmd") = "TransferToMLP" then
	rs.open "select * from " &  varTblNameCorp_Act_table & " where id =" & request("id") ,con
	if not rs.eof then
		symbol = rs("symbol")
		subsymbol = rs("subsymbol")
		operation = rs("operation")
		dateoperation = rs("dateoperation")
		expirationdate = rs("expirationdate")
		affect = rs("affect")
		strclose = rs("close")
		notes = rs("notes")
		rs.close
		rs.open "select * from " &  varTblNameMlp_Compliance_Table & " where symbol ='" & symbol & "' and subsymbol = '" & subsymbol & "' and operation = '" & operation & "' and dateoperation = '" & dateoperation & "' and expirationdate = '" & expirationdate & "' and affect = '" & affect & "'",con
		if rs.eof then
			rs.close
			con.execute "insert into " &  varTblNameMlp_Compliance_Table & " values('" & symbol & "','" & subsymbol & "','" & operation & "','" & dateformat(dateoperation) & "','" & dateformat(expirationdate) & "','','" & affect & "','" & strclose & "','" & notes & "')"
			message = "Entry inserted in mlp compliance table"
		else
			rs.close
			message = "Entry already present in mlp compliance table"
		end if
	end if
	if rs.state = 1 then rs.close
end if


if request("cmdSubmit") = "Add Row" then
	rs.open "select * from " &  varTblNameCorp_Act_table & " where symbol ='" & request("txtsymbol") & "' and subsymbol = '" & request("txtsubsymbol") & "' and operation = '" & request("lstoperation") & "' and dateoperation = '" & request("txtdateoprtn") & "' and expirationdate = '" & request("txtexpdate") & "' and affect = '" & request("lstaffect") & "'",con
	if rs.eof then
		rs.close
		con.execute "insert into " &  varTblNameCorp_Act_table & " values('" & request("txtsymbol") & "','" & request("txtsubsymbol") & "','" & request("lstoperation") & "','" & request("txtdateoprtn") & "','" & request("txtexpdate") & "','','" & request("lstaffect") & "','F','" & request("txtnotes") & "')"
	else
		rs.close
		message = "Entry already present"
	end if
end if

if request("view") = "" then
	viewType = "All"
elseif request("view") = "Open" then
	viewType = "Open"
elseif request("view") = "Close" then
	viewType = "Close"
else
	viewType = "All"
end if

if request("sort") = "" then
	sortby = "symbol"
else
	sortby = request("sort")	
end if

select case viewType
case "All"
	rs.open "select * from " &  varTblNameCorp_Act_table & " order by " & sortby & " " & request("order"),con
case "Open"
	rs.open "select * from " &  varTblNameCorp_Act_table & " where close <> 'T' order by " & sortby & " " & request("order"),con
case "Close"
	rs.open "select * from " &  varTblNameCorp_Act_table & " where close = 'T' order by " & sortby & " " & request("order"),con
end select


if request("order") = "" then
	order = "desc"
else
	order = ""
end if
dim str
function dateformat(str)
	dim a
	a = split(str,"/")
	dateformat= a(2) & "-" & a(0) & "-" & a(1) 
end function
%>
<html>
<head>
<title>CORP ACTIONS TABLE</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<LINK href="include/style.css" type="text/css" rel="stylesheet">
</head>
<body>
<font color="blue"><h4 align="center">CORP ACTIONS TABLE</h4></font>
Note :- <li>Input date format: yyyymmdd Display date format: mm/dd/yyyy</li>
<li>Click the columns heading for sorting</li>
<p align="center"><font color="red"><b><%=message%></b></font></p>
<table width="100%" border="0" cellspacing="1" cellpadding="1" bordercolor="#000000" >
  <tr> 
    <td width="14%"><strong>Display Rows :</strong></td>
    <td width="63%"><strong> 
      <input type="radio" name="optAll" value="All" onClick="window.location = 'corp_action.asp?view=All'">
      All &nbsp;&nbsp;&nbsp;
      <input type="radio" name="optOpen" value="Open" onClick="window.location = 'corp_action.asp?view=Open'">
      Open &nbsp;&nbsp;&nbsp;
      <input type="radio" name="optClose" value="Close" onClick="window.location = 'corp_action.asp?view=Close'">
      Close</strong></td>
	  
    <td width="23%"><strong>Display Type : <%=viewType%></strong></td>
  </tr>
</table>
<table width="100%" border="1" cellspacing="1" cellpadding="1" bordercolor="#000000" >
  <tr> 
    <td ><strong><a href="corp_action.asp?sort=symbol&order=<%=order%>&view=<%=viewType%>">Symbol</a></strong></td>
    <td><strong><a href="corp_action.asp?sort=subsymbol&order=<%=order%>&view=<%=viewType%>">SubSymbol</a></strong></td>
    <td><strong><a href="corp_action.asp?sort=operation&order=<%=order%>&view=<%=viewType%>">Operation</a></strong></td>
	<td><strong><a href="corp_action.asp?sort=affect&order=<%=order%>&view=<%=viewType%>">Affect</a></strong></td>
    <td><strong><a href="corp_action.asp?sort=dateoperation&order=<%=order%>&view=<%=viewType%>">DateOperation</a></strong></td>
    <td><strong><a href="corp_action.asp?sort=expirationdate&order=<%=order%>&view=<%=viewType%>">ExpirationDate</a></strong></td>
    <td><strong>DaysSince</strong></td>
	<td><strong>DaysLeft</strong></td>
	<td><strong>Command</strong></td>
	<td><strong>Notes</strong></td>
  </tr>
  <tr><td colspan="10"><hr color="#000000" noshade size="0"></td></tr>
  <%
if not rs.eof then
	while not rs.eof
		if rs("close") = "T" and viewType <> "Close" then
	%>
	<tr> 
    <td><strike><%=rs("symbol")%></strike></td>
    <td><strike><%=rs("subsymbol")%></strike></td>
    <td><strike><%=rs("operation")%></strike></td>
	<td><strike><%=rs("affect")%></strike></td>
    <td><strike><%=rs("dateoperation")%></strike></td>
    <td><strike><%=rs("expirationdate")%></strike></td>
    <td><strike><%=datediff("d",rs("dateoperation"),date)%></strike></td>
	<td><strike><%=datediff("d",date,rs("expirationdate"))%></strike></td>
	<td>
	<select name="cmd" class="formTextbox" onChange="getLocation(this.value,<%=rs("id")%>);">
		  <option value="delete">Delete</option>
		  <option value="close">Close</option>
		  <option value="TransferToIMS">TransferToIMS</option>
		  <option value="TransferToMLP">TransferToMLP</option>
        </select></td>
	<td><%=rs("Notes")%></td>
  </tr>
  <tr><td colspan="10"><hr color="#000000" noshade size="0"></td></tr>
	<%else%>
  <tr> 
    <td><%=rs("symbol")%></td>
    <td><%=rs("subsymbol")%></td>
    <td><%=rs("operation")%></td>
	<td><%=rs("affect")%></td>
    <td><%=rs("dateoperation")%></td>
    <td><%=rs("expirationdate")%></td>
    <td><%=datediff("d",rs("dateoperation"),date)%></td>
	<td><%=datediff("d",date,rs("expirationdate"))%></td>
	<td>
	<select name="cmd" class="formTextbox" onChange="getLocation(this.value,<%=rs("id")%>);">
	<option value="none">Select</option>
	  <option value="delete">Delete</option>
	  <option value="close">Close</option>
	  <option value="TransferToIMS">TransferToIMS</option>
	  <option value="TransferToMLP">TransferToMLP</option>
        </select></td>
	<td><%=rs("Notes")%></td>
  </tr>
   <tr><td colspan="10"><hr color="#000000" noshade size="0"></td></tr>
  <%
  	end if
	 rs.movenext
	 wend
else
%>
  <tr> 
    <td colspan="6" align="center">No records found currently</td>
  </tr>
  <%
end if
 %>
  <form action="corp_action.asp" method="post">
    <tr> 
      <td><input type="text" name="txtsymbol" class="formTextbox"></td>
      <td><input type="text" name="txtsubsymbol" class="formTextbox"></td>
      <td><select name="lstoperation" class="formTextbox">
		  <option value="result">Result</option>
		  <option value="OptionExpiry">OptionExpiry</option>
		  <option value="QuarterEnd">QuarterEnd</option>
		  <option value="RateHike">RateHike</option>
		  <option value="Misc">Misc</option>
        </select></td>
		<td>
		<select name="lstaffect" class="formTextbox">
		  <option value="result">Result</option>
		  <option value="OptionExpiry">OptionExpiry</option>
		  <option value="QuarterEnd">QuarterEnd</option>
		  <option value="RateHike">RateHike</option>
		  <option value="Misc">Misc</option>
        </select></td>
      <td><input type="text" name="txtdateoprtn" class="formTextbox"></td>
      <td><input type="text" name="txtexpdate" class="formTextbox"></td>
	  <td><font size="-7">Notes:</font>
<input type="text" name="txtnotes" class="formTextbox"></td>
      <td colspan="2" align="center"><input type="submit" name="cmdSubmit" value="Add Row" class="formbutton"></td>
    </tr>
  </form>
</table>

</body>
</html>
<SCRIPT>
function getLocation(value,id)
{
if (value != 'none')
{
	if(confirm('Do you really want to commit the action'))
		window.location = 'corp_action.asp?cmd=' + value + '&id=' + id ;
}
}
</SCRIPT>
<%
if rs.state =1 then rs.close
set rs = nothing
con.close
set con = nothing
%>