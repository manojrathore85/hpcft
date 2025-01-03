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
set rs = server.CreateObject("adodb.recordset")

if request("cmd") = "close" then
	con.execute "update " &  varTblNameIms_Compliance_Table & " set close = 'T' where id =" & request("id")
end if

if request("cmd") = "delete" then
	con.execute "delete from " &  varTblNameIms_Compliance_Table & " where id =" & request("id")
end if

if request("cmdSubmit") = "Add Row" then
	rs.open "select * from " &  varTblNameIms_Compliance_Table & " where symbol ='" & request("txtsymbol") & "' and subsymbol = '" & request("txtsubsymbol") & "' and operation = '" & request("lstoperation") & "' and dateoperation = '" & request("txtdateoprtn") & "' and expirationdate = '" & request("txtexpdate") & "' and affect = '" & request("lstaffect") & "'",con
	if rs.eof then
		rs.close
		con.execute "insert into " &  varTblNameIms_Compliance_Table & " values('" & request("txtsymbol") & "','" & request("txtsubsymbol") & "','" & request("lstoperation") & "','" & request("txtdateoprtn") & "','" & request("txtexpdate") & "','','" & request("lstaffect") & "','F','" & request("txtnotes") & "')"
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
	rs.open "select * from " &  varTblNameIms_Compliance_Table & " order by " & sortby & " " & request("order"),con
case "Open"
	rs.open "select * from " &  varTblNameIms_Compliance_Table & " where close <> 'T' order by " & sortby & " " & request("order"),con
case "Close"
	rs.open "select * from " &  varTblNameIms_Compliance_Table & " where close = 'T' order by " & sortby & " " & request("order"),con
end select

if request("order") = "" then
	order = "desc"
else
	order = ""
end if
%>
<html>
<head>
<title>IMS COMPLIANCE TABLE</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<LINK href="include/style.css" type="text/css" rel="stylesheet">
</head>
<body>
<font color="blue">
<h4 align="center">IMS COMPLIANCE PAGE</h4>
</font>
Note :- <li>Input date format: yyyymmdd Display date format: mm/dd/yyyy</li>
<li>Click the columns heading for sorting</li>
<p align="center"><font color="red"><b><%=message%></b></font></p>
<table width="100%" border="0" cellspacing="1" cellpadding="1" bordercolor="#000000" >
  <tr> 
    <td width="14%"><strong>Display Rows :</strong></td>
    <td width="63%"><strong> 
      <input type="radio" name="optAll" value="All" onClick="window.location = 'ims_compliance.asp?view=All'">
      All &nbsp;&nbsp;&nbsp;
      <input type="radio" name="optOpen" value="Open" onClick="window.location = 'ims_compliance.asp?view=Open'">
      Open &nbsp;&nbsp;&nbsp;
      <input type="radio" name="optClose" value="Close" onClick="window.location = 'ims_compliance.asp?view=Close'">
      Close</strong></td>
	  
    <td width="23%"><strong>Display Type : <%=viewType%></strong></td>
  </tr>
</table>
<table width="100%" border="1" cellspacing="1" cellpadding="1" bordercolor="#000000" >
  <tr> 
    <td ><strong><a href="ims_compliance.asp?sort=symbol&order=<%=order%>&view=<%=viewType%>">Symbol</a></strong></td>
    <td><strong><a href="ims_compliance.asp?sort=subsymbol&order=<%=order%>&view=<%=viewType%>">SubSymbol</a></strong></td>
    <td><strong><a href="ims_compliance.asp?sort=operation&order=<%=order%>&view=<%=viewType%>">Operation</a></strong></td>
	<td><strong><a href="ims_compliance.asp?sort=affect&order=<%=order%>&view=<%=viewType%>">Affect</a></strong></td>
    <td><strong><a href="ims_compliance.asp?sort=dateoperation&order=<%=order%>&view=<%=viewType%>">DateOperation</a></strong></td>
    <td><strong><a href="ims_compliance.asp?sort=expirationdate&order=<%=order%>&view=<%=viewType%>">ExpirationDate</a></strong></td>
    <td><strong>DaysSince</strong></td>
	<td><strong>DaysLeft</strong></td>
	<td><strong>Delete</strong></td>
	<td><strong>Close</strong></td>
	<td><strong>Notes</strong></td>
  </tr>
  <tr><td colspan="8"><hr color="#000000" noshade size="0"></td></tr>
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
	<td><a href="ims_compliance.asp?cmd=delete&id=<%=rs("id")%>">Delete</a></td>
	<td><a href="ims_compliance.asp?cmd=Close&id=<%=rs("id")%>">Close</a></td>
	<td><%=rs("Notes")%></td>
  </tr>
  <tr><td colspan="8"><hr color="#000000" noshade size="0"></td></tr>
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
	<td><a href="ims_compliance.asp?cmd=delete&id=<%=rs("id")%>">Delete</a></td>
	<td><a href="ims_compliance.asp?cmd=close&id=<%=rs("id")%>">Close</a></td>
	<td><%=rs("Notes")%></td>
  </tr>
   <tr><td colspan="8"><hr color="#000000" noshade size="0"></td></tr>
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
  <form action="ims_compliance.asp" method="post">
    <tr> 
      <td><input type="text" name="txtsymbol" class="formTextbox"></td>
      <td><input type="text" name="txtsubsymbol" class="formTextbox"></td>
      <td><select name="lstoperation" class="formTextbox">
          <option value="buy">Buy</option>
          <option value="sell">Sell</option>
		  <option value="result">Result</option>
		  <option value="OptionExpiry">OptionExpiry</option>
		  <option value="QuarterEnd">QuarterEnd</option>
		  <option value="RateHike">RateHike</option>
		  <option value="GTCBuy">GTCBuy</option>
		  <option value="GTCSell">GTCSell</option>
		  <option value="DayBuy">DayBuy</option>
		  <option value="DaySell">DaySell</option>
		  <option value="Misc">Misc</option>
        </select></td>
		<td>
		<select name="lstaffect" class="formTextbox">
          <option value="buy">Buy</option>
          <option value="sell">Sell</option>
		  <option value="result">Result</option>
		  <option value="OptionExpiry">OptionExpiry</option>
		  <option value="QuarterEnd">QuarterEnd</option>
		  <option value="RateHike">RateHike</option>
		  <option value="GTCBuy">GTCBuy</option>
		  <option value="GTCSell">GTCSell</option>
		  <option value="DayBuy">DayBuy</option>
		  <option value="DaySell">DaySell</option>
		  <option value="Misc">Misc</option>
        </select></td>
<td><input type="text" name="txtdateoprtn" class="formTextbox" value="<%=year(date)%><% if len(month(date)) = 1 then response.write("0")%><%=month(date)%><% if len(day(date)) = 1 then response.write("0")%><%=day(date)%>"></td>
      <td><input type="text" name="txtexpdate" class="formTextbox" value="<%=year(date)%><% if len(month(date)) = 1 then response.write("0")%><%=month(date)%><% if len(day(date)) = 1 then response.write("0")%><%=day(date)%>"></td>
	  <td><font size="-7">Notes:</font>
<input type="text" name="txtnotes" class="formTextbox"></td>
      <td colspan="2" align="center"><input type="submit" name="cmdSubmit" value="Add Row" class="formbutton"></td>
    </tr>
  </form>
</table>

</body>
</html>
<%
if rs.state =1 then rs.close
set rs = nothing
con.close
set con = nothing
%>