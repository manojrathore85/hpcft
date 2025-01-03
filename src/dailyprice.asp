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
set rs = server.CreateObject("adodb.recordset")
if request("cmd") = "delete" then
	con.execute "delete from " &  varTblNameDailyData & " where symbol ='" & request("symbol") & "' and ddate = '" & request("date") & "'"
end if

if request("cmdSubmit") = "Add Row" then
	rs.open "select * from " &  varTblNameDailyData & " where symbol ='" & request("symbol") & "' and ddate = '" & request("date") & "'",con
	if rs.eof then
		rs.close
		con.execute "insert into " &  varTblNameDailyData & " values('" & request("source") & "','" & request("symbol") & "','" & request("subsymbol") & "','" & request("date") & "'," & request("open") & "," & request("low") & "," & request("high") & "," & request("close") & "," & request("volume") & ")"
	else
		rs.close
		message = "Entry already present"
	end if
end if

if request("txtsort") = "" then
	sortby = "source"
else
	sortby = request("txtsort")	
end if

dim sql

sql = "select * from " &  varTblNameDailyData & " where 1"

if request("smonth") <> "" and request("syear") <> "" and request("sday") <> "" and request("emonth") <> "" and request("eyear") <> "" and request("eday") <> "" then
	sql = sql & " and ddate between '" & request("syear") & "-" & request("smonth") & "-" & request("sday") & "' and '" & request("eyear") & "-" & request("emonth") & "-" & request("eday") & "'"
end if

if request("lstsymbol") <> "" then
	sql = sql & " and symbol = '" & request("lstsymbol") & "'"
end if
if request("lstsubsymbol") <> "" then
	sql = sql & " and subsymbol = '" & request("lstsubsymbol") & "'"
end if

rs.open sql & " order by " & sortby & " " & request("txtorder"),con

if request("txtorder") = "" then
	order = "desc"
else
	order = ""
end if
%>
<html>
<head>
<title>MLP COMPLIANCE TABLE</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<LINK href="include/style.css" type="text/css" rel="stylesheet">
<script>
function funcsort(sortby,order)
{
	frm123.txtsort.value = sortby;
	frm123.txtorder.value = order;
	frm123.submit();
}
</script>
</head>

<body>
<font color="blue"><h4 align="center">DAILY PRICE</h4></font>
Note :- 
<li>Input date format: yyyymmdd Display date format: mm/dd/yyyy&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href="seldlyprc.asp"><strong>Click Here to select the date</strong></a></li>
<li>Click the columns heading for sorting&nbsp;</li>
<p align="center"><font color="red"><b><%=message%></b></font></p>
<table width="100%" border="1" cellspacing="0" cellpadding="0" bordercolor="#000000">
  <tr> 
	<td><strong><a href="#" onClick="funcsort('source','<%=order%>')">Source</a></strong></td>
    <td><strong><a href="#" onClick="funcsort('symbol','<%=order%>')">Symbol</a></strong></td>
    <td><strong><a href="#" onClick="funcsort('subsymbol','<%=order%>')">SubSymbol</a></strong></td>
    <td><strong><a href="#" onClick="funcsort('ddate','<%=order%>')">Date</a></strong></td>
    <td><strong><a href="#" onClick="funcsort('open','<%=order%>')">Open</a></strong></td>
    <td><strong><a href="#" onClick="funcsort('low','<%=order%>')">Low</a></strong></td>
	<td><strong><a href="#" onClick="funcsort('high','<%=order%>')">High</a></strong></td>
	<td><strong><a href="#" onClick="funcsort('close','<%=order%>')">Close</a></strong></td>
	<td><strong><a href="#" onClick="funcsort('volume','<%=order%>')">Volume</a></strong></td>
	<td align="center"><strong>Delete</strong></td>
  </tr>
  <tr><td colspan="10"><hr color="#000000" noshade size="0"></td></tr>
    <form action="dailyprice.asp" method="post">
  
    <tr> 
	  <td><input type="text" name="source" class="formTextbox" width="25" value="Manual"></td>
      <td><input type="text" name="symbol" class="formTextbox" width="25"></td>
      <td><input type="text" name="subsymbol" class="formTextbox" width="25"></td>
	  <td><input type="text" name="date" class="formTextbox" width="25"></td>
	  <td><input type="text" name="open" class="formTextbox" width="25"></td>
	  <td><input type="text" name="low" class="formTextbox" width="25"></td>
	  <td><input type="text" name="high" class="formTextbox" width="25"></td>
	  <td><input type="text" name="close" class="formTextbox" width="25"></td>
	  <td><input type="text" name="volume" class="formTextbox" width="25"></td>
      <td align="center"><input type="submit" name="cmdSubmit" value="Add Row" class="formbutton"></td>
    </tr>
  </form>
  <%
if not rs.eof then
	while not rs.eof
	%>
  <tr> 
	<td><%=rs("source")%></td>
    <td><%=rs("symbol")%></td>
    <td><%=rs("subsymbol")%></td>
    <td><%=rs("ddate")%></td>
    <td><%=rs("open")%></td>
    <td><%=rs("low")%></td>
    <td><%=rs("high")%></td>
	<td><%=rs("close")%></td>
	<td><%=rs("volume")%></td>
	<td align="center"><a href="dailyprice.asp?cmd=delete&symbol=<%=rs("symbol")%>&date=<%=rs("ddate")%>">Delete</a></td>
  </tr>
   <tr><td colspan="10"><hr color="#000000" noshade size="0"></td></tr>
  <%
	 rs.movenext
	 wend
end if
 %>

</table>
<form action="dailyprice.asp" method="post" name="frm123">
<input type="hidden" name="syear" value="<%=request("syear")%>">
   <input type="hidden" name="smonth" value="<%=request("smonth")%>">
    <input type="hidden" name="sday" value="<%=request("sday")%>">
	  <input type="hidden" name="eyear" value="<%=request("eyear")%>">
   <input type="hidden" name="emonth" value="<%=request("emonth")%>">
    <input type="hidden" name="eday" value="<%=request("eday")%>">
	  <input type="hidden" name="lstsymbol" value="<%=request("lstsymbol")%>">
   <input type="hidden" name="lstsubsymbol" value="<%=request("lstsubsymbol")%>">
   <input type="hidden" name="txtsort">
   <input type="hidden" name="txtorder">
</form>

</body>
</html>
<%
if rs.state =1 then rs.close
set rs = nothing
con.close
set con = nothing
%>