<%
if session("user") = "" then
	response.Write("Please check the session")
end if 
%>
<!-- #include file ="../connect.asp" -->
<%
if session("user") <> "shivankz@gmail.com" and session("user") <> "saurabhz@gmail.com" then
	response.Write("You are not a valid user for this page.")
	con.close
	set con = nothing
	response.End()
end if
dim sql
dim flag
dim prev_item
dim prev_month
dim totalHard
dim totalSoft
dim totalMultiplier
dim totalTotal

dim rs
set rs =server.CreateObject("adodb.recordset")
if request("Submit") = "Go" then
	sql = "select * from statement_tool_table stt where stt.item = '" & request("lstItem") & "'"
	if request("lstItem") = "All" then
		if request("txtStrategy") = "All" then
			sql = "select * from statement_tool_table stt where 1 "
		else
			sql = "select stt.* from statement_tool_table stt,item_list_table ilt where ilt.strategy ='" & request("txtStrategy") & "' and ilt.item = stt.item "
		end if
	end if
	if request("lstView") = "Latest Month" then
		dim latestmonth
		dim latestyear
		if request("lstItem") <> "All" then
			rs.open "select year(max(l_date)) as latestyear from statement_tool_table where item = '" & request("lstItem") & "'",con
		else
			rs.open "select year(max(l_date)) as latestyear from statement_tool_table",con
		end if
		latestyear = rs("latestyear")
		rs.close
		if request("lstItem") <> "All" then
			rs.open "select month(max(l_date)) as latestmonth from statement_tool_table where item ='" & request("lstItem") & "'",con
		else
			rs.open "select month(max(l_date)) as latestmonth from statement_tool_table",con
		end if
		latestmonth = rs("latestmonth")
		rs.close
		sql = sql & " and month(stt.l_date) = " & latestmonth & " and year(stt.l_date) = " & latestyear
	elseif request("lstView") = "Previous Month" then
		dim prevMonth
		dim prevYear
		dim maxLdate
		if request("lstItem") <> "All" then
			rs.open "select DATE_FORMAT(max(l_date),'%Y%m%d') as maxdate from statement_tool_TABLE where item ='" & request("lstItem") & "'",con
			maxLdate = rs("maxdate")
			rs.close
			rs.open "select year(max(l_date)) as prevYear from statement_tool_table where l_date <> '" & maxLdate & "' and item ='" & request("lstItem") & "'",con
			prevYear = rs("prevYear")
			rs.close
			rs.open "select month(max(l_date)) as prevMonth from statement_tool_table where l_date <> '" & maxLdate & "' and item ='" & request("lstItem") & "'",con
			prevMonth = rs("prevMonth")
			rs.close
		else
			rs.open "select DATE_FORMAT(max(l_date),'%Y%m%d') as maxdate from statement_tool_TABLE",con
			maxLdate = rs("maxdate")
			rs.close
			rs.open "select year(max(l_date)) as prevYear from statement_tool_table where l_date <> '" & maxLdate & "'",con
			prevYear = rs("prevYear")
			rs.close
			rs.open "select month(max(l_date)) as prevMonth from statement_tool_table where l_date <> '" & maxLdate & "'",con
			prevMonth = rs("prevMonth")
			rs.close
		end if
		sql = sql & " and month(stt.l_date) = " & prevMonth & " and year(stt.l_date) = " & prevYear
	else
		'sql = sql & " and l_date = (select max(l_date) from statement_tool_table)"
	end if
	rs.open sql & " order by stt.item,stt.l_date",con
end if
dim rsItem
set rsItem =server.CreateObject("adodb.recordset")
dim rsView
set rsView =server.CreateObject("adodb.recordset")
dim rsStrategy
set rsStrategy =server.CreateObject("adodb.recordset")

rsStrategy.open "select distinct strategy from item_list_table where strategy <> ''",con
if request("lstStrategy") = "All" or request("lstStrategy") = "" then
	rsItem.open "select item from item_list_table where item <> ''",con
elseif request("lstStrategy") = "" then
	rsItem.open "select item from item_list_table where item <> ''",con
else
	rsItem.open "select item from item_list_table where item <> '' and strategy ='" & request("lstStrategy") & "'",con
end if
	
rsView.open "select views from item_list_table where Views <> ''",con

%>
<html>
<head>
<title>Display Section</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<body topmargin="0">
<!-- #include file="header.inc" -->
<form name="frmTemp">
<h2 align="center">Display Section for Strategy 

  <select name="lstStrategy" onChange="javascript:window.location = 'ExpenseReport.asp?lstStrategy=' + this.value">
    <option value="All" selected>All</option>
    <%
	while not rsStrategy.eof
	%>
    <option value="<%=rsStrategy("strategy")%>"><%=rsStrategy("strategy")%></option>
    <%
	rsStrategy.movenext
	wend
	%>
  </select>
</h2>
</form>
<p align="center"><font color="red"><b><%=message%></b></font></p>
<form action="ExpenseReport.asp" name="frmReport" method="post">
<input type="hidden" name="txtStrategy" value="<%=request("lstStrategy")%>">
<table width="100%" border="0">
  <tr>
    <td width="27%"><select name="lstItem">
        <option value="none" selected>Select Item</option>
        <%
	while not rsItem.eof
	%>
        <option value="<%=rsItem("item")%>"><%=rsItem("item")%></option>
        <%
	rsItem.movenext
	wend
	%>
      </select></td>
    <td width="23%"><select name="lstView">
        <option value="none" selected>Select View</option>
        <%
	while not rsView.eof
	%>
        <option value="<%=rsView("Views")%>"><%=rsView("Views")%></option>
        <%
	rsView.movenext
	wend
	%>
      </select></td>
    <td width="50%"><input type="submit" name="Submit" value="Go"></td>
  </tr>
</table>
  <table width="100%" border="0" bgcolor="#999999">
    <tr> 
      <td><strong>Date</strong></td>
      <td><strong>Hard Limit</strong></td>
      <td><strong>Soft Limit</strong></td>
      <td><strong>Multiplier</strong></td>
      <td><strong>Total</strong></td>
      <td><strong>Notes</strong></td>
    </tr>
    <%
if rs.state = 1 then
	if rs.eof then
%>
    <TR> 
      <td colspan="6" align="center">No Rows to display</td>
    </TR>
    <%
else
	while not rs.eof
		flag = true
		if (prev_month <> "" and prev_month <> month(rs("l_date"))) or (prev_month = month(rs("l_date")) and prev_item <> rs("item")) then
			prev_month = month(rs("l_date"))
			%>
    <tr> 
      <td><font color="#FF0000"><strong>Total</strong></font></td>
      <td><strong><%=totalHard%></strong></td>
      <td><strong><%=totalSoft%></strong></td>
      <td><strong><%=totalMultiplier%></strong></td>
      <td><strong><%=totalTotal%></strong></td>
      <td>&nbsp;</td>
    </tr>
    <!--<TR><td colspan="6" align="center">Item :- <%=rs("item")%> , Month :- <%=month(rs("l_date"))%></td></TR>-->
    <%
			if prev_item = rs("item") then
			%>
    <TR> 
      <td colspan="6" align="center" bgcolor="#FF66FF">Item :- <%=rs("item")%> 
        , Month :- <%=month(rs("l_date"))%></td>
    </TR>
    <%
			end if
			totalHard = 0
			totalSoft = 0
			totalMultiplier = 0
			totalTotal = 0
		end if
		if prev_item <> rs("item") then
			prev_item = rs("item")
			%>
    <TR> 
      <td colspan="6" align="center" bgcolor="#FF66FF">Item :- <%=rs("item")%> 
        , Month :- <%=month(rs("l_date"))%></td>
    </TR>
    <%
		end if
%>
    <tr> 
      <td><%=rs("l_date")%></td>
      <td><%=rs("hardlimit")%></td>
      <td><%=rs("softlimit")%></td>
      <td><%=rs("multiplier")%></td>
      <td><%=rs("total")%></td>
      <td><%=rs("notes")%></td>
    </tr>
    <%
		totalHard = totalHard + rs("HardLimit")
		totalSoft = totalSoft + rs("SoftLimit")
		totalMultiplier = totalMultiplier + rs("Multiplier")
		totalTotal = totalTotal + rs("total")
		prev_month = month(rs("l_date"))
		rs.movenext
	wend
	if flag = true then
	%>
    <tr> 
      <td><font color="#FF0000"><strong>Total</strong></font></td>
      <td><strong><%=totalHard%></strong></td>
      <td><strong><%=totalSoft%></strong></td>
      <td><strong><%=totalMultiplier%></strong></td>
      <td><strong><%=totalTotal%></strong></td>
      <td>&nbsp;</td>
    </tr>
    <%	
	end if
end if
end if
%>
  </table>
</form>
</body>
</html>
<script>
<% if request("lstStrategy") <>""  then %>
   for (i=0;i<=frmTemp.lstStrategy.options.length-1;i++)
     if (frmTemp.lstStrategy.options[i].value=='<%=request("lstStrategy")%>')
	    frmTemp.lstStrategy.selectedIndex=i;
 <% end if %>
 

</script>
