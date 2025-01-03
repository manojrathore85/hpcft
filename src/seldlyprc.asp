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
dim rsSymbol
dim rsSubSymbol
set rsSymbol = server.CreateObject("adodb.recordset")
set rsSubSymbol = server.CreateObject("adodb.recordset")
rsSymbol.open "select distinct symbol from " &  varTblNameDailyData & "",con
rsSubSymbol.open "select distinct subsymbol from " &  varTblNameDailyData & "",con
%>
<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<h3 align="center">Select Daily Price</h3>
</head>
<body>
<center>
<form action="dailyprice.asp" method="post">
  <table width="50%" border="0" cellspacing="0" cellpadding="0">
    <tr> 
      <td width="27%" align="right">Start Date&nbsp;&nbsp;:</td>
      <td width="73%"><select name="smonth">
          <option value="01">Jan</option>
          <option value="02">Feb</option>
          <option value="03">Mar</option>
          <option value="04">Apr</option>
          <option value="05">May</option>
          <option value="06">Jun</option>
          <option value="07">Jul</option>
          <option value="08">Aug</option>
          <option value="09">Sep</option>
          <option value="10">Oct</option>
          <option value="11">Nov</option>
          <option value="12">Dec</option>
        </select> <input type="text" name="sday" size="2" maxlength="2"> <input type="text" name="syear" size="4" maxlength="4">
        Eg. Jan 1, 2003</td>
    </tr>
    <tr> 
      <td align="right">End Date&nbsp;&nbsp;:</td>
      <td><select name="emonth">
          <option value="01">Jan</option>
          <option value="02">Feb</option>
          <option value="03">Mar</option>
          <option value="04">Apr</option>
          <option value="05">May</option>
          <option value="06">Jun</option>
          <option value="07">Jul</option>
          <option value="08">Aug</option>
          <option value="09">Sep</option>
          <option value="10">Oct</option>
          <option value="11">Nov</option>
          <option value="12">Dec</option>
        </select> <input type="text" name="eday" size="2" maxlength="2"> <input type="text" name="eyear" size="4" maxlength="4"></td>
    </tr>
    <tr> 
      <td align="right">Symbol&nbsp;&nbsp;:</td>
      <td><select name="lstsymbol">
	  <%
	  while not rsSymbol.eof
	  %>
	  <option value="<%=rsSymbol("symbol")%>"><%=rsSymbol("symbol")%></option>
	  <%
	  rsSymbol.movenext
	  wend
	  rsSymbol.close
	  set rsSymbol = nothing
	  %>
	  </select></td>
    </tr>
    <tr> 
      <td align="right">Sub Symbol&nbsp;&nbsp;:</td>
      <td><select name="lstsubsymbol">
	 <%
	  while not rsSubSymbol.eof
	  %>
	  <option value="<%=rsSubSymbol("subsymbol")%>"><%=rsSubSymbol("subsymbol")%></option>
	  <%
	  rsSubSymbol.movenext
	  wend
	  rsSubSymbol.close
	  set rsSubSymbol = nothing
	  con.close
	  set con = nothing
	  %>
	  </select></td>
    </tr>
    <tr> 
      <td>&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
    <tr align="center"> 
      <td colspan="2"><input type="submit" value="Get Prices" name="cmdSubmit"></td>
    </tr>
  </table>
  </form>
</center>
</body>
</html>
