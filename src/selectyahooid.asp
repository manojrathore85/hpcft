<script>
function check()
{
	if(temp.yahooid.value == '')
	{
		alert('Enter the yahoo-id to be created');
		return false;
	}
}
</script>
<table width="95%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td class="SectionHead">Administration<br>
      &nbsp; </td>
  </tr>
  <tr> 
    <td>Enter the Yahoo-id you want for the user :- <b><%=firstname%>&nbsp;<%=lastname%></b><br>
      <br>
    </td>
  </tr>
</table>
<BR>
  <form name="temp" onSubmit="return check();" action="createyahooid.asp" target="_blank">
  <input type="hidden" name="mainlink" value="<%=request("mainlink")%>">
  <input type="hidden" name="id" value="<%=request("id")%>">
<table width="95%" border="0" cellpadding="0" cellspacing="0" bgcolor="#ffffff">
<tr>
   <td align="left"><input type="text" name="yahooid" size="40">&nbsp;&nbsp;&nbsp;&nbsp;<input type="Submit" value="Create Yahoo Id"></td>
</tr>
</table>
</form>

