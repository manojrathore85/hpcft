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

dim rs
dim message
dim sortby
dim order
dim viewType
set rs = server.CreateObject("adodb.recordset")

if request("cmd") = "delete" then
	con.execute "delete from item_list_table where item ='" & request("item") & "'"
end if

if request("cmdSubmit") = "Update Item" then
	rs.open "select * from item_list_table where item ='" & request("txtItem") & "'",con
	if rs.eof then
		rs.close
		con.execute "update item_list_table set item = '" & request("txtItem") & "',strategy = '" & request("txtStrategy") & "' where item = '" & request("eitem") & "'"
		message = "Updation made successfully."
	else
		rs.close
		con.execute "update item_list_table set strategy = '" & request("txtStrategy") & "' where item = '" & request("eitem") & "'"
		message = "Strategy Updated"
	end if
end if

if request("cmdSubmit") = "Add Item" then
	rs.open "select * from item_list_table where item ='" & request("txtItem") & "'",con
	if rs.eof then
		rs.close
		con.execute "insert into item_list_table values('" & request("txtItem") & "','','" & request("txtStrategy") & "')"
	else
		rs.close
		message = "Entry already present"
	end if
end if
rs.open "select item,strategy from item_list_table where item <> ''",con
%>
<html>
<head>
<title>Add-Edit Item</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<LINK href="include/style.css" type="text/css" rel="stylesheet">
</head>
<body topmargin="0">
<!-- #include file="header.inc" -->
<h2 align="center">Item Edit Section</h2>
<p align="center"><font color="red"><b><%=message%></b></font></p>
<table width="100%" border="0" cellspacing="0" cellpadding="0" bordercolor="#000000" >
  <tr> 
    <td width="30%" ><strong>Item</strong></td>
	<td width="29%" ><strong>Stratey</strong></td>
    <td width="19%"><strong>Delete</strong></td>
    <td width="22%"><strong>Edit</strong></td>
  </tr>
  <tr> 
    <td colspan="4"><hr color="#000000" noshade size="0"></td>
  </tr>
  <%
if not rs.eof then
	while not rs.eof
	%>
  <tr> 
    <td><%=rs("Item")%></td>
	<td><%=rs("strategy")%></td>
    <td><a href="addedititem.asp?cmd=delete&item=<%=rs("item")%>">Delete</a></td>
    <td><a href="#" onClick="edititem('<%=rs("item")%>','<%=rs("strategy")%>');">Edit</a></td>
  </tr>
  <tr> 
    <td colspan="4"><hr color="#000000" noshade size="0"></td>
  </tr>
<%
	 rs.movenext
	 wend
else
%>
  <tr> 
    <td align="center" colspan="4">No records found currently</td>
  </tr>
  <%
end if
 %>
  <form action="AddEditItem.asp" method="post" name="frmEdit">
  <input type="hidden" value="" name="eitem">
    <tr> 
      <td><input type="text" name="txtItem" class="formTextbox"></td>
	  <td><input type="text" name="txtStrategy" class="formTextbox"></td>
      <td align="left">
<input type="submit" name="cmdSubmit" value="Add Item" class="formbutton"></td>
	  <td align="left">
<input type="button" name="cmdCancel" value="Cancel" class="formbutton" style="visibility:hidden" onClick="cancel();"></td>
    </tr>
  </form>
</table>

</body>
</html>
<script>

function edititem(eitem,strategy)
{
frmEdit.eitem.value = eitem;
frmEdit.txtItem.value = eitem;
frmEdit.txtStrategy.value = strategy;
frmEdit.cmdSubmit.value = 'Update Item';
frmEdit.cmdCancel.style.visibility = 'visible';
frmEdit.txtItem.focus();
}

function cancel()
{
frmEdit.txtItem.value = "";
frmEdit.txtStrategy.value = "";
frmEdit.cmdSubmit.value = 'Add Item';
frmEdit.cmdCancel.style.visibility = 'hidden';
}

</script>
<%
if rs.state =1 then rs.close
set rs = nothing
con.close
set con = nothing
%>