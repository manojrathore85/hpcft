<!-- #include file="checksession.asp" -->
<!-- #include file="connect.asp" --> <!-- including the connection file -->

<%
' ********************* CLIENTS CANNOT CREATE CLIENT ***************************************************

if varSpecialclientid <> "" then ' THE VARIABLE 'varSpecialclientid' IS CONTAINED IN CONNECT.ASP. VARIABLE IS NOT EMPTY WHEN ASP SCRIPTS ARE ACCESSED BY CLIENT
	response.Write("<script>alert('Clients are not allowed to update client.');window.history.go(-1);</script>")
	con.close
	set con = nothing
	response.End()
end if

'********************************************************************************************************
' ********************************************** CHECKING PERMISSION ************************************

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
' ********************* CHECKING PERMISSION ENDS*********************************************************
dim message
message = "Update the selected client."
if request("cmdSubmit") = "Update" and request("lstClients") <> "" then
	set fs = server.CreateObject("scripting.filesystemobject")
	if not fs.folderexists(server.MapPath(".") & "\clients\old") then fs.createfolder(server.MapPath(".") & "\clients\old")
	if not fs.folderexists(server.MapPath(".") & "\clients\old\" & request("lstClients")) then fs.createfolder(server.MapPath(".") & "\clients\old\" & request("lstClients"))
	Dim ffile
	Dim ffolder
	Dim f1folder
	Set ffolder = fs.GetFolder(server.MapPath(".") & "\clients\" & request("lstClients"))
	For Each ffile In ffolder.Files
		fs.copyfile ffile,server.MapPath(".") & "\clients\old\" & request("lstClients") & "\",true
	Next

	fs.copyfolder server.MapPath(".") & "\clients\" & request("lstClients") & "\monthlystattool",server.MapPath(".") & "\clients\old\" & request("lstClients") & "\monthlystattool",true
	fs.copyfolder server.MapPath(".") & "\clients\" & request("lstClients") & "\GUI_files",server.MapPath(".") & "\clients\old\" & request("lstClients") & "\GUI_files",true
	fs.copyfolder server.MapPath(".") & "\clients\" & request("lstClients") & "\include",server.MapPath(".") & "\clients\old\" & request("lstClients") & "\include",true
		
	Set ffolder = fs.GetFolder(server.MapPath("."))
	For Each ffile In ffolder.Files
		fs.copyfile ffile,server.MapPath(".") & "\clients\" & request("lstClients") & "\",true
	Next
	
	fs.copyfolder server.MapPath(".") & "\monthlystattool",server.MapPath(".") & "\clients\" & request("lstClients") & "\monthlystattool",true
	fs.copyfolder server.MapPath(".")  & "\GUI_files",server.MapPath(".") & "\clients\" & request("lstClients") & "\GUI_files",true
	fs.copyfolder server.MapPath(".") & "\include",server.MapPath(".") & "\clients\" & request("lstClients") & "\include",true
	
	message = "Clients updated successfully."
end if

%>


<%
     dim i ' as incremental variable
   dim objview
   set objview =server.CreateObject("adodb.recordset") ' creating the recordset 
   objview.open "select * from clientids order by createdate desc ",con
%>
<form action="viewprojects.asp" name="frmUpdateClient">
<input type="hidden" value="updateClient" name="mainlink">
<table width="95%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td class="SectionHead">Administration<br>
      &nbsp; </td>
  </tr>
  <tr> 
    <td> 
	<br>
      ----Update Client : <select name="lstClients">
	  <% 
    if objview.bof=false and objview.eof =false then ' checking whether the record is exist or not
	   while not objview.eof ' 
%>	 
	<option value="<%=objview("Clientid")%>"><%=objview("Clientid")%></option>

   <%
	    objview.movenext 
		wend
	end if
	%>
	  </select><!-- <input type="text" name="txtClient"> --> <input type="submit" name="cmdSubmit" value="Update" class="formbutton">
  	</td>
  </tr>
</table>
</form>
<BR>
<table width="95%" border="0" cellpadding="0" cellspacing="0" bgcolor="#CCCCCC">
<tr>
    <td align="center"><p align="center" class="redbold"><%=message%></p></td>
</tr>
</table>
<%
  'nullifying all the objects
  objview.close
  set objview= nothing
  con.close
  set con = nothing
 %> 
