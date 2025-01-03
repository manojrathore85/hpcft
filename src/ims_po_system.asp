<%'@Transaction = "Required"%>
<!-- #include file="checksession.asp" -->
<!-- #include file ="generalfunctions.asp" -->
<!-- #include file ="connect.asp" -->

<%
if request("issueid") = "" then
	response.write("Issueid missing for Tracking System")
	response.end
end if
dim permRead,permWrite,permAdd,permDelete
dim rsglobal
set rsglobal = server.CreateObject("adodb.recordset")
call checkPermission

'response.Write("Read = " & permRead & "Write = " & permWrite & "Add = " & permAdd & "Delete = " &permDelete )

if request("cmd_submit") = "Save" or request("cmd_submit") = "Update" then
	call save_data
end if

if request("mode") = "delete" then
	con.execute "delete from " & varTblNameIms_po_system_table & " where col_rowid=" & request("id")
end if

if request("mode") = "delCol" then
	deleteColumn()
end if

sub deleteColumn()
	dim query
	dim columns
	dim lr,fr,i
	dim fieldid
	fieldid = getFieldId(request("colname"))
	columns = array()
	query = "select distinct col_rowid from "& varTblNameIms_po_system_table & " where " & _
			"col_field_id =" & getFieldId("issueid") & " and col_column_value = '" & request("issueid") & "'"
	columns = returndata(query)
	fr = lbound(columns,2)
	lr = ubound(columns,2)
	for i = fr to lr	
		con.execute "delete from "& varTblNameIms_po_system_table & "  where col_rowid=" & columns(0,i) & " and col_field_id =" & fieldid
	next
end sub
' ********************* CHECKING PERMISSION ***************************
sub checkPermission()
	dim rsperm
	set rsperm = server.CreateObject("adodb.recordset")
	rsperm.open "select perm.* from " &  varTblNamePermissions & " perm, " &  varTblNameUsers & " u, " &  varTblNameissues & _
	 " i where i.projectid = u.projectid and u.permissionid = perm.permissionid and i.issueid = " &  request("issueid") & _
	 " and u.email ='" &  session("user") & "'",con
	if not rsperm.eof then
		if rsperm("pread") = "T" then
			permRead = true
			if rsperm("pwrite") = "T" then permWrite = true
			if rsperm("padd") = "T" then permAdd = true
			if rsperm("pdelete") = "T" then permDelete = true
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

sub save_data
	dim query
	dim rs
	dim btn_submit
	dim maxrowid
	set rs = server.CreateObject("adodb.recordset")
	'query = "select distinct col_column_name from " & varTblNameIms_po_system_table & ""
	
	query = "select distinct f.col_field_name,f.col_field_id from " & varTblNameIms_po_system_table & " i,"& _
			" " & varTblNameIms_po_system_table & " j, " & varTblNameIms_po_fields_table & " f" & _
			" where" & _
			" i.col_rowid = j.col_rowid and j.col_field_id =" & getFieldId("issueid") & " and" & _
			" j.col_column_value = '" & request("issueid") & "' and f.col_field_id = i.col_field_id"
	
	maxrowid = max_id("select if (max(col_rowid) is null ,1 ,max(col_rowid) + 1) from " & varTblNameIms_po_system_table & "")

	on error resume next
	
	
	
	rs.open query,con,3,2

	if not rs.eof then
		'query = "insert into " & varTblNameIms_po_system_table & " values(" & maxrowid & ",'issueid','" & replace(request(rs(0)),"'","''") & "')"
		'con.execute query
		while not rs.eof
			query = "insert into " & varTblNameIms_po_system_table & " values(" & maxrowid & ",'" & rs(1) & "','" & replace(request(rs(0)),"'","''") & "')"
			con.execute query
			rs.movenext
		wend
	else
			query = "insert into " & varTblNameIms_po_system_table & " values(" & maxrowid & ",'" & getFieldId("issueid") & "','" & replace(request("issueid"),"'","''") & "')"
			con.execute query
	end if
	
	if request("cmd_submit") = "Update" then
		con.execute "delete from " & varTblNameIms_po_system_table & " where col_rowid=" & request("id")
	end if
	
	'err.clear
	' adding dynamcally added column and its value
	for each key in request.Form
		if instr(1,key,"col_name") then
			arr = split(key,"_")
			if request(key) <> "" then
				query = "insert into " & varTblNameIms_po_system_table & " values(" & maxrowid & "," & addField(request(key)) & ",'" & replace(request("col_value_" & arr(ubound(arr))),"'","''") & "')"
				con.execute query
			end if	
		end if
	next
	
	msg = "Data saved, screened duplicate records(if any)."
	
	rs.close
	set rs = nothing
	
end sub

function max_id(query)

	dim rs
	set rs = server.CreateObject("adodb.recordset")
	rs.open query,con
	if not (rs.eof or rs.bof) then
		max_id = rs(0)
	else
		max_id = 1
	end if

	rs.close
	set rs = nothing

end function

sub showdata()
	on error resume next
	dim query
	dim columns
	dim fr
	dim lr
	dim i
	dim rowid
	dim my_fields
	dim edit_delete_link
	set my_fields = server.createobject("scripting.dictionary")
	columns = array()
	'query = "select distinct col_column_name from " & varTblNameIms_po_system_table & ""
	
	'query = "select distinct i.col_column_name from " & varTblNameIms_po_system_table & " i,"& _
	'	" " & varTblNameIms_po_system_table & " j where" & _
	'	" i.col_rowid = j.col_rowid and j.col_column_name = 'issueid' and" & _
	'	" j.col_column_value = '" & request("issueid") & "'"
	
	query = "select distinct f.col_field_id,f.col_field_name from " & varTblNameIms_po_system_table & " i,"& _
			" " & varTblNameIms_po_system_table & " j, " & varTblNameIms_po_fields_table & " f" & _
			" where" & _
			" i.col_rowid = j.col_rowid and j.col_field_id =" & getFieldId("issueid") & " and" & _
			" j.col_column_value = '" & request("issueid") & "' and f.col_field_id = i.col_field_id"
			
	columns = returndata(query)
	
	fr = lbound(columns,2)
	lr = ubound(columns,2)	
	
	' CREATING QUERY TO SHOW DISTINCT COLUMNS IN COLUMN OF TABLE
	
	query = "SELECT i.col_rowid,"
	
		
	for i = fr to lr
		
		'response.Write(columns(0,i) & VBTAB) ' printing headers of columns
		
		if i = lr then 
			query = query & "max(CASE WHEN i.col_field_id =" & columns(0,i) & " THEN i.col_column_value END) AS `" & columns(1,i) & "`" 
		else
			query = query & "max(CASE WHEN i.col_field_id =" & columns(0,i) & " THEN i.col_column_value END) AS `" & columns(1,i) & "`,"
		end if
		
	next

	query = query & " FROM " & varTblNameIms_po_system_table & " i," & _
	" " & varTblNameIms_po_system_table & " j where i.col_rowid = j.col_rowid and " & _
	" j.col_field_id =" & getFieldId("issueid") & " and j.col_column_value = '" & request("issueid") & "' group BY i.col_rowid"

	'response.Write("<br>")
	
	columns = returndataandfields(query,my_fields)
	
	'my_fields.remove("col_rowid") ' removing rowid from dictionary, so that it doesn't get prints in table
	
	fr = lbound(columns,2)
	lr = ubound(columns,2)
	
	'print headers of the table
	response.Write("<table width='100%' border='0' cellpadding='0' cellspacing='0' bgcolor=""#CCCCCC""><tr><td>")
	response.Write("<table width='100%' border='0' cellpadding='1' cellspacing='1' id='data_table'><tr class='head'>")
	
	' print headers of the data table
	for each key in my_fields
		if key <> "col_rowid" and key <> "issueid"  then response.Write("<td>" & key & "</td>") 'not showing id column of table
	next	
	
	response.Write("<td>Command</td></tr>")
	
	for i = fr to lr

		response.Write("<tr bgcolor='#FFFFFF'>")

		for each key in my_fields
			
			if key <> "col_rowid" and key <> "issueid" then 
				response.Write("<td>") 
				if isarray(columns(my_fields(key),i)) then
					response.binaryWrite(columns(my_fields(key),i))
				else
					response.Write(columns(my_fields(key),i))
				end if 
				response.Write("&nbsp;</td>")
			else
				if permWrite = true then
					edit_delete_link = "<a href=ims_po_system.asp?mode=edit&id=" & columns(my_fields("col_rowid"),i) & "&issueid=" & request("issueid")
					'if isarray(columns(my_fields("issueid"),i)) then
					'	edit_delete_link = edit_delete_link & SimpleBinaryToString(columns(my_fields(key),i))
					'else
					'	edit_delete_link = edit_delete_link & columns(my_fields("issueid"),i) 
					'end if
					'iif isarray(columns(my_fields("issueid"),i)) , edit_delete_link = edit_delete_link &  SimpleBinaryToString(columns(my_fields("issueid"),i)), edit_delete_link = edit_delete_link & columns(my_fields("issueid"),i) '& _ 
					edit_delete_link =  edit_delete_link & ">Edit</a> "
				else
					edit_delete_link = "Edit "
				end if
				if permDelete = true then
					edit_delete_link = edit_delete_link & "| <a href=ims_po_system.asp?mode=delete&id=" & columns(my_fields("col_rowid"),i) & "&issueid=" & request("issueid")
					'if isarray(columns(my_fields("issueid"),i)) then
					'	edit_delete_link = edit_delete_link & SimpleBinaryToString(columns(my_fields(key),i))
					'else
					'	edit_delete_link = edit_delete_link & columns(my_fields("issueid"),i) 
					'end if
					'iif isarray(columns(my_fields("issueid"),i)) , edit_delete_link = edit_delete_link &  SimpleBinaryToString(columns(my_fields("issueid"),i)), edit_delete_link = edit_delete_link & columns(my_fields("issueid"),i) 
					edit_delete_link = edit_delete_link &  " onClick=""javascript:return confirm('Do you really want to delete this entry');"">Delete</a>"
				else
					edit_delete_link = edit_delete_link & "| Delete "
				end if
'				edit_delete_link = "<a href=ims_po_system.asp?mode=edit&id=" & columns(my_fields("col_rowid"),i) & "&issueid=" & columns(my_fields("issueid"),i) &">Edit</a> | <a href=ims_po_system.asp?mode=delete&id=" & columns(my_fields("col_rowid"),i) & "&issueid=" & columns(my_fields("issueid"),i) & " onClick=""javascript:return confirm('Do you really want to delete this entry');"">Delete</a>"
			end if
			
		next
		
		response.Write("<td>" & edit_delete_link & "</td></tr>")
		
		'response.Write(columns(0,i) & VBTAB & columns(1,i) & VBTAB & columns(2,i) & vbcrlf)

	next
	response.Write("</table></td></tr></table>")
	
end sub
sub getdata()
	'on error resume next
	dim query
	dim rs
	dim btn_submit
	dim i
	i = 0
	set rs = server.CreateObject("adodb.recordset")
	'query = "select distinct col_column_name from " & varTblNameIms_po_system_table & ""
	
	'query = "select distinct i.col_column_name from " & varTblNameIms_po_system_table & " i,"& _
	'	" " & varTblNameIms_po_system_table & " j where" & _
	'	" i.col_rowid = j.col_rowid and j.col_column_name = 'issueid' and" & _
	'	" j.col_column_value = '" & request("issueid") & "'"
	
	query = "select distinct f.col_field_id,f.col_field_name from " & varTblNameIms_po_system_table & " i,"& _
			" " & varTblNameIms_po_system_table & " j, " & varTblNameIms_po_fields_table & " f" & _
			" where" & _
			" i.col_rowid = j.col_rowid and j.col_field_id =" & getFieldId("issueid") & " and" & _
			" j.col_column_value = '" & request("issueid") & "' and f.col_field_id = i.col_field_id"

	response.Write("<div id='mydiv'><table width='100%' border='1'>")

	if request("mode") = "edit" and request("id") <> "" then
		response.Write("<input type='hidden' name='id' value='" & request("id") & "'>")
	end if	

	rs.open query,con
	if not rs.eof then
		response.Write("<tr class=head><td>Hide/Display Column</td><td>Column Name<td>Column Value</td><td>Delete Column</td></tr>")
		while not rs.eof
		
			if rs(1) <> "issueid" then ' we don't want to show issueid column
				if request("mode") = "edit" and request("id") <> "" then
					response.Write("<tr><td><input type='checkbox' onclick=""show_hide_column(" & i & ", this.checked);""></td><td>" & rs(1) & "&nbsp;</td><td><input type='text' name='" & rs(1) & "' value='" & getvalue(rs(0)) & "'></td>")
					' show delete column link only when user has permission
					if permAdd = true and permDelete =true then 
						response.write("<td><a href='ims_po_system.asp?mode=delCol&colname=" & rs(1) & "&issueid=" & request("issueid") & "' onClick=""javascript:return confirm('Do you really want to delete this column');"">Delete Column</td>")
					end if
					response.write("</tr>")
				else
					response.Write("<tr><td><input type='checkbox' onclick=""show_hide_column(" & i & ", this.checked);""></td><td>" & rs(1) & "&nbsp;</td><td><input type='text' name='" & rs(1) & "'></td>")
					if permAdd = true and permDelete =true then
						response.write("<td><a href='ims_po_system.asp?mode=delCol&colname=" & rs(1) & "&issueid=" & request("issueid") & "' onClick=""javascript:return confirm('Do you really want to delete this column');"">Delete Column</td>")
					end if
					response.write("</tr>")
'					<td><a href='ims_po_system.asp?mode=delCol&colname=" & rs(0) & "&issueid=" & request("issueid") & "' onClick=""javascript:return confirm('Do you really want to delete this column');"">Delete Column</td></tr>")
				end if
			i = i + 1
			end if
			rs.movenext
			
		wend
	end if
	response.Write("</table></div>")
	
	if permWrite <> true then
		btnSaveStatus = "disabled"
	end if
	if permAdd <> true then
		btnAddColumn = "disabled"
	end if

	if request("mode") = "" or request("mode") = "delCol" or request("mode") = "delete"then
		btn_submit = "<input type='submit' name='cmd_submit' value='Save' " & btnSaveStatus & " >"
	else
		btn_submit = "<input type='submit' name='cmd_submit' value='Update' " & btnSaveStatus & " >"
	end if
	
	response.Write("<table><tr><td colspan=2 align=center><input type='button' name='btn_add' value='Add column and value' " & btnAddColumn & " onclick='addrow();'/>&nbsp;&nbsp;&nbsp;" & btn_submit & "&nbsp;&nbsp;&nbsp;<input type=reset name=cmd_reset value='Cancel' onclick=""window.location='ims_po_system.asp?issueid=" & request("issueid") & "';""></td></tr></table>")'<input type=reset name=cmd_reset value='Cancel' onclick='frm.cmd_submit.value=""Save""';frm.reset();'>
	
	rs.close
	set rs = nothing

end sub

function addField(fieldname)
	dim query
	dim fldid
	fldid = getFieldId(fieldname)
	if fldid <> "" then 
		addField = fldid ' retrun fieldid already in varTblNameIms_po_fields_table for the fieldname
	else 'select maxid + 1 and insert the new field in table
		rsglobal.open "select if (max(col_field_id) is null ,1 ,max(col_field_id) + 1) from " & varTblNameIms_po_fields_table	,con
		addField = rsglobal(0)
		con.execute "insert into " & varTblNameIms_po_fields_table & " values(" & rsglobal(0) & ",'" & fieldname & "')"
	end if
	if rsglobal.state = 1 then rsglobal.close
end function

function getFieldId(fieldname)
	on error resume next
	dim query
	if rsglobal.state = 1 then rsglobal.close
	query = "select col_field_id from " & varTblNameIms_po_fields_table & " where col_field_name='"  & fieldname & "'"
	rsglobal.open query,con,3,2
	getFieldId = rsglobal(0)
	rsglobal.close
end function
function getvalue(fieldname)
	on error resume next
	dim query
	if rsglobal.state = 1 then rsglobal.close
	query = "select col_column_value from " & varTblNameIms_po_system_table & " where col_rowid=" & request("id") & " and col_field_id=" & fieldname 
	rsglobal.open query,con,3,2
	getvalue = rsglobal(0)
	rsglobal.close
end function

Function SimpleBinaryToString(Binary)
  'SimpleBinaryToString converts binary data (VT_UI1 | VT_ARRAY Or MultiByte string)
  'to a string (BSTR) using MultiByte VBS functions
  Dim I, S
  For I = 1 To LenB(Binary)
    S = S & Chr(AscB(MidB(Binary, I, 1)))
  Next
  SimpleBinaryToString = S
End Function


%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<title>Tracking System</title>
<script>
var counter = 0;
var flag=0;
function addrow()
{
	
	mydiv.innerHTML = mydiv.innerHTML + 
	'<table width="100%" border="1">' + ((flag == 0) ? '<tr><td>Column Name</td><td>Column Value</td></tr>' : '') +
	'<tr>' +
	'<td><input type=text name=col_name_' + counter + '></td>' +
	'<td><input type=text name=col_value_' + counter + '></td>' + 
	'</tr></table>';
	flag = 1;
	counter++;
}

function show_hide_column(col_no, do_show) {
//	alert(col_no + ' - ' +  do_show);

    var stl;
    if (!do_show) stl = 'block'
    else         stl = 'none';

    var tbl  = document.getElementById('data_table');
    var rows = tbl.getElementsByTagName('tr');

    for (var row=0; row<rows.length;row++) {
      var cels = rows[row].getElementsByTagName('td')
      cels[col_no].style.display=stl;
    }
  }

</script>
<LINK 
href="include/style.css" type=text/css rel=stylesheet>
<style type="text/css">
</style>
<script>
self.moveTo(0,0);
</script>
</head>
<body bgcolor="" leftmargin="0" topmargin="0">
<table width="100%" border="0" cellspacing="0" cellpadding="0" height="100%" bgcolor="#FFFFFF">
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
                  <td valign="top"> <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
                      <tr> 
                        <td width="66%">&nbsp;</td>
                        <td width="34%" align="left"><!--include file="include/TopRightNavBar.asp" --></td>
                      </tr>
                      <tr> 
						<td colspan="2" align="Center"><h4>Tracking System</h3><!-- include file="include/GeneralTopNavBar.asp" --></td>
                      </tr>
                    </table>
                    <font color="#FFFFFF">-</font><br>
<!-- 
#################################################         
		IMS PURCHASE ORDER SYSTEM 
#################################################
-->		<p align="left"><a href="track_data_excel.asp?issueid=<%=request("issueid")%>&prjid=<%=request("prjid")%>" target="_blank">Save to excel</a>
					<table width="100%" border="0">
					  <tr id="showbody">
						<td><%showdata%></td>
					  </tr>
					  <form name="frm" method="post" action="ims_po_system.asp">
					  <input type="hidden" value="<%=request("issueid")%>" name="issueid" />
					  <tr id="addtext">
						<td>&nbsp;
						  <%getdata%></td>
					  </tr>
					  </form>
					</table>					
<!-- 
#################################################         
		END IMS PURCHASE ORDER SYSTEM 
#################################################
-->					
                  </td>
                </tr>
              </table>
              
            </div></td>
        </tr>
        <tr> 
          <td width="45%">&nbsp;</td>
          <td width="55%" align="center" valign="middle">&nbsp;</td>
        </tr>
        <tr> 
          <td>&nbsp;</td>
          <td align="center" valign="middle">&nbsp;</td>
        </tr>
      </table>
   </td>
  </tr>
</table>
</body>
</html>