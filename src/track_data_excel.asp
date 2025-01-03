<!-- #include file="checksession.asp" -->
<!-- #include file ="generalfunctions.asp" -->
<!-- #include file ="connect.asp" -->

<%
 Response.ContentType = "application/vnd.ms-excel"
mymonth = month(date)
if(len(mymonth) = 1) then mymonth = "0" & mymonth
myday = day(date)
if(len(myday) = 1) then myday = "0" & myday

Response.AddHeader "Content-Disposition", "attachment; filename=" & getSummary & "-" & request("issueid") & "-" & request("prjid") & "-" & year(date) & mymonth & myday & ".xls"

if request("issueid") = "" then
	response.write("Issueid missing for Tacking System")
	response.end
end if
dim permRead,permWrite,permAdd,permDelete
dim rsglobal
set rsglobal = server.CreateObject("adodb.recordset")
call checkPermission

'call showdata
%>

<%
function getSummary()
	dim rs
	set rs = server.CreateObject("adodb.recordset")
	rs.open "select summary from issues where issueid = " & request("issueid"),con
	if not rs.eof then
		getSummary = rs(0)
		'getSummary = replace(getSummary,chr(32),"-")
	end if
	rs.close
	set rs = nothing
end function
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
	response.Write("<table width='100%' border='1' cellpadding='0' cellspacing='0' bgcolor=""#CCCCCC""><tr>")'<td>")
	'response.Write("<table width='100%' border='0' cellpadding='1' cellspacing='1' id='data_table'><tr class='head'>")
	
	' print headers of the data table
	for each key in my_fields
		if key <> "col_rowid" and key <> "issueid"  then response.Write("<td>" & key & "</td>") 'not showing id column of table
	next	
	
	response.Write("</tr>")
	
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
		
			end if
			
		next
		
		'response.Write("<td>" & edit_delete_link & "</td></tr>")
		
		'response.Write(columns(0,i) & VBTAB & columns(1,i) & VBTAB & columns(2,i) & vbcrlf)

	next
	response.Write("</table>")'</td></tr></table>")
	
end sub
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<title>Untitled Document</title>

</head>

<body>
<table width="100%" border="0">
					  <tr id="showbody">
						<td><%showdata%></td>
					  </tr></table>
</body>
</html>
<script>
//window.close();
</script>