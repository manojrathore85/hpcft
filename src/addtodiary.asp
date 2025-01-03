<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!-- #include file="connect.asp" -->
<%
dim message
dim java_script
sub addtodiary()
	dim sql
	sql = "insert into " & varTblNameIms_MyDiary_table & "(email,projectid,issueid,date_of_creation,comments,extra_field) values('" & session("user") & "'," & request("prjid") & "," & request("issueid") & ",now(),'" & request("comments") & "','')"
	if entry_exist = false then 
		con.execute sql
		message = "Issue added to your diary"
	else
		message = "Issue already in your diary"
	end if
	con.close
	set con = nothing
end sub

sub remove_from_diary()
	dim sql
	sql = "delete from " & varTblNameIms_MyDiary_table & " where email ='" & session("user") & "' and projectid=" & request("prjid") & " and issueid=" & request("issueid")
	if entry_exist = true then 
		con.execute sql
		message = "Issue removed from your diary"
		java_script = "opener.location.reload();"
	else
		message = "Issue does not exist in your diary"
	end if
	con.close
	set con = nothing
end sub

function entry_exist
	entry_exist = false
	dim rs
	dim sql
	set rs =server.CreateObject("adodb.recordset")
	sql = "select 1 from " & varTblNameIms_MyDiary_table & " where email ='" & session("user") & "' and projectid=" & request("prjid") & " and issueid=" & request("issueid")
	'response.Write(sql)
	'response.End()
	rs.open  sql ,con
	if not rs.eof then
		entry_exist = true
	end if
	rs.close
	set rs = nothing
end function

if request("mode") = "remove" then
	call remove_from_diary
elseif request("mode") = "add" then
	call addtodiary
end if
%>
<script>
<%=java_script%>
alert('<%=message%>');
self.close();
</script>