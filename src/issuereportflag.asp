<!-- #include file="connect.asp" -->
<%
dim rs
dim tempStr
dim reportUsers
dim newReportUsers
dim flag
dim i
i = 0

if request("Flag") <> "" then
	dim rschk
	set rschk = server.CreateObject("adodb.recordset")
	rschk.open "select reportUsers from " &  varTblNameIssues & " where issueid =" & request("issueid"),con
	if not rschk.eof then
		if request("Flag") = "true" then
			if isnull(rschk(0)) = false and isempty(rschk(0)) = false then
				rschk.movefirst
				con.execute "update " &  varTblNameIssues & " set reportUsers ='" & rschk(0) & ","  & session("user") & "' where issueid =" & request("issueid")
			else
				con.execute "update " &  varTblNameIssues & " set reportUsers ='" & session("user") & "' where issueid =" & request("issueid")
			end if
		else
			if isnull(rschk("reportUsers")) = false and isempty(rschk("reportUsers")) = false then
			rschk.movefirst
				reportUsers = split(rschk("reportUsers"),",")
			end if
			if ubound(reportUsers) >= 0 then
				for i = 0 to ubound(reportUsers)
					if reportUsers(i) = session("user") then
						' do nothing
					else
						if i = ubound(reportUsers) then
							newReportUsers = newReportUsers & reportUsers(i)
						else
							newReportUsers = newReportUsers & reportUsers(i) & ","
						end if
					end if
				next
		 	end if
			if newReportUsers <> "" then
				 If Mid(newReportUsers, Len(newReportUsers), Len(newReportUsers)) = "," Then
					newReportUsers = Mid(newReportUsers, 1, Len(newReportUsers) - 1)
				 End If
			end if
			con.execute "update " &  varTblNameIssues & " set reportUsers ='" & newReportUsers & "' where issueid =" & request("issueid")
		end if	
	end if
	rschk.close
	set rschk = nothing
	con.close
	set con = nothing
	response.Redirect("issuedetails.asp?issueid=" & request("issueid"))
else
	set rs = server.CreateObject("adodb.recordset")
	rs.open "select reportUsers from " &  varTblNameIssues & " where issueid =" & request("issueid"),con
	if not rs.eof then
		if isnull(rs(0)) = False and isEmpty(rs(0)) = False then
			rs.movefirst
			'response.Write(rs("reportUsers") & "asdf")
			tempStr = rs("reportUsers")
			reportUsers = split(tempStr,",")
			for i = 0 to ubound(reportUsers)
				if reportUsers(i) = session("user") then
					flag = true
					i = ubound(reportUsers)
				end if
			next
		end if
	end if
	rs.close
	set rs = nothing
	con.close
	set con = nothing
	if flag = true then
		response.Write("<input type=checkbox name=chkFlag checked onclick='setflag();'>")
	else
		response.Write("<input type=checkbox name=chkFlag onclick='setflag();'>")
	end if
end if	
%>