<%
'************************ FUNCTION TO SET THE TITLE ********************************
function setTitle
	dim strPath
	dim pos,pos1
	dim titlePrefix
	titlePrefix = "GTSIMS"
	strPath = Request.ServerVariables("SCRIPT_NAME")
	pos = instrrev(strpath,".asp")
	pos1 = instrRev(strpath,"/")
	strPath = Mid(strPath, pos1 + 1 ,pos - pos1) & "asp"

	Select case strPath
		case "addClient.asp"
			setTitle = titlePrefix & ":Administration:Add New Client"
		case "addPermission.asp"
			setTitle = titlePrefix & ":Administration:Add New Permission"
		case "addProject.asp"
			setTitle = titlePrefix & ":Administration:Add New Project"
		case "addsetPermission.asp"
			setTitle = titlePrefix & ":Administration:Set Users Permission"
		case "adduser.asp"
			setTitle = titlePrefix & ":Administration:Add New User"
		case "addwatcher.asp"
			setTitle = titlePrefix & ":Administration:Add Project Watcher"
		case "changepass.asp"
			setTitle = titlePrefix & ":Change My Password"
		case "changepassall.asp"
			setTitle = titlePrefix & ":Change User Password"
		case "corp_action.asp"
			setTitle = titlePrefix & ":Corporate Action"
		case "createissue.asp"
			setTitle = titlePrefix & ":Create New Issue"
		case "dailyprice.asp"
			setTitle = titlePrefix & ":Daily Price"
		case "editissue.asp"
			setTitle = titlePrefix & ":Edit issue"
		case "editpermission.asp"
			setTitle = titlePrefix & ":Administration:Edit Permission"
		case "editprofile.asp"
			setTitle = titlePrefix & ":Edit Profile"
		case "editProject.asp"
			setTitle = titlePrefix & ":Administration:Edit Project"
		case "editsetpermission.asp"
			setTitle = titlePrefix & ":Administration:Edit User Permission"
		case "edituser.asp"
			setTitle = titlePrefix & ":Administration:Edit User Profile"
		case "findissue.asp"
			'setTitle = titlePrefix & ":" & getProjectName & ":Find Issue"  'comment for cr20070101 
			setTitle = titlePrefix & ":Find Issue"
		case "ims_compliance.asp"
			setTitle = titlePrefix & ":Compliance Page"
		case "ims_index.asp"
			setTitle = titlePrefix & ":" & getProjectName & ":Project Details"
		case "issuedetails.asp"
			setTitle = titlePrefix & ":Issue Details"
		case "issuenavigator.asp"
			setTitle = titlePrefix & ":All Issues"
		case "mlp_compliance.asp"
			setTitle = titlePrefix & ":MLP Compliance"
		case "permissionlist.asp"
			setTitle = titlePrefix & ":Administration:List Permission"
		case "projectlist.asp"
			setTitle = titlePrefix & ":Administration:List Projects"
		case "projectwatchlist.asp"
			setTitle = titlePrefix & ":Administration:List Project Watcher"
		case "restoreclient.asp"
			setTitle = titlePrefix & ":Administration:Client Restore"
		case "seeusage.asp"
			setTitle = titlePrefix & ":Administration:Usage Details"
		case "seldlyprc.asp"
			setTitle = titlePrefix & ":Daily Price"
		case "selectProject.asp"
			setTitle = titlePrefix & ":Select Project"
		case "setPermission.asp"
			setTitle = titlePrefix & ":Adminitration:Set Permission"
		case "setPermissionlist.asp"
			setTitle = titlePrefix & ":Administration:List Permission Used"
		case "updateclient.asp"
			setTitle = titlePrefix & ":Administration:Update Client"
		case "userlist.asp"
			setTitle = titlePrefix & ":Administration:List Users"
		case "viewproject.asp"
			setTitle = titlePrefix & ":Administration:Project Details"
		case "viewprojects.asp"
			setTitle = titlePrefix & ":Administration Tasks"
		case "viewuseradmin.asp"
			setTitle = titlePrefix & ":Administration:User Details"
	end select
end function

'CR20060902
'Moved from insideissuedetails because called in cons_search_report.asp,insideissuedetails.asp ( at 2 places ).
function check_file_ondisk(filename)
	dim fso
	set fso = server.CreateObject("scripting.filesystemobject")
	'response.Write(server.MapPath(".") & "\" & "upload" & "\" & filename)
	'response.Write(server.MapPath("."))
	If fso.FileExists(server.MapPath(".") & "\" & "upload" & "\" & filename) Then
		'fso.deletefile(server.MapPath(".") & "\" & "upload" & "\" & filename)
		check_file_ondisk =  filename 
	else
		check_file_ondisk = "<strike>" & filename & "</strike>"
	end if
	set fso = nothing
end function

'CR20060902
'Moved from insideissuedetails because called in cons_search_report.asp,insideissuedetails.asp ( at 2 places ).
function strike_out_deleted_file(my_string)	' CR20060826
	dim pos,len_my_string,filename
	len_my_string = len(my_string)
	pos = InStrRev(my_string, "/", len_my_string - 4) ' 4 is the length of </a> and here we get the position of second last "/"
	filename = Mid(my_string, pos + 1, len_my_string - pos - 4)
	if filename <> check_file_ondisk(filename) then
		pos = instr(1,my_string,"Attached file path :")
		my_string = left(my_string,pos + len("Attached file path :") - 1) & "<strike>" & right(my_string,len_my_string - (pos + len("Attached file path :") - 1)) & "</strike>"
	end if
	strike_out_deleted_file = my_string
end function


'CR20060902
'Moved from insideissuedetails because called in cons_search_report.asp,insideissuedetails.asp ( at 2 places ).
function formatData(data)
dim i,pos,pos1,pos2,pos3,stringFind,stringRep,result,from_len	' CR20060812
	'data=replace(data,"<","&lt;")
	'data= replace(data,">","&gt;")
	'data = replace(data,chr(10),"<br>")
	if IsHTMLContent(data) Then
		exit function
	end if
	data = server.HTMLEncode(data)
	data = replace(data,chr(10),"<br>")
	
	For i = 1 To Len(data)
		'pos1 = InStr(i, data, "http://")
		pos1 = InStr(data, "http://")
		If pos1 = 0 Then 	
			result = result & data	' CR20060812
			Exit For
		end if	' CR20060812
		pos2 = InStr(pos1, data, " ")
		pos3 = InStr(pos1, data, "<br>")
		If pos2 = 0 and pos3 = 0 Then
			pos2 = Len(data) + 1
		End If
		if pos2 > pos3 and pos3 <> 0 then
			pos2 = pos3
		elseif pos2 = 0 and pos3 <> 0 then
			pos2 = pos3
		End If
		
		stringFind = Mid(data, pos1, pos2 - pos1)
		
		stringFind_len = len(stringFind)	' CR20060812
		stringRep = "<a href=""" & stringFind & """ target=""_blank"">" & stringFind & "</a>"
		result = result & Left(data, pos1 - 1) & stringRep	' CR20060812
		data = Mid(data, pos1 + stringFind_len)	' CR20060812
		'data = Replace(data, stringFind, stringRep,i)
		'i = pos2 + (Len(stringRep) - Len(stringFind))
	Next
	' for https:// string CR20070129
	data = result
	result = ""
	For i = 1 To Len(data)
		'pos1 = InStr(i, data, "http://")
		pos1 = InStr(data, "https://")
		If pos1 = 0 Then 	
			result = result & data	' CR20060812
			Exit For
		end if	' CR20060812
		pos2 = InStr(pos1, data, " ")
		pos3 = InStr(pos1, data, "<br>")
		If pos2 = 0 and pos3 = 0 Then
			pos2 = Len(data) + 1
		End If
		if pos2 > pos3 and pos3 <> 0 then
			pos2 = pos3
		elseif pos2 = 0 and pos3 <> 0 then
			pos2 = pos3
		End If
		
		stringFind = Mid(data, pos1, pos2 - pos1)
		
		stringFind_len = len(stringFind)	' CR20060812
		stringRep = "<a href=""" & stringFind & """ target=""_blank"">" & stringFind & "</a>"
		result = result & Left(data, pos1 - 1) & stringRep	' CR20060812
		data = Mid(data, pos1 + stringFind_len)	' CR20060812
		'data = Replace(data, stringFind, stringRep,i)
		'i = pos2 + (Len(stringRep) - Len(stringFind))
	Next
	
	if instr(1,result,"Attached file path :") <> 0 then		' CR20060826
		result =  strike_out_deleted_file(result)
	end if
	
	pos = 0
	pos = instr(1,result,"-------Files Attached-------")
	'response.write(result)
	if pos <> 0 then		' CR20060826
		result =  format_multiple_files_attached(pos,len("-------Files Attached-------"),result)
	end if

	formatData = result

end function

function format_multiple_files_attached(pos,str_len,result)
'response.Write("coming here")
	dim filename
	dim counter
	filename = mid(result,(pos + str_len))
'response.Write(filename)
	result = mid(result,1,(pos + str_len)-1)
	
	format_multiple_files_attached = result & create_file_links(filename,"for_ims")
end function

function create_file_links(filename,for_what)
	dim result
	dim counter
	filename = split(filename,",")
	if isArray(filename) then
		if ubound(filename) >= 0 then
			for counter=0 to ubound(filename)
			'if counter > 0 then
			'	response.Write(", ")
			'end if
			if for_what = "for_ims" then
				result = result & "<br>>> <a href='" & varSiteSpecURL & Mid(request.ServerVariables("PATH_INFO"), 1, InStrRev(request.ServerVariables("PATH_INFO"), "/")) & "getfile.asp?fn=" & filename(counter) & "' target='_blank'>" & check_file_ondisk(trim(filename(counter)))& "</a>"
			else
				result = result & varSiteSpecURL & Mid(request.ServerVariables("PATH_INFO"), 1, InStrRev(request.ServerVariables("PATH_INFO"), "/")) & "upload/" & filename(counter) & vbcrlf
			end if
			
			next
		end if
	end if
	create_file_links = result
end function

' CR20060902
' function to highlight a string in given data
' not working correctly
function highlight(data,text_to_highlight)
	dim i,pos,pos1,pos2,pos3,stringFind,stringRep,result,from_len

	For i = 1 To Len(data)
		pos1 = InStr(1,data, text_to_highlight,1)
		If pos1 = 0 Then 	
			result = result & data
			Exit For
		end if	
		pos2 = pos1 + len(text_to_highlight)

		If pos2 = 0  Then
			pos2 = Len(data) + 1
		End If
				
		stringFind_len = len(text_to_highlight)	
		stringRep = "<span style='background-color:yellow'>" & text_to_highlight &  "</span>" 
		result = result & Left(data, pos1 - 1) & stringRep	
		data = Mid(data, pos1 + stringFind_len)	
	Next
	highlight = result
end function

function returnData(query)		' return the recordset in an array
	dim rs
	dim arr 
	arr = array()
	set rs = server.createobject("adodb.recordset")
	rs.open query,con
	if not rs.eof then
		arr = rs.getrows()
	end if
	returnData = arr
	rs.close
	set rs =nothing
end function

' requires query string and a dictionary object
' and return a array of data
function returnDataAndFields(query,objFields)
	dim rs
	dim x
	dim arrData
	dim fldcount
	arrData = array()
	set rs = server.CreateObject("adodb.recordset")
	rs.open query,con
	if not rs.eof then
		arrData = rs.getrows
		fldcount = rs.Fields.Count
		for x = 0 to fldcount-1
			objFields.add rs.Fields(x).Name,x
		Next
	end if
	returnDataAndFields = arrData
	rs.close
	set rs = nothing
	'set connect = nothing
end function

function date_format(str_date)
	if isempty(str_date) or isnull(str_date) then
		str_date = date
	end if
	ddyear = year(str_date)
	ddmonth = month(str_date)
	ddday = day(str_date)
	if len(ddmonth) = 1 then
		ddmonth = "0" & ddmonth
	end if
	if len(ddday) = 1 then
		ddday = "0" & ddday
	end if
	str_date = ddyear & "-" & ddmonth & "-" & ddday
    date_format = str_date
end function

function receipent_is_active(mailid)
	dim rsUserProfile,sqlisactive
	set rsUserProfile = server.CreateObject("adodb.recordset")
	sqlisactive = "SELECT 1 from userprofile where email = '" & mailid & "' and status = 1"
	rsUserProfile.open sqlisactive, con
	if not rsUserProfile.eof then
		receipent_is_active = true
	else 
		receipent_is_active = false
	end if
	rsUserProfile.close
	set rsUserProfile = nothing
end function 

function chomp(str)
	
	on error resume next
	
	chomp = replace(replace(replace(str,chr(13),""),chr(32),""),chr(10),"")
	
end function 

Function IsHTMLContent(str)
    ' Check if the string contains HTML tags or characters
    If InStr(str, "<") > 0 And InStr(str, ">") > 0 Then
        IsHTMLContent = True
    Else
        IsHTMLContent = False
    End If
End Function

%>
