<%
' Modified Date		Modified By			Comments
' 20060328			Atul Agrawal		if file already exist, it will be datetime stamp further in format yyyyMMddhhmmss 
' 20070202			atul agrawal		manually providing issueid in issues table
%>
<!-- #include file="checksession.asp" -->
<!--  < #include file="connect.asp"> -->
<!-- #include file="generalfunctions.asp" -->
<!-- #include file ="sendmail.asp" -->
<%

' ********************* CHECKING PERMISSION ***************************

dim rsperm
dim maxissueid
dim varServerVariablePath_info
varServerVariablePath_info = request.ServerVariables("PATH_INFO")
set rsperm = server.CreateObject("adodb.recordset")
rsperm.open "select * from " &  varTblNamePermissions & " where permissionid = '" & session("permission") & "'",con
if not rsperm.eof then
	if rsperm("padd") = "T" then
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

' ********************* CHECKING PERMISSION ENDS***************************
%>
<%
response.cachecontrol = "no-cache"
response.ExpiresAbsolute = now
response.AddHeader "pragma","no-cache"
response.expires = -1
Session.CodePage  = 65001
%>
<!-- #include file="freeaspupload.asp" -->

	<%
	
	Dim uploadsDirVar
	dim validExtensions
	uploadsDirVar = server.MapPath(".") & "\Upload" '"G:\PleskVhosts\gtss.in\gtsims.com\upload" 		
	validExtensions = split(varSiteSpecValidExtensions,",")
	
	dim filename
	dim message
	' Create the FileUploader
	Dim Uploader, File

	Set Uploader = New FreeASPUpload
	
	'******************************************
	' Use [FileUploader object].Form to access 
	' additional form variables submitted with
	' the file upload(s). (used below)
	'******************************************
	
	Uploader.Save(uploadsDirVar)
	
	If Err.Number<>0 then
		response.Write(err.description & "<br>")
		response.end
	end if

	Dim  fileSize, ks,  fileKey, counter
	counter = 0
	ks = Uploader.UploadedFiles.keys
    if (UBound(ks) <> -1) then
        SaveFiles = "<B>Files uploaded:</B> "
        for each fileKey in Uploader.UploadedFiles.keys
			filename = Uploader.UploadedFiles(fileKey).FileName
			redim preserve filenames(counter)
			filenames(counter) = filename
			counter = counter  + 1
        next
    end if
	
	If Err.Number<>0 then
		response.Write(err.description & "<br>")
		response.end
	end if
	
	
	%>
<%
if Request.ServerVariables("REQUEST_METHOD") = "POST" then

call set_max_issueid '20070202

%>
	<%
	  dim objrs
	  set objrs =server.CreateObject("adodb.recordset")

	  ' '' TO SELECT THE WATCHER LIST FROM PROJECT TABLE AND ADDING IT TO NEW ISSUE TO BE CREATED
	'  objrs.open "select * from projects where projectid = " & session("projectid"),con
	'  watchers = objrs("projectwatchlist")
	 ' objrs.close
	  objrs.CursorLocation = 3 'adUseClient
	  'con.begintrans
	  objrs.Open "select * from " &  varTblNameIssues & " where 1=2",con,3,3
	  objrs.AddNew 
	  objrs("projectId") = session("projectid")'uploader.form("txtproject")
	  objrs("issueid") = maxissueid ' 20070202
	  objrs("issuetype") = trim(uploader.form("lstIssTypes"))
	  objrs("severity") = trim(uploader.form("lstSeverity"))
	  objrs("summary") = trim(uploader.form("txtIssSummary"))
	  objrs("AssignTo") = trim(uploader.form("lstUsers"))
	  objrs("description") = trim(uploader.form("txtIssDesc"))
	  objrs("createDate") = now()
	  objrs("reporter") = trim(uploader.form("txtreporter"))
	  objrs("updatedate") = now() ' added on 20061212 atul
	  '' CHECKING WETHER USER LOGGED ON EXIST IN THE PROJECT WATCHER LIST IF YES THEN NO NEED TO 
	  '' ADD THE USER IN REPORT USER OF ISSUE CREATED
	  'if instr(1,watchers,session("user")) > 0 then 
	 ' 	objrs("reportusers") = watchers
	 ' else
	  '	if watchers <> "" then
	'	  	objrs("reportusers") = watchers & "," & session("user")
'		else
	  objrs("reportusers") = session("user")
'		end if
'	  end if
	  
	  if filename<>"" then
		   objrs("AttachedFilePath") =  filename
	  end if
	  objrs.Update
	  objrs.close
	  
	  'con.committrans
	  set objrs = nothing
	  call sendmail
	  con.close
	  response.redirect("issuenavigator.asp?search=all")
end if 

if con.state = 1 then
	con.close
	set con = nothing
end if 

private sub sendmail_old()
	Dim rstemp
	dim reportUsers
	dim projectwatchlist ' this variable is used to keep the value in recordset, for it is does not come after checking the recordset for eof
    Set rstemp = Server.CreateObject("adodb.recordset")
    rstemp.Open "select projectname,projectwatchlist from " &  varTblNameProjects & " where projectid =" & session("projectid"), con
    projectwatchlist = rstemp("projectwatchlist")
    If Not rstemp.EOF Then
		
		if not isnull(projectwatchlist) and projectwatchlist <> "" then
			reportUsers = split(projectwatchlist,",")
		end if
        Dim Mailer
        Set Mailer = Server.CreateObject("SMTPsvg.Mailer")
        Mailer.FromName = session("name")
        Mailer.FromAddress = varSiteSpecFromAddress
		Mailer.ReplyTo = session("user")
        Mailer.RemoteHost = varSiteSpecRemoteHost
        Mailer.AddRecipient "", Trim(uploader.Form("lstUsers"))
		if isarray(reportUsers) then
			for i = 0 to ubound(reportUsers)
				Mailer.AddRecipient "",replace(replace(replace(reportUsers(i),chr(13),""),chr(32),""),chr(10),"")
			next
		end if
		if varSiteSpecAddCC <> "" then
			mailer.addCC "", varSiteSpecAddCC
		end if
		'Mailer.Subject = varSiteSpecMailSubjectPrefix & "<" & varSpecialclientid & ">" & "<" & Trim(uploader.Form("txtIssSummary")) & "><" & rstemp("projectname") & ">"
		Mailer.Subject = "(" & varSiteSpecMailSubjectPrefix & ")" & varSpecialclientid & " | " & Trim(uploader.Form("txtIssSummary")) & " | " & rstemp("projectname")	'CR20060821Atul
        Mailer.BodyText = "Project :-" & rstemp("projectname") & vbCrLf & "Issue Created Summary :-" & Trim(uploader.Form("txtIssSummary")) & vbCrLf & "This issue has been assigned to you." & vbCrLf & varSiteSpecURL &  Mid(varServerVariablePath_info, 1, InStrRev(varServerVariablePath_info, "/")) & "/issuedetails.asp?issueid=" &  maxissueid & "&prjid=" & session("projectid") & vbcrlf & "Description :" & trim(uploader.form("txtIssDesc")) 
        Mailer.SendMail		
		set mailer = nothing
    End If
    rstemp.Close
    Set rstemp = Nothing
end sub

private sub sendmail()
	Dim rstemp
	dim reportUsers
	dim projectwatchlist ' this variable is used to keep the value in recordset, for it is does not come after checking the recordset for eof
    Set rstemp = Server.CreateObject("adodb.recordset")
    rstemp.Open "select projectname,projectwatchlist from " &  varTblNameProjects & " where projectid =" & session("projectid"), con
    projectwatchlist = rstemp("projectwatchlist")
    If Not rstemp.EOF Then
		
		if not isnull(projectwatchlist) and projectwatchlist <> "" then
			reportUsers = split(projectwatchlist,",")
		end if
  
        FromName = session("name")
		ReplyTo = session("user")
		dim Receipent()
		redim Receipent(0)  
		Receipent(0) = Trim(uploader.Form("lstUsers"))
		
		if isarray(reportUsers) then
			for i = 0 to ubound(reportUsers)
				redim preserve Receipent(ubound(Receipent)+1)    
				Receipent(ubound(Receipent)) = reportUsers(i)
				'Mailer.AddRecipient "",replace(replace(replace(reportUsers(i),chr(13),""),chr(32),""),chr(10),"")
			next
		end if
		
		Subject = "(" & varSiteSpecMailSubjectPrefix & ")" & varSpecialclientid & " | " & Trim(uploader.Form("txtIssSummary")) & " | " & rstemp("projectname")	'CR20060821Atul
        Body = "Project :-" & rstemp("projectname") & vbCrLf & "Issue Created Summary :-" & Trim(uploader.Form("txtIssSummary")) & vbCrLf & "This issue has been assigned to you." & vbCrLf & varSiteSpecURL &  Mid(varServerVariablePath_info, 1, InStrRev(varServerVariablePath_info, "/")) & "/issuedetails.asp?issueid=" &  maxissueid & vbcrlf & "Description :" & trim(uploader.form("txtIssDesc")) & vbcrlf 
		 ' & "Due Date :" & trim(uploader.form("due_date")) 

		BCC = ""

		mail FromName,ReplyTo,Receipent,varSiteSpecAddCC,BCC,Subject,Body

	End If
    rstemp.Close
    Set rstemp = Nothing
end sub

'20070202
sub set_max_issueid
	dim rs
	set rs = server.CreateObject("adodb.recordset")
	rs.open "select ifnull(max(issueid),0) + 1 as maxid from " & varTblNameIssues , con
	if not rs.eof then
		maxissueid = rs("maxid")
	else
		maxissueid = 1
	end if

	rs.close
	set rs = nothing

end sub

%>