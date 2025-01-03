<%
' Modified Date		Modified By			Comments
' 20060328			Atul Agrawal		if file already exist, it will be datetime stamp further in format yyyyMMddhhmmss 

%>
<!-- #include file="checksession.asp" -->
<!--  < #include file="connect.asp"> -->
<!-- #include file="generalfunctions.asp" -->
<!-- #include file="sendmail.asp" -->

<%

' ********************* CHECKING PERMISSION ***************************

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

' ********************* CHECKING PERMISSION ENDS***************************
%>

<%
response.cachecontrol = "no-cache"
response.ExpiresAbsolute = now
response.AddHeader "pragma","no-cache"
response.expires = -1
%>
<!-- #include file="freeaspupload.asp" -->
	<%
	Dim uploadsDirVar
	dim validExtensions
	uploadsDirVar = server.MapPath(".") & "\Upload" '"G:\PleskVhosts\gtss.in\gtsims.com\ims\upload"
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
%>
	<%
	  dim objrs
	  set objrs =server.CreateObject("adodb.recordset")
	  objrs.CursorLocation = 3 'adUseClient
	  objrs.Open "select * from " &  varTblNameProjects & " where 1=2",con,3,3
	  objrs.AddNew 
	  objrs("projectName") = trim(uploader.form("txtproject"))
	  objrs("LeadDeveloper") = trim(uploader.form("txtLeadDeveloper"))
	  objrs("Description") = trim(uploader.form("txtIssDesc"))
	  objrs("createDate") = now()
	  objrs("ProjectWatchList") = session("user") & "," & trim(uploader.form("txtLeadDeveloper"))
	  if filename<>"" then
		   objrs("DesignDocPath") =  filename
	  end if
	  objrs.Update
	 ' con.committrans
	  objrs.Close 
	  set objrs = nothing
	  call sendmail
	  con.close
	  response.Redirect("viewprojects.asp")
	  'response.write("Project Created Successfully.")
end if 
if con.state = 1 then
	con.close
	set con = nothing
end if 


private sub sendmail_old()
	dim Mailer
	Set Mailer = Server.CreateObject("SMTPsvg.Mailer")
	Mailer.FromName   = session("name")
	Mailer.FromAddress = varSiteSpecFromAddress 
	Mailer.ReplyTo = session("user")
	Mailer.RemoteHost = varSiteSpecRemoteHost
	Mailer.AddRecipient "",trim(uploader.form("txtLeadDeveloper"))
	if varSiteSpecAddCC <> "" then
		mailer.addCC "", varSiteSpecAddCC
	end if
	Mailer.Subject    = varSiteSpecMailSubjectPrefix & "<" & varSpecialclientid & ">" & " Project Added"
	Mailer.BodyText   = "Project :-" & trim(uploader.form("txtproject")) & vbcrlf & "You are lead developer in it." & vbcrlf & vbcrlf & "Description :" & vbcrlf &  trim(uploader.form("txtIssDesc"))
	Mailer.SendMail
end sub

private sub sendmail()
	dim FromName
	dim ReplyTo
	dim BCC
	dim Subject
	dim Body
	dim vari
	FromName   = session("name")
	ReplyTo = session("user")
	BCC = ""
	dim receipent(2)
	'redim Receipent(1)
	Receipent(0) = trim(uploader.form("txtLeadDeveloper"))
	Receipent(1) = session("user")
	'Receipent(2) = varSiteSpecAddCC
	
'	for vari = 0 to ubound(usersPermitted)
'		redim preserve Receipent(ubound(Receipent)+1)  
'		Receipent(ubound(Receipent)) = usersPermitted(vari)
'	next
	
	Subject    = varSiteSpecMailSubjectPrefix & " Project Added"
	Body   = "Project :-" & trim(uploader.form("txtproject")) & vbcrlf & trim(uploader.form("txtLeadDeveloper")) & " is lead developer in it." & vbcrlf & vbcrlf & "Description :" & vbcrlf &  trim(uploader.form("txtIssDesc"))
	
	mail FromName,ReplyTo,Receipent,varSiteSpecAddCC,BCC,Subject,Body
	
end sub

%>