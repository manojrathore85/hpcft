<%
' Modified Date		Modified By			Comments
' 20060328			Atul Agrawal		if file already exist, it will be datetime stamp further in format yyyyMMddhhmmss 
' 20060524			Atul				Sending the url of attached file in comments and in mail
%>
<!-- #include file="checksession.asp" -->
<!--  < #include file="connect.asp"> -->
<!-- #include file="freeaspupload.asp" -->
<!-- #include file="generalFunctions.asp" -->
<!-- #include file ="sendmail.asp" -->
<%
Session.CodePage  = 65001
		' Create the FileUploader
Dim uploadsDirVar
dim validExtensions
uploadsDirVar = server.MapPath(".") & "\upload" '"G:\PleskVhosts\gtss.in\gtsims.com\upload" 		
validExtensions = split(varSiteSpecValidExtensions,",")
	
	
on error resume next
if Request.ServerVariables("REQUEST_METHOD") = "POST" then


	%>
	
	<%
	dim fileName,filenames
	filenames = array()
	
	dim objAdd 'for 
	set objAdd =server.CreateObject("adodb.recordset") 'for updating issuechangescomments
	
	dim varServerVariablePath_info
	varServerVariablePath_info = request.ServerVariables("PATH_INFO")
	call check_permission
	call no_cache
	Dim Uploader, File , issueId

	Set Uploader = New FreeASPUpload
	

	call upload_files
	If Err.Number<>0 then
		response.Write(err.description & "<br>")
		response.End()
	end if
	call update_database
	If Err.Number<>0 then
		response.Write(err.description & "<br>")
		response.End()
	end if
	call sendmail ' send the assigne of the issue that issue is updated
	If Err.Number<>0 then
		response.Write(err.description & "<br>")
	end if
	if con.state = 1 then
		con.close
	end if 
	
	if objAdd.state = 1 then objAdd.close
	set objAdd = nothing
	set con = nothing

end if
%>	
<script language="JavaScript">
	self.close();
	opener.location.reload();
</script>	

<%	
   'Response.Redirect("issueDetails.asp?issueid=" & uploader.form("issueid")) ' for redirect 
'end if 


' ********************* PROCEDURES ***************************

'######################################################################################################
'		FUNCTION TO CHECK THE PERMISSION ON FILE
'######################################################################################################

sub check_permission
	dim rsperm
	dim sstr
	set rsperm = server.CreateObject("adodb.recordset")
	sstr = "select perm.* from " &  varTblNamePermissions & " perm, " &  varTblNameUsers & " u, " &  varTblNameissues & _
	 " i where i.projectid = u.projectid and u.permissionid = perm.permissionid and i.issueid = " &  uploader.form("issueid") & _
	 " and u.email ='" &  session("user") & "'"
	
	 rsperm.open sstr,con
	 
	if not rsperm.eof then
		if rsperm("pwrite") = "T" then
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
'######################################################################################################
'		CLEAR THE CACHE FOR THE PAGE
'######################################################################################################
sub no_cache
	response.cachecontrol = "no-cache"
	response.ExpiresAbsolute = now
	response.AddHeader "pragma","no-cache"
	response.expires = -1
end sub
  
'######################################################################################################
'		FUNCTION TO SAVE UPLOADED FILES ON SERVER
'######################################################################################################
sub upload_files

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
		filename = join(filenames,",")
    end if
	If Err.Number<>0 then
		response.Write(err.description & "<br>")
		response.end
	end if
	
end sub

'######################################################################################################
'		FUNCTION TO SEND THE FULL HISTORY OF ISSUE TO USERS EMAIL
'######################################################################################################
private function fullhistory() 
	dim rs
	set rs = server.CreateObject("adodb.recordset")
	rs.open "select description from " &  varTblNameIssues & " where issueid = " & uploader.form("issueid") ,con
	if not rs.eof then
		fullhistory = "Description :-" & rs("description") & vbcrlf & vbcrlf
	end if
	rs.close
	rs.open "select changesby,updatedate,comments from " &  varTblNameIssueChangesComments & " where issueid = " & uploader.form("issueid") ,con
	while not rs.eof
		fullhistory = fullhistory & "Changes By :-" & rs("changesby") & "   Update Date :-" & rs("updatedate") & vbcrlf & "Comments :-" & rs("comments") & vbcrlf
		rs.movenext
	wend	
	rs.close
	set rs = nothing
end function

'######################################################################################################
'		FUNCTION TO SAVE CHANGES IN BASIC FIELDS OF ISSUE
'######################################################################################################

private sub AddvalueInIssuechanges (byval projectId,byval issueId,byval field,byval newvalue,byval comments,byval oldvalue)
	'if objAdd.state = 1 then objAdd.close
	objAdd.CursorLocation =3 ' adUSeclient
		 objAdd.Open "select * from " &  varTblNameIssueChangesComments & " where 1=2",con,3,3
		  objAdd.AddNew 
			objAdd("projectid")=projectID
			objAdd("issueid")= issueId
			objAdd("field")= field
			objAdd("changesby")= session("user") 
			objAdd("updatedate")=now()
			objAdd("newvalue")=newvalue
		'objAdd("comments")=comments
			objAdd("lastvalue")=oldvalue
	objAdd.Update 
	objAdd.Close 
	'set objAdd= nothing 
end sub

'######################################################################################################
'		FUNCTION TO SEND EMAIL TO USERS OF THIS ISSUE
'######################################################################################################

private sub sendmail_old()
	dim rstemp
	dim i
	dim reportUsers ' users added in reportusers of a issue
	dim reportUsers1 ' users added in projectwatchlist
	dim reportUsers2 ' users added in the form 'One Time Email'
	dim rptuser ' this variable is used to keep the value in recordset, for it is does not come after checking the recordset for eof
	dim projectwatchlist ' this variable is used to keep the value in recordset, for it is does not come after checking the recordset for eof
	dim temp
	
	dim str_comments
	str_comments = uploader.form("txtcomments")
	
	if filename <> "" then
		str_comments = str_comments & vbcrlf & "Files Attached :- " & vbcrlf & create_file_links(filename,"for_mail")
	end if
	
	set rstemp = server.CreateObject("adodb.recordset")
	rstemp.open "select p.projectname,p.projectwatchlist,i.summary,i.assignto,i.reportUsers from " &  varTblNameProjects & " p," &  varTblNameIssues & " i where i.projectid = p.projectid and i.issueid =" & uploader.form("issueid") & " and i.projectid = " & uploader.form("txtprojectid"),con
	projectwatchlist = rstemp("projectwatchlist")
	rptuser = rstemp("reportUsers")
	if not rstemp.eof then
		if not isnull(rptuser) and rptuser <> "" then
			reportUsers = split(rptuser,",")
		end if
		if not isnull(projectwatchlist) and projectwatchlist <> "" then
			reportUsers1 = split(projectwatchlist,",")
		end if
		if  uploader.form("txtMailsTo") <> "" then
			reportUsers2 = split(uploader.form("txtMailsTo"),",")
		end if
		dim Mailer
		Set Mailer = Server.CreateObject("SMTPsvg.Mailer")
		Mailer.FromName   = session("name")
		Mailer.FromAddress = varSiteSpecFromAddress
		Mailer.ReplyTo = session("user")
		Mailer.RemoteHost = varSiteSpecRemoteHost
		Mailer.AddRecipient "",rstemp("assignto")
		if isarray(reportUsers) then
			for i = 0 to ubound(reportUsers)
				Mailer.AddRecipient "",replace(replace(replace(reportUsers(i),chr(13),""),chr(32),""),chr(10),"")
			next
		end if
		if isarray(reportUsers1) then
			for i = 0 to ubound(reportUsers1)
				Mailer.AddRecipient "",replace(replace(replace(reportUsers1(i),chr(13),""),chr(32),""),chr(10),"")
			next
		end if
		if isarray(reportUsers2) then
			for i = 0 to ubound(reportUsers2)
				Mailer.AddRecipient "",replace(replace(replace(reportUsers2(i),chr(13),""),chr(32),""),chr(10),"")
			next
		end if
		if varSiteSpecAddCC <> "" then
			Mailer.AddCC "", varSiteSpecAddCC
		end if
		'Mailer.Subject    =  varSiteSpecMailSubjectPrefix & "<" & varSpecialclientid & ">" & " <" & rstemp("summary") & "><" & rstemp("projectname") & ">"
		Mailer.Subject    = "(" & varSiteSpecMailSubjectPrefix & ")" & varSpecialclientid & " | " &  rstemp("summary") & " | " & rstemp("projectname") 'CR20060821Atul
		if uploader.form("chkSendHistory") <> ""  then
			'Mailer.BodyText   = "Project :-" & rstemp("projectname") & vbcrlf & "Issue Summary:-" & rstemp("summary") & vbcrlf & varSiteSpecURL & Mid(varServerVariablePath_info, 1, InStrRev(varServerVariablePath_info, "/")) & "issuedetails.asp?issueid=" & issueid & vbcrlf & "Comments :" & trim(uploader.form("txtcomments")) & vbcrlf & fullhistory
			Mailer.BodyText   = "Project :-" & rstemp("projectname") & vbcrlf & "Issue Summary:-" & rstemp("summary") & vbcrlf & varSiteSpecURL & Mid(varServerVariablePath_info, 1, InStrRev(varServerVariablePath_info, "/")) & "issuedetails.asp?issueid=" & uploader.form("issueid") & "&prjid=" & uploader.form("txtprojectid") & vbcrlf & "Comments :" & vbcrlf & vbcrlf & str_comments & vbcrlf & fullhistory		' 24May06Atul
		else
			Mailer.BodyText   = "Project :-" & rstemp("projectname") & vbcrlf & "Issue Summary:-" & rstemp("summary") & vbcrlf & varSiteSpecURL & Mid(varServerVariablePath_info, 1, InStrRev(varServerVariablePath_info, "/")) & "issuedetails.asp?issueid=" & uploader.form("issueid") & "&prjid=" & uploader.form("txtprojectid") & vbcrlf & "Comments :" & vbcrlf & vbcrlf & str_comments	' 24May06Atul
		end if
		Mailer.SendMail
	end if
	rstemp.close
	set rstemp = nothing
end sub

private sub sendmail()

	dim rstemp
	dim i
	dim reportUsers ' users added in reportusers of a issue
	dim reportUsers1 ' users added in projectwatchlist
	dim reportUsers2 ' users added in the form 'One Time Email'
	dim rptuser ' this variable is used to keep the value in recordset, for it is does not come after checking the recordset for eof
	dim projectwatchlist ' this variable is used to keep the value in recordset, for it is does not come after checking the recordset for eof
	dim temp
	set rstemp = server.CreateObject("adodb.recordset")
	dim qry
	qry ="select p.projectname,p.projectwatchlist,i.summary,i.assignto,i.reportUsers from " &  varTblNameProjects & " p," &  varTblNameIssues & " i where i.projectid = p.projectid and i.issueid =" & issueid & " and i.projectid = " & uploader.form("txtprojectid")

	rstemp.open qry,con
	projectwatchlist = rstemp("projectwatchlist")
	rptuser = rstemp("reportUsers")
	summary = rstemp("summary")
	
	dim str_comments
	str_comments = uploader.form("txtcomments")
	
	if filename <> "" then
		str_comments = str_comments & vbcrlf & "Files Attached :- " & vbcrlf & create_file_links(filename,"for_mail")
	end if
	
	if not rstemp.eof then
	
		if not isnull(rptuser) and rptuser <> "" then
			reportUsers = split(rptuser,",")
		end if
		if not isnull(projectwatchlist) and projectwatchlist <> "" then
			reportUsers1 = split(projectwatchlist,",")
		end if
		if  uploader.form("txtMailsTo") <> "" then
			reportUsers2 = split(uploader.form("txtMailsTo"),",")
		end if
		FromName   = session("name")
		ReplyTo = session("user")
		dim Receipent()
		redim Receipent(0)  
		Receipent(0) = rstemp("assignto")

		if isarray(reportUsers) then
			for i = 0 to ubound(reportUsers)
				redim preserve Receipent(ubound(Receipent)+1)    
				Receipent(ubound(Receipent)) = reportUsers(i)
			next
		end if
		if isarray(reportUsers1) then
			for i = 0 to ubound(reportUsers1)
				redim preserve Receipent(ubound(Receipent)+1)    
				Receipent(ubound(Receipent)) = reportUsers1(i)
			next
		end if
		if isarray(reportUsers2) then
			for i = 0 to ubound(reportUsers2)
				redim preserve Receipent(ubound(Receipent)+1)    
				Receipent(ubound(Receipent)) = reportUsers2(i)
			next
		end if

		Subject    = "(" & varSiteSpecMailSubjectPrefix & ")" & varSpecialclientid & " | " &  summary & " | " & rstemp("projectname") 'CR20060821Atul

		if uploader.form("chkSendHistory") <> ""  then
			Body   = "Project :-" & rstemp("projectname") & vbcrlf & "Issue Summary:-" & summary & vbcrlf & "Due Date:-"  & vbcrlf & varSiteSpecURL & Mid(varServerVariablePath_info, 1, InStrRev(varServerVariablePath_info, "/")) & "/issuedetails.asp?issueid=" & issueid & "&prjid=" & uploader.form("txtprojectid") & vbcrlf & "Comments :" & vbcrlf & vbcrlf & str_comments & vbcrlf & fullhistory		' 24May06Atul
		else
			Body = "Project :-" & rstemp("projectname") & vbcrlf & "Issue Summary:-" & summary & vbcrlf & "URL: -" &  varSiteSpecURL & Mid(varServerVariablePath_info, 1, InStrRev(varServerVariablePath_info, "/")) & "/issuedetails.asp?issueid=" & issueid & "&prjid=" & uploader.form("txtprojectid") & vbcrlf & "Comments :" & vbcrlf & vbcrlf & str_comments	' 24May06Atul
		end if
		
		BCC = ""
		
		mail FromName,ReplyTo,Receipent,varSiteSpecAddCC,BCC,Subject,Body
		
	end if
	rstemp.close
	set rstemp = nothing
	
end sub

sub update_database
	dim issuetype,severity,summary,assignto,description,projectId,filenamesindbase 'varible use for storing the initial values 
   
	dim str
    dim objview ' for record set
	issueid = uploader.form("issueid")
	set objview = server.CreateObject("adodb.recordset") ' createing the object
'   			objview.Open "select * from issues where issueId="& request("txtissueId")
		 objview.Open "select * from " &  varTblNameIssues & " where issueId="& issueid,con
				 if objview.EOF  =false and objview.BOF  =false then
						 projectid=objview("projectId")
						 issuetype=objview("issuetype")
						 severity =objview("severity")
						 summary =objview("summary")
						 assignto= objview("Assignto")
						 description =objview("description")
						 str = objview("attachedfilepath")
						 if not isnull(objview("attachedfilepath")) then
							 filenamesindbase = str
						 end if
				  end if 
		  objview.Close 
	  set objview = nothing
	'dim objIssupdat 'variable for issue update
			'	if  issuetype is changed then we insert this value in issuechangescomments
	dim newIssDesc 
	' handling all the special characters of mysql for update statement
	newIssDesc = replace(trim(uploader.form("txtIssDesc")),"'","\'")
	newIssDesc = replace(newIssDesc,chr(34),"\""")  'chr(34) is for double quote ( " )
	'newIssDesc = replace(newIssDesc,"_","\_")
	'newIssDesc = replace(newIssDesc,"%","\%")

	con.begintrans

		if issuetype <> uploader.form("lstIssTypes") then
			AddvalueInIssuechanges uploader.form("txtprojectId"), issueid, "Issuetype",trim(uploader.form("lstIssTypes")), trim(uploader.form("txtcomments")),trim(issuetype)
			'call sendmail trim(uploader.form("lstUsers")),"Issue Updated",trim(uploader.form("txtIssSummary")) & "<br><br>Comments :- " & trim(uploader.form("txtcomments"))
		end if  

' if severity is changed then we insert this value in issuechangescomments
		if trim(severity) <> trim(uploader.form("lstSeverity")) then
			AddvalueInIssuechanges uploader.form("txtprojectId"), issueid, "severity",trim(uploader.form("lstSeverity")), trim(uploader.form("txtcomments")),trim(severity) 'calling the procedure
			'call sendmail trim(uploader.form("lstUsers")), "Issue Updated",  trim(uploader.form("txtIssSummary")) & "<br><br>Comments :- " & trim(uploader.form("txtcomments"))
		end if  
'if summary is changed then we insert in this value in issuechangescomments
		if trim(summary) <> trim(uploader.form("txtIssSummary")) then
			AddvalueInIssuechanges uploader.form("txtprojectId"), issueid, "summary",trim(uploader.form("txtIssSummary")), trim(uploader.form("txtcomments")),trim(summary) 'calling the procedure
			'call sendmail trim(uploader.form("lstUsers")) ,"Issue Updated", trim(uploader.form("txtIssSummary")) & "<br><br>Comments :- " & trim(uploader.form("txtcomments"))		
		end if  

'if assign to is changed then we insert this value in issuechangescomments

		if trim(assignto) <> trim(uploader.form("lstUsers")) then
				 AddvalueInIssuechanges uploader.form("txtprojectId"), issueid, "assignto",trim(uploader.form("lstUsers")), trim(uploader.form("txtcomments")),trim(assignto) 'calling the procedure
				 'call sendmail trim(uploader.form("lstUsers")) ,"You are assigned an issue in project", trim(uploader.form("txtIssSummary"))
		end if  

'if description to is changed then we insert this value in issuechangescomments

		if trim(description) <> trim(uploader.form("txtIssDesc")) then
			AddvalueInIssuechanges uploader.form("txtprojectId"), issueid, "description",trim(uploader.form("txtIssDesc")), trim(uploader.form("txtcomments")),trim(description) 'calling the procedure
			'call sendmail trim(uploader.form("lstUsers")), "Issue Updated", trim(uploader.form("txtIssSummary")) & "<br><br>Comments :- " & trim(uploader.form("txtcomments"))
	 	end if  

'	 ************************** adding the comment in the fields  ***
	 if  trim(uploader.form("txtcomments"))<>"" or filename <> "" then  		'24May06Atul 	added filename condition
		dim objcomment

			set objcomment =server.CreateObject("adodb.recordset")
				objcomment.CursorLocation=3
				objcomment.Open "select * from " &  varTblNameIssueChangesComments & " where 1=2",con,3,3
				objcomment.AddNew


				objcomment("projectid")=uploader.form("txtprojectId")
				objcomment("issueid")= issueId
				strComment = trim(uploader.form("txtcomments"))			'24May06Atul
				if filename <> "" then			'24May06Atul
					'strComment = strComment & vbcrlf & "Attached file path :" & varSiteSpecURL & "/upload/" & filename		'24May06Atul
					 strComment = strComment & vbcrlf & "-------Files Attached-------" & filename '& vbcrlf & varSiteSpecURL & "/upload/" & filename		'24May06Atul
				end if
				objcomment("comments")= strComment 			'24May06Atul 
				objcomment("updateDate")=now()
				objcomment("changesBy")=session("user")
				objcomment.Update
				objcomment.Close 
				set objcomment = nothing
	 end if  

    ' adding the data or updateing the data in issues
     
   

'set objIssupdat =server.CreateObject ("adodb.recordset")


  dim strsql
  if filenamesindbase = "" and filename = "" then
	  strsql="update " &  varTblNameIssues & " set issuetype='"& trim(uploader.form("lstIssTypes")) & "',Severity='"& trim(uploader.form("lstSeverity"))&"',AssignTo='" &  trim(uploader.form("lstUsers")) & "',summary='"&  trim(uploader.form("txtIssSummary")) &"',description='"& trim(newIssDesc) &"',updateDate= now() where issueId=" & issueid
  elseif filenamesindbase = "" and filename <> "" then
  	  strsql="update " &  varTblNameIssues & " set issuetype='"& trim(uploader.form("lstIssTypes")) & "',Severity='"& trim(uploader.form("lstSeverity"))&"',AssignTo='" &  trim(uploader.form("lstUsers")) & "',summary='"&  trim(uploader.form("txtIssSummary")) &"',description='"& trim(newIssDesc) &"',updateDate= now(),attachedfilepath='" & filename & "' where issueId=" & issueid
  elseif filenamesindbase <> "" and filename <> "" then
  	  strsql="update " &  varTblNameIssues & " set issuetype='"& trim(uploader.form("lstIssTypes")) & "',Severity='"& trim(uploader.form("lstSeverity"))&"',AssignTo='" &  trim(uploader.form("lstUsers")) & "',summary='"&  trim(uploader.form("txtIssSummary")) &"',description='"& trim(newIssDesc) &"',updateDate= now(),attachedfilepath='" & filenamesindbase & "," & filename & "' where issueId=" & issueid
	elseif filenamesindbase <> "" and filename = "" then
  	  strsql="update " &  varTblNameIssues & " set issuetype='"& trim(uploader.form("lstIssTypes")) & "',Severity='"& trim(uploader.form("lstSeverity"))&"',AssignTo='" &  trim(uploader.form("lstUsers")) & "',summary='"&  trim(uploader.form("txtIssSummary")) &"',description='"& trim(newIssDesc) &"',updateDate=now() where issueId=" & issueid	  
  end if
  if strsql <> "" then
	  con.execute strsql 
  end if
  con.committrans
   'Response.Write("record is saved successfully") 
   'con.close
end sub
' ********************* END PROCEDURES ***************************
%>
