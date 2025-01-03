<%@ Language=VBScript %>

<!-- #include file="checksession.asp" -->
<!-- #include file="connect.asp" -->
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
%>
<%
	dim strsql
	dim filenamesindbase
	dim rs
	dim str
	dim newPrjDesc 
	' handling all the special characters of mysql for update statement
	newPrjDesc = replace(trim(uploader.form("txtIssDesc")),"'","\'") 
	newPrjDesc = replace(newPrjDesc,chr(34),"\""")  'chr(34) is for double quote ( " )
	
	set rs = server.CreateObject("adodb.recordset")
	rs.open "select designdocpath from " &  varTblNameProjects & " where projectid=" & uploader.form("txtprojectId"),con
	str = rs(0)
	'filenamesindbase = rs(0)
	if not isnull(rs(0)) then
		filenamesindbase = str
	end if
	rs.close
	set rs = nothing
	if filenamesindbase = "" and filename = "" then
	  strsql= "Update " &  varTblNameProjects & " set projectName='"& trim(uploader.form("txtProject")) & "',LeadDeveloper='" & trim(uploader.form("txtLeadDeveloper")) &"', Description='" & trim(newPrjDesc) & "' where projectId=" & uploader.form("txtprojectId")
    elseif filenamesindbase = "" and filename <> "" then
  	  strsql= "Update " &  varTblNameProjects & " set projectName='"& trim(uploader.form("txtProject")) & "',LeadDeveloper='" & trim(uploader.form("txtLeadDeveloper")) &"', Description='" & trim(newPrjDesc) & "',designdocpath='" & filename & "' where projectId=" & uploader.form("txtprojectId")
    elseif filenamesindbase <> "" and filename <> "" then
  	  strsql= "Update " &  varTblNameProjects & " set projectName='"& trim(uploader.form("txtProject")) & "',LeadDeveloper='" & trim(uploader.form("txtLeadDeveloper")) &"', Description='" & trim(newPrjDesc) & "',designdocpath='" & filenamesindbase & "," & filename & "' where projectId=" & uploader.form("txtprojectId")
    end if
   if not strsql = "" then
	   con.execute strsql
	end if
   con.close()
   response.Redirect("viewprojects.asp")
 end if
 
 if con.state = 1 then
 	con.close
	set con = nothing
 end if
 

%>