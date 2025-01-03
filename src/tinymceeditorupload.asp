<!--#INCLUDE FILE="genvariables.asp"-->
<!--#INCLUDE FILE="freeaspupload.asp"-->

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

	Dim  fileSize, ks,  fileKey, counter,filenames1
	counter = 0
	ks = Uploader.UploadedFiles.keys
    if (UBound(ks) <> -1) then
        'SaveFiles = "<B>Files uploaded:</B> "
        for each fileKey in Uploader.UploadedFiles.keys
			filename = Uploader.UploadedFiles(fileKey).FileName
			redim preserve filenames(counter)
			filenames(counter) = filename
			'filenames1&="https:\/\/ims.gtssolutions.net\/upload\"&filename&"""
			counter = counter  + 1
        next
		filenames1 = Join( filenames , "', 'https:\/\/ims.gtssolutions.net\/upload/\" )  
		'filenames
    end if
	strJSN = "{""location"":""http:\/\/ims.gtssolutions.net\/upload\/"&filename&"""}"
	'strJSN = "{""location"":"""&filenames1&"""}"
	response.write(strJSN)
	If Err.Number<>0 then
		response.Write(err.description & "<br>")
		response.end
	end if
	
	
	%>