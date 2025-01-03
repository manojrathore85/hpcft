<%
if session("user") = "" then
	session("referrer") = Request.ServerVariables("QUERY_STRING")
	response.Redirect("http://gtsims.com/ims/login_index.asp")
	'response.Write("You are not authorized for this file, please login.")
	'response.End()
end if
'For Each Item In Request.ServerVariables
'	Response.Write Item & " = " & Request.ServerVariables(Item) & "<br>"
'Next

'response.End()

	if session("referrer") <> "" then
	     orginalqstring = session("referrer")
	 else
	     orginalqstring = Request.ServerVariables("QUERY_STRING")
	 end if

	qrystr = replace(orginalqstring,"404;http://gtsims.com:80/","")
	filename = replace(orginalqstring,"404;http://gtsims.com:80/upload/","")
	strFilePath = Request.ServerVariables("APPL_PHYSICAL_PATH")  &  replace(qrystr,"/","\")

	Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
	If objFSO.FileExists(strFilePath) Then
		Set objFile = objFSO.GetFile(strFilePath)
		intFileSize = objFile.Size
		Set objFile = Nothing
		
		strFileName = filename
		strFileName = replace(filename," ","-")
		Response.AddHeader "Content-Disposition","attachment; filename=" & strFileName
	
		Response.ContentType = "application/x-msdownload"
		Response.AddHeader "Content-Length", intFileSize
	
		Set objStream = Server.CreateObject("ADODB.Stream")
		objStream.Open
		objStream.Type = 1 'adTypeBinary
		objStream.LoadFromFile strFilePath
		Do While Not objStream.EOS And Response.IsClientConnected
			Response.BinaryWrite objStream.Read(1024)
			Response.Flush()
		Loop
		objStream.Close
		Set objStream = Nothing
	Else
		Response.write "Error finding file."
	End if
	
	Set objFSO = Nothing
%>
