<%
'Cross domain allow
Call Response.AddHeader("Access-Control-Allow-Origin", "*")
Call Response.AddHeader("Access-Control-Allow-Credentials", "true")
'if 9dfcd5e558dfa04aaf37f137a1d9d3e5
'9DFCD5E558DFA04AAF37F137A1D9D3E5
'for each x in Request.ServerVariables
  'response.write(x & " : " & Request.ServerVariables(x) & "<br>" )
'next
'<p class="font_7" style="line-height:normal; font-size:20px;"><span class="color_11">Details to be fetched from Princeton University&nbsp; Issue</span></p>
' get summary of all issues under AODM Project
'Sanjay: to convert string to MD5 hash 
function stringToUTFBytes(aString)
    Dim UTF8
    Set UTF8 = CreateObject("System.Text.UTF8Encoding")
    stringToUTFBytes = UTF8.GetBytes_4(aString)
end function
function bytesToHex(aBytes)
    dim hexStr, x
    for x=1 to lenb(aBytes)
        hexStr= hex(ascb(midb( (aBytes),x,1)))
        if len(hexStr)=1 then hexStr="0" & hexStr
        bytesToHex=bytesToHex & hexStr
    next
end function
function md5hashBytes(aBytes)
    Dim MD5
    set MD5 = CreateObject("System.Security.Cryptography.MD5CryptoServiceProvider")
    MD5.Initialize()
    'Note you MUST use computehash_2 to get the correct version of this method, and the bytes MUST be double wrapped in brackets to ensure they get passed in correctly.
    md5hashBytes = MD5.ComputeHash_2( (aBytes) )
end function
'Sanjay: Get all Issues from projectID=55
function getAllIssues()
	dim objissue ' issue change comments 
	set objissue =server.CreateObject("adodb.recordset") 
	objissue.open "select IssueId,summary from issues where projectid = 55 order by summary asc",con
	response.write("<br/><br/><p class='font_7' style='line-height:normal; font-size:20px;'>")
	while not objissue.eof
		response.write("<span class='color_11'><a  href='https://artofdecisionmaking.org/ddata?issueid=" & bytesToHex(md5hashBytes(stringToUTFBytes(objissue("IssueId")))) & "' target='_self'>" & objissue("summary") & "</a></span><br/>")
		objissue.movenext
	wend
	response.write("</p>")
end function
%>
<!-- #include file="generalFunctions.asp" -->
<!-- #include file ="Connect.asp" -->
<%
'Sanjay: this is the entry point for ASP
if request("issueid") = "" then  'For the list access
	call getAllIssues()
	response.end()
end if
'Sanjay: if issueId was received...
dim md5_issueid
dim IssueForAODM_API
dim issueMatch
issueMatch = false
For Each item In myArray 'myArray included and contains list of all issueids that are allowed to be shown/security arrangement
   md5_issueid = bytesToHex(md5hashBytes(stringToUTFBytes(item)))
   if md5_issueid = request("issueid") then
		IssueForAODM_API = item
		issueMatch = true
	end if
Next
if issueMatch = false then
	response.end  'TODO : for future handling and warning people not to attempt improper things on this website. Their attempt has been noted and would be reported.
end if
'response.write(md5_issueid & "<>" & request("issueid")) 
'response.write(request.ServerVariables("QUERY_STRING"))
dim TextStream,configText,lineArray
'Atul : Function to read config file and keep lines in TextStream variable
Function ReadConfigFile
	Const Filename = "aodm_config.txt"    ' file to read
	Const ForReading = 1, ForWriting = 2, ForAppending = 3
	Const TristateUseDefault = -2, TristateTrue = -1, TristateFalse = 0
	' Create a filesystem object
	Dim FSO
	set FSO = server.createObject("Scripting.FileSystemObject")
	' Map the logical path to the physical system path
	Dim Filepath
	Filepath = Server.MapPath(Filename)
	If FSO.FileExists(Filepath) Then
		' Get a handle to the file
		Dim file    
		set file = FSO.GetFile(Filepath)
		' Open the file
		Set TextStream = file.OpenAsTextStream(ForReading, TristateUseDefault)
		configText=TextStream.ReadAll
		lineArray = split(configText,vbCrLf)
	Else
		response.write("Config file not found")
	End If
End Function
'Atul : Function to find and replace a pattern in given string
Function RegExResults(strReplace, strPattern, strToReplaceIn)
    Set regEx = New RegExp
    regEx.Pattern = strPattern
    regEx.Global = true
	RegExResults = regEx.Replace(strToReplaceIn,strReplace)
    Set regEx = Nothing
End Function
'Atul : Format the data as per the config file
Function DataForWix(data)
	Dim Line
	For Each Line In lineArray
		myLine = Split(Line,"|")
		myLine(0) = replace(myLine(0),"__TXT__","(.*?)")
		myLine(1) = replace(myLine(1),"__TXT__","$1")
		data = RegExResults( myLine(1), myLine(0),data)
    Next
	DataForWix  = data
End Function
Call ReadConfigFile ' Reading config data
con.execute "insert into ims_usage (Email,u_datetime,action,subaction) values('sanjayz@gmail.com','" & now() & "','" & request.ServerVariables("PATH_INFO") & "','" & request.ServerVariables("QUERY_STRING") & "')"
		   dim objissue ' issue change comments 
		   set objissue =server.CreateObject("adodb.recordset") 'setting the record oobhject for issue chages goes here 
			   objissue.open "select *,if(comments is null,'',comments) newcomments from " &  varTblNameIssueChangesComments & " where issueid=" & IssueForAODM_API & " order by updateDate",con
		   dim dattemp
		   dim newcomments
		   dim flag
		   	dattemp = ""
			counter = 1
			flag = false
			if objissue.eof =false then 
				while not objissue.eof
 			 	   updatedate = objissue("UpdateDate")
					newcomments = objissue("newcomments")
					changesby = objissue("ChangesBy")
					myfield = objissue("Field")
					newvalue = objissue("NewValue")
					lastvalue = objissue("LastValue")
					rowid = objissue("Rowid")
					if  dattemp <> updatedate and newcomments = "" then
					dattemp =updatedate	 
         	 %>
					
					<%
					do
						if objissue.eof = true then
							exit do
						end if
						if (newcomments = "") and (dattemp = updatedate) then				
							dattemp = updatedate
							%>
						  
						   <% 
					   else
							exit do
					   end if
					   objissue.movenext
  						if objissue.eof = true then
							exit do
						end if
					   updatedate = objissue("UpdateDate")
					newcomments = objissue("newcomments")
					changesby = objissue("ChangesBy")
					myfield = objissue("Field")
					newvalue = objissue("NewValue")
					lastvalue = objissue("LastValue")
					rowid = objissue("Rowid")
				   loop
				   'while (not )
				   flag = true
				   
				   %>
				<% elseif newcomments <> "" then%>
					<%flag = false%>
					   <%=DataForWix(newcomments)%>	
				  	<% 
					response.Flush()	
					counter = counter + 1
					end if 
					%>
				<%
				if objissue.eof = false and flag <> true  then
		         	 objissue.movenext
				end if
				
			  wend
			  
			end if
			objissue.close
			set objissue = nothing	 
			%>