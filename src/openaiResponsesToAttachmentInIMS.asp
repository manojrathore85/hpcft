<%
Dim fso, csvFile, fileName, inputData

' Function to check if the file exists and rename if needed
Function GetUniqueFilePath(baseFilePath)
    Dim fileExists, newFilePath, counter, extension, fileName, currentTimestamp
    counter = 1
    extension = ".csv"
    
    ' Extract the base file name without extension
    fileName = Left(baseFilePath, Len(baseFilePath) - Len(extension))

    ' Check if file exists
    Do While fso.FileExists(baseFilePath)
        ' Append timestamp or counter to create a new unique file name
        currentTimestamp = Replace(Replace(Now(), "/", ""), ":", "") ' Remove slashes and colons
        newFilePath = fileName & "_" & currentTimestamp & "_" & counter & extension
        baseFilePath = newFilePath
        counter = counter + 1
    Loop

    GetUniqueFilePath = baseFilePath
End Function

' Get the unique file path (renaming if necessary)
baseFilePath = "/upload/"


' Create FileSystemObject to check if the file exists
Set fso = Server.CreateObject("Scripting.FileSystemObject")

if request.form("username") <> "" and request.form("csvData") <> "" Then
	baseFilePath = Server.MapPath("/upload/" & request.form("username"))
	if not fso.FolderExists(baseFilePath) then 
		fso.createfolder(baseFilePath)
	end if
	csvFilePath = GetUniqueFilePath(baseFilePath & "/ResultsWindow.csv")
	' Create the new file and write the data
	Set csvFile = Server.CreateObject("ADODB.Stream")
	csvFile.Type = 2 ' Text mode
	csvFile.Charset = "UTF-8" ' Set UTF-8 encoding
	csvFile.Open
	csvFile.WriteText request.form("csvData") ' Write the data to the stream

	' Save the stream to a file
	'csvFile.SaveToFile csvFilePath, 1 ' 2 = Create new file, 1 = Overwrite
	csvFile.SaveToFile csvFilePath, 1 ' 2 = Create new file, 1 = Overwrite
	csvFile.Close

	' Respond back to the user
	Response.Write "File saved on server to the user: " & Request.Form("username") 

	' Clean up
	Set csvFile = Nothing
	Set fso = Nothing
end if 	


' Read the incoming JSON data from the request body
'inputData = BinaryToString(Request.BinaryRead(Request.TotalBytes))






' Function to convert Binary data to String
Function BinaryToString(binaryData)
    Dim objStream
    Set objStream = Server.CreateObject("ADODB.Stream")
    objStream.Type = 1 ' adTypeBinary
    objStream.Open
    objStream.Write binaryData
    objStream.Position = 0
    objStream.Type = 2 ' adTypeText
    objStream.Charset = "utf-8"
    BinaryToString = objStream.ReadText
    objStream.Close
    Set objStream = Nothing
End Function

%>
