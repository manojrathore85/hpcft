<% 'Option Explicit
response.write("<a href='https://www.artofdecisionmaking.org/'>click me</a>")
response.end
Function RegExResults(strTarget, strPattern,strToReplace)

    Set regEx = New RegExp
    regEx.Pattern = strPattern
    regEx.Global = true
    'Set RegExResults = regEx.Execute(strTarget)
	RegExResults = regEx.Replace(strToReplace,strTarget)
    Set regEx = Nothing
	
	response.write(server.HTMLEncode(RegExResults) & "<br><br>")

End Function

Const Filename = "aodm_config.txt"    ' file to read
Const ForReading = 1, ForWriting = 2, ForAppending = 3
Const TristateUseDefault = -2, TristateTrue = -1, TristateFalse = 0

' Create a filesystem object
Dim FSO
set FSO = server.createObject("Scripting.FileSystemObject")
' Map the logical path to the physical system path
Dim Filepath
Filepath = Server.MapPath(Filename)

if FSO.FileExists(Filepath) Then
    ' Get a handle to the file
    Dim file    
    set file = FSO.GetFile(Filepath)

    ' Get some info about the file
    Dim FileSize
    FileSize = file.Size

    Response.Write "<p><b>File: " & Filename & " (size " & FileSize  &_
                   " bytes)</b></p><hr>"
    Response.Write "<pre>"

    ' Open the file
    Dim TextStream
    Set TextStream = file.OpenAsTextStream(ForReading, TristateUseDefault)

    ' Read the file line by line
	strtoreplace = "<header1>Princeton University1</header1><table version=1>"&_
"<row> <name>Financial Information</name> <value> Please visit website https://www.youtube.com// </value> </row>"&_

"<row>"&_
"<name> Admissions </name> <value> https://admission.princeton.edu/ </value> </row>"&_

"<row>"&_
"<name> Popular Alumni </name> <value> https://studyabroad.shiksha.com/list-of-princeton-university-notable-alumni-articlepage-2065#:~:text=Princeton%20University%20has%20a%20notable,and%20Commander%20of%20Apollo%2012). </value> </row>"&_

"<row>"&_
"<name> Campus Tour </name> <value> https://www.youtube.com/watch?v=ccRnkLnwUZg"&_
"</value> </row>"&_
"<row>"&_
"<name> Academic programs </name> <value> https://admission.princeton.edu/academics/degrees-departments"&_
"</value> </row>"&_
"</table>"

strtoreplace = "<row>"&_
"<name> Clubs and Organizations </name> <value>https://admission.princeton.edu/community/clubs-and-organizations"&_
"</value> </row>"

    Do While Not TextStream.AtEndOfStream
        Dim Line, strtoreplace
        Line = TextStream.readline
    
        ' Do something with "Line"
        'Line = Line & vbCRLF
        'Response.write Line 
		
		'Pass the original string and pattern into the function and get a collection object back'
		
		
		
		myLine = Split(Line,"|")
		myLine(0) = replace(myLine(0),"__TXT__","(.*?)")
		myLine(1) = replace(myLine(1),"__TXT__","$1")
		'Set arrResults = RegExResults(myLine(0), myLine(1),strtoreplace)
		strtoreplace = RegExResults( myLine(1), myLine(0),strtoreplace)
		'In your pattern the answer is the first group, so all you need is'
		'For each result in arrResults
		'	Response.Write(result.Submatches(0))
		'Next

		
    Loop


    Response.Write strtoreplace

    Set TextStream = nothing
    
Else

    Response.Write "<h3><i><font color=red> File " & Filename &_
                       " does not exist</font></i></h3>"

End If

Set FSO = nothing
%>