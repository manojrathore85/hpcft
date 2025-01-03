<!-- #include file="genvariables.asp" -->
<%

' ******************************************************************************************************

' *****************      CONNECTING TO DATABASE ********************************************************
dim con
set con = server.CreateObject("adodb.connection")
con.connectionstring = "Driver={MySQL ODBC 9.1 Unicode Driver}; Server=localhost; uid=root; pwd=; database=cms_ims;"
con.open 
'con.Charset = "UTF-8"
'*********************************************************************************************************

'***********************   SEARCHING CLIENT ID  **********************************************************

Dim varSpecialsttr
Dim varSpecialclientid
Dim varSpecialstartPosClientId
Dim varSpecialendPosClientId
Dim varSpeciallast5CharOfDomain
Dim myArray(1)
myArray(0) = 311
myArray(1) = 313
'IssueForAODM_API= "311,313"

varSpeciallast5CharOfDomain = "ents/"   ' http://shivankz.gtsims.com/ims/clients/....
varSpecialsttr = request.ServerVariables("PATH_INFO")
varSpecialstartPosClientId = InStr(1, varSpecialsttr, varSpeciallast5CharOfDomain)
if varSpecialstartPosClientId <> 0 then
	varSpecialendPosClientId = InStr(varSpecialstartPosClientId + Len(varSpeciallast5CharOfDomain), varSpecialsttr, "/")
else
	varSpecialendPosClientId = 0
end if
if varSpecialendPosClientId <> 0 then
	varSpecialclientid = Mid(varSpecialsttr, varSpecialstartPosClientId + Len(varSpeciallast5CharOfDomain), varSpecialendPosClientId - (varSpecialstartPosClientId + Len(varSpeciallast5CharOfDomain)))
	dim rsSearchClient
	set rsSearchClient = server.CreateObject("adodb.recordset")
	rsSearchClient.open "select clientid from clientids where clientid = '" & varSpecialclientid & "'",con
	if rsSearchClient.eof then
		varSpecialclientid = ""
	end if
	rsSearchClient.close
	set rsSearchClient = nothing
else
	varSpecialclientid = ""
end if
if varSpecialclientid <> "" then
	varSpecialclientid = varSpecialclientid & "_"
end if

'*********************************************************************************************************
' ******************* VARIABLE TABLE NAMES *************************************************************
	dim varTblNameClientids
	dim varTblNameIssues 
	dim varTblNameCorp_Act_table 
	dim varTblNameDailyData
	dim varTblNameGrand_Total_Table 
	dim varTblNameIms_Compliance_Table 
	dim varTblNameIms_Usage 
	dim varTblNameIssueChangesComments
	dim varTblNameItem_List_table 
	dim varTblNameMlp_Compliance_Table 
	dim varTblNamePermissions 
	dim varTblNameProjects 
	dim varTblNameStatement_Tool_Table 
	dim varTblNameUserProfile 
	dim varTblNameUsers
	dim varTblNameReminder_table
	dim varTblNameIms_MyDiary_table
	dim varTblNameIms_po_system_table
	dim varTblNameIms_po_fields_table
	
	varTblNameIssues = varSpecialclientid  & "issues"
	varTblNameCorp_Act_table = varSpecialclientid  & "corp_act_table"
	varTblNameDailyData = varSpecialclientid  & "dailydata"
	varTblNameGrand_Total_Table = varSpecialclientid  & "grand_total_table"
	varTblNameIms_Compliance_Table = varSpecialclientid  & "ims_compliance_table"
	varTblNameIms_Usage = varSpecialclientid  & "ims_usage"
	varTblNameIssueChangesComments= varSpecialclientid  & "issuechangescomments"
	varTblNameItem_List_table = varSpecialclientid  & "item_list_table"
	varTblNameMlp_Compliance_Table = varSpecialclientid  & "mlp_compliance_table"
	varTblNamePermissions = varSpecialclientid  & "permissions"
	varTblNameProjects = varSpecialclientid  & "projects"
	varTblNameStatement_Tool_Table = varSpecialclientid  & "statement_tool_table"
	varTblNameUserProfile = varSpecialclientid  & "userprofile"
	varTblNameUsers = varSpecialclientid  & "users"
	varTblNameReminder_table = varSpecialclientid  & "reminder_table"
	varTblNameIms_MyDiary_table = varSpecialclientid  & "ims_mydiary"
	varTblNameIms_po_system_table = varSpecialclientid  & "ims_po_system"
	varTblNameIms_po_fields_table = varSpecialclientid  & "ims_po_fields"
	
'******************************************************************************************************
'*******************     ACTIVITY LOGGING PROCESS  ****************************************************

if instr(request.ServerVariables("PATH_INFO"),"loginprocess.asp") = 0 and instr(request.ServerVariables("PATH_INFO"),"loginprocess_so.asp") = 0 then
dim rsIms_Usage1
dim varConnectProjectId
if not isempty(session("projectid")) then
	varConnectProjectId = "Projectid=" & session("projectid") & ","
end if
set rsIms_Usage1 = server.CreateObject("adodb.recordset")
rsIms_Usage1.cursorlocation = 3
rsIms_Usage1.open "select * from " &  varTblNameIms_Usage & " where 1=2",con,3,3
    rsIms_Usage1.AddNew 
	rsIms_Usage1("email") = session("user")
	rsIms_Usage1("u_datetime") = now()
	if instr(request.ServerVariables("PATH_INFO"),"logout.asp") <> 0 then  
		rsIms_Usage1("action")= "logout"
	else
		rsIms_Usage1("action")= request.ServerVariables("PATH_INFO")
	end if
	
	rsIms_Usage1("subaction") = varConnectProjectId &  request.ServerVariables("QUERY_STRING")
	rsIms_Usage1.update
	rsIms_Usage1.close
	set rsIms_Usage1 = nothing
	' CHECKING IF USER ACCESSING THE SAME CLIENT OR NOT ( AS IS LOGGED ON )
	if session("client") <> varSpecialclientid then
		con.close
		set con = nothing
'		UrlToRedirect = varSiteSpecURL & Mid(request.ServerVariables("PATH_INFO"), 1, InStrRev(request.ServerVariables("PATH_INFO"), "/") & "login_index.asp")
		response.Redirect(varSiteSpecURL & Mid(request.ServerVariables("PATH_INFO"), 1, InStrRev(request.ServerVariables("PATH_INFO"), "/")) & "login_index.asp?message=You are currently logged in as '" & session("client") & "'." & "<br>" & "You are trying to access '" & varSpecialclientid & "'<br>" & "Please login again.")
	end if
	
	'con.execute "insert into ims_usage (Email,u_datetime,action,subaction) values('" & session("user") & "','date(" & now() & ")','" & request.ServerVariables("PATH_INFO") & "','" & request.ServerVariables("QUERY_STRING") & "')"
	
end if

'************** FUNCTION TO GET THE PROJECT NAME ******************************************
function getProjectName
	dim rsProjectName
	set rsProjectName = server.CreateObject("adodb.recordset")
	rsProjectName.open "select projectname from projects where projectid = " & session("projectid"),con
	if not rsProjectName.eof then
		getProjectName = rsProjectName("projectname")
	end if
	rsProjectName.close
	set rsProjectName = nothing
end function	
%>

