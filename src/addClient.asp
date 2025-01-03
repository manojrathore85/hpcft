<!-- #include file="checksession.asp" -->
<!-- #include file="connect.asp" --> <!-- including the connection file -->

<%
' ********************* CLIENTS CANNOT CREATE CLIENT ***************************************************

if varSpecialclientid <> "" then ' THE VARIABLE 'varSpecialclientid' IS CONTAINED IN CONNECT.ASP. VARIABLE IS NOT EMPTY WHEN ASP SCRIPTS ARE ACCESSED BY CLIENT
	response.Write("<script>alert('Clients are not allowed to add client.');window.history.go(-1);</script>")
	con.close
	set con = nothing
	response.End()
end if

'********************************************************************************************************
' ********************************************** CHECKING PERMISSION ************************************

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
' ********************* CHECKING PERMISSION ENDS*********************************************************
dim message
if request("cmdSubmit") = "Create" then
	'con.begintrans
	dim rs
	set rs = server.CreateObject("adodb.recordset")
	rs.open "select * from clientids where clientid='" & request("txtClient") & "'",con
	if rs.eof then
		rs.close
		' CREATING COPY OF MAIN TABLES FOR THE CLIENT BEING CREATED
		'con.execute "create table " & request("txtClient") & "_" & varTblNameIssues & " like " & varTblNameIssues
		'con.execute "create table " & request("txtClient") & "_" & varTblNameCorp_Act_table & " like " & varTblNameCorp_Act_table
		'con.execute "create table " & request("txtClient") & "_" & varTblNameDailyData & " like " & varTblNameDailyData
		'con.execute "create table " & request("txtClient") & "_" & varTblNameGrand_Total_Table & " like " & varTblNameGrand_Total_Table
		'con.execute "create table " & request("txtClient") & "_" & varTblNameIms_Compliance_Table & " like " & varTblNameIms_Compliance_Table
		'con.execute "create table " & request("txtClient") & "_" & varTblNameIms_Usage & " like " & varTblNameIms_Usage
		'con.execute "create table " & request("txtClient") & "_" & varTblNameIssueChangesComments & " like " & varTblNameIssueChangesComments
		'con.execute "create table " & request("txtClient") & "_" & varTblNameItem_List_table & " like " & varTblNameItem_List_table
		'con.execute "create table " & request("txtClient") & "_" & varTblNameMlp_Compliance_Table & " like " & varTblNameMlp_Compliance_Table
		'con.execute "create table " & request("txtClient") & "_" & varTblNamePermissions & " like " & varTblNamePermissions
		'con.execute "create table " & request("txtClient") & "_" & varTblNameProjects & " like " & varTblNameProjects
		'con.execute "create table " & request("txtClient") & "_" & varTblNameStatement_Tool_Table & " like " & varTblNameStatement_Tool_Table
		'con.execute "create table " & request("txtClient") & "_" & varTblNameUserProfile & " like " & varTblNameUserProfile
		'con.execute "create table " & request("txtClient") & "_" & varTblNameUsers & " like " & varTblNameUsers
		
		call createTables
		' ENTERING ADMIN PROFILE AS A USER 
		con.execute "insert into " & request("txtClient") & "_" & varTblNameUserProfile & " select * from userprofile where email = 'imsadmin@gtsims.com'"
		' ENTERING DUMMY PROJECT
		con.execute "insert into " & request("txtClient") & "_" & varTblNameProjects & "  values('','dummy','imsadmin@gtsims.com','Just to get to administration part','','','')"
'		' ALLOWING ALL THE RIGHTS TO ADMIN FOR THE DUMMY PROJECT
		con.execute "insert into " & request("txtClient") & "_" & varTblNameUsers & " values('imsadmin@gtsims.com',1,'all')"
'		' COPYING ALL THE PERMISSIONS FROM MAIN TABLES TO CLIENTS TABLE
		con.execute "insert into " & request("txtClient") & "_" & varTblNamePermissions & " select * from permissions"
'		
'		' ENTERING THE RECORD OF CLIENT CREATED IN CIENTIDS TABLE
		con.execute "insert into clientids(clientid,createDate) values('" & request("txtClient") & "',now())"
'		rs.CursorLocation = 3 'adUseClient
'	   	rs.Open "select * from clientids where 1=2",con,3,3
'	  	rs.AddNew
'		rs("clientid") = request("txtClient")
'		rs("createDate") =  now()
'		rs.update
'		rs.close
		message = "Clientid created successfully."
	else
		message = "Clientid already exists."
	end if
'	
	if rs.state = 1 then rs.close
	set rs = nothing
'	'con.committrans
	set fs = server.CreateObject("scripting.filesystemobject")
	if not fs.folderexists(server.MapPath(".") & "\clients\" & request("txtClient")) then fs.createfolder(server.MapPath(".") & "\clients\" & request("txtClient"))
	Dim ffile
	Dim ffolder
	Dim f1folder
	Set ffolder = fs.GetFolder(server.MapPath("."))
	For Each ffile In ffolder.Files
		fs.copyfile ffile,server.MapPath(".") & "\clients\" & request("txtClient") & "\"
	Next
	fs.createfolder(server.MapPath(".") & "\clients\" & request("txtClient") & "\upload")
	'fs.copyfolder server.MapPath(".") & "\monthlystattool",server.MapPath(".") & "\clients\" & request("txtClient") & "\monthlystattool"
	'fs.copyfolder server.MapPath(".") & "\GUI_files",server.MapPath(".") & "\clients\" & request("txtClient") & "\GUI_files"
	fs.copyfolder server.MapPath(".") & "\include",server.MapPath(".") & "\clients\" & request("txtClient") & "\include"
end if

sub createTables()
	dim sql
	' CREATING ISSUES TABLE	
	sql = "CREATE TABLE  IF NOT EXISTS `" & request("txtclient") & "_" & varTblNameIssues & "` ("
	sql = sql & " `Projectid` int(10) unsigned NOT NULL default '0',"
	sql = sql & " `IssueId` int(10) unsigned NOT NULL auto_increment,"
	sql = sql & " `CreateDate` datetime NOT NULL default '0000-00-00 00:00:00',"
	sql = sql & " `UpdateDate` datetime default '0000-00-00 00:00:00',"
	sql = sql & " `IssueType` varchar(32) NOT NULL default '',"
	sql = sql & " `Severity` varchar(16) NOT NULL default '',"
	sql = sql & " `Summary` text NOT NULL,"
	sql = sql & " `AssignTo` varchar(64) NOT NULL default '',"
	sql = sql & " `Reporter` varchar(64) NOT NULL default '',"
	sql = sql & " `Description` text,"
	sql = sql & " `Status` varchar(32) NOT NULL default 'Opened',"
	sql = sql & " `AttachedFilePath` text,"
	sql = sql & " `reportUsers` text,"
	sql = sql & " PRIMARY KEY  (`IssueId`)"
	sql = sql & " ) TYPE=MyISAM;"
	con.execute sql
	'CREATING Corp_Act_table 
	sql = "CREATE TABLE  IF NOT EXISTS `" & request("txtclient") & "_" & varTblNameCorp_Act_table & "` ("
	sql = sql & " `Symbol` varchar(8) NOT NULL default '',"
	sql = sql & " `SubSymbol` varchar(16) NOT NULL default '',"
	sql = sql & " `Operation` varchar(15) NOT NULL default '',"
	sql = sql & " `DateOperation` datetime NOT NULL default '0000-00-00 00:00:00',"
	sql = sql & " `ExpirationDate` datetime NOT NULL default '0000-00-00 00:00:00',"
	sql = sql & " `id` int(10) unsigned NOT NULL auto_increment,"
	sql = sql & " `Affect` varchar(15) NOT NULL default '',"
	sql = sql & " `close` char(1) default NULL,"
	sql = sql & " `Notes` text,"
	sql = sql & "  PRIMARY KEY  (`id`),"
	sql = sql & "  KEY `Index_1` (`id`,`Symbol`)"
	sql = sql & ") TYPE=MyISAM;"
	con.execute sql
	
	sql = "CREATE TABLE  IF NOT EXISTS `" & request("txtclient") & "_" & varTblNameDailyData & "` ("
	sql = sql & " `Source` varchar(32) NOT NULL default '',"
	sql = sql & " `Symbol` varchar(32) NOT NULL default '',"
	sql = sql & " `SubSymbol` varchar(32) default NULL,"
	sql = sql & " `DDate` datetime NOT NULL default '0000-00-00 00:00:00',"
	sql = sql & " `Open` float default NULL,"
	sql = sql & " `Low` float default NULL,"
	sql = sql & " `High` float default NULL,"
	sql = sql & " `Close` float default NULL,"
	sql = sql & " `Volume` decimal(10,0) default NULL,"
	sql = sql & "  PRIMARY KEY  (`Symbol`,`DDate`)"
	sql = sql & " ) TYPE=MyISAM;"
	con.execute sql
	
	sql = "CREATE TABLE  IF NOT EXISTS `" & request("txtclient") & "_" & varTblNameGrand_Total_Table & "` ("
	sql = sql & " `strategy` varchar(32) NOT NULL default '',"
	sql = sql & " `item` varchar(128) NOT NULL default '',"
	sql = sql & " `s_date` datetime NOT NULL default '0000-00-00 00:00:00',"
	sql = sql & " `HardLimit` double NOT NULL default '0',"
	sql = sql & " `SoftLimit` double NOT NULL default '0',"
	sql = sql & " `Multiplier` double NOT NULL default '0',"
	sql = sql & " `Total` double NOT NULL default '0',"
	sql = sql & " `Notes` text NOT NULL,"
	sql = sql & " `rowid` int(10) unsigned NOT NULL auto_increment,"
	sql = sql & " PRIMARY KEY  (`rowid`)"
	sql = sql & " ) TYPE=MyISAM;"
	con.execute sql
	
	sql = "CREATE TABLE  IF NOT EXISTS `" & request("txtclient") & "_" & varTblNameIms_Compliance_Table & "` ("
	sql = sql & " `Symbol` varchar(8) NOT NULL default '',"
	sql = sql & " `SubSymbol` varchar(16) NOT NULL default '',"
	sql = sql & " `Operation` varchar(15) NOT NULL default '',"
	sql = sql & " `DateOperation` datetime NOT NULL default '0000-00-00 00:00:00',"
	sql = sql & " `ExpirationDate` datetime NOT NULL default '0000-00-00 00:00:00',"
	sql = sql & " `id` int(10) unsigned NOT NULL auto_increment,"
	sql = sql & " `Affect` varchar(15) NOT NULL default '',"
	sql = sql & " `Close` char(1) default NULL,"
	sql = sql & " `Notes` text,"
	sql = sql & "  PRIMARY KEY  (`id`),"
	sql = sql & "  KEY `Index_1` (`id`,`Symbol`)"
	sql = sql & " ) TYPE=MyISAM;"
	con.execute sql
	
	sql = "CREATE TABLE  IF NOT EXISTS `" & request("txtclient") & "_" & varTblNameIms_Usage & "` ("
	sql = sql & " `Email` varchar(64) default NULL,"
	sql = sql & " `U_DateTime` datetime NOT NULL default '0000-00-00 00:00:00',"
	sql = sql & " `Action` varchar(64) default NULL,"
	sql = sql & " `SubAction` text,"
	sql = sql & " `IdKey` int(10) unsigned NOT NULL auto_increment,"
	sql = sql & "  PRIMARY KEY  (`IdKey`)"
	sql = sql & " ) TYPE=MyISAM;"
	con.execute sql
	
	
	sql = "CREATE TABLE  IF NOT EXISTS `" & request("txtclient") & "_" & varTblNameIssueChangesComments & "` ("
	sql = sql & " `ProjectId` int(10) unsigned NOT NULL default '0',"
	sql = sql & " `IssueId` int(10) unsigned NOT NULL default '0',"
	sql = sql & " `Field` varchar(32) NOT NULL default '',"
	sql = sql & " `ChangesBy` varchar(64) NOT NULL default '',"
	sql = sql & " `UpdateDate` datetime NOT NULL default '0000-00-00 00:00:00',"
	sql = sql & " `NewValue` text NOT NULL,"
	sql = sql & " `Comments` text,"
	sql = sql & " `LastValue` text NOT NULL,"
	sql = sql & " `Rowid` int(10) unsigned NOT NULL auto_increment,"
	sql = sql & " PRIMARY KEY  (`Rowid`)"
	sql = sql & " ) TYPE=MyISAM;"
	con.execute sql
	
	sql = "CREATE TABLE  IF NOT EXISTS `" & request("txtclient") & "_" & varTblNameItem_List_table & "` ("
	sql = sql & " `Item` varchar(128) default NULL,"
	sql = sql & " `Views` varchar(64) default NULL,"
	sql = sql & " `Strategy` varchar(32) default NULL"
	sql = sql & " ) TYPE=MyISAM;"
	con.execute sql
	
	sql = "CREATE TABLE  IF NOT EXISTS `" & request("txtclient") & "_" & varTblNameMlp_Compliance_Table & "` ("
	sql = sql & " `Symbol` varchar(8) NOT NULL default '',"
	sql = sql & " `SubSymbol` varchar(16) NOT NULL default '',"
	sql = sql & " `Operation` varchar(15) NOT NULL default '',"
	sql = sql & " `DateOperation` datetime NOT NULL default '0000-00-00 00:00:00',"
	sql = sql & " `ExpirationDate` datetime NOT NULL default '0000-00-00 00:00:00',"
	sql = sql & " `id` int(10) unsigned NOT NULL auto_increment,"
	sql = sql & " `Affect` varchar(15) NOT NULL default '',"
	sql = sql & " `close` char(1) default NULL,"
	sql = sql & " `Notes` text,"
	sql = sql & " PRIMARY KEY  (`id`),"
	sql = sql & " KEY `Index_1` (`id`,`Symbol`)"
	sql = sql & " ) TYPE=MyISAM;"
	con.execute sql
	
	sql = "CREATE TABLE IF NOT EXISTS  `" & request("txtclient") & "_" & varTblNamePermissions & "` ("
	sql = sql & " `PermissionId` varchar(64) NOT NULL default '',"
	sql = sql & " `PRead` char(1) default NULL,"
	sql = sql & " `PWrite` char(1) default NULL,"
	sql = sql & " `PAdd` char(1) default NULL,"
	sql = sql & " `PDelete` char(1) default NULL,"
	sql = sql & " `Description` text,"
	sql = sql & " PRIMARY KEY  (`PermissionId`)"
	sql = sql & " ) TYPE=MyISAM;"
	con.execute sql
	
	sql = "CREATE TABLE IF NOT EXISTS  `" & request("txtclient") & "_" & varTblNameProjects & "` ("
	sql = sql & " `ProjectId` int(10) unsigned NOT NULL auto_increment,"
	sql = sql & " `ProjectName` varchar(128) NOT NULL default '',"
	sql = sql & " `LeadDeveloper` varchar(64) NOT NULL default '',"
	sql = sql & " `Description` text,"
	sql = sql & " `DesignDocPath` text,"
	sql = sql & " `CreateDate` datetime NOT NULL default '0000-00-00 00:00:00',"
	sql = sql & " `ProjectWatchList` text,"
	sql = sql & "  PRIMARY KEY  (`ProjectId`)"
	sql = sql & " ) TYPE=MyISAM;"
	con.execute sql
	
	sql = "CREATE TABLE IF NOT EXISTS  `" & request("txtclient") & "_" & varTblNameStatement_Tool_Table & "` ("
	sql = sql & " `Item` varchar(128) NOT NULL default '',"
	sql = sql & " `l_date` datetime NOT NULL default '0000-00-00 00:00:00',"
	sql = sql & " `HardLimit` double NOT NULL default '0',"
	sql = sql & " `SoftLimit` double NOT NULL default '0',"
	sql = sql & " `Multiplier` double NOT NULL default '0',"
	sql = sql & " `Total` double NOT NULL default '0',"
	sql = sql & " `Notes` text NOT NULL,"
	sql = sql & " `rowid` int(10) unsigned NOT NULL auto_increment,"
	sql = sql & " PRIMARY KEY  (`rowid`)"
	sql = sql & " ) TYPE=MyISAM;"
	con.execute sql
	
	sql = "CREATE TABLE  IF NOT EXISTS `" & request("txtclient") & "_" & varTblNameUserProfile & "` ("
	sql = sql & " `Email` varchar(64) NOT NULL default '',"
	sql = sql & " `Password` varchar(16) NOT NULL default '',"
	sql = sql & " `FirstName` varchar(64) NOT NULL default '',"
	sql = sql & " `LastName` varchar(64) NOT NULL default '',"
	sql = sql & " `Address` varchar(255) default NULL,"
	sql = sql & " `Mobile` varchar(16) default NULL,"
	sql = sql & " `HPhone` varchar(16) default NULL,"
	sql = sql & " `OPhone` varchar(16) default NULL,"
	sql = sql & " `YahooId` varchar(64) default NULL,"
	sql = sql & "  PRIMARY KEY  (`Email`)"
	sql = sql & " ) TYPE=MyISAM;"
	con.execute sql
	
	sql = "CREATE TABLE IF NOT EXISTS  `" & request("txtclient") & "_" & varTblNameUsers & "` ("
	sql = sql & " `Email` varchar(64) NOT NULL default '',"
	sql = sql & " `ProjectId` int(10) unsigned NOT NULL default '0',"
	sql = sql & " `PermissionId` varchar(64) NOT NULL default '',"
	sql = sql & "  PRIMARY KEY  (`Email`,`ProjectId`)"
	sql = sql & " ) TYPE=MyISAM;"
	con.execute sql
	
	sql = "CREATE TABLE IF NOT EXISTS  `" & request("txtclient") & "_" & varTblNameReminder_table & "` ("
	  sql = sql & " `rem_idCol` int(10) unsigned NOT NULL auto_increment,"
	  sql = sql & " `issueidCol` int(10) unsigned NOT NULL default '0',"
	  sql = sql & " `TypeCol` varchar(32) NOT NULL default '',"
	  sql = sql & " `Due_DateTimeCol` datetime NOT NULL default '0000-00-00 00:00:00',"
	  sql = sql & " `TimeZoneCol` char(3) NOT NULL default '',"
	  sql = sql & " `Date_enteredCol` datetime NOT NULL default '0000-00-00 00:00:00',"
	  sql = sql & " `Person_ResponsibleCol` varchar(64) NOT NULL default '',"
	  sql = sql & " `AlertFrequencyCol` varchar(16) NOT NULL default '',"
	  sql = sql & " `StateCol` varchar(16) NOT NULL default '',"
	  sql = sql & " `NotesCol` text NOT NULL,"
	  sql = sql & " `DescCol` text NOT NULL,"
	  sql = sql & " `DaysOfWeekCol` varchar(40) NOT NULL default '',"
	  sql = sql & " `MonthsOfYearCol` varchar(60) NOT NULL default '',"
	  sql = sql & " PRIMARY KEY  (`rem_idCol`)"
	  sql = sql & " ) TYPE=MyISAM;"
	con.execute sql
	
	sql = "CREATE TABLE IF NOT EXISTS  `" & request("txtclient") & "_" & varTblNameIms_MyDiary_table & "` ("
	sql = sql & " `email` varchar(128) NOT NULL default '',"
	sql = sql & " `projectid` int(10) unsigned NOT NULL default '0',"
    sql = sql & " `issueid` int(10) unsigned NOT NULL default '0',"
	sql = sql & " `date_of_creation` datetime NOT NULL default '0000-00-00 00:00:00',"
	sql = sql & " `comments` text NOT NULL,"
	sql = sql & " `extra_field` text NOT NULL,"
	sql = sql & " PRIMARY KEY  (`email`,`projectid`,`issueid`)"
	sql = sql & " ) TYPE=MyISAM;"
	con.execute sql

  'sql = "CREATE TABLE IF NOT EXISTS `ims_po_system` ("
  'sql = sql & " `col_rowid` int(10) unsigned NOT NULL default '1',"
  'sql = sql & " `col_field_id` int(10) unsigned NOT NULL,"
  'sql = sql & " `col_column_value` text,"
  'sql = sql & " PRIMARY KEY  (`col_rowid`,`col_field_id`)"
  'sql = sql & " ) TYPE=MyISAM;"
  'con.execute sql

  'sql = "CREATE TABLE IF NOT EXISTS `ims`.`ims_po_fields` ("
  'sql = sql & " `col_field_id` int(10) unsigned NOT NULL default '1',"
  'sql = sql & " `col_field_name` varchar(45) NOT NULL,"
  'sql = sql & " PRIMARY KEY  (`col_field_id`)"
  'sql = sql & " ) TYPE=MyISAM;"
  'con.execute sql
  	
end sub
%>


<%
     dim i ' as incremental variable
   dim objview
   set objview =server.CreateObject("adodb.recordset") ' creating the recordset 
   objview.open "select * from clientids order by createdate desc ",con
%>
<form action="viewprojects.asp" name="frmAddClient">
<input type="hidden" value="addClient" name="mainlink">
<table width="95%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td class="SectionHead">Administration<br>
      &nbsp; </td>
  </tr>
  <tr> 
    <td> 
	<p align="center" class="redbold"><%=message%></p>
	<br>
      ----Add Client : <input type="text" name="txtClient"> <input type="submit" name="cmdSubmit" value="Create" class="formbutton">
  	</td>
  </tr>
</table>
</form>
<BR>
<table width="95%" border="0" cellpadding="0" cellspacing="0" bgcolor="#CCCCCC">
<tr>
  <td><table width="100%" border="0" cellpadding="1" cellspacing="1">
	  <tr bgcolor="#FFFFFF" class=head> 
		<td id=head>&nbsp;Client Id</td>
		<td id=head>&nbsp;Create Date</td>
	  </tr>
 <% 
    if objview.bof=false and objview.eof =false then ' checking whether the record is exist or not
	   while not objview.eof ' checkout all the user in database
%>	 
	  <tr bgcolor="#FFFFFF"> 
		  <td>&nbsp;<%=objview("Clientid")%></td>
		  <td>&nbsp;<%=DateAdd("h", -5, objview("createDate"))%></td>
	  </tr>
   <%
	    objview.movenext 
		wend
	else
	%>
	<tr bgcolor="#FFFFFF"> 
    <td colspan="4" align="center">NO CLIENTS CURRENTLY</td></tr>
	<%

    end if		
   %>

	</table></td>
</tr>
</table>
<%
  'nullifying all the objects
  objview.close
  set objview= nothing
  con.close
  set con = nothing
 %> 
