<!--  Description this the file which show the project which is show   Created on date 24 march 2k5 rajat-->
<!-- #include file="Connect.asp" --> <!-- includeing the connection file -->
<!-- checking the session -->
<!-- #include file=" checksession.asp" --> <!-- it checks the session whether the valid user or not -->

<!-- ******************************************************** script language  Vb script coding goes here**************** -->
<%
     dim objview 'for viewing the projects
	 set objview =server.CreateObject("adodb.recordset") 
	 objview.open "select * from " &  varTblNameProjects & " where projectid=" & session("projectid") ,con ' recodset of projects 
	  
%>
<!-- ***************************************  html coding goes here ************************************ -->
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<LINK href="include/style.css" type="text/css" rel="stylesheet">
</HEAD>
<BODY>
	
		
		 
					<% 
					  if objview.eof =false and  objview.bof =false then
					   while not objview.eof
					 %>  
					 <table width="50%" border="1" cellpadding="1" cellspacing="1">
								 <tr bgcolor="#CCCCCC"> 
										<td  width="30%">&nbsp;Project</td>
										<td class="SectionHead" align="center" width="70%"><%=objview("projectName")%></td>
								</tr>
								<tr>
										 <td>&nbsp;Lead</td>
										<td><%=objview("LeadDeveloper")%></td>
								</tr>			 
								<tr>
										 <td>&nbsp;Description</td>
										  <td ><%=objview("description")%></td>
						
								</tr>
								<tr>
								    <td><a href="createissue.asp?projectId=<%=objview("projectId")%>">New Issue</a> </td>
								</tr>
					</table>
					<p></p>
				<%
				    objview.movenext
					wend
					end if
					objview.close
					set objview = nothing
				%>
					 
		
</BODY>
</HTML>
