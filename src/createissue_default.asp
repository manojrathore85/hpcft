<!-- #include file="connect.asp" -->
<!-- this the file which is basically for the  default  issue  entry scree -->


<!-- #include file="checkSession.asp" --> <!-- including the session checking file -->

<!--  ************************************* asp coding  ******************************* rajat -->
 <%
   dim projectId  'for storing the selected project id
   
	projectID=request("projectID") 'assigning the project Id
   
   dim objproject ' for viewing the project
   set objproject =server.CreateObject("adodb.recordset")
   objproject.Open "select * from " &  varTblNameProjects & " order by projectName" ,con
    
 
 %>

<!-- ******************************************* html part  statred here -- 24 march 2k45  ************ -->
<html>
<head>
<title></title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<LINK href="include/style.css" type="text/css" rel="stylesheet">
</head>

<body>
 
<table bgcolor="#CCCCCC" width="57%"  align="center">
  <form name="frm" method="post" action="createissue.asp">
			  <tr>
				 <td colspan="2"> <p class="SectionHead">Create Issue</p><p class="regularText"> Step 1 of 2: Choose the project and issue type... </p> </td>
			  </tr>
			  <tr>
				 <td width="24%" class="regularText">* Project:  </td>
				 
      <td width="76%">&nbsp;&nbsp;
<select class="formTextbox" name="lstProjects">
					               
					                
					                
					                       <% 
												if objproject.EOF=false and objproject.BOF =false then ' checking whether the records is exist or not
													 while not objproject.EOF 
													 
												%>
					                        
													<option value=<%=objproject("projectId")%>><%=objproject("projectName")%></option>
													
												 <%
												     objproject.MoveNext 
												     wend 
												     end if
											     objproject.Close 
											     set objproject = nothing
											     con.close
											     set con = nothing
											 %>
					                         
					              </select></td>

			  </tr>
			  <tr>
			          <td>* Issue Type:</td>
					   <td height="35">
					   		;&nbsp; <select class="formTextbox" name="lstIssTypes">
								        
								        <option value="BUG">Bug</option>
								        <option value="NEW">New Feature</option>
								        <option value="TASK">Task</option>
								        <option value="IMPROVEMENT">Improvement</option>
								      </select>
					  </td>
			  </tr>
			  <tr>
			    <td colspan="2" align="center"><input type="submit" name="cmdnext" value="  Next >>"> </td>
			  </tr>
		  </form>    
 </table>

</body>
</html>
<!-- SCript  language java script for highlighting the selected project in lstprojects *************** 24 march 2k5 Rajat Jaiswal -->
<script language ="javascript">
 <% if projectId<>"" then %>
  for(i=1;i<=frm.lstProjects.options.length-1;i++)
    if(frm.lstProjects.options[i].value==<%=projectId%>)
       frm.lstProjects.selectedIndex=i;
   <% end if %>         
</script>