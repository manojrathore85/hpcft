<%
' CR20060629	gtsAtul		Modifies session("projectid") for selected issueid 
' CR20060706	gtsAtul		Javascript adjustment for breaking a long script for visibility enhancements 
' CR20060713	gtsAtul		Added link to edit any particular comment
' CR20060812	gtsAtul		correction made, for more then one link in a comment not showing correctly
' CR20060826	gtsAtul		file deleted physically are shown strike out
' CR20060902	gtsAtul		functions(formatData , strike_out_deleted_file , check_file_ondisk) moved in generalFunctions.asp

'Response.CharSet = "utf-8"
%>
<!--  ******************************************************* issue details goes here ******************* date 1 apr 2k5 --->
<!-- #include file="checksession.asp" -->
<!-- #include file ="Connect.asp" --> <!-- includeing the database fie -->

<%
Session.CodePage  = 65001
 response.Buffer = true
' ********************* CHECKING PERMISSION ***************************

dim rsperm
set rsperm = server.CreateObject("adodb.recordset")
rsperm.open "select perm.* from " &  varTblNamePermissions & " perm, " &  varTblNameUsers & " u, " &  varTblNameissues & _
 " i where i.projectid = u.projectid and u.permissionid = perm.permissionid and i.issueid = " &  request("issueid") & _
 " and u.email ='" &  session("user") & "'",con
if not rsperm.eof then
	if rsperm("pread") = "T" then
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
if request("issueId")<>"" then
	
	  dim projectname,createdate,updatedate,issuetype,severity,summary,assignto,reporter,desciption,status,attachedfilepath ' genral variable for the issue details 
      dim counter, filename
	  dim objview  'define a variable for issues details from issue and project name  
	  dim objproject
	  set objview =server.CreateObject("adodb.recordset") 'define the object of view 
	  set objproject =server.CreateObject("adodb.recordset") 'define the object of view 
	  dim strsql
	  
	  ' After copy issue feature urls in old email should work
	  if request("prjid") = "" then
		  strsql = "select I.*,p.projectid,p.projectname from " &  varTblNameIssues & " I," &  varTblNameProjects & " p where I.projectId=p.projectId and issueId=" & request("issueId") '& " and I.projectid = " & request("prjid")
	  else
		  strsql = "select I.*,p.projectid,p.projectname from " &  varTblNameIssues & " I," &  varTblNameProjects & " p where I.projectId=p.projectId and issueId=" & request("issueId") & " and I.projectid = " & request("prjid")
	  end if	
	  'response.Write(strsql)
	  'response.End()
	  
 			 objview.open  strsql,con		'5Jul06Atul
			if objview.eof =false then
					   projectname =objview("projectname")
					   createdate=DateAdd("h", +1, objview("createdate")) 'for showing eastern standard time as it is 1 hrs + then American/Chicago time ( GMT - 6 HRS )
					   updatedate =DateAdd("h", +1, objview("updateDate"))'for showing eastern standard time as it is 1 hrs + then American/Chicago time ( GMT - 6 HRS )
					   issuetype=objview("issuetype")
					   severity=objview("severity")
					   summary =objview("summary")
					   assignto =objview("assignto")
					   reporter=objview("reporter")
					   attachedfilepath = objview("attachedfilepath")
					   if isnull(attachedfilepath) = false then
							filename = split(attachedfilepath,",")
					   end if
					   description= objview("description")
					   status =objview("status")
					    session("projectid") = objview("projectid")	'5Jul06Atul
		  end if
		' now nullifying the object 
		   objview.close
		   set objview =nothing
		' recordset for filling the combo with project name
		' for moving issue to selected project option 
		   objproject.Open "select p.projectid, projectname from " &  varTblNameProjects & " p," &  varTblNameUsers & " u where p.projectid =u.projectid and u.email = '" & session("user") & "' and (permissionid = 'project_mgnt' or permissionid = 'project_user' or permissionid = 'all')",con
		   
end if

sub list_user_projects
	on error resume next
	objproject.movefirst	
	if not objproject.eof then
		while not objproject.eof
	%>
			<option value=<%=objproject("projectId")%>><%=objproject("projectName")%></option>
	<%
			 objproject.MoveNext 
		 wend 
	end if
end sub
%>


                    <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
                      <tr> 
                        <td width="18%" valign="top">
						<table width="90%" border="0" align="right" cellpadding="0" cellspacing="0">
                            <tr> 
                              <td>&nbsp;</td>
                            </tr>
                            <tr><td><a href="ims_index.asp">Project Home</a><br>&nbsp;</td></tr>
                            <tr> 
                              <td><a href="#" onClick="PopUp('editissue.asp?issueid=<%=request("issueid")%>&prjid=<%=request("prjid")%>','EditIssue','status=yes,scrollbars=yes,width=800,height=550')"><b>Edit 
                                Issue</b></a><br>&nbsp;</td>
                            </tr>
							<tr> 
                              <td><a href="#" onClick="PopUp('editreminder.asp?issueid=<%=request("issueid")%>&prjid=<%=request("prjid")%>','EditReminders','status=yes,scrollbars=yes,width=' + screen.width + ',height=' + screen.height )"><b>Edit 
                                Reminders</b></a><br>&nbsp;</td>
                            </tr>
                            <tr> 
                              <td><a href="deleteissue.asp?issueid=<%=request("issueid")%>&prjid=<%=request("prjid")%>" onClick="return checkDeleteissue();">Delete 
                                Issue</a><br>&nbsp;</td>
                            </tr>
                            
	   					    <% if status = "Opened" or status = "Reopened" or status = "Stop Progress" then %>
							<tr><td><a href="changestatusissue.asp?issueid=<%=request("issueid")%>&prjid=<%=request("prjid")%>&status=startprogress">Start Progress</a><br>&nbsp;</td></tr>
							<tr><td><a href="changestatusissue.asp?issueid=<%=request("issueid")%>&prjid=<%=request("prjid")%>&status=resolve">Resolve Issue</a><br>&nbsp;</td></tr>
							<tr><td><a href="changestatusissue.asp?issueid=<%=request("issueid")%>&prjid=<%=request("prjid")%>&status=close">Close Issue</a><br>&nbsp;</td></tr>
							<% elseif status = "In Progress" then%>
							<tr><td><a href="changestatusissue.asp?issueid=<%=request("issueid")%>&prjid=<%=request("prjid")%>&status=stopprogress">Stop Progress</a><br>&nbsp;</td></tr>
							<tr><td><a href="changestatusissue.asp?issueid=<%=request("issueid")%>&prjid=<%=request("prjid")%>&status=resolve">Resolve Issue</a><br>&nbsp;</td></tr>
							<tr><td><a href="changestatusissue.asp?issueid=<%=request("issueid")%>&prjid=<%=request("prjid")%>&status=close">Close Issue</a><br>&nbsp;</td></tr>
							<%end if%>
							<% if status = "Resolved" or status = "Closed" then %>
							<tr>
                              <td><a href="changestatusissue.asp?issueid=<%=request("issueid")%>&prjid=<%=request("prjid")%>&status=reopen">Reopen 
                                Issue</a><br>&nbsp;</td>
                            </tr>
							<% end if %>
							<tr>
                              <!--<td><a href="addtodiary.asp?issueid=<%=request("issueid")%>&prjid=<%=request("prjid")%>" target="_blank">EditMyDiary</a><br>&nbsp;</td>-->
							  <td><a href="#" onClick="PopUp('ims_po_system.asp?issueid=<%=request("issueid")%>&prjid=<%=request("prjid")%>','EditIssue','status=yes,scrollbars=yes,width=800,height=550')">Tracking System</a><br>&nbsp;</td>
                            </tr>
							<tr>
                              <!--<td><a href="addtodiary.asp?issueid=<%=request("issueid")%>&prjid=<%=request("prjid")%>" target="_blank">EditMyDiary</a><br>&nbsp;</td>-->
							  <td><a href="#" onclick="editmydiary();">Add to diary</a><br>&nbsp;</td>
                            </tr>
							<tr>
							  <td><a href="chat.html" target="_blank">Chat GPT</a><br>&nbsp;</td>
                            </tr>
							<tr id="id_edit_my_diary" style="visibility:hidden">
								<td>
									<form name="frm_diary" method="post" action="addtodiary.asp" target="_blank" onsubmit="id_edit_my_diary.style.visibility='hidden'">
									<input type="hidden" name="issueid" value="<%=request("issueid")%>" />
									<input type="hidden" name="prjid" value="<%=request("prjid")%>" />
									<input type="hidden" name="mode" value="add" />
									<textarea name="comments" cols="10" rows="5" class="textarea"></textarea><br />
									<input type="submit" name="cmd_ok" value="Ok" />
									<input type="button" name="cmd_cancel" value="Cancel" onclick="cancel_diary()" />
									</form>
								</td>
							</tr>

                          </table>
						</td>
						<td width="82%" valign="top" align="center"><br>
						<form name="frm" method="post" action="moveissue.asp">
						<input type="hidden" value="<%=request("issueid")%>" name="issueid" />
						<table	border="0" cellspacing="0" cellpadding="0" width="100%">
			 <tr>
			    		<td class="bluebold" colspan="2" align="LEFT"><font size="+1"><%=projectname%></font>
						<br>
						&nbsp;&nbsp;
						<span class="regularText"><%=summary%></span><br>
						<span class="regularTextSmall">Created:</span><span class="regularTextBlueSmall">&nbsp;<%=createdate%> EST</span><span class="regularTextSmall">&nbsp;&nbsp;Updated:</span>&nbsp;<span class="regularTextBlueSmall"><%=updateDate%> EST</span><br>&nbsp;
						</td> <!-- displaying the project name on which the issue is created  -->
			 </tr>
			 <tr>
						<td class="regularTextBold">&nbsp;&nbsp;Issue Type </td> <!-- issue type of the issue -->
						<td>:&nbsp;&nbsp;<%=issuetype%> </td> <!-- severity of the issue goes here -->
			 </tr>
			 <tr>
					   	<td class="regularTextBold">&nbsp;&nbsp;Severity</td> <!-- assign to details goes here -->
						<td>:&nbsp;&nbsp;<%=severity%> </td>
			 </tr>
			  <tr>
			           <td class="regularTextBold">&nbsp;&nbsp;Reporter</td>
						<td>:&nbsp;&nbsp;<%=reporter%></td>
			 </tr>
			 <tr>  
			            <td class="regularTextBold">&nbsp;&nbsp;Assign To</td>
						<td>:&nbsp;&nbsp;<%=Assignto%></td>
			 </tr>
			 <tr>  
			            <td class="regularTextBold">&nbsp;&nbsp;Report Updation to me</td>
						<td>:&nbsp;&nbsp;<%server.Execute("issuereportflag.asp")%></td>
			 </tr>			 
			 <tr>  
			            <td class="regularTextBold">&nbsp;&nbsp;Status</td>
						<td>:&nbsp;&nbsp;<span class="redbold"><%=status%></span></td>
			 </tr>
			 <% response.Flush()%>
 			 <tr>  
			            <td class="regularTextBold">&nbsp;&nbsp;Move Issue to</td>
						<td>:&nbsp;&nbsp;<span class="redbold"><select name="selProject" class="listbox">
						<option value="none" selected="selected">Select Project</option>
						<%
							call list_user_projects
						%>
						<option>
						</select>
						&nbsp;&nbsp;<input type="submit" name="cmdMove_Issue" value=" Move " onClick="return checkListValue(frm.selProject);"></span></td>
			 </tr>
  			 <tr>  
			            <td class="regularTextBold">&nbsp;&nbsp;Copy Issue to</td>
						<td>:&nbsp;&nbsp;<span class="redbold"><select name="selProject_to_copy" class="listbox">
						<option value="none" selected="selected">Select Project</option>
						<%
							call list_user_projects
						%>
						<option>
						</select>
						&nbsp;&nbsp;<input type="submit" name="cmdCopy_Issue" value=" Copy " onClick="return checkListValue(frm.selProject_to_copy);"></span></td>
			 </tr>
			 <tr>  
			            <td class="regularTextBold">&nbsp;&nbsp;File Attached</td>
						<td valign="middle">
						<table width="100%" cellpadding="0" cellspacing="0">
						 <tr>
						 <td width="2%">:&nbsp;&nbsp;</td>
							<td width="98%"><%
				if isArray(filename) then
					if ubound(filename) >= 0 then
						for counter=0 to ubound(filename)
						if counter > 0 then
							response.Write(", ")
						end if
						%>
						<a href="getfile.asp?fn=<%=filename(counter)%>" target="_blank"><%=check_file_ondisk(filename(counter))%><%'=filename(counter)%></a>
						<%
						next
					end if
				end if
					%>
			</td>
						 </tr>
						 </table>
						
				</td>
			 </tr>
 			 <tr>  
			            <td class="regularTextBold">&nbsp;&nbsp;Decrypt Key</td>
						<td>:&nbsp;&nbsp;<span class="redbold"><input type="password" name="deKey" size="30">&nbsp;&nbsp;<input type="button" name="" value="Decode" onClick="decodeIssue();"></span></td>
			 </tr>
			 <tr>
			 
			             
    <td class="regularTextBold" colspan="2">&nbsp;&nbsp;Description :- </td>
			</tr>
			<tr>
						 <td colspan="2">
						 <table width="100%" cellpadding="0" cellspacing="0">
						 <tr>
						 	<td width="3%">&nbsp;</td>
							<td width="97%"><%
							if IsHTMLContent(description) then
									response.write(description)
								else
									response.write(formatdata(description))
								end if
							%></td>
						 </tr>
						 </table>
						 </td>
			 </tr>
			 
		 
	</table></form><br>

	<!-- *************************** showing the all the comment of the issue if exist  order by date of the update which is goes here  ************* -->
	<%
		   dim objissue ' issue change comments 
		   set objissue =server.CreateObject("adodb.recordset") 'setting the record oobhject for issue chages goes here 
			   objissue.open "select *,if(comments is null,'',comments) newcomments from " &  varTblNameIssueChangesComments & " where issueid=" & request("issueId") & " order by updateDate",con
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
					<table border=1 width="100%" cellpadding="0" cellspacing="0">
					   <tr class="regularTextBold"> 	
						<td colspan="3"> Change By : &nbsp; <%=changesby%> &nbsp; &nbsp; [ <%= DateAdd("h", +1, updatedate)%> EST ] &nbsp;</td>
						</tr>
						<tr  bgcolor="#CCCCCC" class="regularTextBold">
								<td width="20%">Field</td>
								<td width="40%"> Old Value </td>
								<td width="40%"> New Value</td>
						</tr>
					<%
					do
						if objissue.eof = true then
							exit do
						end if
						if (newcomments = "") and (dattemp = updatedate) then				
							dattemp = updatedate
							%>
						   <tr>
								<td><%=myfield%></td> 
								<td><%=formatdata(lastvalue)%></td>
								<td><%=formatdata(newvalue)%></td>
						   </tr>
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
				   response.Write("</table><br>")
				   response.Flush()
				   %>
				<% elseif newcomments <> "" then%>
					<%flag = false%>
					   <table border=1 width="100%">
					   <tr bgcolor="#CCCCCC" class="regularTextBold"> 	
						<td colspan="3"> Comments By : &nbsp; <%=changesby%> &nbsp; &nbsp; [ <%= DateAdd("h", +1, updatedate)%> EST ] &nbsp;<a href="deleteissuecomments.asp?id=<%=rowid%>&issueid=<%=request("issueId")%>" onClick="return checkDelete();">Delete</a> | <a href="#" onClick="PopUp('editissuecomments.asp?id=<%=rowid%>&issueid=<%=request("issueid")%>','EditReminders','status=yes,scrollbars=yes,width=700,height=400')"><b>Edit</b></a> | <a href="#" onclick="Decrypt(this)">Decrypt</a></td>
						</tr>
						<tr>
						<td colspan="3" id="comnt<%'=counter%>" style="overflow:scroll;"><%
						if IsHTMLContent(newcomments) then
									response.write(newcomments)
								else
									response.write(formatdata(newcomments))
								end if
						
						%></td>	
						</tr></table><br>	
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
			  %>
			  <div id="dynamicComment"></div>
              <table border=0 width="100%" cellpadding="0" cellspacing="0">
					   <tr class="regularTextBold">
						<td colspan="3">
                        <!--<input type="button" value="Add Comment" id="btn_add_comment"/> -->
                        <div id="addcomment" style="visibility:visible;width:100%">
                      <form id="myForm" action="comment.asp" method="post"> 
    <textarea name="txtcomments" id="txtcomments" class="noclass" style="width:100%;" rows="10"></textarea> <br>
    <input type="hidden" name="projectid" value="<%=session("projectid")%>" />
	<input type="hidden" name="userid" value="<%=session("user")%>" />
	<input type="hidden" name="issueid" value="<%=request("issueId")%>" />
	
    <input type="submit" value="Submit Comment" id="submit_comment" name="submit_comment" /> 
</form></div></td></tr></table>
			  <%
			 else
			 %>
             <div id="dynamicComment"></div>
              <table border=0 width="100%" cellpadding="0" cellspacing="0">
					   <tr class="regularTextBold">
						<td colspan="3">
                        <!--<input type="button" value="Add Comment" id="btn_add_comment"/> -->
                        <div id="addcomment" style="visibility:visible;width:100%">
                      <form id="myForm" action="comment.asp" method="post"> 
    <textarea name="txtcomments" id="txtcomments" class="noclass" style="width:100%;" rows="10"></textarea> <br>
	<input type="hidden" name="projectid" value="<%=session("projectid")%>" />
	<input type="hidden" name="userid" value="<%=session("user")%>" />
    <input type="hidden" name="issueid" value="<%=request("issueId")%>" />
    <input type="submit" value="Submit Comment" id="submit_comment" name="submit_comment" /> 
</form></div></td></tr></table>
             <% 
			end if
			objissue.close
			set objissue = nothing	  
			
			objproject.Close 
			set objproject = nothing
    	  %>			

						
                        </td>
                      </tr>
					  </table>

<script>
function editmydiary()
{
	id_edit_my_diary.style.visibility = 'visible';	
	frm_diary.comments.value = 'Add your comments here';
	frm_diary.comments.select();
	
}	
function cancel_diary()
{
	id_edit_my_diary.style.visibility = 'hidden';	
	frm_diary.comments.value = 'Add your comments here';
}
//CR20060706
// The below javascript keep a copy of all the comments in array or single variable for further use.
//var ab= "asdflasjdf lkjas";
//alert(ab.indexOf(' '));
varCmnt = "";
varCmntCopy = "";
mySingleComment = ""
if(typeof document.all("comnt").length == "undefined")	{	
	mySingleComment = document.all("comnt").innerText;
	varCmnt = document.all("comnt").innerText;
	if(varCmnt.length >= 100) {
		if(varCmnt.indexOf(' ') == -1) {
			for(j=0;j<varCmnt.length;j+=100) {
				varCmntCopy += varCmnt.substr(j+1,100) + "-";
			}
			document.all("comnt").innerText = varCmntCopy ;
		}
	}
}
else {
	/*var myArrComments = new Array(document.all("comnt").length);
	for(i=0;i<document.all("comnt").length;i++)	{
		varCmnt = document.all("comnt")(i).innerText;
		myArrComments[i] = varCmnt; 
		varCmntCopy = "";
		if(varCmnt.length >= 100) {
			if(varCmnt.indexOf(' ') == -1) {
				for(j=0;j<varCmnt.length;j+=100) {
					varCmntCopy += varCmnt.substr(j,100) + "-";
				}
				document.all("comnt")(i).innerText = varCmntCopy ;
			}
		}
	}
	*/
}
function AddRemoveTinyMce(editorId) {
			if(tinyMCE.get(editorId)) 
			{
				tinyMCE.EditorManager.execCommand('mceFocus', false, editorId);                    
				tinyMCE.EditorManager.execCommand('mceRemoveEditor', true, editorId);

			} else {
				tinymce.EditorManager.execCommand('mceAddEditor', false, editorId);				
			}
		}
</script>
<script>
// this function is called from issuereportflag.asp
function setflag()
{
	//window.location= 'issuereportflag.asp';
	if(frm.chkFlag.checked == true)
		window.location = 'issuereportflag.asp?Flag=true&issueid=<%=request("issueid")%>';
	else
		window.location = 'issuereportflag.asp?Flag=false&issueid=<%=request("issueid")%>';
}
function checkListValue(obj)
{
	//frm.action = 
	if(obj.name == "selProject_to_copy")
	{
		frm.action = "copyissue.asp";
	}
	else { frm.action = "moveissue.asp"; } 
	//alert(frm.action);
	//return false;
	if(obj.value == 'none')
	{
		alert('Please select a project');
		return false;
	}

	if (confirm('Do u really want to ' + (obj.name == "selProject_to_copy" ? "copy" : "move") + ' this issue.'))
		return true;
	else
		return false;
}
function checkDelete()
{
	if (confirm('Do u really want to delete this comment'))
		return true;
	else
		return false;
}
function checkDeleteissue()
{
	if (confirm('Do u really want to delete this issue.'))
		return true;
	else
		return false;
}




 
</script>
<script src="https://cdn.tiny.cloud/1/jfhvmzju5ge1ljiv30evfvsh676206icpk2z7i430ny5ouoo/tinymce/6/tinymce.min.js" referrerpolicy="origin"></script>
<script>


tinymce.init({
  selector: '#txtcomments',  // change this value according to your HTML
  plugins: 'a_tinymce_plugin autolink image lists table',
  images_upload_url: 'tinymceeditorupload.asp',
  a_plugin_option: true,
  toolbar: 'undo redo | styles | bold italic | alignleft aligncenter alignright alignjustify | ' +
      'bullist numlist outdent indent | link image | table | print preview media fullscreen | fontsize | ' +
      'forecolor backcolor emoticons | help',
  a_configuration_option: 400,
  menubar: '',
});

</script>
