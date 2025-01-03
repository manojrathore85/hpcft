<table width="100%" border="0">
<tr>
    <td><b><a href="selectproject.asp">Home</a> | <a href="selectProject.asp">Change 
      Project</a> | <a href="ims_index.asp">Filter</a> | <a href="findissue.asp">Find 
      Issue</a> | <a href="issuenavigator.asp?search=all">All Issues</a> | <a href="createissue.asp">Create 
      New Issue</a> | <a href="viewReminder.asp">Reminders</a> | <a href="viewprojects.asp">Administration</a> | <a href="mydiary.asp">My Diary 
      </a> <!--| <a href="projectworks.html">Project 
      Works</a>--> </b> </td>
	  <div id="msgDiv"></div>
</tr>
</table>
<script type="text/javascript" src="js/jquery-latest.min.js"></script> 
<script type="text/javascript" src="js/jquery.form.min.js"></script> 
<script>
function loadlink(){
    $('#msgDiv').load('checksession.asp',function () {
         //$(this).unwrap();
    });
}

loadlink(); // This will run on page load
setInterval(function(){
    loadlink() // this will run after every 5 seconds
}, 15000);
</script>