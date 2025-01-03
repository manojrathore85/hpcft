<!-- #include file="connect.asp" --> <!-- including the connection file -->
<%
if request("mytextarea") <> "" then
	<!-- response.write(request("mytextarea")) -->
	dim objrs
	  set objrs =server.CreateObject("adodb.recordset")

	  
	  objrs.CursorLocation = 3 'adUseClient
	  'con.begintrans
	  objrs.Open "select * from hindi where 1=2",con,3,3
	  objrs.AddNew 
	  objrs("data") = request("mytextarea")
	  'objrs("reportusers") = session("user")
	  
	  'if filename<>"" then
	'	   objrs("AttachedFilePath") =  filename
	'  end if
	  objrs.Update
	  objrs.close
	  
	  'con.committrans
	  set objrs = nothing
	 ' call sendmail
	  con.close
end if
%>
<!DOCTYPE html>

<html lang="en">
  <head>
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1">

    <script src="https://cdn.tiny.cloud/1/jfhvmzju5ge1ljiv30evfvsh676206icpk2z7i430ny5ouoo/tinymce/6/tinymce.min.js" referrerpolicy="origin"></script>
  </head>

  <body>
  <h1>TinyMCE Quick Start Guide</h1>
    <form method="post">
      <textarea name="mytextarea" id="mytextarea"><%= request("mytextarea") %></textarea>
	  <input type="submit" value="Submit">
    </form>
  </body>
  <script>
  /*
  tinymce.init({
    selector: 'textarea',
    plugins: 'anchor autolink charmap codesample emoticons image link lists media searchreplace table visualblocks wordcount checklist mediaembed casechange export formatpainter pageembed linkchecker a11ychecker tinymcespellchecker permanentpen powerpaste advtable advcode editimage advtemplate ai mentions tinycomments tableofcontents footnotes mergetags autocorrect typography inlinecss',
    toolbar: 'undo redo | blocks fontfamily fontsize | bold italic underline strikethrough | link image media table mergetags | addcomment showcomments | spellcheckdialog a11ycheck typography | align lineheight | checklist numlist bullist indent outdent | emoticons charmap | removeformat',
    tinycomments_mode: 'embedded',
    tinycomments_author: 'Author name',
    mergetags_list: [
      { value: 'First.Name', title: 'First Name' },
      { value: 'Email', title: 'Email' },
    ],
    ai_request: (request, respondWith) => respondWith.string(() => Promise.reject("See docs to implement AI Assistant")),
  });
  */
  tinymce.init({
  selector: 'textarea',  // change this value according to your HTML
  plugins: 'a_tinymce_plugin autolink image',
  images_upload_url: 'tinymceeditorupload.asp',
  a_plugin_option: true,
  a_configuration_option: 400
});
</script>
</html>


<!-- Place the following <script> and <textarea> tags your HTML's <body> -->

