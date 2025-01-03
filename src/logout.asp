<%
' HERE CONNECTION IS USED TO MARK THE ENTRY IN IMS_USAGE TABLE
%>
<!-- #include file="connect.asp" -->
<%
   session("user")=""
   session.Abandon()
   con.close
   set con = nothing
   
   response.Redirect("website_index.asp")
%>

