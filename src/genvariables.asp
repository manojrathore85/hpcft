<%
' *******************   SITE SPECIFIC ITEMS ***********************************************************
'	dim UrlToRedirect
	dim varSiteSpecFromAddress
	dim varSiteSpecRemoteHost
	dim varSiteSpecAddCC
	dim varSiteSpecURL
	dim varSiteSpecTitle  
	dim varSiteSpecMailSubjectPrefix
	dim varSiteSpecNUM_USAGE_DAYS
	dim varSiteSpecValidExtensions
	varSiteSpecFromAddress = "info@hpchft.ai"  
	'varSiteSpecRemoteHost = "184.168.224.165" 'Related to gtsims.com
	varSiteSpecRemoteHost = "localhost" 'Related to mail server for gtsims.com
	varSiteSpecAddCC = "" ' mail address to which each and every mail for any updation in IMS will be mailed.
	varSiteSpecURL = "https://ims.hpchft.ai"
	varSiteSpecTitle = "Home"
	varSiteSpecMailSubjectPrefix = "CLIENTIMS"
	varSiteSpecValidExtensions = "docx,DOCX,XLSX,xlsx,jpg,jpeg,bmp,gif,img,pdf,doc,rtf,txt,zip,tar.gz,gz,tgz,xls,xml,rar,JPG,BMP,GIF,IMG,PDF,DOC,RTF,TXT,ZIP,TAR.GZ,GZ,TGZ,XLS,XML,RAR,png,PNG,JPEG"
	varSiteSpecNUM_USAGE_DAYS = 15
	dim users_ims_po_systems 
	users_ims_po_systems = "agrawalatul@gmail.com"
	
	server.ScriptTimeout = 600  ' asp scriptimeout set to 10 minutes
%>	