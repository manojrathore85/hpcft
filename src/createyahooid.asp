<!-- #include file="checksession.asp" -->
<!-- #include file="Connect.asp" -->
<%

' ********************* CHECKING PERMISSION ***************************

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
		con.close
		set con = nothing
		response.End()
	end if
else
		server.Execute("permissiondenied.asp")
		rsperm.close
		set rsperm = nothing
		con.close
		set con = nothing
		response.End()
end if

' ********************* CHECKING PERMISSION ENDS***************************
%>

<%
    ' *********** DERIVING NEW UNIQUE YAHOO	MAIL ID FOR CREATION ***********************
	
	Dim sstr, posAt, posDot, newmailid
	sstr = request("id") 
	posAt = InStr(1, sstr, "@")
	posDot = InStr(1, sstr, ".")
	newmailid = Mid(sstr, 1, posAt - 1) & "_"
	newmailid = newmailid & Mid(sstr, posAt + 1, (posDot - posAt) - 1) & "_com"
	
   '****************   COMPLETED *******************************************************
   
   dim strsql 
   dim email
   dim firstname
   dim lastname
   dim password
   strsql = "select email,firstname,lastname,password from " &  varTblNameUserProfile & " where email ='" & request("id") & "'"
   dim rs
   set rs = server.CreateObject("adodb.recordset")
   rs.open strsql,con
   if not rs.eof then
   	email = rs("email")
	firstname = rs("firstname")
	lastname = rs("lastname")
	password = rs("password")
   end if
   rs.close
   set rs = nothing
   con.close
   set con = nothing
%>
<html>
<head>
	<title>Yahoo! Registration</title>
<script language="javascript" src="http://us.i1.yimg.com/us.yimg.com/lib/reg/yih2.js" type="text/javascript"></script>
<script language="javascript" src="http://us.i1.yimg.com/us.yimg.com/lib/common/yg_csstare.js" type="text/javascript"></script>
<script language="javascript" src="http://us.i1.yimg.com/us.yimg.com/lib/g/yg_dom.js" type="text/javascript"></script> 
<script language="javascript" type="text/javascript">
		<!--
		function yreg_pop(s){
		window.open(s,"yec_pop3","width=520,height=500,scrollbars=yes,resizable=yes");
		}
		// -->
</script> 
<script language="JavaScript" type="text/javascript"> 
  <!--
  //ID Helper DHTML module toggle
  function toggleInstructions() {
  	if(oBw.ie5||oBw.ns6){  //Disable function in unsupported browsers
    	Istate = document.getElementById('instructions').style.display;
    		if (Istate=='none') {
      			Istate='block';
    		}
    		else {
      			Istate='none';
    		}
    	document.getElementById('instructions').style.display=Istate;
  		}
  	}
  function toggleInstructionsOn() {
  	if(oBw.ie5||oBw.ns6){  //Disable function in unsupported browsers
    	Istate = document.getElementById('instructions').style.display;
    		if (Istate=='none') {
      			Istate='block';
    		}
    		else {
      			Istate='block';
    		}
    	document.getElementById('instructions').style.display=Istate;
		}
  	}
   //Toggle Alt Email Mandatory notation
   //switch layers for different browsers
	var ie4 = (document.all) ? true : false;
	var ns4 = (document.layers) ? true : false;
	var ns6 = (document.getElementById && !document.all) ? true : false;
	function hidelayer(lay) {
	if (ie4) {document.all[lay].style.visibility = "hidden";}
	if (ns4) {document.layers[lay].visibility = "hide";}
	if (ns6) {document.getElementById([lay]).style.display = "none";}
	}
	function showlayer(lay) {
	if (ie4) {document.all[lay].style.visibility = "visible";}
	if (ns4) {document.layers[lay].visibility = "show";}
	if (ns6) {document.getElementById([lay]).style.display = "block";}
	}
   // -->
</script>
<link rel="stylesheet" href="http://us.i1.yimg.com/us.yimg.com/i/reg/yreg_lite_v5.css" type="text/css">
<style type="text/css"> 
      .yregfloathelp { 
              border: 1px solid #FFC30E; 
			  padding: 5px 5 5 5px; 
              background-color: #FFFBB8; 
              text-align: left; 
              color: #9C7600; 
              width: 18em; 
              font-size: 11px; 
              font-family: arial, sans-serif; 
      } 
      input {font-family:Arial, Helvetica, sans-serif; font-size: 12px;} /*n4 hack -buttons*/
</style> 
<script language="JavaScript" type="text/javascript">
/*
 * A JavaScript implementation of the RSA Data Security, Inc. MD5 Message
 * Digest Algorithm, as defined in RFC 1321.
 * Copyright (C) Paul Johnston 1999 - 2000.
 * Updated by Greg Holt 2000 - 2001.
 * See http://pajhome.org.uk/site/legal.html for details.
 */

/*
 * Convert a 32-bit number to a hex string with ls-byte first
 */
var hex_chr = '0123456789abcdef';
function rhex(num)
{
  str = "";
  for(j = 0; j <= 3; j++)
    str += hex_chr.charAt((num >> (j * 8 + 4)) & 0x0F) +
           hex_chr.charAt((num >> (j * 8)) & 0x0F);
  return str;
}

/*
 * Convert a string to a sequence of 16-word blocks, stored as an array.
 * Append padding bits and the length, as described in the MD5 standard.
 */
function str2blks_MD5(str)
{
  nblk = ((str.length + 8) >> 6) + 1;
  blks = new Array(nblk * 16);
  for(i = 0; i < nblk * 16; i++) blks[i] = 0;
  for(i = 0; i < str.length; i++)
    blks[i >> 2] |= str.charCodeAt(i) << ((i % 4) * 8);
  blks[i >> 2] |= 0x80 << ((i % 4) * 8);
  blks[nblk * 16 - 2] = str.length * 8;
  return blks;
}

/*
 * Add integers, wrapping at 2^32. This uses 16-bit operations internally 
 * to work around bugs in some JS interpreters.
 */
function add(x, y)
{
  var lsw = (x & 0xFFFF) + (y & 0xFFFF);
  var msw = (x >> 16) + (y >> 16) + (lsw >> 16);
  return (msw << 16) | (lsw & 0xFFFF);
}

/*
 * Bitwise rotate a 32-bit number to the left
 */
function rol(num, cnt)
{
  return (num << cnt) | (num >>> (32 - cnt));
}

/*
 * These functions implement the basic operation for each round of the
 * algorithm.
 */
function cmn(q, a, b, x, s, t)
{
  return add(rol(add(add(a, q), add(x, t)), s), b);
}
function ff(a, b, c, d, x, s, t)
{
  return cmn((b & c) | ((~b) & d), a, b, x, s, t);
}
function gg(a, b, c, d, x, s, t)
{
  return cmn((b & d) | (c & (~d)), a, b, x, s, t);
}
function hh(a, b, c, d, x, s, t)
{
  return cmn(b ^ c ^ d, a, b, x, s, t);
}
function ii(a, b, c, d, x, s, t)
{
  return cmn(c ^ (b | (~d)), a, b, x, s, t);
}

/*
 * Take a string and return the hex representation of its MD5.
 */
function MD5(str)
{
  x = str2blks_MD5(str);
  var a =  1732584193;
  var b = -271733879;
  var c = -1732584194;
  var d =  271733878;
 
  for(i = 0; i < x.length; i += 16)
  {
    var olda = a;
    var oldb = b;
    var oldc = c;
    var oldd = d;

    a = ff(a, b, c, d, x[i+ 0], 7 , -680876936);
    d = ff(d, a, b, c, x[i+ 1], 12, -389564586);
    c = ff(c, d, a, b, x[i+ 2], 17,  606105819);
    b = ff(b, c, d, a, x[i+ 3], 22, -1044525330);
    a = ff(a, b, c, d, x[i+ 4], 7 , -176418897);
    d = ff(d, a, b, c, x[i+ 5], 12,  1200080426);
    c = ff(c, d, a, b, x[i+ 6], 17, -1473231341);
    b = ff(b, c, d, a, x[i+ 7], 22, -45705983);
    a = ff(a, b, c, d, x[i+ 8], 7 ,  1770035416);
    d = ff(d, a, b, c, x[i+ 9], 12, -1958414417);
    c = ff(c, d, a, b, x[i+10], 17, -42063);
    b = ff(b, c, d, a, x[i+11], 22, -1990404162);
    a = ff(a, b, c, d, x[i+12], 7 ,  1804603682);
    d = ff(d, a, b, c, x[i+13], 12, -40341101);
    c = ff(c, d, a, b, x[i+14], 17, -1502002290);
    b = ff(b, c, d, a, x[i+15], 22,  1236535329);    

    a = gg(a, b, c, d, x[i+ 1], 5 , -165796510);
    d = gg(d, a, b, c, x[i+ 6], 9 , -1069501632);
    c = gg(c, d, a, b, x[i+11], 14,  643717713);
    b = gg(b, c, d, a, x[i+ 0], 20, -373897302);
    a = gg(a, b, c, d, x[i+ 5], 5 , -701558691);
    d = gg(d, a, b, c, x[i+10], 9 ,  38016083);
    c = gg(c, d, a, b, x[i+15], 14, -660478335);
    b = gg(b, c, d, a, x[i+ 4], 20, -405537848);
    a = gg(a, b, c, d, x[i+ 9], 5 ,  568446438);
    d = gg(d, a, b, c, x[i+14], 9 , -1019803690);
    c = gg(c, d, a, b, x[i+ 3], 14, -187363961);
    b = gg(b, c, d, a, x[i+ 8], 20,  1163531501);
    a = gg(a, b, c, d, x[i+13], 5 , -1444681467);
    d = gg(d, a, b, c, x[i+ 2], 9 , -51403784);
    c = gg(c, d, a, b, x[i+ 7], 14,  1735328473);
    b = gg(b, c, d, a, x[i+12], 20, -1926607734);
    
    a = hh(a, b, c, d, x[i+ 5], 4 , -378558);
    d = hh(d, a, b, c, x[i+ 8], 11, -2022574463);
    c = hh(c, d, a, b, x[i+11], 16,  1839030562);
    b = hh(b, c, d, a, x[i+14], 23, -35309556);
    a = hh(a, b, c, d, x[i+ 1], 4 , -1530992060);
    d = hh(d, a, b, c, x[i+ 4], 11,  1272893353);
    c = hh(c, d, a, b, x[i+ 7], 16, -155497632);
    b = hh(b, c, d, a, x[i+10], 23, -1094730640);
    a = hh(a, b, c, d, x[i+13], 4 ,  681279174);
    d = hh(d, a, b, c, x[i+ 0], 11, -358537222);
    c = hh(c, d, a, b, x[i+ 3], 16, -722521979);
    b = hh(b, c, d, a, x[i+ 6], 23,  76029189);
    a = hh(a, b, c, d, x[i+ 9], 4 , -640364487);
    d = hh(d, a, b, c, x[i+12], 11, -421815835);
    c = hh(c, d, a, b, x[i+15], 16,  530742520);
    b = hh(b, c, d, a, x[i+ 2], 23, -995338651);

    a = ii(a, b, c, d, x[i+ 0], 6 , -198630844);
    d = ii(d, a, b, c, x[i+ 7], 10,  1126891415);
    c = ii(c, d, a, b, x[i+14], 15, -1416354905);
    b = ii(b, c, d, a, x[i+ 5], 21, -57434055);
    a = ii(a, b, c, d, x[i+12], 6 ,  1700485571);
    d = ii(d, a, b, c, x[i+ 3], 10, -1894986606);
    c = ii(c, d, a, b, x[i+10], 15, -1051523);
    b = ii(b, c, d, a, x[i+ 1], 21, -2054922799);
    a = ii(a, b, c, d, x[i+ 8], 6 ,  1873313359);
    d = ii(d, a, b, c, x[i+15], 10, -30611744);
    c = ii(c, d, a, b, x[i+ 6], 15, -1560198380);
    b = ii(b, c, d, a, x[i+13], 21,  1309151649);
    a = ii(a, b, c, d, x[i+ 4], 6 , -145523070);
    d = ii(d, a, b, c, x[i+11], 10, -1120210379);
    c = ii(c, d, a, b, x[i+ 2], 15,  718787259);
    b = ii(b, c, d, a, x[i+ 9], 21, -343485551);

    a = add(a, olda);
    b = add(b, oldb);
    c = add(c, oldc);
    d = add(d, oldd);
  }
  return rhex(a) + rhex(b) + rhex(c) + rhex(d);
}

function isValidPasswd(pw, login, fn, ln, pwqa_a){
  if(pw.length == 0){
    alert("Please choose a password.");
    return -1;
  } else if (pw.length > 32) {
    alert("Your new password must be no greater than 32 characters.");
    return -1;
  }else if (pw.length < 6) {
    alert("Your new password must be at least 6 characters.");
    return  -1;
  } else if (login.length != 0 && (login.indexOf(pw) >= 0 || pw.indexOf(login) >= 0)){
    alert("Your new password is too similar to your Yahoo! ID.");
    return -1;
  } else if ((fn.length > 2) && (pw.indexOf(fn) >= 0) ){
    alert("Your new password is too similar to your first or last name.");
    return -1;
  } else if ( (ln.length > 2) && (pw.indexOf(ln) >=0) ){
    alert("Your new password is too similar to your first or last name.");
    return -1;
  } else if ( pwqa_a == pw ){
    alert("Your Security Answer is too similar to your password. Please choose a different answer for your Security Question.");
    return -1;
  } else {
    return 0;
  }
}
function GetSelectedValue(box){
  var boxText = "";
  for(var i=0; i<box.options.length; i++){
    if(box.options[i].selected && box.options[i].value != ""){
      boxText = box.options[i].value;
    }
  }
  return boxText;
}

function clearMD5Flag(form) {
  form['.md5'].value = "" ;
  return true ;
}

function hash(form) {
    // rudimentary check for a 4.x brower. should catch IE4+ and NS4.*
    // netscape BSD does not allow JS to change pw field.
    if (navigator.userAgent.indexOf('Mozilla/4')==0 && (navigator.userAgent.indexOf('X11') < 0)) {
      var pw = form['.pw'].value;
      var pw2 = form['.pw2'].value;
      var fn = "";
      var ln = "";
      var login = "";
      var pwqa_a = "";
      if(form.login){
        login = form.login.value;
      } else if (form['.login_other']){
        login = form['.login_other'].value;
      }
      // is pw already md5?
      var md5 = form['.md5'].value;
      if(form['.fn']){
        fn = form['.fn'].value;
      }
      if(form['.ln']){
        ln = form['.ln'].value;
      }
      if(form['.pw_a']){
        pwqa_a = form['.pw_a'].value;
      }
      if(pw != pw2){
        alert("Your password entries did not match.");
        return false;
      }
      if(isValidPasswd(pw, login, fn, ln, pwqa_a) == 0){
        var pwHash = pw;
        var prntHash = ""; 
        var prntPW = ""; 
        if(form['.prntPW']){
          prntPW = form['.prntPW'].value;
          prntHash = prntPW;
        }
        md5 = parseInt(md5);
        if (!md5) md5 = 0;
        if(md5 == 0){ 
          pwHash = MD5(pw);
          // MD5 parent pw exist, it's a family account
          if(prntHash.length > 0){
            var challenge = "";
            if(form['.challenge']){
              challenge = form['.challenge'].value;
            }
            prntHash = MD5(MD5(prntPW)+challenge);
          }
        }
        form.action = 'http://edit.yahoo.com/config/register?';
        if(form['.pw'] && form['.pw'].type != 'hidden'){
          form['.pw'].value = pwHash;
        } 
        if(form['.pw2'] && form['.pw2'].type != 'hidden'){
          form['.pw2'].value = pwHash;
        } 
        if(form['.prntPW'] && form['.prntPW'].type != 'hidden'){
          form['.prntPW'].value = prntHash;
        } 
        md5 ++;
        form['.md5'].value = md5;
        // prevent from running this again. Allow the server response to submit the form directly
        form.onsubmit=null;

        return true;
      }
      return false;
    } 
    // allow normal form submission
    return true;
}

</script>

</head>
<body>
<div class="masthead" align="center">
<table cellspacing="0" cellpadding="0" width="720" border="0" summary="Header">
 	<tr>
		<td width=1%>
		</td>
		<td width=1% align=left>
			<table width=1% cellspacing=0 cellpadding=0 border=0 summary="null">
				<tr>
 					<td width=1% align=left nowrap>
<img src="http://us.i1.yimg.com/us.yimg.com/i/us/nt/ma/ma_mail_1.gif" width="196" height="33" border="0" alt="Yahoo! Mail">
					</td>
				</tr>
			</table>
		</td>
		<td width=100% align=right nowrap valign=top>
			<table width=100% cellspacing=0 cellpadding=0 border=0 summary="null">
				<tr>		
					<td width=99% align=right nowrap valign=top><font class="yregtopmodulefont"><a href="http://www.yahoo.com " target="www" title="Click here to go to Yahoo! homepage">Yahoo!</a> - <a href="http://help.yahoo.com/help/us/edit/" target="yhelp" title="Click here for help">Help</a></font></td>
				</tr>
			</table>
		</td>
	</tr>
</table></div><div class="topmodule" align="center">
<table cellpadding="0" cellspacing="0" border="0" width="720" summary="null">
	<tr>
		<td width="720" height="5" nowrap colspan="2"><img src="http://us.i1.yimg.com/us.yimg.com/i/reg/yreg_rounded_top.gif" alt="null" width="720" height="5" vspace="0" hspace="0"></td>
	</tr>
	<tr>
		<td bgcolor="#A9A9A9" width="1" nowrap><spacer type="block" width="1"></td>
		<td width="718" nowrap valign="top">
<table cellpadding="5" cellspacing="0" border="0" width="718" summary="Click on Sign-In link if you already have an ID or a Yahoo! Mail address.">
	<tr>
		<td valign="top" class="yregtopmodulefont"><img src="http://us.i1.yimg.com/us.yimg.com/i/reg/orange_arrow.gif" border="0" width="14" height="14" align="absmiddle" alt="null">&nbsp;<strong>
		Already have an ID or a Yahoo! Mail address?
		<a href="http://us.rd.yahoo.com/reg/topbar_login/v2/us/*http://login.yahoo.com/config/login?.src=ym&partner=&.v=&.u=6gbu48d17hdkc&.intl=us&.done=http%3a//mail.yahoo.com" title="Click here to sign-in if you already have an ID">Sign In</a>.
		</strong></td>
	</tr>
</table>
		</td> 
		<td bgcolor="#A9A9A9" width="1" nowrap><spacer type="block" width="1"></td>
	</tr>
	<tr>
		<td width="720" height="6" nowrap colspan="2"><img src="http://us.i1.yimg.com/us.yimg.com/i/reg/yreg_rounded_top_reverse.gif" alt="null" width="720" height="5" vspace="0" hspace="0"></td> 
	</tr>
</table>
</div>
<div align="center">


<table cellpadding="5" cellspacing="0" border="0" width="720" summary="null">
	<tr>
		<td nowrap class="yregmicrofont">Fields marked with an asterisk <span class="yregasterisk">*</span> are required.<spacer type="block" width="1" height="10"><a href="http://add.yahoo.com/fast/help/us/edit/cgi_access"><img src="http://us.i1.yimg.com/us.yimg.com/i/space.gif" border="0" width="1" height="1" alt="Attention Blind or Visually Impaired Users. To complete this form you must enter a word that is part of an image. If you can't read the image, Yahoo is happy to help you create your account. A representative from customer care will need to contact you. To request assistance with registration, please read the Yahoo! Terms of Service located at http://docs.yahoo.com/info/terms. Once you have reviewed our policies, please provide your phone number and email address and send your request by visiting this URL - http://add.yahoo.com/fast/help/us/edit/cgi_access"></a></td>
	</tr>
</table>

<table cellpadding="0" cellspacing="0" border="0" width="720" summary="null">
	<tr>
	<!-- Left Col -->
		<td width="720" nowrap align="left" valign="top">
		<!-- Left Col Content -->
<table cellpadding="0" cellspacing="0" border="0" width="720" summary="null">
	<tr>
		<td width="720" height="5" nowrap colspan="2"><img src="http://us.i1.yimg.com/us.yimg.com/i/reg/yreg_rounded_top.gif" alt="null" width="720" height="5" vspace="0" hspace="0"></td>
	</tr>   
<form method="post" action="https://edit.yahoo.com/config/register" name="IOS" onsubmit="return hash(this)">
<input type="hidden" name=".save" value="1">
<input type=hidden name=".accept" value="0" >
<input type=hidden name=".demog" value="" >
<input type=hidden name=".done" value="http://mail.yahoo.com" >
<input type=hidden name=".fam" value="" >
<input type=hidden name=".i" value="" >
<input type=hidden name=".last" value="" >
<input type=hidden name=".src" value="ym" >
<input type=hidden name=".regattempts" value="0" >
<input type=hidden name=".partner" value="" >
<input type=hidden name="promo" value="" >
<input type=hidden name=".ignore" value="ind,job,spe" >
<input type=hidden name=".pwtoken" value="" >
<input type=hidden name=".u" value="6gbu48d17hdkc" >
<input type=hidden name=".v" value="" >
<input type=hidden name=".md5" value="" >
<input type=hidden name=".testid" value="none" >
<input type=hidden name=".branch" value="" >
<input type=hidden name=".t" value="uyjhR75ugVuyBi4xElXeIMeba3v2gxHnsn.W87.iz3AJWo6nrBb4mkFl52XikK6edL2Po7ryGfSp6yIgEc29eh0ZvS3h4d_cqA_VhHRQboHD.CRglbkwhagh79dvApswT2.RMWGvLKnB_A3fDDvylEJdqdjk1Zxw0s_JNzKW_F5upOtqc_XX.DG66z22h.hniuGEBcmsmuqeGwHmsa_5PvREuYqGRNtBcd9IymRasYv.UdO0a3aljy5oSmAwR3Y7ytR0fhzXkQuiH0bSoDX.2Brn0y03wR6M3RoEuln90UyST4eV_WDMGuZwD5W4D80jhvVGtOqMMnREf.ayaFXai3knKfvvYQkpN_ulpVJP.MVseg--" >
<tr>
	<td bgcolor="#A9A9A9" width="1" nowrap><spacer type="block" width="1"></td>
	<td width="718" nowrap>
<table cellpadding="0" cellspacing="0" border="0" width="718" summary="Create Your Yahoo! ID">
	<tr>
		<td colspan="3"><h1>Create Your Yahoo! ID</h1></td>
	</tr>
	<tr>
		<td colspan="3" nowrap height="10"><spacer type="vertical" height="10"></td>
	</tr>
	<tr  >
		<td align="right" width="150" class="yregfieldtitle"><span class="yregasterisk">*</span>&nbsp;<LABEL for="fname" accesskey="F">First name:</LABEL></td>
		<td nowrap width="10"><spacer type="horizontal" width="10"></td>
		<td align="left" width="560"><input type=text name=".fn" value="<%=firstname%>" size="30" maxlength="50" id="fname"></td>
	</tr>
	<tr>
		<td colspan="3" nowrap height="5"><spacer type="vertical" height="5"></td>
	</tr>
	<tr >
		<td align="right" width="150" class="yregfieldtitle"><span class="yregasterisk">*</span>&nbsp;<LABEL for="lname" accesskey="L">Last name:</LABEL></td>
		<td nowrap width="10"><spacer type="horizontal" width="10"></td>
		<td align="left" width="560"><input type=text name=".ln" value="<%=lastname%>" size="30" maxlength="50" id="lname"></td>
	</tr>
	<tr>
		<td colspan="3" nowrap height="5"><spacer type="vertical" height="5"></td>
	</tr>
	<tr >
		<td align="right" width="150" class="yregfieldtitle"><span class="yregasterisk">*</span>&nbsp;<LABEL for="intl" accesskey="I">Preferred content:</LABEL></td>
		<td nowrap width="10"><spacer type="horizontal" width="10"></td>
		<td align="left" width="560" class="yregbelowfield">
<script language="JavaScript" type="text/javascript">
<!--
function yreg_changeIntl(intl){
	var fn		= document.IOS.elements['.fn'].value;
	var ln 		= document.IOS.elements['.ln'].value;
	var sx 		= document.IOS.elements['.sx'].value;
	var login 	= document.IOS.elements['login'].value;
	self.location.href("http://us.rd.yahoo.com/reg/v2/swintl/*http://edit.yahoo.com/config/eval_register?.intl=" + intl + "&.fn=" + fn + "&.ln=" + ln + "&.sx=" + sx + "&login=" + login + "&.src=ym&.u=6gbu48d17hdkc&.v=&.partner=&.p=&promo=&new=1&.done=http%3a//mail.yahoo.com&.last="); 
}
//-->
</script>
<select name=".intl"onchange="yreg_changeIntl(this.options[this.selectedIndex].value)" id="intl">
<option selected value="us">Yahoo! U.S.
<option value="e1">Yahoo! U.S. in Spanish
<option value="b5">Yahoo! U.S. in Chinese
<option value="ru">Yahoo! U.S. in Russian
<option value="cn">Yahoo! China
<option value="uk">Yahoo! United Kingdom
<option value="ar">Yahoo! Argentina
<option value="aa">Yahoo! Asia
<option value="au">Yahoo! Australia 
<option value="br">Yahoo! Brazil
<option value="ca">Yahoo! Canada in English
<option value="cf">Yahoo! Canada in French
<option value="dk">Yahoo! Denmark
<option value="fr">Yahoo! France 
<option value="de">Yahoo! Germany
<option value="gr">Yahoo! Greece
<option value="hk">Yahoo! Hong Kong
<option value="in">Yahoo! India
<option value="ie">Yahoo! Ireland
<option value="it">Yahoo! Italy
<option value="kr">Yahoo! Korea 
<option value="mx">Yahoo! Mexico
<option value="nz">Yahoo! New Zealand
<option value="no">Yahoo! Norway
<option value="sg">Yahoo! Singapore
<option value="es">Yahoo! Spain
<option value="se">Yahoo! Sweden
<option value="tw">Yahoo! Taiwan

</select>&nbsp;
<script language="JavaScript" type="text/javascript">
	<!--
	image3 = new Image();
	image3.src = "http://us.i1.yimg.com/us.yimg.com/i/reg/ctx_question_active.gif";
	//-->
</script>
<script language="JavaScript" type="text/javascript">
if(oBw.ie5||oBw.ns6){
document.write("<span onmouseover=\"showlayer('intl_help');image3.src='http://us.i1.yimg.com/us.yimg.com/i/reg/ctx_question_active.gif';\" onmouseout=\"hidelayer('intl_help');image3.src='http://us.i1.yimg.com/us.yimg.com/i/reg/ctx_question_inactive.gif';\"><img src=\"http://us.i1.yimg.com/us.yimg.com/i/reg/ctx_question_inactive.gif\" name=\"image3\" width=\"17\" height=\"17\" border=\"0\"></span>"); 
document.write("&nbsp;&nbsp;<span class=yregfloathelp style='position: absolute;' id=\"intl_help\">Please select the local Yahoo! site from this list that best meets your content needs. Note: If you change this setting, you may receive a country-specific Yahoo! email address instead of one that ends in @yahoo.com.</span>");
} else {
document.write("<a href=\"http://us.rd.yahoo.com/reg/exp_bday/us/*http://edit.yahoo.com/config/form?.branch=&.form=bday_explain&.intl=us\" onclick=\"yreg_pop('http://us.rd.yahoo.com/reg/exp_bday/us/*http://edit.yahoo.com/config/form?.branch=&.form=bday_explain&.intl=us');return false;\" target=\"_blank\">More info</a>");
}
</script>
<noscript><a href="http://us.rd.yahoo.com/reg/exp_intl/us/*http://edit.yahoo.com/config/form?.branch=&.form=explain_intl&.intl=us" target="info" title="More information about Preferred Content">More info</a></noscript>
		</td>
	</tr>
<script language="JavaScript" type="text/javascript"> 
  <!-- 
  hidelayer('intl_help') 
  //--> 
</script> 
	<tr>
		<td colspan="3" nowrap height="5"><spacer type="vertical" height="5"></td>
	</tr>

	<tr  >
		<td align="right" width="150" class="yregfieldtitle"><span class="yregasterisk">*</span>&nbsp;<LABEL for="sex" accesskey="G">Gender:</LABEL></td>
		<td nowrap width="10"><spacer type="horizontal" width="10"></td>
		<td align="left" width="560">
<select name=".sx"id="sex"><option value="">[Select] 
<option value="m" selected>
			Male
<option value="f">
			Female</select>
		</td>
	</tr>
	<tr>
		<td colspan="3" nowrap height="5"><spacer type="vertical" height="5"></td>
	</tr>
	<tr  >
		<td align="right" width="150" class="yregfieldtitle"><span class="yregasterisk">*</span>&nbsp;<LABEL for="yid" accesskey="Y">Yahoo! ID:</LABEL></td>
		<td nowrap width="10"><spacer type="horizontal" width="10"></td>
		<td align="left" width="560">
<input type=text name="login" value="<%=newmailid%>" autocomplete="off" id="yid" size="30" maxlength="32" onfocus="toggleInstructionsOn()">
<span class="yregmaildomain">@yahoo.com</span>
		</td>
	</tr>
	<tr>
		<td align="left" width="150"></td>
		<td nowrap width="10"><spacer type="horizontal" width="10"></td>
		<td align="left" width="560" class="yregbelowfield">ID may consist of a-z, 0-9 and underscores.</td>
	</tr>
	<tr>
		<td align="right" width="150"></td>
		<td nowrap width="10"><spacer type="horizontal" width="10"></td>
		<td align="left" width="560">
		<script language="javascript" type="text/javascript">if(oBw.ie5||oBw.ns6){document.write("<div id=\"instructions\" style=\"border: 1px solid #708090; background: #fff8dc; padding:4px; width: 200px;margin:0 0 0px 0;display:block\" align=\"center\"><input class=\"yihbtn\" type=\"button\" name=\"checkAvail\" value=\"Check Availability of This ID\" style=\"width:15em\" onclick=\"yih_idpop('/config/id_check?.intl=us&.src=ym');return false;\" target=\"_blank\"></div>");}</script>
		</td>
	</tr>
<script language="JavaScript" type="text/javascript">
	<!--
	if(oBw.ie5||oBw.ns6){
		toggleInstructions()   
		}
	//-->
</script>
	<tr>
		<td colspan="3" nowrap height="5"><spacer type="vertical" height="5"></td>
	</tr>
	<tr  >
		<td align="right" width="150" class="yregfieldtitle"><span class="yregasterisk">*</span>&nbsp;<LABEL for="pw" accesskey="P">Password:</LABEL></td>
		<td nowrap width="10"><spacer type="horizontal" width="10"></td>
		<td align="left" width="560"><input type=password name=".pw" value="<%=password%>" onChange="return clearMD5Flag(this.form)" autocomplete="off" id="pw" size="30" maxlength="32"></td>
	</tr>
	<tr>
		<td align="right" width="150"></td>
		<td nowrap width="10"><spacer type="horizontal" width="10"></td>
		<td align="left" width="560" class="yregbelowfield">Six characters or more; capitalization matters!</td>
	</tr>
	<tr>
		<td colspan="3" nowrap height="5"><spacer type="vertical" height="5"></td>
	</tr>
	<tr  >
		<td align="right" width="150" class="yregfieldtitle"><span class="yregasterisk">*</span>&nbsp;<LABEL for="pw2" accesskey="R">Re-type password:</LABEL></td>
		<td nowrap width="10"><spacer type="horizontal" width="10"></td>
		<td align="left" width="560"><input type=password name=".pw2" value="<%=password%>" onChange="return clearMD5Flag(this.form)" autocomplete="off" id="pw2" size="30" maxlength="32"></td>
	</tr>
	<tr>
		<td colspan="3" nowrap height="5"><spacer type="vertical" height="5"></td>
	</tr>
<input type=hidden name=".ym_signup" value="1">
	<tr>
		<td colspan="3" nowrap height="10"><spacer type="vertical" height="10"></td>
	</tr>
</table>
	</td>
	<td bgcolor="#A9A9A9" width="1" nowrap><spacer type="block" width="1"></td>
</tr>
	<tr>
		<td bgcolor="#A9A9A9" width="1" nowrap><spacer type="block" width="1"></td>
		<td width="718" height="1" nowrap bgcolor="#A9A9A9"><spacer type="block" width="718" height="1"></td> 
		<td bgcolor="#A9A9A9" width="1" nowrap><spacer type="block" width="1"></td>
	</tr>

 <tr>
	<td bgcolor="#A9A9A9" width="1" nowrap><spacer type="block" width="1"></td>
	<td width="718" nowrap>
<table cellpadding="0" cellspacing="0" border="0" width="718" summary="Account Recovery Information">
	<tr>
		<td colspan="3"><h1>If You Forget Your Password...</h1></td>
	</tr>
	<tr>
		<td colspan="3" nowrap height="10"><spacer type="vertical" height="10"></td>
	</tr>
	<tr  >
		<td align="right" width="150" class="yregfieldtitle"><span class="yregasterisk">*</span>&nbsp;<LABEL for="pwq" accesskey="Q">Security question:</LABEL></td>
		<td nowrap width="10"><spacer type="horizontal" width="10"></td>
		<td align="left" width="560"><select name=".pw_q"id="pwq"><option value="">[Select a Question]
											<option value="What is your pets name?" >What is your pet's name?
											<option value="What was the name of your first school?" >What was the name of your first school?
											<option value="Who was your childhood hero?" selected>Who was your childhood hero?
											<option value="What is your favorite pastime?" >What is your favorite pastime?
											<option value="What is your all-time favorite sports team?" >What is your all-time favorite sports team?
											<option value="What is your fathers middle name?" >What is your father's middle name?
											<option value="What was your high school mascot?" >What was your high school mascot?
											<option value="What make was your first car or bike?" >What make was your first car or bike?
											<option value="Where did you first meet your spouse?" >Where did you first meet your spouse?</select>
		</td>
	</tr>
	<tr>
		<td colspan="3" nowrap height="5"><spacer type="vertical" height="5"></td>
	</tr>
	<tr  >
		<td align="right" width="150" class="yregfieldtitle"><span class="yregasterisk">*</span>&nbsp;<LABEL for="pwa" accesskey="A">Your answer:</LABEL></td>
		<td nowrap width="10"><spacer type="horizontal" width="10"></td>
		<td align="left" width="560"><input type=text name=".pw_a" value="rambo" size="30" maxlength="30" id="pwa" autocomplete=off ></td>
	</tr>
	<tr>
		<td align="right" width="150"></td>
		<td nowrap width="10"><spacer type="horizontal" width="10"></td>
		<td align="left" width="560" class="yregbelowfield">Four characters or more. Make sure your answer is memorable for you
		but hard for others to guess!</td>
	</tr>
	<tr>
		<td colspan="3" nowrap height="5"><spacer type="vertical" height="5"></td>
	</tr>
	<tr  >
		<td align="right" width="150" class="yregfieldtitle"><span class="yregasterisk">*</span>&nbsp;<LABEL for="bday" accesskey="B">Birthday:</LABEL></td>
		<td nowrap width="10"><spacer type="horizontal" width="10"></td>
		<td align="left" width="560" class="yregbelowfield">
<select name=".bmon"><option value="">[Select a Month]<option value="0" selected>January<option value="1">February<option value="2">March<option value="3">April<option value="4">May<option value="5">June<option value="6">July<option value="7">August<option value="8">September<option value="9">October<option value="10">November<option value="11">December</select>
&nbsp;<input type="Text" maxlength="2" name=".bday" size="2" value="05" id="bday" autocomplete=off >&nbsp;,&nbsp;<input type="Text" onfocus="this.value=''" maxlength="4" name=".byear" size="4" value="1978" autocomplete=off >
&nbsp;
<script language="JavaScript" type="text/javascript">
	<!--
	if(oBw.ie5||oBw.ns6){
		image1 = new Image();
		image1.src = "http://us.i1.yimg.com/us.yimg.com/i/reg/ctx_question_active.gif";
	}
	//-->
</script>
<script language="JavaScript" type="text/javascript">
if(oBw.ie5||oBw.ns6){
document.write("<span onmouseover=\"showlayer('bday_help');image1.src='http://us.i1.yimg.com/us.yimg.com/i/reg/ctx_question_active.gif';\" onmouseout=\"hidelayer('bday_help');image1.src='http://us.i1.yimg.com/us.yimg.com/i/reg/ctx_question_inactive.gif';\"><img src=\"http://us.i1.yimg.com/us.yimg.com/i/reg/ctx_question_inactive.gif\" name=\"image1\" width=\"17\" height=\"17\" border=\"0\"></span>"); 
document.write("&nbsp;&nbsp;<span class=yregfloathelp style='position: absolute;' id=\"bday_help\">Please provide an accurate birthdate for your own protection. We ask your birthdate to verify your account if you ever forget your Yahoo! ID or password. (Yahoo! will never request your password or ID in an unsolicited email or phone call.)</span>");
} else {
document.write("<a href=\"http://us.rd.yahoo.com/reg/exp_bday/us/*http://edit.yahoo.com/config/form?.branch=&.form=bday_explain&.intl=us\" onclick=\"yreg_pop('http://us.rd.yahoo.com/reg/exp_bday/us/*http://edit.yahoo.com/config/form?.branch=&.form=bday_explain&.intl=us');return false;\" target=\"_blank\">Why we ask for this?</a>");
}
</script>
<noscript><a href="http://us.rd.yahoo.com/reg/exp_bday/us/*http://edit.yahoo.com/config/form?.branch=&.form=bday_explain&.intl=us" target="info" title="Why we ask for your birthday">Why we ask for this?</a></noscript>
			</td>
	</tr>
<script language="JavaScript" type="text/javascript"> 
 <!-- 
 if(oBw.ie5||oBw.ns6){
 	hidelayer('bday_help') 
 }
 //--> 
</script> 
	<tr>
		<td colspan="3" nowrap height="5"><spacer type="vertical" height="5"></td>
	</tr>
	<tr  >
		<td align="right" width="150" class="yregfieldtitle">
		<span class="yregasterisk">*</span>
		&nbsp;<LABEL for="zip" accesskey="Z">ZIP/Postal code:</LABEL></td>
		<td nowrap width="10"><spacer type="horizontal" width="10"></td>
		<td align="left" width="560"><input type=text name=".pc" value="10001" size="15" maxlength="15" id="zip" autocomplete=off ></td>
	</tr>
	<tr>
		<td colspan="3" nowrap height="5"><spacer type="vertical" height="5"></td>
	</tr>
<input type=hidden name=".co" value="" >
	<tr>
		<td colspan="3" nowrap height="5"><spacer type="vertical" height="5"></td>
	</tr>
	<tr  >
		<td align="right" width="150" class="yregfieldtitle">
		&nbsp;&nbsp;&nbsp;&nbsp;Alternate Email:
		</td>
		<td nowrap width="10"><spacer type="horizontal" width="10"></td>
		<td align="left" width="560" class="yregbelowfield"><input type="text" name=".em" value="<%=email%>" maxlength="120" size="40" id="email" autocomplete=off >&nbsp;
<script language="JavaScript" type="text/javascript">
	<!--
	if(oBw.ie5||oBw.ns6){
		image2 = new Image();
		image2.src = "http://us.i1.yimg.com/us.yimg.com/i/reg/ctx_question_active.gif";
	}
	//-->
</script>
<script language="JavaScript" type="text/javascript">
if(oBw.ie5||oBw.ns6){
document.write("<span onmouseover=\"showlayer('email_help');image2.src='http://us.i1.yimg.com/us.yimg.com/i/reg/ctx_question_active.gif';\" onmouseout=\"hidelayer('email_help');image2.src='http://us.i1.yimg.com/us.yimg.com/i/reg/ctx_question_inactive.gif';\"><img src=\"http://us.i1.yimg.com/us.yimg.com/i/reg/ctx_question_inactive.gif\" name=\"image2\" width=\"17\" height=\"17\" border=\"0\"></span>"); 
document.write("&nbsp;&nbsp;<span class=yregfloathelp style='position: absolute;' id=\"email_help\">We use your alternate email address to send information about your account (including new password requests) and important news about your Yahoo! services.</span>");
} else {
document.write("<a href=\"http://us.rd.yahoo.com/reg/exp_av/us/*http://edit.yahoo.com/config/form?.branch=&.form=explain_avinfo&.intl=us\" onclick=\"yreg_pop('http://us.rd.yahoo.com/reg/exp_av/us/*http://edit.yahoo.com/config/form?.branch=&.form=explain_avinfo&.intl=us');return false;\" target=\"_blank\">More info</a>");
}
</script>
<noscript><a href="http://us.rd.yahoo.com/reg/exp_av/us/*http://edit.yahoo.com/config/form?.branch=&.form=explain_avinfo&.intl=us" target="info" title="More information about Alternate Email">More info</a></noscript>
		</td>
	</tr>
<script language="JavaScript" type="text/javascript"> 
  <!-- 
  if(oBw.ie5||oBw.ns6){
  	hidelayer('email_help') 
  }
  //--> 
</script> 
	<tr>
		<td colspan="3" nowrap height="10"><spacer type="vertical" height="10"></td>
	</tr>
</table>
	</td>
	<td bgcolor="#A9A9A9" width="1" nowrap><spacer type="block" width="1"></td>
</tr>

     <tr>   
         <td bgcolor="#A9A9A9" width="1" nowrap><spacer type="block" width="1"></td>   
         <td width="718" height="1" nowrap bgcolor="#A9A9A9"><spacer type="block" width="718" height="1"></td>   
         <td bgcolor="#A9A9A9" width="1" nowrap><spacer type="block" width="1"></td>   
     </tr>   
<script language="javascript" type="text/javascript">
NS4 = (document.layers) ? true : false;
delta = 1;
var jobArray = new Array();
var titleArray = new Array();
var job;
var ind;

coreTitles = new Array(
"Account Manager/Sales Rep;21",
"Analyst;22",
"Associate/Individual Contributor/Specialist;23",
"Assistant/Support Staff/CSR;24",
"Buyer;25",
"Chairman/CEO/COO;26",
"CFO;27",
"CIO/CTO;28",
"Counsel/Legal;29",
"Database Administrator;30",
"Systems Administrator;31",
"Network Administrator;32",
"Systems Architect;33",
"Programmer/Developer;34",
"Designer;35",
"Director/Sr Mgr;36",
"Engineer (non-IT);37",
"HR;38",
"Inspector;39",
"Producer;40",
"Project Mgr;41",
"Recruiter;42",
"Regional/General Mgr;43",
"Researcher/Scientist;44",
"Mgr/Supervisor;45",
"VP/Sr VP/Exec VP;46",
"Other;47"
);

function loadDefaults(i, j, s)
{
  if (i!='') {
      document.IOS['.ind'].options[getValueIndex(document.forms['IOS']['.ind'], i)].selected=true;
      setInd(document.forms['IOS']['.ind']);
      document.forms['IOS']['.job'].options[getValueIndex(document.forms['IOS']['.job'], j)].selected=true;
      setJob(document.forms['IOS']['.job']);
      document.forms['IOS']['.spe'].options[getValueIndex(document.forms['IOS']['.spe'], s)].selected=true;
  }
}
function loadDefaultJobs(i,j,s){
   if( i != '' ){
      document.forms['IOS']['.job'].options[getValueIndex(document.forms['IOS']['.job'], j)].selected=true;
      fillTitle(i);
      setJob(document.forms['IOS']['.job']);
      document.forms['IOS']['.spe'].options[getValueIndex(document.forms['IOS']['.spe'], s)].selected=true;
   }
}
function loadDefaultSpecs(i,j,s){
   if( i != '' && j != '' ){
      document.forms['IOS']['.spe'].options[getValueIndex(document.forms['IOS']['.spe'], s)].selected=true;
      fillJob(j);
   }
}
function removeElement(arr, element)
{
	for(i=0; i<arr.length; i++){
		if(arr[i].search(';'+element)!=-1){
			arr1 = arr.slice(0, i);
			arr2 = arr.slice(i+1);
			arr = arr1.concat(arr2);
			return arr;
		}
	}
}
function getTitle(industry){
	titleArr = new Array(); 	
	titleArr=titleArr.concat(coreTitles);
	if(industry =='36'){
		titleArr=removeElement(titleArr, '25'); 
		titleArr=removeElement(titleArr, '39'); 
		titleArr=removeElement(titleArr, '44'); 
		var i0Arr = new Array(
"Art Director;48",
"Artist;49",
"Copywriter;50",
"Media Planner;51",
"Photographer;52",
"Creative Director;53");
		titleArr=i0Arr.concat(titleArr);
		return titleArr;
	}
	if(industry == '18'){
		titleArr=removeElement(titleArr, '40');	
		return titleArr;
	}
	if(industry == '19'){
		titleArr=removeElement(titleArr, '40');	
		titleArr=removeElement(titleArr, '35');	
		titleArr=removeElement(titleArr, '45');	
		var i2Arr = new Array(
"Laborer;54",
"Farmer;55",
"Rancher;56",
"Vintner;57");
		titleArr=i2Arr.concat(titleArr);
		return titleArr;
	
	}
	if(industry == '20') {
		titleArr=removeElement(titleArr, '40');	
		return titleArr;
	}
	if(industry == '21'){
		var i4Arr = new Array(
"Technical Support;58",
"Technical Writer;59");
		titleArr=i4Arr.concat(titleArr);
		return titleArr;
	}
	if(industry == '22'){
		titleArr=removeElement(titleArr, '45');	
		var i5Arr = new Array(
"Carpenter;60",
"Developer;61",
"Electrician;62",
"General Contractor;63",
"Laborer;54",
"Plumber;64");
		titleArr=i5Arr.concat(titleArr);
		return titleArr;
	}
	if(industry == '23'){
		titleArr=removeElement(titleArr, '44'); 
		var i6Arr = new Array(

"Artist;49",
"Copywriter;50",
"Media Planner;51",
"Photographer;52");
		titleArr=i6Arr.concat(titleArr);
		return titleArr;
	}
	if(industry == '24'){
		titleArr=removeElement(titleArr, '21');
                titleArr=removeElement(titleArr, '22'); 
		var i7Arr = new Array(
"Administration;65",
"Chancellor/Provost/Dean;66",
"Librarian;67",
"Professor;68",
"Student - College;69",
"Student - Graduate;70",
"Student - K-12;71",
"Teacher - High School;72",
"Teacher - K-8;73");
		titleArr=i7Arr.concat(titleArr);
		return titleArr;
	}
	if(industry == '25'){
		titleArr=removeElement(titleArr, '35');	
		titleArr=removeElement(titleArr, '40');	
		var i8Arr = new Array(
"Maintenance Worker;74",
"Miner;75",
"Driller;76");

		titleArr=i8Arr.concat(titleArr);
		return titleArr;
	}
	if(industry == '26'){
		titleArr=removeElement(titleArr, '39'); 
		titleArr=removeElement(titleArr, '40');	
		titleArr=removeElement(titleArr, '41');	
		titleArr=removeElement(titleArr, '44'); 
		var i9Arr = new Array(
"Accountant;77",
"Actuary;78",
"Banker;79",
"Broker/Dealer;80",
"Financial Planner;81",
"Insurance Agent - Commercial;82",
"Insurance Agent - Personal;83",
"Real Estate Developer;84",
"Realtor - Commercial;85",
"Realtor - Residential;86",
"Trader;87",
"Underwriter;88",
"Venture Capitalist;89");
		titleArr=i9Arr.concat(titleArr);
		return titleArr;
	}
	if(industry == '27'){
		titleArr=removeElement(titleArr, '21');
		titleArr=removeElement(titleArr, '27');
		titleArr=removeElement(titleArr, '28');
		titleArr=removeElement(titleArr, '35');	
		titleArr=removeElement(titleArr, '39'); 
		titleArr=removeElement(titleArr, '40');	
		titleArr=removeElement(titleArr, '43');	
		titleArr=removeElement(titleArr, '46');	
		var i10Arr = new Array(
"Assessor/Recorder;90",
"Enlisted Person;91",
"Firefighter;92",
"Judge;93",
"Legislator;94",
"Mayor;95",
"Officer;96",
"Police Officer/Sheriff;97");
		titleArr=i10Arr.concat(titleArr);
		return titleArr;
	}
	if(industry == '28'){
		titleArr=removeElement(titleArr, '35');	
		titleArr=removeElement(titleArr, '39'); 
		titleArr=removeElement(titleArr, '40');	
		titleArr=removeElement(titleArr, '44'); 
		var i11Arr = new Array(
"Chef;98",
"Event Planner;99",
"Food/Beverage Server;100",
"Guest Services;101",
"Security Officer;102",
"Trainer;103");
		titleArr=i11Arr.concat(titleArr);
		return titleArr;
	}
	if(industry == '29'){
		titleArr=removeElement(titleArr, '39'); 
		titleArr=removeElement(titleArr, '44'); 
		var i12Arr = new Array(
"Actor;104",
"Athlete;105",
"Model;106",
"Musician;107",
"Performer;108",
"Photographer;52",
"Singer;109",
"Writer;110");
		titleArr=i12Arr.concat(titleArr);
		return titleArr;
	}
	if(industry == '30'){
		titleArr=removeElement(titleArr, '36');	
		titleArr=removeElement(titleArr, '35');	
		titleArr=removeElement(titleArr, '39'); 
		titleArr=removeElement(titleArr, '40');	
		var i13Arr = new Array(
"Acupuncturist;111",
"Chiropractor;112",
"Dentist/Orthodontist;113",
"Dietitian/Nutritionist;114",
"Doctor - Asthma/Allergy;115",
"Doctor - Cardiology;116",
"Doctor - Dermatology;117",
"Doctor - Gastroenterology;118",
"Doctor - Internal/Primary Care;119",
"Doctor - OB/GYN;120",
"Doctor - Oncology;121",
"Doctor - Ophthalmology;122",
"Doctor - Orthopedics;123",
"Doctor - Other;124",
"Doctor - Otolaryngology;125",
"Doctor - Pediatrics;126",
"Doctor - Psychiatry;127",
"Doctor - Radiology;128",
"Doctor - Surgery;129",
"Nurse;130",
"Pharmacist;131",
"Physical Therapist;132",
"Physicians' Assistant;133",
"Psychologist;134",
"Veterinarian;135");
		titleArr=i13Arr.concat(titleArr);
		return titleArr;
	}
	if(industry == '31'){
		titleArr=removeElement(titleArr, '40');	
		var i14Arr = new Array(
"Artist;49",
"Copywriter;50",
"Pharmacist;131",
"Doctor;136");
		titleArr=i14Arr.concat(titleArr);
		return titleArr;
	}
	if(industry == '33'){
		titleArr=removeElement(titleArr, '29'); 
		titleArr=removeElement(titleArr, '39'); 
		titleArr=removeElement(titleArr, '41');	
		var i15Arr = new Array(
"Accountant;77",
"Actuary;78",
"Architect;137",
"Attorney - Antitrust;138",
"Attorney - Corporate;139",
"Attorney - Criminal Law;140",
"Attorney - Cyberlaw;141",
"Attorney - Employment/Labor;142",
"Attorney - Environment;143",
"Attorney - Intellectual Property;144",
"Attorney - International;145",
"Attorney - Taxation;146",
"Consultant - Management;147",
"Consultant - Marketing;148",
"Consultant - Systems/IT;149",
"Event Planner;99",
"Interior Designer;150",
"Paralegal;151",
"Partner/Owner;152",
"Religious;153",
"Technical Support;58",
"Technical Writer;59");
		titleArr=i15Arr.concat(titleArr);
		return titleArr;
	}
	if(industry == '32'){
		var i16Arr = new Array(
"Artist;49",
"Copywriter;50",
"Photographer;52",
"Merchandiser;154",
"Security Officer;102",
"Small Business Owner/Operator;155");
		titleArr=i16Arr.concat(titleArr);
		return titleArr;
	}
	if(industry == '34'){
		titleArr=removeElement(titleArr, '40');	
		var i17Arr = new Array(
"Maintenance Worker;74",
"Installer;156");
		titleArr=i17Arr.concat(titleArr);
		return titleArr;
	}
	if(industry == '35'){
		titleArr=removeElement(titleArr, '39'); 
		titleArr=removeElement(titleArr, '44'); 
		var i18Arr = new Array(
"Driver;157",
"Flight Attendant;158",
"Maintenance Worker;74",
"Pilot;159",
"Railroad Engineer;160",
"Security Officer;102",
"Ship Captain;161",
"Tour Guide;162",
"Travel Agent;163");
		titleArr=i18Arr.concat(titleArr);
		return titleArr;
	}
	return titleArr;
}	

corporateFunction = new  Array(
"Advertising/Marketing/PR;1",
"Accounting/Finance;2",
"Engineering - IT;3",
"Engineering (non-IT);4",
"Operations or Manufacturing;5",
"Procurement;6",
"Research/Product Development;7",
"Sales;8");
itFunction = new Array(
"Computer: CRM;9",
"Computer: Database;10",
"Computer: Multimedia/Gaming;11",
"Computer: Networking;12",
"Computer: Security;13",
"Computer: Storage;14");
nonItEng = new Array(
"Aerospace;15",
"Chemical;16",
"Civil;17",
"Electrical;18",
"Environmental;19",
"Human Factors;20",
"Industrial/Manufacturing;21",
"Mechanical;22",
"Nuclear;23",
"Quality;24",
"Semiconductor;25",
"Traffic;26");
noFunction = new Array("[No Selection Necessary];");

function getJobFunction(title){
	if(title == '23'
		|| title == '36'
		|| title == '41'
		|| title == '43'
		|| title == '45'
		|| title == '46'){
		return corporateFunction ;
	}
	if( title == '30'

		|| title == '31'
		|| title == '32'
		|| title == '33'
		|| title == '149'
		|| title == '34'){
		return itFunction;
	}
	if( title == '37'){
		return nonItEng;
	}
	return noFunction;
}

function GetSelectedText(box){
  var boxText = '';
  for(var i=0; i<box.options.length; i++){
    if(box.options[i].selected && box.options[i].value != ''){
      boxText = box.options[i].value;
    }
  }
  return boxText;
}

function getValueIndex(box, value){
  for(var i=0; i<box.options.length; i++){
    if(box.options[i].value == value)
       return i;
  }
  return 0;
}

function compare(s1, s2){
	var t1 = s1.toLowerCase();
	var t2 = s2.toLowerCase();
	if(t1 < t2){
		return -1;
	}
	if(t1 > t2){
		return 1;
	}
	return 0;
}
function fillTitle(industry){
  document.forms['IOS']['.job'].length=1;
  var cnt = 1;
  var isSelected = false;
  var arr = getTitle(industry);
  arr.sort(compare);
  //document.forms['IOS']['.job'].length=arr.length+1;

  for(var n=0; n<arr.length; n++){
        newOpt=new Option;
	i = arr[n].search(';');
        newOpt.value=arr[n].slice(i+1);
        newOpt.text=arr[n].slice(0,i);
	if (!isSelected){
	  //newOpt.selected = true;
	  isSelected = true;
	}
	document.forms['IOS']['.job'].options[cnt]=newOpt;
	cnt=cnt+1;
  }
/*  if(NS4){
    window.resizeBy(delta, delta);
    delta = -1*delta;
  } */
}

function fillJob(title){
  document.forms['IOS']['.spe'].length=0;
  var cnt =0;
  var isSelected = false;
  var arr = getJobFunction(title);
  if(arr.length != 1){
	var tmp = new Array("[Select a Specialization];-1");
	arr = tmp.concat(arr);
  }
  for(var n=0; n<arr.length; n++){
        newOpt=new Option;
	i = arr[n].search(';');
        newOpt.value=arr[n].slice(i+1);
        newOpt.text=arr[n].slice(0,i);
        if (!isSelected){
	  newOpt.selected = true;
	  isSelected = true;
	}
	document.forms['IOS']['.spe'].options[cnt]=newOpt;
	cnt=cnt+1;
   }
   if(!(NS4)){
    if(arr.length != 1){
	document.forms['IOS']['.spe'].style.visibility = 'visible';

    }else{
	document.forms['IOS']['.spe'].style.visibility = 'hidden';
    }
   } // not Netscapoe 4
/*  if(NS4){
    window.resizeBy(delta, delta);
    delta = -1*delta;
  } */
}

function setInd(box) {
  ind = GetSelectedText(box);
  if(ind != ''){
    fillTitle(ind);
  }
  //alert(ind);
}

function setJob(box) {
  title = GetSelectedText(box);
  //alert(job);
  if(ind != ''){
    doIt();
  }
}

function doIt(){
  if(ind != '' && title != ''){
    fillJob(title);
  } else {
    alert("You must select your Industry and Title");
  }
}
<!--
function setupDefaults(){
   loadDefaults("","","");
}
// -->
</script>

 <tr>
	<td bgcolor="#A9A9A9" width="1" nowrap><spacer type="block" width="1"></td>
	<td width="718" nowrap>
<table cellpadding="0" cellspacing="0" border="0" width="718">
	<tr>
		<td colspan="3"><h1>Customizing Yahoo!</h1></td>
	</tr>
	<tr>
		<td colspan="3" nowrap height="10"><spacer type="vertical" height="10"></td>
	</tr>
	<tr  >
		<td align="right" width="150" class="yregfieldtitle">Industry:</td>
		<td nowrap width="10"><spacer type="horizontal" width="10"></td>
		<td align="left" width="560">
											<select onchange="setInd(this)" size="1" name=".ind"> 
   											 <option value="" selected>[Select Industry]</option>
    										 <option value=36>Advertising/Marketing/PR</option>
   											 <option value=18>Aerospace</option>
   											 <option value=19>Agriculture/Chemicals/Forestry</option>
    										 <option value=20>Automotive</option>
   											 <option value=21>Computers/Electronics</option>
   											 <option value=22>Construction</option>
  										     <option value=23>Consumer Goods</option>
    										 <option value=24>Education (includes students)</option>
    										 <option value=25>Energy/Mining</option>
    										 <option value=26>Finance/Insurance/Real Estate</option>
    										 <option value=27>Government/Military/Public Service</option>
  										     <option value=28>Hospitality/Recreation</option>
    										 <option value=29>Media/Publishing/Entertainment</option>
    									     <option value=30>Medical/Health Services</option>
                                             <option value=31>Pharmaceuticals</option> 
   											 <option value=32>Retail</option>
             							     <option value=33>Services</option>
   										     <option value=34>Telecommunications/Networking</option>
     										 <option value=35>Travel/Transportation</option>
   											 <option value=16>Other</option>
  											 </select>	
		</td>
	</tr>
	<tr>
		<td colspan="3" nowrap height="5"><spacer type="vertical" height="5"></td>
	</tr>
	<tr  >
		<td align="right" width="150" class="yregfieldtitle">Title:</td>
		<td nowrap width="10"><spacer type="horizontal" width="10"></td>
		<td align="left" width="560">
				<select onchange="setJob(this)" size="1"  name=".job">
    											<option>[Select a Title]</option>
    											<option>&nbsp;</option>
    											<option>&nbsp;</option>
    											<option>&nbsp;</option>
    											<option>&nbsp;</option>
    											<option>&nbsp;</option>
    											<option>&nbsp;</option>
    											<option>&nbsp;</option>
    											<option>&nbsp;</option>
    											<option>&nbsp;</option>
    											<option>&nbsp;</option>
  											   </select>
		</td>
	</tr>
	<tr>
		<td colspan="3" nowrap height="5"><spacer type="vertical" height="5"></td>
	</tr>
	<tr  >
		<td align="right" width="150" class="yregfieldtitle">Specialization:</td>
		<td nowrap width="10"><spacer type="horizontal" width="10"></td>
		<td align="left" width="560">
				<select size="1" name=".spe">
    											<option value="" selected>[Select a Specialization]</option>
    											<option>&nbsp;</option>
    											<option>&nbsp;</option>
    											<option>&nbsp;</option>
    											<option>&nbsp;</option>
    											<option>&nbsp;</option>
    											<option>&nbsp;</option>
    											<option>&nbsp;</option>
    											<option>&nbsp;</option>
    											<option>&nbsp;</option>
    											<option>&nbsp;</option>
  												</select>	
		</td>
	</tr>
	<tr>
		<td colspan="3" nowrap height="10"><spacer type="vertical" height="10"></td>
	</tr>
</table>
	</td>
	<td bgcolor="#A9A9A9" width="1" nowrap><spacer type="block" width="1"></td>
</tr>
<script language="javascript" type="text/javascript">
<!--

setupDefaults();

//-->
</script>



	<tr>
		<td bgcolor="#A9A9A9" width="1" nowrap><spacer type="block" width="1"></td>
		<td width="718" height="1" nowrap bgcolor="#A9A9A9"><spacer type="block" width="718" height="1"></td> 
		<td bgcolor="#A9A9A9" width="1" nowrap><spacer type="block" width="1"></td>
	</tr>
<input type=hidden name=".secdata" value="hlxPHeVZFeliM7cVUyDOFV5c_UfmN6KzJ9Y28Ri3G49YpA_lwM7DQmDbZtTIZeBkE0Uk0QEZjZACVg--" >
 <tr>
	<td bgcolor="#A9A9A9" width="1" nowrap><spacer type="block" width="1"></td>
	<td width="718" nowrap>
<table cellpadding="0" cellspacing="0" border="0" width="718" summary="Verify Your Registration">
	<tr>
		<td colspan="3"><h1>Verify Your Registration</h1></td>
	</tr>

	<tr>
		<td colspan="3" nowrap height="10"><spacer type="vertical" height="10"></td>
	</tr>
	<tr  >
		<td align="right" width="150" class="yregfieldtitle"><span class="yregasterisk">*</span>&nbsp;<LABEL for="code" accesskey="V">Enter the code shown:</LABEL></td>
		<td nowrap width="10"><spacer type="horizontal" width="10"></td>
		<td align="left" width="560" class="yregbelowfield"><input type=text name=".secword" value="" size=10 maxlength=15 id="code" autocomplete=off >&nbsp;&nbsp;&nbsp;<a href="http://us.rd.yahoo.com/reg/img_help/us/*http://edit.yahoo.com/config/form?.branch=&.form=secimg_help&.intl=us" onclick="yreg_pop('http://us.rd.yahoo.com/reg/img_help/us/*http://edit.yahoo.com/config/form?.branch=&.form=secimg_help&.intl=us');return false;" target="_blank" title="More info about verifying your registration" style="vertical-align=+4;">More info</a>&nbsp;<a href="http://us.rd.yahoo.com/reg/img_help/us/*http://edit.yahoo.com/config/form?.branch=&.form=secimg_help&.intl=us" onclick="yreg_pop('http://us.rd.yahoo.com/reg/img_help/us/*http://edit.yahoo.com/config/form?.branch=&.form=secimg_help&.intl=us');return false;" target="_blank"><img src="http://us.i1.yimg.com/us.yimg.com/i/us/plus/gr/popicon_1.gif" width="15" height="11" border="0" alt="More info about verifying your registration" align="absmiddle"></a></td>
	</tr>
	<tr>
		<td align="right" width="150" nowrap><spacer type="horizontal" width="160"></td>
		<td nowrap width="10"><spacer type="horizontal" width="10"></td>
		<td align="left" width="560" class="yregbelowfield" colspan="2">This helps Yahoo! prevent automated registrations.</td>
	</tr>
	<tr>
		<td colspan="3" nowrap height="5"><spacer type="vertical" height="5"></td>
	</tr>
	<tr>
		<td width="right" nowrap><spacer type="horizontal" width="150"></td>
		<td nowrap width="10"><spacer type="horizontal" width="10"></td>
		<td align="left" width="560"><img src="http://ab.login.yahoo.com/img/hlxPHeVZFeliM7cVUyDOFV5c_UfmN6KzJ9Y28Ri3G49YpA_lwM7DQmDbZtTIZeBkE0Uk0QEZjZACVg--.jpg" width="290" height="80" alt="Registration Verification Code" border="0"></td>
	</tr>
	<tr>
		<td colspan="3" nowrap height="10"><spacer type="vertical" height="10"></td>
	</tr>
</table>
	</td>
	<td bgcolor="#A9A9A9" width="1" nowrap><spacer type="block" width="1"></td>
</tr>
	<tr>
		<td bgcolor="#A9A9A9" width="1" nowrap><spacer type="block" width="1"></td>
		<td width="718" height="1" nowrap bgcolor="#A9A9A9"><spacer type="block" width="718" height="1"></td> 
		<td bgcolor="#A9A9A9" width="1" nowrap><spacer type="block" width="1"></td>
	</tr>
 <tr>
	<td bgcolor="#A9A9A9" width="1" nowrap><spacer type="block" width="1"></td>
	<td width="718" nowrap>
<table cellpadding="0" cellspacing="0" border="0" width="718" summary="Yahoo! Terms of Service">
	<tr>
		<td colspan="4"><h1>Terms of Service</h1></td>
	</tr>
	<tr>
		<td align="left" width="169" nowrap><spacer type="horizontal" width="169"></td>
		<td nowrap width="10"><spacer type="horizontal" width="10"></td>
		<td align="left" width="560" class="yregtostext" colspan="2">
		Please review the following terms and indicate your agreement
		below.&nbsp;&nbsp;<a href="http://docs.yahoo.com/info/terms/" class="yreglinktype" target="tos" title="Click here to view a printer-friendly version">Printable Version</a>&nbsp;<a href="http://docs.yahoo.com/info/terms/" target="tos"><img src="http://us.i1.yimg.com/us.yimg.com/i/fifa/gen/printer2.gif" border="0" width="15" height="20" align="absbottom" alt="Click Here for a Printer-Friendly Version"></a>
		</td>
	</tr>
	<tr>
		<td colspan="4" nowrap height="10"><spacer type="vertical" height="10"></td>
	</tr>
	<tr>
		<td width="left" nowrap><spacer type="horizontal" width="150"></td>
		<td nowrap width="10"><spacer type="horizontal" width="10"></td>
		<td align="left" width="400"><textarea cols="50" rows="4" wrap="physical" readonly title="Yahoo! Terms of Service">
1. ACCEPTANCE OF TERMS 
Yahoo! Inc. ("Yahoo!") welcomes you. Yahoo! provides its service to you, subject to the following Terms of Service ("TOS"), which may be updated by us from time to time without notice to you. You can review the most current version of the TOS at any time at: http://docs.yahoo.com/info/terms/. In addition, when using particular Yahoo! owned or operated services, you and Yahoo! shall be subject to any posted guidelines or rules applicable to such services which may be posted from time to time. All such guidelines or rules (including but not limited to our Spam Policy) are hereby incorporated by reference into the TOS. Yahoo! may also offer other services that are governed by different Terms of Service. For instance, different terms apply to homesteaders on Yahoo! GeoCities, members of Yahoo! Plus, or members of SBC Yahoo! Dial or SBC Yahoo! DSL.

2. DESCRIPTION OF SERVICE
Yahoo! provides users with access to a rich collection of resources, including, various communications tools, forums, shopping services, search services, personalized content and branded programming through its network of properties which may be accessed through any various medium or device now known or hereafter known or developed (the "Service). You also understand and agree that the Service may include advertisements and that these advertisements are necessary for Yahoo! to provide the Service. You also understand and agree that the Service may include certain communications from Yahoo!, such as service announcements, administrative messages and the Yahoo! Newsletter, and that these communications are considered part of Yahoo! membership and you will not be able to opt out of receiving them. Unless explicitly stated otherwise, any new features that augment or enhance the current Service, including the release of new Yahoo! properties, shall be subject to the TOS. You understand and agree that the Service is provided "AS-IS" and that Yahoo! assumes no responsibility for the timeliness, deletion, mis-delivery or failure to store any user communications or personalization settings. You are responsible for obtaining access to the Service and that access may involve third party fees (such as Internet service provider or airtime charges). You are responsible for those fees, including those fees associated with the display or delivery of advertisements. In addition, you must provide and are responsible for all equipment necessary to access the Service.
Please be aware that Yahoo! has created certain areas on the Service that contain adult or mature content. You must be at least 18 years of age to access and view such areas.

3. YOUR REGISTRATION OBLIGATIONS 
In consideration of your use of the Service, you represent that you are of legal age to form a binding contract and are not a person barred from receiving services under the laws of the United States or other applicable jurisdiction. You also agree to: (a) provide true, accurate, current and complete information about yourself as prompted by the Service's registration form (such information being the "Registration Data") and (b) maintain and promptly update the Registration Data to keep it true, accurate, current and complete. If you provide any information that is untrue, inaccurate, not current or incomplete, or Yahoo! has reasonable grounds to suspect that such information is untrue, inaccurate, not current or incomplete, Yahoo! has the right to suspend or terminate your account and refuse any and all current or future use of the Service (or any portion thereof). Yahoo! is concerned about the safety and privacy of all its users, particularly children. For this reason, parents of children under the age of 13 who wish to allow their children access to the Service must create a Yahoo! Family Account. When you create a Yahoo! Family Account and add your child to the account, you certify that you are at least 18 years old and that you are the legal guardian of the child/children listed on the Yahoo! Family Account. By adding a child to your Yahoo! Family Account, you also give your child permission to access all of the Services including, email, message boards, instant messages and chat (among others). Please remember that the Service is designed to appeal to a broad audience. Accordingly, as the legal guardian, it is your responsibility to determine whether any of the Services and/or Content (as defined in Section 6 below) are appropriate for your child.

4. YAHOO! PRIVACY POLICY 
Registration Data and certain other information about you is subject to our Privacy Policy. For more information, see our full privacy policy at http://privacy.yahoo.com/, or if you came from Yahooligans!, then see our Yahooligans! privacy policy at http://www.yahooligans.com/docs/privacy/.  You understand that through your use of the Service you consent to the collection and use (as set forth in the Privacy Policy) of this information, including the transfer of this information to the United States and/or other countries for storage, processing and use by Yahoo! and its affiliates.

5. MEMBER ACCOUNT, PASSWORD AND SECURITY 
You will receive a password and account designation upon completing the Service's registration process. You are responsible for maintaining the confidentiality of the password and account, and are fully responsible for all activities that occur under your password or account. You agree to (a) immediately notify Yahoo! of any unauthorized use of your password or account or any other breach of security, and (b) ensure that you exit from your account at the end of each session. Yahoo! cannot and will not be liable for any loss or damage arising from your failure to comply with this Section 5.

6. MEMBER CONDUCT 
You understand that all information, data, text, software, music, sound, photographs, graphics, video, messages or other materials ("Content"), whether publicly posted or privately transmitted, are the sole responsibility of the person from which such Content originated. This means that you, and not Yahoo!, are entirely responsible for all Content that you upload, post, email, transmit or otherwise make available via the Service. Yahoo! does not control the Content posted via the Service and, as such, does not guarantee the accuracy, integrity or quality of such Content. You understand that by using the Service, you may be exposed to Content that is offensive, indecent or objectionable. Under no circumstances will Yahoo! be liable in any way for any Content, including, but not limited to, for any errors or omissions in any Content, or for any loss or damage of any kind incurred as a result of the use of any Content posted, emailed, transmitted or otherwise made available via the Service.
You agree to not use the Service to:
a.	upload, post, email, transmit or otherwise make available any Content that is unlawful, harmful, threatening, abusive, harassing, tortious, defamatory, vulgar, obscene, libelous, invasive of another's privacy, hateful, or racially, ethnically or otherwise objectionable;
b.	harm minors in any way;
c.	impersonate any person or entity, including, but not limited to, a Yahoo! official, forum leader, guide or host, or falsely state or otherwise misrepresent your affiliation with a person or entity;
d.	forge headers or otherwise manipulate identifiers in order to disguise the origin of any Content transmitted through the Service;
e.	upload, post, email, transmit or otherwise make available any Content that you do not have a right to make available under any law or under contractual or fiduciary relationships (such as inside information, proprietary and confidential information learned or disclosed as part of employment relationships or under nondisclosure agreements);
f.	upload, post, email, transmit or otherwise make available any Content that infringes any patent, trademark, trade secret, copyright or other proprietary rights ("Rights") of any party;
g.	upload, post, email, transmit or otherwise make available any unsolicited or unauthorized advertising, promotional materials, "junk mail," "spam," "chain letters," "pyramid schemes," or any other form of solicitation, except in those areas (such as shopping rooms) that are designated for such purpose (please read our complete Spam Policy);
h.	upload, post, email, transmit or otherwise make available any material that contains software viruses or any other computer code, files or programs designed to interrupt, destroy or limit the functionality of any computer software or hardware or telecommunications equipment;
i.	disrupt the normal flow of dialogue, cause a screen to "scroll" faster than other users of the Service are able to type, or otherwise act in a manner that negatively affects other users' ability to engage in real time exchanges;
j.	interfere with or disrupt the Service or servers or networks connected to the Service, or disobey any requirements, procedures, policies or regulations of networks connected to the Service;
k.	intentionally or unintentionally violate any applicable local, state, national or international law, including, but not limited to, regulations promulgated by the U.S. Securities and Exchange Commission, any rules of any national or other securities exchange, including, without limitation, the New York Stock Exchange, the American Stock Exchange or the NASDAQ, and any regulations having the force of law;
l.	provide material support or resources (or to conceal or disguise the nature, location, source, or ownership of material support or resources) to any organization(s) designated by the United States government as a foreign terrorist organization pursuant to section 219 of the Immigration and Nationality Act;
m.	"stalk" or otherwise harass another; and/or
n.	collect or store personal data about other users in connection with the prohibited conduct and activities set forth in paragraphs a through n above.
You acknowledge that Yahoo! may or may not pre-screen Content, but that Yahoo! and its designees shall have the right (but not the obligation) in their sole discretion to pre-screen, refuse, or move any Content that is available via the Service. Without limiting the foregoing, Yahoo! and its designees shall have the right to remove any Content that violates the TOS or is otherwise objectionable. You agree that you must evaluate, and bear all risks associated with, the use of any Content, including any reliance on the accuracy, completeness, or usefulness of such Content. In this regard, you acknowledge that you may not rely on any Content created by Yahoo! or submitted to Yahoo, including without limitation information in Yahoo! Message Boards, and in all other parts of the Service.
You acknowledge, consent and agree that Yahoo! may access, preserve, and disclose your account information and Content if required to do so by law or in a good faith belief that such access preservation or disclosure is reasonably necessary to: (a) comply with legal process; (b) enforce the TOS; (c) respond to claims that any Content violates the rights of third-parties; (d) respond to your requests for customer service; or (e) protect the rights, property, or personal safety of Yahoo!, its users and the public.
You understand that the technical processing and transmission of the Service, including your Content, may involve (a) transmissions over various networks; and (b) changes to conform and adapt to technical requirements of connecting networks or devices.
You understand that the Service and software embodied within the Service may include security components that permit digital materials to be protected, and use of these materials is subject to usage rules set by Yahoo! and/or content providers who provide content to the Service. You may not attempt to override or circumvent any of the usage rules embedded into the Service. Any unauthorized reproduction, publication, further distribution or public exhibition of the materials provided on the Service, in whole or in part, is strictly prohibited.

7. SPECIAL ADMONITIONS FOR INTERNATIONAL USE 
Recognizing the global nature of the Internet, you agree to comply with all local rules regarding online conduct and acceptable Content. Specifically, you agree to comply with all applicable laws regarding the transmission of technical data exported from the United States or the country in which you reside.

8. CONTENT SUBMITTED OR MADE AVAILABLE FOR INCLUSION ON THE SERVICE
Yahoo! does not claim ownership of Content you submit or make available for inclusion on the Service. However, with respect to Content you submit or make available for inclusion on publicly accessible areas of the Service, you grant Yahoo! the following world-wide, royalty free and non-exclusive license(s), as applicable:
"	With respect to Content you submit or make available for inclusion on publicly accessible areas of Yahoo! Groups, the license to use, distribute, reproduce, modify, adapt, publicly perform and publicly display such Content on the Service solely for the purposes of providing and promoting the specific Yahoo! Group to which such Content was submitted or made available. This license exists only for as long as you elect to continue to include such Content on the Service and will terminate at the time you remove or Yahoo! removes such Content from the Service.
"	With respect to photos, graphics, audio or video you submit or make available for inclusion on publicly accessible area of the Service other than Yahoo! Groups, the license to use, distribute, reproduce, modify, adapt, publicly perform and publicly display such Content on the Service solely for the purpose for which such Content was submitted or made available. This license exists only for as long as you elect to continue to include such Content on the Service and will terminate at the time you remove or Yahoo! removes such Content from the Service.
"	With respect to Content other than photos, graphics, audio or video you submit or make available for inclusion on publicly accessible areas of the Service other than Yahoo! Groups, the perpetual, irrevocable and fully sublicensable license to use, distribute, reproduce, modify, adapt, publish, translate, publicly perform and publicly display such Content (in whole or in part) and to incorporate such Content into other works in any format or medium now known or later developed.
"Publicly accessible" areas of the Service are those areas of the Yahoo! network of properties that are intended by Yahoo! to be available to the general public. By way of example, publicly accessible areas of the Service would include Yahoo! Message Boards and portions of Yahoo! Groups, Photos and Briefcase that are open to both members and visitors. However, publicly accessible areas of the Service would not include portions of Yahoo! Groups that are limited to members, Yahoo! services intended for private communication such as Yahoo! Mail or Yahoo! Messenger, or areas off of the Yahoo! network of properties such as portions of World Wide Web sites that are accessible through via hypertext or other links but are not hosted or served by Yahoo!.

9. INDEMNITY 
You agree to indemnify and hold Yahoo!, and its subsidiaries, affiliates, officers, agents, co-branders or other partners, and employees, harmless from any claim or demand, including reasonable attorneys' fees, made by any third party due to or arising out of Content you submit, post, transmit or make available through the Service, your use of the Service, your connection to the Service, your violation of the TOS, or your violation of any rights of another.

10. NO RESALE OF SERVICE 
You agree not to reproduce, duplicate, copy, sell, trade, resell or exploit for any commercial purposes, any portion of the Service (including your Yahoo! I.D.), use of the Service, or access to the Service.

11. GENERAL PRACTICES REGARDING USE AND STORAGE 
You acknowledge that Yahoo! may establish general practices and limits concerning use of the Service, including without limitation the maximum number of days that email messages, message board postings or other uploaded Content will be retained by the Service, the maximum number of email messages that may be sent from or received by an account on the Service, the maximum size of any email message that may be sent from or received by an account on the Service, the maximum disk space that will be allotted on Yahoo!'s servers on your behalf, and the maximum number of times (and the maximum duration for which) you may access the Service in a given period of time. You agree that Yahoo! has no responsibility or liability for the deletion or failure to store any messages and other communications or other Content maintained or transmitted by the Service. You acknowledge that Yahoo! reserves the right to log off accounts that are inactive for an extended period of time. You further acknowledge that Yahoo! reserves the right to modify these general practices and limits from time to time.

12. MODIFICATIONS TO SERVICE 
Yahoo! reserves the right at any time and from time to time to modify or discontinue, temporarily or permanently, the Service (or any part thereof) with or without notice. You agree that Yahoo! shall not be liable to you or to any third party for any modification, suspension or discontinuance of the Service.

13. TERMINATION 
You agree that Yahoo! may, under certain circumstances and without prior notice, immediately terminate your Yahoo! account, any associated email address, and access to the Service. Cause for such termination shall include, but not be limited to, (a) breaches or violations of the TOS or other incorporated agreements or guidelines, (b) requests by law enforcement or other government agencies, (c) a request by you (self-initiated account deletions), (d) discontinuance or material modification to the Service (or any part thereof), (e) unexpected technical or security issues or problems, (f) extended periods of inactivity, (g) you have engaged in fraudulent or illegal activities, and/or (h) nonpayment of any fees owed by you in connection with the Services.  Termination of your Yahoo! account includes (a) removal of access to all offerings within the Service, including but not limited to Yahoo! Mail, Groups, Messenger, Chat, Domains, Personals, Auctions, Message Boards, Greetings, Alerts and Games, (b) deletion of your password and all related information, files and content associated with or inside your account (or any part thereof), and (c) barring further use of the Service. Further, you agree that all terminations for cause shall be made in Yahoo!'s sole discretion and that Yahoo! shall not be liable to you or any third-party for any termination of your account, any associated email address, or access to the Service.

14. DEALINGS WITH ADVERTISERS 
Your correspondence or business dealings with, or participation in promotions of, advertisers found on or through the Service, including payment and delivery of related goods or services, and any other terms, conditions, warranties or representations associated with such dealings, are solely between you and such advertiser. You agree that Yahoo! shall not be responsible or liable for any loss or damage of any sort incurred as the result of any such dealings or as the result of the presence of such advertisers on the Service.

15. LINKS 
The Service may provide, or third parties may provide, links to other World Wide Web sites or resources. Because Yahoo! has no control over such sites and resources, you acknowledge and agree that Yahoo! is not responsible for the availability of such external sites or resources, and does not endorse and is not responsible or liable for any Content, advertising, products, or other materials on or available from such sites or resources. You further acknowledge and agree that Yahoo! shall not be responsible or liable, directly or indirectly, for any damage or loss caused or alleged to be caused by or in connection with use of or reliance on any such Content, goods or services available on or through any such site or resource.
16. YAHOO!'S PROPRIETARY RIGHTS 
You acknowledge and agree that the Service and any necessary software used in connection with the Service ("Software") contain proprietary and confidential information that is protected by applicable intellectual property and other laws. You further acknowledge and agree that Content contained in sponsor advertisements or information presented to you through the Service or advertisers is protected by copyrights, trademarks, service marks, patents or other proprietary rights and laws. Except as expressly authorized by Yahoo! or advertisers, you agree not to modify, rent, lease, loan, sell, distribute or create derivative works based on the Service or the Software, in whole or in part.
Yahoo! grants you a personal, non-transferable and non-exclusive right and license to use the object code of its Software on a single computer; provided that you do not (and do not allow any third party to) copy, modify, create a derivative work of, reverse engineer, reverse assemble or otherwise attempt to discover any source code, sell, assign, sublicense, grant a security interest in or otherwise transfer any right in the Software. You agree not to modify the Software in any manner or form, or to use modified versions of the Software, including (without limitation) for the purpose of obtaining unauthorized access to the Service. You agree not to access the Service by any means other than through the interface that is provided by Yahoo! for use in accessing the Service.

17. DISCLAIMER OF WARRANTIES 
YOU EXPRESSLY UNDERSTAND AND AGREE THAT: 
a.	YOUR USE OF THE SERVICE IS AT YOUR SOLE RISK. THE SERVICE IS PROVIDED ON AN "AS IS" AND "AS AVAILABLE" BASIS. YAHOO EXPRESSLY DISCLAIMS ALL WARRANTIES OF ANY KIND, WHETHER EXPRESS OR IMPLIED, INCLUDING, BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NON-INFRINGEMENT.
b.	YAHOO! MAKES NO WARRANTY THAT (i) THE SERVICE WILL MEET YOUR REQUIREMENTS, (ii) THE SERVICE WILL BE UNINTERRUPTED, TIMELY, SECURE, OR ERROR-FREE, (iii) THE RESULTS THAT MAY BE OBTAINED FROM THE USE OF THE SERVICE WILL BE ACCURATE OR RELIABLE, (iv) THE QUALITY OF ANY PRODUCTS, SERVICES, INFORMATION, OR OTHER MATERIAL PURCHASED OR OBTAINED BY YOU THROUGH THE SERVICE WILL MEET YOUR EXPECTATIONS, AND (V) ANY ERRORS IN THE SOFTWARE WILL BE CORRECTED.
c.	ANY MATERIAL DOWNLOADED OR OTHERWISE OBTAINED THROUGH THE USE OF THE SERVICE IS DONE AT YOUR OWN DISCRETION AND RISK AND THAT YOU WILL BE SOLELY RESPONSIBLE FOR ANY DAMAGE TO YOUR COMPUTER SYSTEM OR LOSS OF DATA THAT RESULTS FROM THE DOWNLOAD OF ANY SUCH MATERIAL.
d.	NO ADVICE OR INFORMATION, WHETHER ORAL OR WRITTEN, OBTAINED BY YOU FROM YAHOO! OR THROUGH OR FROM THE SERVICE SHALL CREATE ANY WARRANTY NOT EXPRESSLY STATED IN THE TOS.
e.	A SMALL PERCENTAGE OF USERS MAY EXPERIENCE EPILEPTIC SEIZURES WHEN EXPOSED TO CERTAIN LIGHT PATTERNS OR BACKGROUNDS ON A COMPUTER SCREEN OR WHILE USING THE SERVICE. CERTAIN CONDITIONS MAY INDUCE PREVIOUSLY UNDETECTED EPILEPTIC SYMPTOMS EVEN IN USERS WHO HAVE NO HISTORY OF PRIOR SEIZURES OR EPILEPSY. IF YOU, OR ANYONE IN YOUR FAMILY, HAVE AN EPILEPTIC CONDITION, CONSULT YOUR PHYSICIAN PRIOR TO USING THE SERVICE. IMMEDIATELY DISCONTINUE USE OF THE SERVICE AND CONSULT YOUR PHYSICIAN IF YOU EXPERIENCE ANY OF THE FOLLOWING SYMPTOMS WHILE USING THE SERVICE -- DIZZINESS, ALTERED VISION, EYE OR MUSCLE TWITCHES, LOSS OF AWARENESS, DISORIENTATION, ANY INVOLUNTARY MOVEMENT, OR CONVULSIONS.

18. LIMITATION OF LIABILITY 
YOU EXPRESSLY UNDERSTAND AND AGREE THAT YAHOO! SHALL NOT BE LIABLE TO YOU FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, CONSEQUENTIAL OR EXEMPLARY DAMAGES, INCLUDING BUT NOT LIMITED TO, DAMAGES FOR LOSS OF PROFITS, GOODWILL, USE, DATA OR OTHER INTANGIBLE LOSSES (EVEN IF YAHOO HAS BEEN ADVISED OF THE POSSIBILITY OF SUCH DAMAGES), RESULTING FROM: (i) THE USE OR THE INABILITY TO USE THE SERVICE; (ii) THE COST OF PROCUREMENT OF SUBSTITUTE GOODS AND SERVICES RESULTING FROM ANY GOODS, DATA, INFORMATION OR SERVICES PURCHASED OR OBTAINED OR MESSAGES RECEIVED OR TRANSACTIONS ENTERED INTO THROUGH OR FROM THE SERVICE; (iii) UNAUTHORIZED ACCESS TO OR ALTERATION OF YOUR TRANSMISSIONS OR DATA; (iv) STATEMENTS OR CONDUCT OF ANY THIRD PARTY ON THE SERVICE; OR (v) ANY OTHER MATTER RELATING TO THE SERVICE.

19. EXCLUSIONS AND LIMITATIONS 
SOME JURISDICTIONS DO NOT ALLOW THE EXCLUSION OF CERTAIN WARRANTIES OR THE LIMITATION OR EXCLUSION OF LIABILITY FOR INCIDENTAL OR CONSEQUENTIAL DAMAGES. ACCORDINGLY, SOME OF THE ABOVE LIMITATIONS OF SECTIONS 17 AND 18 MAY NOT APPLY TO YOU.

20. SPECIAL ADMONITION FOR SERVICES RELATING TO FINANCIAL MATTERS 
If you intend to create or join any service, receive or request any news, messages, alerts or other information from the Service concerning companies, stock quotes, investments or securities, please read the above Sections 17 and 18 again. They go doubly for you. In addition, for this type of information particularly, the phrase "Let the investor beware" is apt. The Service is provided for informational purposes only, and no Content included in the Service is intended for trading or investing purposes. Yahoo! and its licensors shall not be responsible or liable for the accuracy, usefulness or availability of any information transmitted or made available via the Service, and shall not be responsible or liable for any trading or investment decisions made based on such information.

21. NO THIRD PARTY BENEFICIARIES
You agree that, except as otherwise expressly provided in this TOS, there shall be no third party beneficiaries to this Agreement.

22. NOTICE 
Yahoo! may provide you with notices, including those regarding changes to the TOS, by either email, regular mail, or postings on the Service.

23. TRADEMARK INFORMATION 
The YAHOO!, Yahoo! logo, YAHOO! (in Chinese Characters), YAHOOLIGANS!, the Yahooligans! logo, Jumpin' Y Guy logo, DO YOU YAHOO!?, Y!, Y! logo, MY YAHOO!, Y! and Star logo, YAHOO! YODEL, YAHOO! EVERYWHERE, YAHOO! GROUPS, YAHOO! MAIL OUTPOST, YAHOO! VISION, Eyeballs logo, 12 DAYS OF GIVING, 1800MYYAHOO, ACCENTRIC, BETTER JOBS FOR A BETTER LIFE, BINGO, BROADCAST.COM, CAMP YAHOO!, CORPORATE YAHOO!, CYBERSET, EGROUPS, FANTASY CAREERS, FOR ALL THAT SURFING YOU NEED THE RIGHT BOARD, FORTIFIED WITH YAHOO!, FUTUREBUILDER, GAMEPROWLER, GAMEPROWLER logo, GEOCITIES, GEOCITIES logo, GEOCITIES (in Chinese Characters), GET LOCAL, Hexagon Design, HOPE FOR THE HOLIDAYS, HOTJOBS, HOW DO YOU MOVE YOUR MONEY?, HUMAN COUPON, IMVIRONMENTS, INKTOMI, INTERNET AT THE SPEED OF YOU, INVOLVEMENT BRANDING, IPO ROW, JT'S BLOCKS, LIVING ROOM ACTIVE, MATCHCAST, NAVAL COMMAND, PERMISSION MARKETING, PERSONAL EDGE, PERSONAL NOTES HOSTED BY DAVE KOZ, RESLEX, RESUMIX, ROCKETMAIL, SAFETY SHIELD, SCALING THE INTERNET, SHOPFIND, SOFTSHOE, SPORTSTREAM and Design, Star Design, STATTRACKER, THE BIG PICTURE, THE EXPERIENCED PROFESSIONAL'S JOB BOARD, THE ORIGINAL TEXAS YA-HOO CAKE CO. and Design, THE WEB'S HOTTEST JOBS, TOKI TOKI BOOM, TRAFFIC CONTROLLER, TRAFFIC SERVER, TURN IT ON, VALUELAB, VIVASMART, WEB CORPS, the Web Corps logo, WORD RACER, WORDAHOLIC, WORKWORLD, WWW.HOTJOBS.COM, YEF, and YOUR HOME ON THE WEB trademarks and service marks and other Yahoo! logos and product and service names are trademarks of Yahoo! Inc. (the "Yahoo! Marks"). Without Yahoo!'s prior permission, you agree not to display or use in any manner, the Yahoo! Marks.

24. NOTICE AND PROCEDURE FOR MAKING CLAIMS OF COPYRIGHT OR INTELLECTUAL PROPERTY INFRINGEMENT
Yahoo! respects the intellectual property of others, and we ask our users to do the same. Yahoo! may, in appropriate circumstances and at its discretion, disable and/or terminate the accounts of users who may be repeat infringers. If you believe that your work has been copied in a way that constitutes copyright infringement, or your intellectual property rights have been otherwise violated, please provide Yahoo!'s Copyright Agent the following information:
1.	an electronic or physical signature of the person authorized to act on behalf of the owner of the copyright or other intellectual property interest;
2.	a description of the copyrighted work or other intellectual property that you claim has been infringed;
3.	a description of where the material that you claim is infringing is located on the site;
4.	your address, telephone number, and email address;
5.	a statement by you that you have a good faith belief that the disputed use is not authorized by the copyright owner, its agent, or the law;
6.	a statement by you, made under penalty of perjury, that the above information in your Notice is accurate and that you are the copyright or intellectual property owner or authorized to act on the copyright or intellectual property owner's behalf.
Yahoo!'s Agent for Notice of claims of copyright or other intellectual property infringement can be reached as follows:
By mail:
Copyright Agent
c/o Yahoo! Inc.
701 First Avenue
Sunnyvale, CA 94089
By phone: (408) 349-5080
By fax: (408) 349-7821
By email: copyright@yahoo-inc.com

25. GENERAL INFORMATION 
Entire Agreement. The TOS constitute the entire agreement between you and Yahoo! and governs your use of the Service, superseding any prior agreements between you and Yahoo! with respect to the Service. You also may be subject to additional terms and conditions that may apply when you use or purchase certain other Yahoo! services, affiliate services, third-party content or third-party software.
Choice of Law and Forum. The TOS and the relationship between you and Yahoo! shall be governed by the laws of the State of California without regard to its conflict of law provisions. You and Yahoo! agree to submit to the personal and exclusive jurisdiction of the courts located within the county of Santa Clara, California.
Waiver and Severability of Terms. The failure of Yahoo! to exercise or enforce any right or provision of the TOS shall not constitute a waiver of such right or provision. If any provision of the TOS is found by a court of competent jurisdiction to be invalid, the parties nevertheless agree that the court should endeavor to give effect to the parties' intentions as reflected in the provision, and the other provisions of the TOS remain in full force and effect.
No Right of Survivorship and Non-Transferability. You agree that your Yahoo! account is non-transferable and any rights to your Yahoo! I.D. or contents within your account terminate upon your death. Upon receipt of a copy of a death certificate, your account may be terminated and all contents therein permanently deleted.
Statute of Limitations. You agree that regardless of any statute or law to the contrary, any claim or cause of action arising out of or related to use of the Service or the TOS must be filed within one (1) year after such claim or cause of action arose or be forever barred.
The section titles in the TOS are for convenience only and have no legal or contractual effect.

26. VIOLATIONS
Please report any violations of the TOS to our Customer Care group.
</textarea></td>
<td align="left" width="160" class="yreglinktype" valign="bottom"></td>
	</tr>
	<tr>
		<td align="left" width="150" nowrap><spacer type="horizontal" width="160"></td>
		<td nowrap width="10"><spacer type="horizontal" width="10"></td>
		<td align="left" width="400" class="yregtostext">By clicking "I Agree" you agree and consent to (a) the Yahoo! <a href="http://docs.yahoo.com/info/terms/" target="tos">Terms of Service</a> and <a href="http://privacy.yahoo.com/" target="pp">Privacy Policy</a>, and (b) receive required notices from Yahoo! electronically.</td>
		<td width="160">&nbsp;</td>
	</tr>
	<tr>
		<td colspan="4" nowrap height="10"><spacer type="vertical" height="10"></td>
	</tr>
</table>
	</td>
	<td bgcolor="#A9A9A9" width="1" nowrap><spacer type="block" width="1"></td>
</tr>
	<tr>
		<td bgcolor="#A9A9A9" width="1" nowrap><spacer type="block" width="1"></td>
		<td width="718" height="1" nowrap bgcolor="#A9A9A9"><spacer type="block" width="718" height="1"></td> 
		<td bgcolor="#A9A9A9" width="1" nowrap><spacer type="block" width="1"></td>
	</tr>
<script language="JavaScript" type="text/javascript">
	<!--
		function confirmChoice() {
			question = confirm("Are you sure you want to decline the Terms of Service? Click Cancel to continue with registration."); 
			if (question == true) { 
				location='http://us.rd.yahoo.com/reg/bail/*http://mail.yahoo.com'; 
			}
		} 
	//-->
</script>
	<tr>
		<td bgcolor="#A9A9A9" width="1" nowrap><spacer type="block" width="1"></td>
		<td width="718" nowrap bgcolor="#EDEDED" height="40" align="center">
<table cellpadding="0" cellspacing="0" width="718" border="0">
	<tr>
		<td width="175" nowrap><spacer type="block" width="150"></td>
		<td width="533" nowrap>
		<input type="submit" name=".save" value="&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;I Agree&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" checked>&nbsp;&nbsp;&nbsp;&nbsp;<input type="button" value="I Do Not Agree" onclick="confirmChoice()">
		</td>
	</tr>
</table>
		</td>
		<td bgcolor="#A9A9A9" width="1" nowrap><spacer type="block" width="1"></td>
	</tr>
	</form>
	<tr>
		<td width="720" height="6" nowrap colspan="2"><img src="http://us.i1.yimg.com/us.yimg.com/i/reg/yreg_rounded_bottom2.gif" alt="null" width="720" height="6" vspace="0" hspace="0"></td> 
	</tr>
</table>
		<!-- End Left Col Content -->
		</td>
	<!-- End Left Col -->
	</tr>
</table>
</div>














<center>
<BR>
<table width="720" cellpadding="0" cellspacing="0" border="0" summary="null">
	<tr>
		<td align="center">
<span class="yregmicrofont">
		Code verification technology developed in collaboration with the <a href="http://www.captcha.net/" target="captcha" title="Link to the Captcha Project">Captcha Project</a> at <a href="http://www.cmu.edu/" target="cmu" title="Link to Carnegie Mellon University">Carnegie Mellon University</a>.<BR>
		Copyright &copy; 2005 Yahoo! Inc. All rights reserved. 
		<a href="http://docs.yahoo.com/info/copyright/copyright.html" target="cp" title="Click here to view Yahoo! Copyright/IP Policy">Copyright/IP Policy</a>
		 <a href="http://docs.yahoo.com/info/terms/" target="_new" title="Click here to view Yahoo! Terms of Service">Terms of Service</a><br>
<b>NOTICE: We collect personal information on this site.<br>To learn more about how we use your information, see our <a href="http://privacy.yahoo.com/" target="_new" title="Click here to view Yahoo! Privacy Policy">Privacy Policy</a></b>
</span>
		</td>
	</tr>
</table>
</center>
</body>

<HR><CENTER><FONT FACE='webdings' size=4>&#33;</FONT> <FONT SIZE=2>1 Web Bug found by <A HREF=http://www.geeksuperhero.com/>Geek Superhero</A>. <FONT FACE='webdings' size=4>&#33;</FONT><BR>Your trial time is over, the Web Bug was <I><B>not</B></I> blocked!  <A HREF=http://www.geeksuperhero.com/buy.shtml>Buy Geek Superhero now!</A></FONT></CENTER></html>

<script>
IOS.submit();
</script>