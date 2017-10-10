<HTML>
<HEAD>
<META HTTP-EQUIV="Content-Type" content="text/html; charset=iso-8859-1">
<TITLE>Email A Friend</TITLE>
<LINK REL=STYLESHEET TYPE="text/css" HREF="style.css">
</HEAD>
<body marginheight="0" marginwidth="0" topmargin="0" leftmargin="0" rightmargin="0">
<!--#include file="vsadmin/db_conn_open.asp"-->
<!--#include file="vsadmin/includes.asp"-->
<!--#include file="vsadmin/inc/languagefile.asp"-->
<!--#include file="vsadmin/inc/incfunctions.asp"-->
<!--#include file="vsadmin/inc/incemail.asp"-->
<table border='0' cellspacing='4' cellpadding='3' width='100%'>
<% if request.form("posted")="1" then
		if htmlemails=true then emlNl = "<br>" else emlNl=vbCrLf
		Set rs = Server.CreateObject("ADODB.RecordSet")
		Set cnn=Server.CreateObject("ADODB.Connection")
		cnn.open sDSN
		sSQL="SELECT adminEmail,smtpserver,emailUser,emailPass,adminStoreURL,emailObject FROM admin WHERE adminID=1"
		rs.Open sSQL,cnn,0,1
		emailAddr = rs("adminEmail")
		themailhost = Trim(rs("smtpserver")&"")
		theuser = Trim(rs("emailUser")&"")
		thepass = Trim(rs("emailPass")&"")
		adminStoreURL = rs("adminStoreURL")
		if (left(LCase(adminStoreURL),7) <> "http://") AND (left(LCase(adminStoreURL),8) <> "https://") then
			adminStoreURL = "http://" & adminStoreURL
		end if
		if Right(adminStoreURL,1) <> "/" then adminStoreURL = adminStoreURL & "/"
		emailObject = rs("emailObject")
		rs.Close
		seBody = xxEFYF1 & request.form("yourname") & " ("&request.form("youremail")&")" & xxEFYF2
		if Trim(request.form("yourcomments"))<>"" then
			seBody = seBody & xxEFYF3 & emlNl
			seBody = seBody & Trim(request.form("yourcomments")) & emlNl
		else
			seBody = seBody & "." & emlNl
		end if
		if htmlemails=true then
			storeLink = adminStoreURL
			if Trim(Request.Form("id")) <> "" then storeLink = storeLink & "proddetail.asp?prod=" & Trim(Request.Form("id"))
			seBody = seBody & emlNl & "<a href=""" & storeLink & """>" & storeLink & "</a>"
		else
			seBody = seBody & emlNl & adminStoreURL
			if Trim(Request.Form("id")) <> "" then seBody = seBody & "proddetail.asp?prod=" & Trim(Request.Form("id"))
		end if
		seBody = seBody & emlNl
		call DoSendEmailEO(request.form("friendsemail"),emailAddr,request.form("youremail"),request.form("yourname") & xxEFRec,seBody,emailObject,themailhost,theuser,thepass)
		cnn.Close
		set rs = nothing
		set cnn = nothing
%>
	<tr bgcolor="#D8CCE0">
	  <td colspan="2" align="center" width="100%">&nbsp;</td>
  </tr>
	<tr>
	  <td colspan="2" align="center" width="100%"><p>&nbsp;</p>
	  <p><%=xxEFThk%></p>
	  <p><%=xxClkClo%></p>
	  <p>&nbsp;</p>
	  </td>
	</tr>
	<tr>
	  <td colspan="2" align="center" width="100%"><input type="button" name="close" value="<%=xxClsWin%>" onClick="javascript:self.close()">
	  <p>&nbsp;</p>
	  </td>
	</tr>
	<tr bgcolor="#D8CCE0">
	  <td colspan="2" align="center" width="100%">&nbsp;</td>
  </tr>
<% else %>
<script Language="JavaScript">
<!--
function formvalidator(theForm)
{
  if (theForm.yourname.value == "")
  {
    alert("<%=xxPlsEntr%> \"<%=xxEFNam%>\".");
    theForm.yourname.focus();
    return (false);
  }
  if (theForm.youremail.value == "")
  {
    alert("<%=xxPlsEntr%> \"<%=xxEFEm%>\".");
    theForm.youremail.focus();
    return (false);
  }
  if (theForm.friendsemail.value == "")
  {
    alert("<%=xxPlsEntr%> \"<%=xxEFFEm%>\".");
    theForm.friendsemail.focus();
    return (false);
  }
  
  return (true);
}
//-->
</script>
  <form method="POST" action="emailfriend.asp" onSubmit="return formvalidator(this)">
	<input type="hidden" name="posted" value="1">
	<input type="hidden" name="id" value="<%=Request.QueryString("id")%>">
	<tr bgcolor="#D8CCE0">
	  <td colspan="2" align="center" width="100%">&nbsp;</td>
	</tr>
    <tr>
	  <td colspan="2" align="center" width="100%">
		<table border='0' cellspacing='1' cellpadding='1' width='350'>
		  <tr>
			<td width="100%"><%=xxEFBlr%>
			</td>
		  </tr>
		</table>
	  </td>
	</tr>
	<tr>
	  <td width="40%" align="right"><font color="#FF0000">*</font><%=xxEFNam%>:</td><td><input type="text" name="yourname" size="30"></td>
	</tr>
	<tr>
	  <td width="40%" align="right"><font color="#FF0000">*</font><%=xxEFEm%>:</td><td><input type="text" name="youremail" size="30"></td>
	</tr>
	<tr>
	  <td width="40%" align="right"><font color="#FF0000">*</font><%=xxEFFEm%>:</td><td><input type="text" name="friendsemail" size="30"></td>
	</tr>
	<tr>
	  <td colspan="2" align="center" width="100%">
		<table border='0' cellspacing='1' cellpadding='1' width='250'>
		  <tr>
			<td width="100%">&nbsp;<br><%=xxEFCmt%>:<br>
			  <textarea name="yourcomments" cols="30" rows="5"></textarea>
			</td>
		  </tr>
		</table>
	  </td>
	</tr>
	<tr>
	  <td colspan="2" align="center" width="100%"><input type="submit" name="Send" value="<%=xxSend%>">&nbsp;&nbsp;<input type="button" name="close" value="<%=xxClsWin%>" onClick="javascript:self.close()">
	  </td>
	</tr>
	<tr bgcolor="#D8CCE0">
	  <td colspan="2" align="center" width="100%">&nbsp;</td>
	</tr>
  </form>
<% end if %>
</table>
</BODY>
</HTML>
