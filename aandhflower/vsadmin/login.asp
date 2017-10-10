<!--#include file="db_conn_open.asp"-->
<!--#include file="includes.asp"-->
<!--#include file="inc/languageadmin.asp"-->
<!--#include file="inc/incfunctions.asp"-->
<!--#include file="inc/incemail.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html><!-- InstanceBegin template="/Templates/admintemplate.dwt" codeOutsideHTMLIsLocked="false" -->
<head>
<!-- InstanceBeginEditable name="doctitle" -->
<title>Admin Login</title>
<!-- InstanceEndEditable --><link rel="stylesheet" type="text/css" href="adminstyle.css"/>
<meta http-equiv="Content-Type" content="text/html; charset=<%=adminencoding%>"/>
</head>
<body <% if isprinter then response.write "class=""printbody"""%>>
<% if NOT isprinter then %>

<!-- Header section -->
<div id="header1" align="right"><a class="topbar" href="logout.asp"><%=yyLLLogO%></a>&nbsp;&nbsp;</div>
<div id="header"><img src="adminimages/bclogo.gif" width="278" height="53" alt=""/></div>

<!-- Left menus -->
<div id="left1">
<img src="adminimages/administration.jpg" width="150" height="31" alt=""/><br />
  &nbsp;&middot; <a class="topbar" href="admin.asp"><%=yyLLHome%></a><img src="adminimages/hr.gif" alt=""/><br />
  &nbsp;&middot; <a class="topbar" href="adminmain.asp"><%=yyLLMain%></a><img src="adminimages/hr.gif" alt=""/><br />
  &nbsp;&middot; <a class="topbar" href="adminorders.asp"><%=yyLLOrds%></a><img src="adminimages/hr.gif" alt=""/><br />
  &nbsp;&middot; <a class="topbar" href="adminlogin.asp"><%=yyLLPass%></a><img src="adminimages/hr.gif" alt=""/><br />
  &nbsp;&middot; <a class="topbar" href="adminpayprov.asp"><%=yyLLPayP%></a><img src="adminimages/hr.gif" alt=""/><br />
&nbsp;&middot; <a class="topbar" href="adminaffil.asp"><%=yyLLAffl%></a><img src="adminimages/hr.gif" alt=""/><br />
&nbsp;&middot; <a class="topbar" href="adminclientlog.asp"><%=yyLLClLo%></a><img src="adminimages/hr.gif" alt=""/><br />
&nbsp;&middot; <a class="topbar" href="adminordstatus.asp"><%=yyLLOrSt%></a></div>

<div id="left2">
<img src="adminimages/product_admin.jpg" width="150" height="31" alt=""/><br />
  &nbsp;&middot; <a class="topbar" href="adminprods.asp"><%=yyLLProA%></a><img src="adminimages/hr.gif" alt=""/><br />
  &nbsp;&middot; <a class="topbar" href="adminprodopts.asp"><%=yyLLProO%></a><img src="adminimages/hr.gif" alt=""/><br />
  &nbsp;&middot; <a class="topbar" href="admincats.asp"><%=yyLLCats%></a><img src="adminimages/hr.gif" alt=""/><br />
  &nbsp;&middot; <a class="topbar" href="admindiscounts.asp"><%=yyLLDisc%></a><img src="adminimages/hr.gif" alt=""/><br />
&nbsp;&middot; <a class="topbar" href="adminpricebreak.asp"><%=yyLLQuan%></a></div>

<div id="left3"><img src="adminimages/shipping_admin.jpg" width="150" height="31" alt=""/><br />
  &nbsp;&middot; <a class="topbar" href="adminstate.asp"><%=yyLLStat%></a><img src="adminimages/hr.gif" alt=""/><br />
  &nbsp;&middot; <a class="topbar" href="admincountry.asp"><%=yyLLCoun%></a><img src="adminimages/hr.gif" alt=""/><br />
&nbsp;&middot; <a class="topbar" href="adminzones.asp"><%=yyLLZone%></a><img src="adminimages/hr.gif" alt=""/><br />
&nbsp;&middot; <a class="topbar" href="adminuspsmeths.asp"><%=yyLLShpM%></a></div>

<div id="left4"><img src="adminimages/extras.jpg" width="150" height="31" alt=""/><br />
 </div>

<% end if %>
<!-- main content -->
<!-- InstanceBeginEditable name="Body" -->
<div id="main">
      <table width="100%" border="0" cellpadding="2" cellspacing="2" bgcolor="#FFFFFF">
        <tr> 
          <td>
<%
Dim sSQL,rs,alldata,alladmin,success,cnn,errmsg
success=true
Set rs = Server.CreateObject("ADODB.RecordSet")
Set cnn=Server.CreateObject("ADODB.Connection")
on error resume next
cnn.open sDSN
if err.number <> 0 then
	success=false
	errmsg = "<p><strong>Your database connection needs to be set before you can proceed.</strong></font><br /><br /></p>" &_
					"<p>The current setting is:<br />"&sDSN&"</p>" &_
					"<p>The following information may be helpful</p>" &_
					"<p><strong>Path to this directory<br />"&server.mappath("../")&"</strong></p><p>&nbsp;</p>"
end if
on error goto 0
if success then
	if request.form("posted")="1" then
		alreadygotadmin = getadminsettings()
		sSQL = "SELECT adminEmail, adminStoreURL, adminUser, adminPassword FROM admin WHERE adminID=1"
		rs.Open sSQL,cnn,0,1
		alladmin=rs.getrows
		rs.Close
		if notifyloginattempt=TRUE then
			if htmlemails=true then emlNl = "<br />" else emlNl=vbCrLf
			sMessage = "This is notification of a login attempt at your store."  & emlNl
			sMessage = sMessage & storeurl & emlNL
			if (Trim(request.form("user"))=alladmin(2,0) AND Trim(request.form("pass"))=alladmin(3,0)) then
				sMessage = sMessage & "The correct login / password was used." & emlNl
			else
				sMessage = sMessage & "An incorrect login was used." & emlNl & _
					"Username: " & request.form("user") & emlNl & _
					"Password: " & request.form("pass") & emlNl
			end if
			sMessage = sMessage & "User Agent: " & Request.ServerVariables("HTTP_USER_AGENT") & emlNl & _
				"IP: " & Request.ServerVariables("REMOTE_HOST") & emlNl
			Call DoSendEmailEO(emailAddr,emailAddr,"","Login attempt at your store",sMessage,emailObject,themailhost,theuser,thepass)
		end if
		if NOT (Trim(request.form("user"))=alladmin(2,0) AND Trim(request.form("pass"))=alladmin(3,0)) OR disallowlogin=TRUE then
			success = false
			errmsg = yyLogSor
		else
			if storesessionvalue="" then storesessionvalue="virtualstore"
			Session("loggedon") = storesessionvalue
			response.write "<meta http-equiv=""refresh"" content=""3; url=admin.asp"">"
		end if
	end if
end if
if cnn.State=1 then cnn.Close
set rs = nothing
set cnn = nothing
%>
      <table border="0" cellspacing="0" cellpadding="0" width="100%" bgcolor="" align="center">
<% if request.form("posted")="1" AND success then %>
        <tr>
          <td width="100%">
            <table width="100%" border="0" cellspacing="0" cellpadding="3" bgcolor="">
			  <tr> 
                <td width="100%" colspan="2" align="center"><br /><strong><%=yyLogCor%></strong><br /><br /><%=yyNowFrd%><br /><br />
                        <%=yyNoAuto%><a href="admin.asp"><strong><%=yyClkHer%></strong></a>.<br />
                        <br />
				<p align="center"><img src="../images/clearpixel.gif" width="350" height="1" alt="" /> 
                  </p>
                </td>
			  </tr>
			</table>
		  </td>
        </tr>
<% else %>
        <tr>
		  <td width="100%">
			<form method="post" action="login.asp">
			<input type="hidden" name="posted" value="1" />
            <table width="100%" border="0" cellspacing="0" cellpadding="3" bgcolor="">
			  <tr> 
                <td width="100%" colspan="2" align="center"><br /><strong><%=yyEntUna%><br />&nbsp;</strong></td>
			  </tr>
<% if not success then %>
			  <tr> 
                <td width="100%" colspan="2" align="center"><br /><font color="#FF0000"><%=errmsg%></font></td>
			  </tr>
<% end if %>
              <tr> 
                <td width="50%" align="right"><strong><%=yyUname%>: </strong></td>
				<td width="50%" align="left"><input type="text" name="user" size="20" /></td>
			  </tr>
			  <tr> 
                <td width="50%" align="right"><strong><%=yyPass%>: </strong></td>
				<td width="50%" align="left"><input type="password" name="pass" size="20" /></td>
			  </tr>
			  <tr> 
                <td width="100%" colspan="2" align="center"><br />
                          <input name="submit" type="submit" value="<%=yySubmit%>" /> <br />
				<p align="center"><img src="../images/clearpixel.gif" width="300" height="1" alt="" /> 
                  </p>
                </td>
			  </tr>
            </table>
			</form>
          </td>
        </tr>
<% end if %>
      </table></td>
        </tr>
      </table>
</div>
<!-- InstanceEndEditable -->


</body>
<!-- InstanceEnd --></html>
