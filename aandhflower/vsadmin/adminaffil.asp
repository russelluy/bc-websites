<%
Response.Buffer = True
%>
<!--#include file="db_conn_open.asp"-->
<!--#include file="includes.asp"-->
<!--#include file="inc/languageadmin.asp"-->
<!--#include file="inc/incfunctions.asp"-->
<%
if storesessionvalue="" then storesessionvalue="virtualstore"
if Session("loggedon") <> storesessionvalue OR disallowlogin=TRUE then response.redirect "login.asp"
isprinter=false
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html><!-- InstanceBegin template="/Templates/admintemplate.dwt" codeOutsideHTMLIsLocked="false" -->
<head>
<!-- InstanceBeginEditable name="doctitle" -->
<title>Admin Affiliates</title>
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
<!--#include file="inc/incaffil.asp"-->
</div>
<!-- InstanceEndEditable -->


</body>
<!-- InstanceEnd --></html>
