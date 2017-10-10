<%
if storesessionvalue="" then storesessionvalue="virtualstore"
if Session("loggedon") <> storesessionvalue OR disallowlogin=TRUE then response.end
Dim sSQL,rs,alldata,success,cnn,rowcounter,errmsg,index
success=true
Set rs = Server.CreateObject("ADODB.RecordSet")
Set cnn=Server.CreateObject("ADODB.Connection")
cnn.open sDSN
sSQL = ""
alldata=""
if maxloginlevels="" then maxloginlevels=5
if request.form("posted")="1" then
	if request.form("act")="delete" then
		sSQL = "DELETE FROM clientlogin WHERE clientUser='" & request.form("id") & "'"
		cnn.Execute(sSQL)
		response.write "<meta http-equiv=""refresh"" content=""3; url=adminclientlog.asp"">"
	elseif request.form("act")="domodify" then
		sSQL = "UPDATE clientlogin SET clientUser='" & replace(request.form("clientUser"), "'", "''") & "'"
		sSQL = sSQL & "," & "clientPW='" & replace(request.form("clientPW"), "'", "''") & "'"
		sSQL = sSQL & "," & "clientLoginLevel=" & request.form("clientLoginLevel")
		cpd = trim(request.form("clientPercentDiscount"))
		sSQL = sSQL & "," & "clientPercentDiscount=" & IIfVr(IsNumeric(cpd) AND cpd<>"", cpd, 0)
		clientActions=0
		for each objItem in request.form("clientActions")
			clientActions = clientActions + Int(objItem)
		next
		sSQL = sSQL & "," & "clientActions=" & clientActions
		sSQL = sSQL & " WHERE clientUser='" & request.form("id") & "'"
		cnn.Execute(sSQL)
		response.write "<meta http-equiv=""refresh"" content=""3; url=adminclientlog.asp"">"
	elseif request.form("act")="doaddnew" then
		sSQL = "SELECT clientUser FROM clientlogin WHERE clientUser='" & replace(request.form("clientUser"), "'", "''") & "'"
		rs.Open sSQL,cnn,0,1
		if NOT rs.EOF then
			success=false
			errmsg="The login &quot;" & request.form("clientUser") & "&quot; is already in use. Please choose another."
		end if
		rs.Close
		if success then
			sSQL = "INSERT INTO clientlogin (clientUser,clientPW,clientLoginLevel,clientPercentDiscount,clientActions) VALUES ("
			sSQL = sSQL & "'" & replace(request.form("clientUser"), "'", "''") & "'"
			sSQL = sSQL & ",'" & replace(request.form("clientPW"), "'", "''") & "'"
			sSQL = sSQL & "," & request.form("clientLoginLevel")
			cpd = trim(request.form("clientPercentDiscount"))
			sSQL = sSQL & "," & IIfVr(IsNumeric(cpd) AND cpd<>"", cpd, 0)
			clientActions=0
			for each objItem in request.form("clientActions")
				clientActions = clientActions + Int(objItem)
			next
			sSQL = sSQL & "," & clientActions & ")"
			cnn.Execute(sSQL)
			response.write "<meta http-equiv=""refresh"" content=""3; url=adminclientlog.asp"">"
		end if
	end if
end if
%>
<script language="javascript" type="text/javascript">
<!--
function formvalidator(theForm){
if (theForm.clientUser.value == "")
{
alert("<%=yyPlsEntr%> \"<%=yyLiName%>\".");
theForm.clientUser.focus();
return (false);
}
var checkOK = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789_@.-";
var checkStr = theForm.clientUser.value;
var allValid = true;
for (i = 0;  i < checkStr.length;  i++)
{
    ch = checkStr.charAt(i);
    for (j = 0;  j < checkOK.length;  j++)
      if (ch == checkOK.charAt(j))
        break;
    if (j == checkOK.length)
    {
      allValid = false;
      break;
    }
}
if (!allValid)
{
    alert('<%=yyAlpha3%> "<%=yyLiName%>".');
    theForm.clientUser.focus();
    return (false);
}
var checkStr = theForm.clientPW.value;
var allValid = true;
for (i = 0;  i < checkStr.length;  i++)
{
    ch = checkStr.charAt(i);
    for (j = 0;  j < checkOK.length;  j++)
      if (ch == checkOK.charAt(j))
        break;
    if (j == checkOK.length)
    {
      allValid = false;
      break;
    }
}
if (!allValid)
{
    alert("<%=yyOnlyAl%> \"<%=yyPass%>\".");
    theForm.clientPW.focus();
    return (false);
}
if(document.mainform.clientActions.options[3].selected && document.mainform.clientActions.options[4].selected){
    alert("<%=yyWSDsc%>");
    theForm.clientActions.focus();
    return (false);
}
document.mainform.clientPercentDiscount.disabled=false;
return (true);
}
function checkperdisc(){
	document.mainform.clientPercentDiscount.disabled=!document.mainform.clientActions.options[4].selected;
}
//-->
</script>
      <table border="0" cellspacing="0" cellpadding="0" width="100%" bgcolor="" align="center">
<% if request.form("posted")="1" AND request.form("act")="modify" then 
		sSQL = "SELECT clientUser,clientPW,clientLoginLevel,clientActions,clientPercentDiscount FROM clientlogin WHERE clientUser='"&Request.Form("id")&"'"
		rs.Open sSQL,cnn,0,1
%>
        <tr>
		  <form name="mainform" method="post" action="adminclientlog.asp" onsubmit="return formvalidator(this)">
			<td width="100%" align="center">
			<input type="hidden" name="posted" value="1" />
			<input type="hidden" name="act" value="domodify" />
			<input type="hidden" name="id" value="<%=Request.Form("id")%>" />
            <table width="100%" border="0" cellspacing="2" cellpadding="2" bgcolor="">
			  <tr> 
                <td width="100%" colspan="4" align="center"><strong><%=yyLiAdm%></strong><br />&nbsp;</td>
			  </tr>
			  <tr>
				<td align="right"><strong><%=yyLiName%>:</strong></td>
				<td align="left"><input type="text" name="clientUser" size="20" value="<%=Replace(rs("clientUser"),"""","&quot;")%>" /></td>
				<td align="right" rowspan="4" valign="top"><strong><%=yyActns%>:</strong></td>
				<td rowspan="4" align="left" valign="top"><select name="clientActions" size="6" onchange="checkperdisc()" multiple>
				<option value="1"<% if (rs("clientActions") AND 1) = 1 then response.write " selected" %>><%=yyExStat%></option>
				<option value="2"<% if (rs("clientActions") AND 2) = 2 then response.write " selected" %>><%=yyExCoun%></option>
				<option value="4"<% if (rs("clientActions") AND 4) = 4 then response.write " selected" %>><%=yyExShip%></option>
				<option value="8"<% if (rs("clientActions") AND 8) = 8 then response.write " selected" %>><%=yyWholPr%></option>
				<option value="16"<% if (rs("clientActions") AND 16) = 16 then response.write " selected" %>><%=yyPerDis%></option>
				</select></td>
			  </tr>
			  <tr>
				<td align="right"><p><strong><%=yyPass%>:</strong></td>
				<td align="left"><input type="text" name="clientPW" size="20" value="<%=Replace(rs("clientPW"),"""","&quot;")%>" /></td>
			  </tr>
			  <tr>
				<td align="right"><strong><%=yyLiLev%>:</strong></td>
				<td align="left"><select name="clientLoginLevel" size="1">
				<%	for rowcounter=0 to maxloginlevels
						response.write "<option value='"&rowcounter&"'"
						if rowcounter=Int(rs("clientLoginLevel")) then response.write " selected"
						response.write ">&nbsp; "&rowcounter&" </option>"&vbCrLf
					next
				%>
				</select></td>
			  </tr>
			  <tr>
				<td align="right"><strong><%=yyPerDis%>:</strong></td>
				<td align="left"><input type="text" name="clientPercentDiscount" size="10" value="<%=rs("clientPercentDiscount")%>" /></td>
			  </tr>
			  <tr>
                <td width="100%" colspan="4" align="center"><br /><input type="submit" value="<%=yySubmit%>" />&nbsp;<input type="reset" value="<%=yyReset%>" /><br />&nbsp;</td>
			  </tr>
			  <tr> 
                <td width="100%" colspan="4" align="center"><br />
                          <a href="admin.asp"><strong><%=yyAdmHom%></strong></a><br />
                          &nbsp;</td>
			  </tr>
            </table></td>
		  </form>
        </tr>
<script language="javascript" type="text/javascript">
<!--
checkperdisc();
//-->
</script>
<%	rs.Close
elseif request.form("posted")="1" AND request.form("act")="addnew" then %>
        <tr>
		  <form name="mainform" method="post" action="adminclientlog.asp" onsubmit="return formvalidator(this)">
			<td width="100%" align="center">
			<input type="hidden" name="posted" value="1" />
			<input type="hidden" name="act" value="doaddnew" />
            <table width="100%" border="0" cellspacing="2" cellpadding="2" bgcolor="">
			  <tr> 
                <td width="100%" colspan="4" align="center"><strong><%=yyLiAdm%></strong><br />&nbsp;</td>
			  </tr>
			  <tr>
				<td align="right"><strong><%=yyLiName%>:</strong></td>
				<td align="left"><input type="text" name="clientUser" size="20" value="" /></td>
				<td align="right" rowspan="4" valign="top"><strong><%=yyActns%>:</strong></td>
				<td rowspan="4" align="left" valign="top"><select name="clientActions" size="6" onchange="checkperdisc()" multiple>
				<option value="1"><%=yyExStat%></option>
				<option value="2"><%=yyExCoun%></option>
				<option value="4"><%=yyExShip%></option>
				<option value="8"><%=yyWholPr%></option>
				<option value="16"><%=yyPerDis%></option>
				</select></td>
			  </tr>
			  <tr>
				<td align="right"><p><strong><%=yyPass%>:</strong></td>
				<td align="left"><input type="text" name="clientPW" size="20" value="" /></td>
			  </tr>
			  <tr>
				<td align="right"><strong><%=yyLiLev%>:</strong></td>
				<td align="left"><select name="clientLoginLevel" size="1">
				<%	for rowcounter=0 to maxloginlevels
						response.write "<option value='"&rowcounter&"'"
						response.write ">&nbsp; "&rowcounter&" </option>"&vbCrLf
					next
				%>
				</select></td>
			  </tr>
			  <tr>
				<td align="right"><strong><%=yyPerDis%>:</strong></td>
				<td align="left"><input type="text" name="clientPercentDiscount" size="10" value="0" /></td>
			  </tr>
			  <tr>
                <td width="100%" colspan="4" align="center"><br /><input type="submit" value="<%=yySubmit%>" />&nbsp;<input type="reset" value="<%=yyReset%>" /><br />&nbsp;</td>
			  </tr>
			  <tr> 
                <td width="100%" colspan="4" align="center"><br />
                          <a href="admin.asp"><strong><%=yyAdmHom%></strong></a><br />
                          &nbsp;</td>
			  </tr>
            </table></td>
		  </form>
        </tr>
<script language="javascript" type="text/javascript">
<!--
checkperdisc();
//-->
</script>
<% elseif request.form("posted")="1" AND success then %>
        <tr>
          <td width="100%">
			<table width="100%" border="0" cellspacing="0" cellpadding="2" bgcolor="">
			  <tr> 
                <td width="100%" colspan="2" align="center"><br /><strong><%=yyUpdSuc%></strong><br /><br /><%=yyNowFrd%><br /><br />
                        <%=yyNoAuto%> <A href="adminclientlog.asp"><strong><%=yyClkHer%></strong></a>.<br />
                        <br />
				<img src="../images/clearpixel.gif" width="300" height="3" alt="" />
                </td>
			  </tr>
			</table></td>
        </tr>
<% elseif request.form("posted")="1" then %>
        <tr>
          <td width="100%">
			<table width="100%" border="0" cellspacing="0" cellpadding="2" bgcolor="">
			  <tr> 
                <td width="100%" colspan="2" align="center"><br /><font color="#FF0000"><strong><%=yyOpFai%></strong></font><br /><br /><%=errmsg%><br /><br />
				<a href="javascript:history.go(-1)"><strong><%=yyClkBac%></strong></a></td>
			  </tr>
			</table></td>
        </tr>
<% else 
		sSQL = "SELECT clientUser,clientPW,clientLoginLevel,clientActions FROM clientlogin ORDER BY clientUser"
		rs.Open sSQL,cnn,0,1
		if NOT rs.EOF then alldata=rs.getrows
		rs.Close
%>
<script language="javascript" type="text/javascript">
<!--
function modrec(id) {
	document.mainform.id.value = id;
	document.mainform.act.value = "modify";
	document.mainform.submit();
}
function newrec(id) {
	document.mainform.id.value = id;
	document.mainform.act.value = "addnew";
	document.mainform.optType.value = "2";
	document.mainform.submit();
}
function delrec(id) {
cmsg = "<%=yyConDel%>\n"
if (confirm(cmsg)) {
	document.mainform.id.value = id;
	document.mainform.act.value = "delete";
	document.mainform.submit();
}
}
// -->
</script>
        <tr>
		<form name="mainform" method="post" action="adminclientlog.asp">
		  <td width="100%">
			<input type="hidden" name="posted" value="1" />
			<input type="hidden" name="act" value="xxxxx" />
			<input type="hidden" name="id" value="xxxxx" />
			<input type="hidden" name="optType" value="xxxxx" />
            <table width="100%" border="0" cellspacing="0" cellpadding="1" bgcolor="">
			  <tr> 
                <td width="100%" colspan="6" align="center"><strong><%=yyLiAdm%></strong><br />&nbsp;</td>
			  </tr>
			  <tr>
				<td><strong><%=yyLiName%></strong></td>
				<td><strong><%=yyPass%></strong></td>
				<td align="center"><strong><%=yyLiLev%></strong></td>
				<td><strong><%=yyActns%></strong></td>
				<td width="5%" align="center"><strong><%=yyModify%></strong></td>
				<td width="5%" align="center"><strong><%=yyDelete%></strong></td>
			  </tr>
<%	if IsArray(alldata) then
		for rowcounter=0 to UBOUND(alldata,2) %>
			  <tr>
				<td><%=alldata(0,rowcounter)%></td>
				<td><%=alldata(1,rowcounter)%></td>
				<td align="center"><%=alldata(2,rowcounter)%></td>
				<td><%	if (alldata(3,rowcounter) AND 1) = 1 then response.write "STE "
						if (alldata(3,rowcounter) AND 2) = 2 then response.write "CTE "
						if (alldata(3,rowcounter) AND 4) = 4 then response.write "SHE "
						if (alldata(3,rowcounter) AND 8) = 8 then response.write "WSP "
						if (alldata(3,rowcounter) AND 16) = 16 then response.write "PED "
				%>&nbsp;</td>
				<td align="center"><input type=button value="<%=yyModify%>" onclick="modrec('<%=alldata(0,rowcounter)%>')" /></td>
				<td align="center"><input type=button value="<%=yyDelete%>" onclick="delrec('<%=alldata(0,rowcounter)%>')" /></td>
			  </tr>
<%		next
	else
%>
			  <tr>
                <td width="100%" colspan="6" align="center"><br /><%=yyCLNo%><br />&nbsp;</td>
			  </tr>
<%
	end if
%>
			  <tr>
                <td width="100%" colspan="6" align="center"><br /><strong><%=yyPOClk%> </strong>&nbsp;&nbsp;<input type="button" value="<%=yyCLNew%>" onclick="newrec()" /><br />&nbsp;</td>
			  </tr>
			  <tr>
                <td width="100%" colspan="6" align="center"><ul><li><%=yyCLTyp%></li></ul>
                          <a href="admin.asp"><strong><%=yyAdmHom%></strong></a><br />
				<img src="../images/clearpixel.gif" width="300" height="3" alt="" /></td>
			  </tr>
            </table></td>
		  </form>
        </tr>
<% end if
cnn.Close
set rs = nothing
set cnn = nothing
%>
      </table>