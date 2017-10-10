<%
if storesessionvalue="" then storesessionvalue="virtualstore"
if Session("loggedon") <> storesessionvalue OR disallowlogin=TRUE then response.end
Dim sSQL,rs,cnn,success,showaccount,addsuccess,alldata,index,allcountries,rowcounter,sd,ed,errmsg
Dim affilName,affilPW,affilAddress,affilCity,affilState,affilZip,affilCountry,affilEmail,affilInform,smonth
addsuccess = true
success = true
showaccount = true
if dateadjust="" then dateadjust=0
thedate = DateAdd("h",dateadjust,Now())
thedate = DateSerial(year(thedate),month(thedate),day(thedate))
Set rs = Server.CreateObject("ADODB.RecordSet")
Set cnn=Server.CreateObject("ADODB.Connection")
cnn.open sDSN
alreadygotadmin = getadminsettings()
if request.form("editaction")="modify" then
	sSQL = "UPDATE affiliates SET affilPW='"&replace(trim(request.form("affilPW")),"'","''") & "'," & _
		"affilEmail='"&replace(trim(request.form("Email")),"'","''") & "'," & _
		"affilName='"&replace(trim(request.form("Name")),"'","''") & "'," & _
		"affilAddress='"&replace(trim(request.form("Address")),"'","''") & "'," & _
		"affilCity='"&replace(trim(request.form("City")),"'","''") & "'," & _
		"affilState='"&replace(trim(request.form("State")),"'","''") & "'," & _
		"affilCountry='"&replace(trim(request.form("Country")),"'","''") & "'," & _
		"affilZip='"&replace(trim(request.form("Zip")),"'","''") & "',"
		if trim(request.form("affilCommision"))="" then
			sSQL = sSQL & "affilCommision=0,"
		else
			sSQL = sSQL & "affilCommision="&trim(request.form("affilCommision"))&","
		end if
		if trim(request.form("Inform"))="ON" then
			sSQL = sSQL & "affilInform=1 "
		else
			sSQL = sSQL & "affilInform=0 "
		end if
		sSQL = sSQL & "WHERE affilID='" & replace(trim(request.form("affilID")),"'","''") & "'"
		cnn.Execute(sSQL)
elseif request.form("editaction")="delete" then
	sSQL = "DELETE FROM affiliates WHERE affilID='" & replace(trim(request.form("affilID")),"'","''") & "'"
	cnn.Execute(sSQL)
end if
if Trim(request.querystring("id"))<>"" then
	sSQL = "SELECT affilName,affilPW,affilAddress,affilCity,affilState,affilZip,affilCountry,affilEmail,affilInform,affilCommision FROM affiliates WHERE affilID='"&Trim(request.querystring("id"))&"'"
	rs.Open sSQL,cnn,0,1
	if NOT rs.EOF then
		affilName = rs("affilName")
		affilPW = rs("affilPW")
		affilAddress = rs("affilAddress")
		affilCity = rs("affilCity")
		affilState = rs("affilState")
		affilZip = rs("affilZip")
		affilCountry = rs("affilCountry")
		affilEmail = rs("affilEmail")
		affilInform = Int(rs("affilInform"))=1
		affilCommision = rs("affilCommision")
	end if
	rs.Close
%>
<script language="javascript" type="text/javascript">
<!--
function checkform(frm)
{
if(frm.affilid.value==""){
	alert("<%=yyPlsEntr%> \"<%=yyAffId%>\".");
	frm.affilid.focus();
	return (false);
}
var checkOK = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789";
var checkStr = frm.affilid.value;
var allValid = true;
for (i = 0;  i < checkStr.length;  i++){
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
if (!allValid){
    alert("<%=yyOnlyAl%> \"<%=yyAffId%>\".");
    frm.affilid.focus();
    return (false);
}
if(frm.affilpw.value==""){
	alert("<%=yyPlsEntr%> \"<%=yyPass%>\".");
	frm.affilpw.focus();
	return (false);
}
if(frm.name.value==""){
	alert("<%=yyPlsEntr%> \"<%=yyName%>\".");
	frm.name.focus();
	return (false);
}
if(frm.email.value==""){
	alert("<%=yyPlsEntr%> \"<%=yyEmail%>\".");
	frm.email.focus();
	return (false);
}
if(frm.address.value==""){
	alert("<%=yyPlsEntr%> \"<%=yyAddress%>\".");
	frm.address.focus();
	return (false);
}
if(frm.city.value==""){
	alert("<%=yyPlsEntr%> \"<%=yyCity%>\".");
	frm.city.focus();
	return (false);
}
if(frm.state.value==""){
	alert("<%=yyPlsEntr%> \"<%=yyState%>\".");
	frm.state.focus();
	return (false);
}
if(frm.zip.value==""){
	alert("<%=yyPlsEntr%> \"<%=yyZip%>\".");
	frm.zip.focus();
	return (false);
}
var checkOK = "0123456789.";
var checkStr = frm.affilCommision.value;
var allValid = true;
for (i = 0;  i < checkStr.length;  i++){
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
if (!allValid){
    alert("<%=yyOnlyDec%> \"<%=yyCommis%>\".");
    frm.affilCommision.focus();
    return (false);
}
return (true);
}
//-->
</script>
	  <table border="0" cellspacing="0" cellpadding="0" width="100%" bgcolor="" align="center">
        <tr> 
          <td width="100%">
		    <form method="post" action="adminaffil.asp" onsubmit="return checkform(this)">
			  <table width="100%" border="0" cellspacing="0" cellpadding="3" bgcolor="">
				<tr>
				  <td width="100%" align="center" colspan="4"><strong><%=yyAffAdm%></strong></td>
				</tr>
<% if NOT addsuccess then %>
				<tr>
				  <td width="100%" align="center" colspan="4"><strong><font color='#FF0000'><%=yyAffDup%></font></strong></td>
				</tr>
<% end if %>
				<tr>
				  <td width="25%" align="right"><strong><font color='#FF0000'>*</font><%=yyAffId%>:</strong></td>
				  <td width="25%" align="left"><%=Trim(request.querystring("id"))%>
					<input type="hidden" name="affilid" size="20" value="<%=Trim(request.querystring("id"))%>" />
					<input type="hidden" name="editaction" value="modify" /></td>
				  <td width="25%" align="right"><strong><font color='#FF0000'>*</font><%=yyPass%>:</strong></td>
				  <td width="25%" align="left"><input type="text" name="affilpw" size="20" value="<%=affilPW%>" /></td>
				</tr>
				<tr>
				  <td width="25%" align="right"><strong><font color='#FF0000'>*</font><%=yyName%>:</strong></td>
				  <td width="25%" align="left"><input type="text" name="name" size="20" value="<%=affilName%>" /></td>
				  <td width="25%" align="right"><strong><font color='#FF0000'>*</font><%=yyEmail%>:</strong></td>
				  <td width="25%" align="left"><input type="text" name="email" size="25" value="<%=affilEmail%>" /></td>
				</tr>
				<tr>
				  <td width="25%" align="right"><strong><font color='#FF0000'>*</font><%=yyAddress%>:</strong></td>
				  <td width="25%" align="left"><input type="text" name="address" size="20" value="<%=affilAddress%>" /></td>
				  <td width="25%" align="right"><strong><font color='#FF0000'>*</font><%=yyCity%>:</strong></td>
				  <td width="25%" align="left"><input type="text" name="city" size="20" value="<%=affilCity%>" /></td>
				</tr>
				<tr>
				  <td width="25%" align="right"><strong><font color='#FF0000'>*</font><%=yyState%>:</strong></td>
				  <td width="25%" align="left"><input type="text" name="state" size="20" value="<%=affilState%>" /></td>
				  <td width="25%" align="right"><strong><font color='#FF0000'>*</font><%=yyCountry%>:</strong></td>
				  <td width="25%" align="left"><select name="country" size="1">
<%
Sub show_countries(tcountry)
	if NOT IsArray(allcountries) then
		sSQL = "SELECT countryName FROM countries ORDER BY countryOrder DESC, countryName"
		rs.Open sSQL,cnn,0,1
		allcountries=rs.getrows
		rs.Close
	end if
	for rowcounter=0 to UBOUND(allcountries,2)
		response.write "<option value='"&allcountries(0,rowcounter)&"'"
		if tcountry=allcountries(0,rowcounter) then
			response.write " selected"
		end if
		response.write ">"&allcountries(0,rowcounter)&"</option>"&vbCrLf
	next
End Sub
show_countries(affilCountry)
%>
					</select>
				  </td>
				</tr>
				<tr>
				  <td width="25%" align="right"><strong><font color='#FF0000'>*</font><%=yyZip%>:</strong></td>
				  <td width="25%" align="left"><input type="text" name="zip" size="10" value="<%=affilZip%>" /></td>
				  <td width="25%" align="right"><strong>Inform me:</strong></td>
				  <td width="25%" align="left"><input type="checkbox" name="inform" value="ON" <% if affilInform then response.write "checked"%> /></td>
				</tr>
				<tr>
				  <td align="right"><strong><font color='#FF0000'></font><%=yyCommis%>:</strong></td>
				  <% session.LCID = 1033 %>
				  <td colspan="3"><input type="text" name="affilCommision" size="6" value="<%=affilCommision%>" />%</td>
				  <% session.LCID = saveLCID %>
				</tr>
				<tr>
				  <td width="100%" colspan="4">
					<font size="1"><ul><li><%=yyAffInf%></li></ul></font>
				  </td>
				</tr>
				<tr>
				  <td width="50%" align="center" colspan="4"><input type="submit" value="<%=yySubmit%>" /> <input type="reset" value="<%=yyReset%>" /> </td>
				</tr>
			  </table>
			</form>
		  </td>
        </tr>
      </table>
<%
else
	if Request("sd") = "" then
		sd=DateSerial(DatePart("yyyy",thedate),DatePart("m",thedate),1)
	else
		sd=Request("sd")
	end if
	if Request("ed") = "" then
		ed=thedate
	else
		ed=Request("ed")
	end if
	on error resume next
	sd = DateValue(sd)
	ed = Datevalue(ed)
	if err.number <> 0 then
		sd = DateSerial(DatePart("yyyy",thedate),DatePart("m",thedate),1)
		ed = thedate
		success=false
		errmsg=yyDatInv
	end if
	on error goto 0
	tdt = DateValue(sd)
	tdt2 = DateValue(ed)+1
	if mysqlserver=true then
		sSQL = "SELECT affilID,affilName,affilPW,affilEmail,affilCommision,SUM(ordTotal-ordDiscount) FROM affiliates LEFT JOIN orders ON affiliates.affilID=orders.ordAffiliate WHERE ordStatus>=3 AND ordDate BETWEEN "&datedelim&VSUSDate(tdt)&datedelim&" AND "&datedelim&VSUSDate(tdt2)&datedelim&" OR orders.ordAffiliate IS NULL GROUP BY affilID ORDER BY affilID"
	else
		sSQL = "SELECT affilID,affilName,affilPW,affilEmail,affilCommision,(SELECT Sum(ordTotal-ordDiscount) FROM orders WHERE ordStatus>=3 AND ordAffiliate=affilID AND ordDate BETWEEN "&datedelim&VSUSDate(tdt)&datedelim&" AND "&datedelim&VSUSDate(tdt2)&datedelim&") FROM affiliates ORDER BY affilID"
	end if
	rs.Open sSQL,cnn,0,1
	if NOT rs.EOF then alldata=rs.GetRows()
	rs.Close
%>
<script language="javascript" type="text/javascript">
<!--
function delrec(id) {
cmsg = "<%=yyConDel%>\n"
if (confirm(cmsg)) {
	document.mainform.affilID.value = id;
	document.mainform.editaction.value = "delete";
	document.mainform.submit();
}
}
function dumpinventory(){
	document.mainform.action="dumporders.asp";
	document.mainform.act.value = "dumpaffiliate";
	document.mainform.submit();
}
// -->
</script>
	  <table border="0" cellspacing="0" cellpadding="0" width="100%" bgcolor="" align="center">
        <tr> 
          <td width="100%">
			<table width="100%" border="0" cellspacing="0" cellpadding="2" bgcolor="">
			  <tr>
				<td width="100%" align="center" colspan="6"><strong><%=yyAffAdm%></strong><br />&nbsp;</td>
			  </tr>
			  <form method="post" action="adminaffil.asp">
			  <tr> 
                <td width="100%" colspan="6" align="center"><% if NOT success then response.write "<p><font color='#FF0000'>"&errmsg&"</font></p>" %><br /><strong><%=yyAffBet%>:</strong> <input type="text" size="10" name="sd" value="<%=sd%>" /> <strong><%=yyAnd%>:</strong> <input type="text" size="10" name="ed" value="<%=ed%>" /> <input type="submit" value="Go" /><br />&nbsp;</td>
			  </tr>
			  </form>
			  <form method="post" action="adminaffil.asp">
			  <tr> 
                <td width="100%" colspan="6" align="center"><p><strong><%=yyAffFrm%>:</strong> <select name="sd" size="1"><%
					For rowcounter=0 to Day(thedate)-1
						response.write "<option value='"&thedate-rowcounter&"'"
						if thedate-rowcounter=sd then response.write " selected"
						response.write ">"&thedate-rowcounter&"</option>"&vbCrLf
						smonth=thedate-rowcounter
					Next
					For rowcounter=1 to 12
						response.write "<option value='"&DateAdd("m",0-rowcounter,smonth)&"'"
						if DateAdd("m",0-rowcounter,smonth)=sd then response.write " selected"
						response.write ">"&DateAdd("m",0-rowcounter,smonth)&"</option>"&vbCrLf
					Next
				%></select> <strong><%=yyTo%>:</strong> <select name="ed" size="1"><%
					For rowcounter=0 to Day(thedate)-1
						response.write "<option value='"&thedate-rowcounter&"'"
						if thedate-rowcounter=ed then response.write " selected"
						response.write ">"&thedate-rowcounter&"</option>"&vbCrLf
						smonth=thedate-rowcounter
					Next
					For rowcounter=1 to 12
						response.write "<option value='"&DateAdd("m",0-rowcounter,smonth)&"'"
						if DateAdd("m",0-rowcounter,smonth)=ed then response.write " selected"
						response.write ">"&DateAdd("m",0-rowcounter,smonth)&"</option>"&vbCrLf
					Next
				%></select> <input type="submit" value="Go" /><br />&nbsp;</p></td>
			  </tr>
			  </form>
			  <form name="mainform" method="post" action="adminaffil.asp">
				<tr>
				  <td><strong><%=yyAffId%></strong></td>
				  <td><strong><%=yyName%></strong></td>
				  <td><strong><%=yyEmail%></strong></td>
				  <td align="right"><strong><%=yyTotSal%></strong></td>
				  <td align="right"><strong><%=yyCommis%></strong></td>
				  <td align="center"><strong><%=yyDelete%></strong></td>
				</tr>
				<input type="hidden" name="affilID" value="xxx" />
				<input type="hidden" name="editaction" value="xxx" />
				<input type="hidden" name="act" value="xxxxx" />
				<input type="hidden" name="ed" value="<%=DateValue(ed)%>" />
				<input type="hidden" name="sd" value="<%=DateValue(sd)%>" />
<%
	if NOT IsArray(alldata) then
%>
				<tr>
				  <td width="100%" align="center" colspan="6"><br />&nbsp;<br /><strong><%=yyNoAff%></strong><br />&nbsp;</td>
				</tr>
<%
	else
		for index=0 to UBOUND(alldata,2) %>
				<tr>
				  <td><a href="adminaffil.asp?id=<%=alldata(0,index)%>"><strong><%=alldata(0,index)%></strong></a></td>
				  <td><%=alldata(1,index)%></td>
				  <td><a href="mailto:<%=alldata(3,index)%>"><%=alldata(3,index)%></a></td>
				  <td align=right><%if NOT IsNumeric(alldata(5,index)) then response.write "-" else response.write FormatEuroCurrency(alldata(5,index))%></td>
				  <td align=right><%if NOT IsNumeric(alldata(5,index)) OR alldata(4,index)=0 then response.write "-" else response.write FormatEuroCurrency((alldata(4,index)*alldata(5,index)) / 100.0)%></td>
				  <td align="center"><input type=button name=delete value="Delete" onclick="delrec('<%=replace(alldata(0,index),"'","\'")%>')" /></td>
				</tr>
<%
		next %>
				<tr> 
				  <td width="100%" colspan="6" align="center"><input type="button" value="Affiliate Report" onclick="dumpinventory()" /></td>
				</tr>
<%	end if
%>
			  </form>
			</table>
		  </td>
        </tr>
      </table>
<%
end if
cnn.Close
set rs = nothing
set cnn = nothing
%>