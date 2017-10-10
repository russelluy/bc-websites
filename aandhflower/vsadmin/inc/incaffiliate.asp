<%
Dim sSQL,rs,cnn,success,showaccount,addsuccess
addsuccess = true
success = true
showaccount = true
Set rs = Server.CreateObject("ADODB.RecordSet")
Set cnn=Server.CreateObject("ADODB.Connection")
cnn.open sDSN
if request.form("editaction")<>"" then
	sSQL = "SELECT * FROM affiliates WHERE affilID='"&replace(trim(request.form("affilid")),"'","")&"'"
	if request.form("editaction")="modify" then
		rs.Open sSQL,cnn,1,3,&H0001
	elseif request.form("editaction")="new" then
		rs.Open sSQL,cnn,0,1
		if NOT rs.EOF then addsuccess = false
		rs.Close
		if addsuccess then
			rs.Open "affiliates",cnn,1,3,&H0002
			rs.AddNew
			rs.Fields("affilID") = replace(trim(request.form("affilID")),"'","")
			if defaultcommission<>"" then
				rs.Fields("affilCommision") = defaultcommission
				session("affilCommision") = cDbl(defaultcommission)
			else
				session("affilCommision") = 0
			end if
		end if
	end if
	if addsuccess then
		rs.Fields("affilPW")		= replace(trim(request.form("affilPW")),"'","")
		rs.Fields("affilEmail")		= trim(request.form("Email"))
		rs.Fields("affilName")		= trim(request.form("Name"))
		rs.Fields("affilAddress")	= trim(request.form("Address"))
		rs.Fields("affilCity")		= trim(request.form("City"))
		rs.Fields("affilState")		= trim(request.form("State"))
		rs.Fields("affilCountry")	= trim(request.form("Country"))
		rs.Fields("affilZip")		= trim(request.form("Zip"))
		if trim(request.form("Inform"))="ON" then
			rs.Fields("affilInform") = 1
		else
			rs.Fields("affilInform") = 0
		end if
		rs.Update
		rs.Close
		session("affilid") = replace(trim(request.form("affilid")),"'","")
		session("affilpw") = replace(trim(request.form("affilpw")),"'","")
		session("affilName") = trim(request.form("Name"))
		response.write "<meta http-equiv=""Refresh"" content=""0; URL=affiliate.asp"">"
	end if
elseif request.form("affillogin")<>"" then
	sSQL = "SELECT affilID,affilName,affilCommision FROM affiliates WHERE affilID='"&replace(trim(request.form("affilid")),"'","")&"' AND affilPW='"&replace(trim(request.form("affilpw")),"'","")&"'"
	rs.Open sSQL,cnn,0,1
	if NOT rs.EOF then
		session("affilid")=replace(trim(request.form("affilid")),"'","")
		session("affilpw")=replace(trim(request.form("affilpw")),"'","")
		session("affilName") = rs("affilName")
		session("affilCommision") = cDbl(rs("affilCommision"))
		showaccount=false
	else
		success=false
	end if
	rs.Close
	if success then
		response.write "<meta http-equiv=""Refresh"" content=""3; URL=affiliate.asp"">"
%>
	  <table border="0" cellspacing="<%=maintablespacing%>" cellpadding="<%=maintablepadding%>" width="<%=maintablewidth%>" bgcolor="<%=maintablebg%>" align="center">
        <tr> 
          <td width="100%">
		    <form method="post" action="affiliate.asp">
			  <table width="<%=innertablewidth%>" border="0" cellspacing="<%=innertablespacing%>" cellpadding="<%=innertablepadding%>" bgcolor="<%=innertablebg%>">
				<tr>
				  <td width="100%" align="center" colspan="2"><strong><%=xxAffPrg & " " & xxWelcom & " " & session("affilName")%>.</strong></td>
				</tr>
				<tr>
				  <td width="100%" align="center" colspan="2">&nbsp;</td>
				</tr>
				<tr>
				  <td width="100%" align="center" colspan="2"><p><%=xxAffLog%></p>
					<p><%=xxForAut%> <a href="affiliate.asp"><strong><%=xxClkHere%></strong></a>.</p></td>
				</tr>
			  </table>
			</form>
		  </td>
        </tr>
      </table>
<%
	end if
elseif request.form("logout") <> "" then
	session("affilid") = ""
	session("affilpw") = ""
	session("affilName") = ""
end if
if request.form("newaffil")="Go" OR (request.form("editaffil")<>"" AND trim(session("affilid"))<>"") OR NOT addsuccess then
	showaccount=false
%>
<script language="javascript" type="text/javascript">
<!--
function checkform(frm)
{
if(frm.affilid.value=="")
{
	alert("<%=xxPlsEntr%> \"<%=xxAffID%>\".");
	frm.affilid.focus();
	return (false);
}
var checkOK = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789";
var checkStr = frm.affilid.value;
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
    alert("<%=xxAlphaNu%> \"<%=xxAffID%>\".");
    frm.affilid.focus();
    return (false);
}
if(frm.affilpw.value=="")
{
	alert("<%=xxPlsEntr%> \"<%=xxPwd%>\".");
	frm.affilpw.focus();
	return (false);
}
if(frm.name.value=="")
{
	alert("<%=xxPlsEntr%> \"<%=xxName%>\".");
	frm.name.focus();
	return (false);
}
if(frm.email.value=="")
{
	alert("<%=xxPlsEntr%> \"<%=xxEmail%>\".");
	frm.email.focus();
	return (false);
}
if(frm.address.value=="")
{
	alert("<%=xxPlsEntr%> \"<%=xxAddress%>\".");
	frm.address.focus();
	return (false);
}
if(frm.city.value=="")
{
	alert("<%=xxPlsEntr%> \"<%=xxCity%>\".");
	frm.city.focus();
	return (false);
}
if(frm.state.value=="")
{
	alert("<%=xxPlsEntr%> \"<%=xxAllSta%>\".");
	frm.state.focus();
	return (false);
}
if(frm.zip.value=="")
{
	alert("<%=xxPlsEntr%> \"<%=xxZip%>\".");
	frm.zip.focus();
	return (false);
}
return (true);
}
//-->
</script>
<%
	if NOT addsuccess then
		affilName = request.form("Name")
		affilPW = request.form("affilPW")
		affilID = request.form("affilID")
		affilAddress = request.form("Address")
		affilCity = request.form("City")
		affilState = request.form("State")
		affilZip = request.form("Zip")
		affilCountry = request.form("Country")
		affilEmail = request.form("Email")
		affilInform = trim(request.form("Inform"))="ON"
	elseif (request.form("editaffil")<>"" AND trim(session("affilid"))<>"") then
		sSQL = "SELECT affilName,affilPW,affilAddress,affilCity,affilState,affilZip,affilCountry,affilEmail,affilInform FROM affiliates WHERE affilID='"&replace(trim(session("affilid")),"'","")&"' AND affilPW='"&replace(trim(session("affilpw")),"'","")&"'"
		rs.Open sSQL,cnn,1,3,&H0001
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
		end if
		rs.Close
	end if
%>
	  <table border="0" cellspacing="<%=maintablespacing%>" cellpadding="<%=maintablepadding%>" width="<%=maintablewidth%>" bgcolor="<%=maintablebg%>" align="center">
        <tr> 
          <td width="100%">
		    <form method="post" action="affiliate.asp" onsubmit="return checkform(this)">
			  <table width="<%=innertablewidth%>" border="0" cellspacing="<%=innertablespacing%>" cellpadding="<%=innertablepadding%>" bgcolor="<%=innertablebg%>">
				<tr>
				  <td width="100%" align="center" colspan="4"><strong><%=xxAffDts%></strong></td>
				</tr>
<% if NOT addsuccess then %>
				<tr>
				  <td width="100%" align="center" colspan="4"><strong><font color='#FF0000'><%=xxAffUse%></font></strong></td>
				</tr>
<% end if %>
				<tr>
				  <td width="25%" align="right"><strong><font color='#FF0000'>*</font><%=xxAffID%>:</strong></td>
				  <td width="25%" align="left"><%
					if (request.form("editaffil")<>"" AND trim(session("affilid"))<>"") then
						response.write trim(session("affilid"))
						%><input type="hidden" name="affilid" value="<%=trim(session("affilid"))%>" />
						  <input type="hidden" name="editaction" value="modify" /><%
					else
						%><input type="text" name="affilid" size="20" value="<%=affilid%>" />
						  <input type="hidden" name="editaction" value="new" /><%
					end if %></td>
				  <td width="25%" align="right"><strong><font color='#FF0000'>*</font><%=xxPwd%>:</strong></td>
				  <td width="25%" align="left"><input type="password" name="affilpw" size="20" value="<%=affilPW%>" /></td>
				</tr>
				<tr>
				  <td width="25%" align="right"><strong><font color='#FF0000'>*</font><%=xxName%>:</strong></td>
				  <td width="25%" align="left"><input type="text" name="name" size="20" value="<%=affilName%>" /></td>
				  <td width="25%" align="right"><strong><font color='#FF0000'>*</font><%=xxEmail%>:</strong></td>
				  <td width="25%" align="left"><input type="text" name="email" size="25" value="<%=affilEmail%>" /></td>
				</tr>
				<tr>
				  <td width="25%" align="right"><strong><font color='#FF0000'>*</font><%=xxAddress%>:</strong></td>
				  <td width="25%" align="left"><input type="text" name="address" size="20" value="<%=affilAddress%>" /></td>
				  <td width="25%" align="right"><strong><font color='#FF0000'>*</font><%=xxCity%>:</strong></td>
				  <td width="25%" align="left"><input type="text" name="city" size="20" value="<%=affilCity%>" /></td>
				</tr>
				<tr>
				  <td width="25%" align="right"><strong><font color='#FF0000'>*</font><%=xxAllSta%>:</strong></td>
				  <td width="25%" align="left"><input type="text" name="state" size="20" value="<%=affilState%>" /></td>
				  <td width="25%" align="right"><strong><font color='#FF0000'>*</font><%=xxCountry%>:</strong></td>
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
				  <td width="25%" align="right"><strong><font color='#FF0000'>*</font><%=xxZip%>:</strong></td>
				  <td width="25%" align="left"><input type="text" name="zip" size="10" value="<%=affilZip%>" /></td>
				  <td width="25%" align="right"><strong><%=xxInfMe%>:</strong></td>
				  <td width="25%" align="left"><input type="checkbox" name="inform" value="ON" <% if affilInform then response.write "checked"%> /></td>
				</tr>
				<tr>
				  <td width="100%" colspan="4">
					<font size="1"><ul><li><%=xxInform%></li></ul></font>
				  </td>
				</tr>
				<tr>
				  <td width="50%" align="center" colspan="4"><input type="submit" value="<%=xxSubmt%>" /> <input type="reset" value="Reset" /> <%
					if (request.form("editaffil")<>"" AND trim(session("affilid"))<>"") then
					  %><br /><br /><input type="button" value="<%=xxBack%>" onclick="javascript:history.go(-1)" /><%
					end if %></td>
				</tr>
			  </table>
			</form>
		  </td>
        </tr>
      </table>
<%
end if
if showaccount then
	if session("affilid")="" then
%>
	  <table border="0" cellspacing="<%=maintablespacing%>" cellpadding="<%=maintablepadding%>" width="<%=maintablewidth%>" bgcolor="<%=maintablebg%>" align="center">
        <tr> 
          <td width="100%">
			<table width="<%=innertablewidth%>" border="0" cellspacing="<%=innertablespacing%>" cellpadding="<%=innertablepadding%>" bgcolor="<%=innertablebg%>">
			  <form method="post" action="affiliate.asp">
				<tr>
				  <td width="100%" align="center" colspan="2"><strong><%=xxAffPrg%></strong></td>
				</tr>
				<tr>
				  <td width="100%" align="center" colspan="2">&nbsp;</td>
				</tr>
				<tr>
				  <td width="50%" align="right"><%=xxNewAct%>:</td>
				  <td><input type="submit" name="newaffil" value="Go" /></td>
				</tr>
			  </form>
			  <form method="post" action="affiliate.asp">
				<tr>
				  <td width="100%" align="center" colspan="2">&nbsp;</td>
				</tr>
				<tr>
				  <td width="100%" align="center" colspan="2"><strong><%=xxGotAct%></strong></td>
				</tr>
<%		if NOT success then %>
				<tr>
				  <td width="100%" align="center" colspan="2"><font color="#FF0000"><%=xxAffNo%></font></td>
				</tr>
<%		end if %>
				<tr>
				  <td width="50%" align="right"><%=xxAffID%>:</td>
				  <td><input type="text" name="affilid" size="20" value="<%=trim(request.form("affilid"))%>" /></td>
				</tr>
				<tr>
				  <td width="50%" align="right"><%=xxPwd%>:</td>
				  <td><input type="password" name="affilpw" size="20" value="<%=trim(request.form("affilpw"))%>" /></td>
				</tr>
				<tr>
				  <td width="100%" align="center" colspan="2"><input type="submit" name="affillogin" value="<%=xxAffLI%>" /></td>
				</tr>
			  </form>
			</table>
		  </td>
        </tr>
      </table>
<%	else
		totalDay=0.0
		totalYesterday=0.0
		totalMonth=0.0
		totalLastMonth=0.0
		tdt = Date()
		tdt2 = Date()+1
		sSQL = "SELECT Sum(ordTotal-ordDiscount) as theCount FROM orders WHERE ordStatus>=3 AND ordAffiliate='"&trim(replace(session("affilid"),"'",""))&"' AND ordDate BETWEEN "&datedelim&VSUSDate(tdt)&datedelim&" AND "&datedelim&VSUSDate(tdt2)&datedelim
		rs.Open sSQL,cnn,0,1
		if NOT rs.EOF then totalDay = rs("theCount")
		rs.Close
		tdt = Date()-1
		tdt2 = Date()
		sSQL = "SELECT Sum(ordTotal-ordDiscount) as theCount FROM orders WHERE ordStatus>=3 AND ordAffiliate='"&trim(replace(session("affilid"),"'",""))&"' AND ordDate BETWEEN "&datedelim&VSUSDate(tdt)&datedelim&" AND "&datedelim&VSUSDate(tdt2)&datedelim
		rs.Open sSQL,cnn,0,1
		if NOT rs.EOF then totalYesterday = rs("theCount")
		rs.Close
		tdt = DateSerial(year(Date()),Month(Date()),1)
		tdt2 = Date()+1
		sSQL = "SELECT Sum(ordTotal-ordDiscount) as theCount FROM orders WHERE ordStatus>=3 AND ordAffiliate='"&trim(replace(session("affilid"),"'",""))&"' AND ordDate BETWEEN "&datedelim&VSUSDate(tdt)&datedelim&" AND "&datedelim&VSUSDate(tdt2)&datedelim
		rs.Open sSQL,cnn,0,1
		if NOT rs.EOF then totalMonth = rs("theCount")
		rs.Close
		tdt = DateSerial(year(Date()),Month(Date())-1,1)
		tdt2 = DateSerial(year(Date()),Month(Date()),1)
		sSQL = "SELECT Sum(ordTotal-ordDiscount) as theCount FROM orders WHERE ordStatus>=3 AND ordAffiliate='"&trim(replace(session("affilid"),"'",""))&"' AND ordDate BETWEEN "&datedelim&VSUSDate(tdt)&datedelim&" AND "&datedelim&VSUSDate(tdt2)&datedelim
		rs.Open sSQL,cnn,0,1
		if NOT rs.EOF then totalLastMonth = rs("theCount")
		rs.Close
		if IsNull(totalDay) then totalDay=0.0
		if IsNull(totalYesterday) then totalYesterday=0.0
		if IsNull(totalMonth) then totalMonth=0.0
		if IsNull(totalLastMonth) then totalLastMonth=0.0
		alreadygotadmin = getadminsettings()
%>	  <table border="0" cellspacing="<%=maintablespacing%>" cellpadding="<%=maintablepadding%>" width="<%=maintablewidth%>" bgcolor="<%=maintablebg%>" align="center">
        <tr> 
          <td width="100%">
		    <form method="post" action="affiliate.asp">
			  <table width="<%=innertablewidth%>" border="0" cellspacing="<%=innertablespacing%>" cellpadding="<%=innertablepadding%>" bgcolor="<%=innertablebg%>">
				<tr>
				  <td width="100%" align="center" colspan="2"><strong><%=xxAffPrg & " " & xxWelcom & " " & session("affilName")%>.</strong></td>
				</tr>
				<tr>
				  <td width="100%" align="center" colspan="2">&nbsp;</td>
				</tr>
				<tr>
				  <td width="50%" align="right"><strong><%=xxTotTod%>:</strong></td>
				  <td width="50%"><% response.write FormatEuroCurrency(totalDay)
				  if Session("affilCommision")<>0 then response.write " = " & FormatEuroCurrency((totalDay * Session("affilCommision")) / 100.0) & " <strong>" & xxCommis & "</strong>"
				  %></td>
				</tr>
				<tr>
				  <td width="50%" align="right"><strong><%=xxTotYes%>:</strong></td>
				  <td width="50%"><% response.write FormatEuroCurrency(totalYesterday)
				  if Session("affilCommision")<>0 then response.write " = " & FormatEuroCurrency((totalYesterday * Session("affilCommision")) / 100.0) & " <strong>" & xxCommis & "</strong>"%></td>
				</tr>
				<tr>
				  <td width="50%" align="right"><strong><%=xxTotMTD%>:</strong></td>
				  <td width="50%"><% response.write FormatEuroCurrency(totalMonth)
				  if Session("affilCommision")<>0 then response.write " = " & FormatEuroCurrency((totalMonth * Session("affilCommision")) / 100.0) & " <strong>" & xxCommis & "</strong>"%></td>
				</tr>
				<tr>
				  <td width="50%" align="right"><strong><%=xxTotLM%>:</strong></td>
				  <td width="50%"><% response.write FormatEuroCurrency(totalLastMonth)
				  if Session("affilCommision")<>0 then response.write " = " & FormatEuroCurrency((totalLastMonth * Session("affilCommision")) / 100.0) & " <strong>" & xxCommis & "</strong>"%></td>
				</tr>
				<tr>
				  <td width="100%" align="center" colspan="2">&nbsp;</td>
				</tr>
				<tr>
				  <td width="100%" align="center" colspan="2"><input type="submit" name="editaffil" value="<%=xxEdtAff%>" /></td>
				</tr>
				<tr>
				  <td width="100%" align="center" colspan="2">&nbsp;</td>
				</tr>
				<tr>
				  <td width="100%" colspan="2"><font size="1">
				    <ul>
					  <li><%=xxAffLI1%> <strong>products.asp?PARTNER=<%=trim(session("affilid"))%></strong></li>
					  <li><%=xxAffLI2%></li>
					  <% if Session("affilCommision")=0 then %>
					  <li><%=xxAffLI3%></li>
					  <% end if %>
					</ul></font></td>
				</tr>
				<tr>
				  <td width="100%" align="center" colspan="2"><input type="submit" name="logout" value="Logout" /></td>
				</tr>
			  </table>
			</form>
		  </td>
        </tr>
      </table>
<%
	end if
end if
cnn.Close
set rs = nothing
set cnn = nothing
%><br />