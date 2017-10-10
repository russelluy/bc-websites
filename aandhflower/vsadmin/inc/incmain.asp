<%
if storesessionvalue="" then storesessionvalue="virtualstore"
if Session("loggedon") <> storesessionvalue OR disallowlogin=TRUE then response.end
Dim sSQL,rs,alldata,success,cnn,errmsg,index,allcountries
success=true
Set rs = Server.CreateObject("ADODB.RecordSet")
Set cnn=Server.CreateObject("ADODB.Connection")
cnn.open sDSN
if request.form("posted")="1" then
	adminTweaks=0
	for each objItem in request.form("admintweaks")
		adminTweaks = adminTweaks + Int(objItem)
	next
	adminLangSettings=0
	for each objItem in request.form("adminlangsettings")
		adminLangSettings = adminLangSettings + Int(objItem)
	next
	sSQL = "UPDATE admin SET adminEmail='"&request.form("email")&"',adminStoreURL='"&request.form("url")&"',adminProdsPerPage='"&request.form("prodperpage")&"',adminShipping="&request.form("shipping")&",adminIntShipping="&request.form("intshipping")&",adminUSPSUser='"&request.form("USPSUser")&"',adminZipCode='"&request.form("zipcode")&"',adminCountry="&request.form("countrySetting")&",adminDelUncompleted="&request.form("deleteUncompleted")&",adminPacking="&request.form("packing")&",adminStockManage="&request.form("stockManage")&",adminHandling="&request.form("handling")&",adminTweaks="&adminTweaks&",adminDelCC="&request.form("adminDelCC")&",emailObject="&request.form("emailObject")&",smtpserver='"&request.form("smtpserver")&"',emailUser='"&request.form("emailuser")&"',emailPass='"&request.form("emailpass")&"',adminCanPostUser='"&request.form("adminCanPostUser")&"',"
	if request.form("emailconfirm")="ON" then
		sSQL = sSQL & "adminEmailConfirm=1, "
	else
		sSQL = sSQL & "adminEmailConfirm=0, "
	end if
	sSQL = sSQL & "adminUnits=" & (int(request.form("adminUnits")) + int(request.form("adminDims")))
	for index=1 to 3
		if NOT IsNumeric(Trim(request.form("currRate" & index))) then
			sSQL = sSQL & ",currRate" & index & "=0"
		else
			sSQL = sSQL & ",currRate" & index & "=" & request.form("currRate" & index)
		end if
		sSQL = sSQL & ",currSymbol" & index & "='" & Replace(request.form("currSymbol" & index), "'", "''") & "'"
	next
	sSQL = sSQL & ",currLastUpdate=" & datedelim & VSUSDateTime(Now()-10) & datedelim
	sSQL = sSQL & ",currConvUser='" & request.form("currConvUser") & "'"
	sSQL = sSQL & ",currConvPw='" & request.form("currConvPw") & "'"
	sSQL = sSQL & ",adminlanguages=" & request.form("adminlanguages")
	sSQL = sSQL & ",adminlangsettings=" & adminLangSettings & " WHERE adminID=1"
	cnn.Execute(sSQL)
	Application.Lock()
	Application("getadminsettings")=""
	Application.UnLock()
	response.write "<meta http-equiv=""refresh"" content=""2; url=admin.asp"">"
else
	sSQL = "SELECT countryID,countryName,countryCurrency FROM countries WHERE countryLCID<>0 ORDER BY countryOrder DESC, countryName"
	rs.Open sSQL,cnn,0,1
	allcountries=rs.getrows
	rs.Close
	sSQL = "SELECT DISTINCT countryCurrency FROM countries ORDER BY countryCurrency"
	rs.Open sSQL,cnn,0,1
	allcurrencies=rs.getrows
	rs.Close
end if
%>
<script language="javascript" type="text/javascript">
<!--
function formvalidator(theForm)
{
  if(theForm.prodperpage.value == "")
  {
    alert("<%=yyPlsEntr%> \"<%=yyPPP%>\".");
    theForm.prodperpage.focus();
    return (false);
  }
  var checkOK = "0123456789";
  var checkStr = theForm.prodperpage.value;
  var allValid = true;
  for (i = 0;  i < checkStr.length;  i++)
  {
	ch = checkStr.charAt(i);
	for (j = 0;  j < checkOK.length;  j++)
		if (ch == checkOK.charAt(j))
			break;
	if (j == checkOK.length){
		allValid = false;
			break;
	}
  }
  if (!allValid)
  {
	alert("<%=yyOnlyNum%> \"<%=yyPPP%>\".");
	theForm.prodperpage.focus();
	return (false);
  }
for(index=1;index<=3;index++){
  var checkOK = "0123456789.";
  var thisRate = eval("theForm.currRate" + index);
  var checkStr = thisRate.value;
  var allValid = true;
  for (i = 0;  i < checkStr.length;  i++)
  {
	ch = checkStr.charAt(i);
	for (j = 0;  j < checkOK.length;  j++)
		if (ch == checkOK.charAt(j))
			break;
	if (j == checkOK.length){
		allValid = false;
			break;
	}
  }
  if (!allValid)
  {
	alert("<%=yyOnlyDec%> \"<%=yyConRat%> " + index + "\".");
	thisRate.focus();
	return (false);
  }
}

  if(theForm.handling.value == "")
  {
    alert('<%=yyPlsEntr%> "<%=yyHanChg%>". <%=yyNoHan%>');
    theForm.handling.focus();
    return (false);
  }
  var checkOK = "0123456789.";
  var checkStr = theForm.handling.value;
  var allValid = true;
  for (i = 0;  i < checkStr.length;  i++)
  {
	ch = checkStr.charAt(i);
	for (j = 0;  j < checkOK.length;  j++)
		if (ch == checkOK.charAt(j))
			break;
	if (j == checkOK.length){
		allValid = false;
			break;
	}
  }
  if (!allValid)
  {
	alert("<%=yyOnlyDec%> \"<%=yyHanChg%>\".");
	theForm.handling.focus();
	return (false);
  }
  return (true);
}
//-->
</script>
      <table border="0" cellspacing="0" cellpadding="0" width="100%" bgcolor="" align="center">
<% if request.form("posted")="1" AND success then %>
        <tr>
          <td width="100%">
            <table width="100%" border="0" cellspacing="0" cellpadding="3" bgcolor="">
			  <tr> 
                <td width="100%" colspan="2" align="center"><br /><strong><%=yyUpdSuc%></strong><br /><br /><%=yyNowFrd%><br /><br />
                        <%=yyNoAuto%><A href="admin.asp"><strong><%=yyClkHer%></strong></a>.<br />
                        <br />
				<img src="../images/clearpixel.gif" width="300" height="1" alt="" /></td>
			  </tr>
			</table></td>
        </tr>
<%
else
	sSQL = "SELECT adminEmail,adminStoreURL,adminProdsPerPage,adminShipping,adminIntShipping,adminUSPSUser,adminZipCode,adminEmailConfirm,adminCountry,adminUnits,adminDelUncompleted,adminPacking,adminStockManage,adminHandling,adminTweaks,adminDelCC,currRate1,currSymbol1,currRate2,currSymbol2,currRate3,currSymbol3,currConvUser,currConvPw,emailObject,smtpserver,emailUser,emailPass,adminCanPostUser,adminlanguages,adminlangsettings FROM admin WHERE adminID=1"
	rs.Open sSQL,cnn,0,1
%>
        <tr>
		<form method="post" action="adminmain.asp" onsubmit="return formvalidator(this)">
		  <td width="100%">
			<input type="hidden" name="posted" value="1" />
            <table width="100%" border="0" cellspacing="0" cellpadding="3" bgcolor="">
			  <tr> 
                <td width="100%" colspan="2" align="center"><br /><strong><%=yyUpdAdm%></strong><br />&nbsp;</td>
			  </tr>
<%	if not success then %>
			  <tr> 
                <td width="100%" colspan="2" align="center"><br /><font color="#FF0000"><%=errmsg%></font></td>
			  </tr>
<%	end if %>
			  <tr>
				<td width="100%" align="center" colspan="2"><%=yyCsSym%></td>
			  </tr>
			  <tr>
				<td width="50%" align="right"><strong><%=yyCouSet%>: </strong></td>
				<td width="50%" align="left"><select name="countrySetting" size="1">
				  <%
					for index=0 to UBOUND(allcountries,2)
						response.write "<option value='"&allcountries(0,index)&"'"
						if rs("adminCountry")=allcountries(0,index) then response.write " selected"
						response.write ">"&allcountries(1,index)&"</option>"&vbCrLf
					next
				  %>
				  </select></td>
			  </tr>
			  <tr>
				<td width="100%" align="center" colspan="2"><hr width="70%" /><%=yy3CurCon%><br />
				  <font size="1"><%=yyNo3Con%></font></td>
			  </tr>
			  <tr>
				<td width="50%" align="right"><strong><%=yyConv%> 1: </strong></td>
				<td width="50%" align="left">&nbsp;<%=yyRate%> <input type="text" name="currRate1" size="10" value="<% if rs("currRate1")<>0 then response.write rs("currRate1")%>" />&nbsp;&nbsp;&nbsp;Symbol <select name="currSymbol1" size="1"><option value="">None</option>
				  <%
					for index=0 to UBOUND(allcurrencies,2)
						response.write "<option value='"&allcurrencies(0,index)&"'"
						if rs("currSymbol1")=allcurrencies(0,index) then response.write " selected"
						response.write ">"&allcurrencies(0,index)&"</option>"&vbCrLf
					next
				  %></select></td>
			  </tr>
			  <tr>
				<td width="50%" align="right"><strong><%=yyConv%> 2: </strong></td>
				<td width="50%" align="left">&nbsp;<%=yyRate%> <input type="text" name="currRate2" size="10" value="<% if rs("currRate2")<>0 then response.write rs("currRate2")%>" />&nbsp;&nbsp;&nbsp;Symbol <select name="currSymbol2" size="1"><option value="">None</option>
				  <%
					for index=0 to UBOUND(allcurrencies,2)
						response.write "<option value='"&allcurrencies(0,index)&"'"
						if rs("currSymbol2")=allcurrencies(0,index) then response.write " selected"
						response.write ">"&allcurrencies(0,index)&"</option>"&vbCrLf
					next
				  %></select></td>
			  </tr>
			  <tr>
				<td width="50%" align="right"><strong><%=yyConv%> 3: </strong></td>
				<td width="50%" align="left">&nbsp;<%=yyRate%> <input type="text" name="currRate3" size="10" value="<% if rs("currRate3")<>0 then response.write rs("currRate3")%>" />&nbsp;&nbsp;&nbsp;Symbol <select name="currSymbol3" size="1"><option value="">None</option>
				  <%
					for index=0 to UBOUND(allcurrencies,2)
						response.write "<option value='"&allcurrencies(0,index)&"'"
						if rs("currSymbol3")=allcurrencies(0,index) then response.write " selected"
						response.write ">"&allcurrencies(0,index)&"</option>"&vbCrLf
					next
				  %></select></td>
			  </tr>
			  <tr>
				<td width="100%" align="center" colspan="2"><font size="1"><%=yyAutoLogin%></font></td>
			  </tr>
			  <tr>
				<td width="50%" align="right"><strong><%=yyUname%>: </strong></td>
				<td width="50%" align="left"><input type="text" name="currConvUser" size="15" value="<%=rs("currConvUser")%>" /></td>
			  </tr>
			  <tr>
				<td width="50%" align="right"><strong><%=yyPass%>: </strong></td>
				<td width="50%" align="left"><input type="text" name="currConvPw" size="15" value="<%=rs("currConvPw")%>" /></td>
			  </tr>
			  <tr>
				<td width="100%" align="center" colspan="2"><hr width="70%" /><%=yyLikeCE%></td>
			  </tr>
			  <tr>
				<td width="50%" align="right"><strong><%=yyConEm%>: </strong></td>
				<td width="50%" align="left"><input type="checkbox" name="emailconfirm" value="ON" <%if Int(rs("adminEmailConfirm"))=1 then response.write "checked"%> /></td>
			  </tr>
			  <tr>
				<td width="100%" align="center" colspan="2"><hr width="70%" /><%=yyCEAddr%></td>
			  </tr>
			  <tr>
				<td width="50%" align="right"><strong><%=yyEmail%>: </strong></td>
				<td width="50%" align="left"><input type="text" name="email" size="30" value="<%=rs("adminEmail")%>" /></td>
			  </tr>
			  <tr>
				<td width="100%" align="center" colspan="2"><hr width="70%" /><%=yyEmObjs%><br />
				  <font size="1"><%=yyEmCDO%> 
				  <%=yyEmMoInf%> <a href="http://www.beancastle.com" target="_blank"><strong><%=yyHere%></strong></a>. <%=yyEmGen%> <a href="http://www.beancastle.com" target="_blank"><strong><%=yyHere%></strong></a>.</font></td>
			  </tr>
			  <tr>
				<td width="50%" align="right"><strong><%=yyEmailObj%>: </strong></td>
				<td width="50%" align="left"><select name="emailObject" size="1"><option value="99"><%=yyNone%></option>"<%
					gotobject=false
					function checkemail(objnum)
						if objnum=rs("emailObject") then
							checkemail = " selected"
							gotobject=true
						else
							checkemail=""
						end if
					end function
					on error resume next
					err.number=0
					Set EmailObj = Server.CreateObject("CDONTS.NewMail")
					if err.number = 0 then response.write "<option value=""0"""&checkemail(0)&">CDONTS</option>"
					Set EmailObj = nothing
					err.number=0
					Set EmailObj = Server.CreateObject("CDO.Message")
					if err.number = 0 then response.write "<option value=""1"""&checkemail(1)&">CDO</option>"
					Set EmailObj = nothing
					err.number=0
					Set EmailObj = Server.CreateObject("Persits.MailSender")
					if err.number = 0 then response.write "<option value=""2"""&checkemail(2)&">ASP Email (PERSITS)</option>"
					Set EmailObj = nothing
					err.number=0
					Set EmailObj = Server.CreateObject("SMTPsvg.Mailer")
					if err.number = 0 then response.write "<option value=""3"""&checkemail(3)&">ASP Mail (ServerObjects)</option>"
					Set EmailObj = nothing
					err.number=0
					Set EmailObj = Server.CreateObject("JMail.SMTPMail")
					if err.number = 0 then response.write "<option value=""4"""&checkemail(4)&">JMail SMTPMail (Dimac)</option>"
					Set EmailObj = nothing
					err.number=0
					Set EmailObj = Server.CreateObject("SoftArtisans.SMTPMail")
					if err.number = 0 then response.write "<option value=""5"""&checkemail(5)&">SMTPMail (SoftArtisans)</option>"
					Set EmailObj = nothing
					err.number=0
					Set EmailObj = Server.CreateObject("JMail.Message")
					if err.number = 0 then response.write "<option value=""6"""&checkemail(6)&">JMail (Dimac)</option>"
					Set EmailObj = nothing
					on error goto 0
					%></select></td>
			  </tr>
			  <tr>
				<td width="100%" align="center" colspan="2"><hr width="70%" />
				  <font size="1"><%=yySMTPEn%></font></td>
			  </tr>
			  <tr>
				<td width="50%" align="right"><strong><%=yySMTPSe%>: </strong></td>
				<td width="50%" align="left"><input type="text" name="smtpserver" size="30" value="<%=rs("smtpserver")%>" /></td>
			  </tr>
			  <tr>
				<td width="100%" align="center" colspan="2"><hr width="70%" />
				  <font size="1"><%=yySMTPSt%></font></td>
			  </tr>
			  <tr>
				<td width="50%" align="right"><strong><%=yyUname%>: </strong></td>
				<td width="50%" align="left"><input type="text" name="emailuser" size="15" value="<%=rs("emailUser")%>" /></td>
			  </tr>
			  <tr>
				<td width="50%" align="right"><strong><%=yyPass%>: </strong></td>
				<td width="50%" align="left"><input type="text" name="emailpass" size="15" value="<%=rs("emailPass")%>" /></td>
			  </tr>
			  <tr>
				<td width="100%" align="center" colspan="2"><hr width="70%" /><%=yyURLEx & " " & yyExample%><br /> 
						<strong><%
						guessURL ="http://" & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("PATH_INFO")
						wherevs = InStr(guessURL,"vsadmin")
						if wherevs > 0 then
							guessURL = Left(guessURL,wherevs-1)
						else
							guessURL = "http://www.myurl.com/mystore/"
						end if
						response.write guessURL
						%></strong></td>
			  </tr>
			  <tr>
				<td width="50%" align="right"><strong><%=yyStoreURL%>: </strong></td>
				<td width="50%" align="left"><input type="text" name="url" size="35" value="<%=rs("adminStoreURL")%>" /></td>
			  </tr>
			  <tr>
				<td width="100%" align="center" colspan="2"><hr width="70%" /><%=yyHMPPP%></td>
			  </tr>
			  <tr>
				<td width="50%" align="right"><strong><%=yyPPP%>: </strong></td>
				<td width="50%" align="left"><input type="text" name="prodperpage" size="10" value="<%=rs("adminProdsPerPage")%>" /></td>
			  </tr>
			  <tr>
				<td width="100%" align="center" colspan="2"><hr width="70%" /><%=yyHandEx%></td>
			  </tr>
			  <tr>
				<td width="50%" align="right"><strong><%=yyHanChg%>: </strong></td>
				<td width="50%" align="left"><input type="text" name="handling" size="10" value="<%=rs("adminHandling")%>" /></td>
			  </tr>
			  <tr>
				<td width="100%" align="center" colspan="2"><hr width="70%" /><%=yySelShp%></td>
			  </tr>
			  <tr>
				<td width="50%" align="right"><strong><%=yyShpTyp%>: </strong></td>
				<td width="50%" align="left"><select name="shipping" size="1">
					<option value="0"><%=yyNoShp%></option>
					<option value="1" <%if Int(rs("adminShipping"))=1 then response.write "selected"%>><%=yyFlatShp%></option>
					<option value="2" <%if Int(rs("adminShipping"))=2 then response.write "selected"%>><%=yyWghtShp%></option>
					<option value="5" <%if Int(rs("adminShipping"))=5 then response.write "selected"%>><%=yyPriShp%></option>
					<option value="3" <%if Int(rs("adminShipping"))=3 then response.write "selected"%>><%=yyUSPS%></option>
					<option value="4" <%if Int(rs("adminShipping"))=4 then response.write "selected"%>><%=yyUPS%></option>
					<option value="6" <%if Int(rs("adminShipping"))=6 then response.write "selected"%>><%=yyCanPos%></option>
					<option value="7" <%if Int(rs("adminShipping"))=7 then response.write "selected"%>><%=yyFedex%></option>
					</select></td>
			  </tr>
			  <tr>
				<td width="100%" align="center" colspan="2"><hr width="70%" /><%=yySelShI%></td>
			  </tr>
			  <tr>
				<td width="50%" align="right"><strong><%=yyShpTyp%>: </strong></td>
				<td width="50%" align="left"><select name="intshipping" size="1">
					<option value="0"><%=yySamDom%></option>
					<option value="1" <%if Int(rs("adminIntShipping"))=1 then response.write "selected"%>><%=yyFlatShp%></option>
					<option value="2" <%if Int(rs("adminIntShipping"))=2 then response.write "selected"%>><%=yyWghtShp%></option>
					<option value="5" <%if Int(rs("adminIntShipping"))=5 then response.write "selected"%>><%=yyPriShp%></option>
					<option value="3" <%if Int(rs("adminIntShipping"))=3 then response.write "selected"%>><%=yyUSPS%></option>
					<option value="4" <%if Int(rs("adminIntShipping"))=4 then response.write "selected"%>><%=yyUPS%></option>
					<option value="6" <%if Int(rs("adminIntShipping"))=6 then response.write "selected"%>><%=yyCanPos%></option>
					<option value="7" <%if Int(rs("adminIntShipping"))=7 then response.write "selected"%>><%=yyFedex%></option>
					</select></td>
			  </tr>
			  <tr>
				<td width="100%" align="center" colspan="2"><hr width="70%" /><%=yyHowPck%><br /><font size="1"><%=yyOnlyAf%></font></td>
			  </tr>
			  <tr>
				<td width="50%" align="right"><strong><%=yyPackPr%>: </strong></td>
				<td width="50%" align="left"><select name="packing" size="1">
					<option value="0"><%=yyPckSep%></option>
					<option value="1" <%if Int(rs("adminPacking"))=1 then response.write "selected"%>><%=yyPckTog%></option>
					</select></td>
			  </tr>
			  <tr>
				<td width="100%" align="center" colspan="2"><hr width="70%" /><%=yyIfUSPS%><br />
				<font size="1"><%=yyUPSForm%> <a href="adminupslicense.asp"><%=yyHere%></a>.</font></td>
			  </tr>
			  <tr>
				<td width="50%" align="right"><strong><%=yyUname%>: </strong></td>
				<td width="50%" align="left"><input type="text" size="15" name="USPSUser" value="<%=rs("adminUSPSUser")%>" /></td>
			  </tr>
			  <tr>
				<td width="100%" align="center" colspan="2"><hr width="70%" /><%=yyEnMerI%></font></td>
			  </tr>
			  <tr>
				<td colspan="2" align="center"><strong><%=yyRetID%>: </strong><input type="text" size="36" name="adminCanPostUser" value="<%=rs("adminCanPostUser")%>" /></td>
			  </tr>
			  <tr>
				<td width="100%" align="center" colspan="2"><hr width="70%" /><%=yyEntZip%></td>
			  </tr>
			  <tr>
				<td width="50%" align="right"><strong><%=yyZip%>: </strong></td>
				<td width="50%" align="left"><input type="text" name="zipcode" size="10" value="<%=rs("adminZipCode")%>" /></td>
			  </tr>
			  <tr>
				<td width="100%" align="center" colspan="2"><hr width="70%" /><%=yyUPSUnt%></td>
			  </tr>
			  <tr>
				<td width="50%" align="center"><strong><%=yyShpUnt%>: </strong>
				  <select name="adminUnits" size="1">
					<option value="1" <%if (int(rs("adminUnits")) AND 3)=1 then response.write "selected"%>>LBS</option>
					<option value="0" <%if (int(rs("adminUnits")) AND 3)=0 then response.write "selected"%>>KGS</option>
				  </select></td>
				<td width="50%" align="center"><strong><%=yyDims%>: </strong>
				  <select name="adminDims" size="1">
					<option value="0"><%=yyNotSpe%></option>
					<option value="4" <%if (int(rs("adminUnits")) AND 12)=4 then response.write "selected"%>>IN</option>
					<option value="8" <%if (int(rs("adminUnits")) AND 12)=8 then response.write "selected"%>>CM</option>
				  </select></td>
			  </tr>
			  <tr>
				<td width="100%" align="left" colspan="2"><ul>
				  <li><font size="1"><font color="#FF0000">*</font><%=yyUntNote%></font></li>
				  <li><font size="1"><font color="#FF0000">*</font><%=yyUntNo2%></font></li></ul></td>
			  </tr>
			  <tr>
				<td width="100%" align="center" colspan="2"><hr width="70%" /><%=yyStkMgt%><br />
					<font size="1"><%=yyTimUnv%></font></td>
			  </tr>
			  <tr>
				<td width="50%" align="right"><strong><%=yyConUnv%>: </strong></td>
				<td width="50%" align="left"><select name="stockManage" size="1">
					<option value="0"><%=yyNoStk%></option>
					<option value="1" <% if Int(rs("adminStockManage"))=1 then response.write "selected"%>>1 <%=yyHours%></option>
					<option value="2" <% if Int(rs("adminStockManage"))=2 then response.write "selected"%>>2 <%=yyHours%></option>
					<option value="3" <% if Int(rs("adminStockManage"))=3 then response.write "selected"%>>3 <%=yyHours%></option>
					<option value="4" <% if Int(rs("adminStockManage"))=4 then response.write "selected"%>>4 <%=yyHours%></option>
					<option value="6" <% if Int(rs("adminStockManage"))=6 then response.write "selected"%>>6 <%=yyHours%></option>
					<option value="8" <% if Int(rs("adminStockManage"))=8 then response.write "selected"%>>8 <%=yyHours%></option>
					<option value="12" <% if Int(rs("adminStockManage"))=12 then response.write "selected"%>>12 <%=yyHours%></option>
					</select>
				</td>
			  </tr>
			  <tr>
				<td width="100%" align="center" colspan="2"><hr width="70%" /><%=yyHowLan%></td>
			  </tr>
			  <tr>
				<td width="50%" align="right"><strong><%=yyNumLan%>: </strong></td>
				<td width="50%" align="left"><select name="adminlanguages" size="1">
					<option value="0">1</option>
					<option value="1" <% if Int(rs("adminlanguages"))=1 then response.write "selected"%>>2</option>
					<option value="2" <% if Int(rs("adminlanguages"))=2 then response.write "selected"%>>3</option>
					</select>
				</td>
			  </tr>
			  <tr>
				<td width="100%" align="center" colspan="2"><hr width="70%" /><%=yyWhMull%><br />
					<font size="1"><%=yyLonrel%></font></td>
			  </tr>
			  <tr>
				<td width="50%" align="right"><strong><%=yyLaSet%>: </strong></td>
				<td width="50%" align="left"><select name="adminlangsettings" size="5" multiple>
					<option value="1" <% if (Int(rs("adminlangsettings")) AND 1)=1 then response.write "selected"%>><%=yyPrName%></option>
					<option value="2" <% if (Int(rs("adminlangsettings")) AND 2)=2 then response.write "selected"%>><%=yyDesc%></option>
					<option value="4" <% if (Int(rs("adminlangsettings")) AND 4)=4 then response.write "selected"%>><%=yyLnDesc%></option>
					<option value="8" <% if (Int(rs("adminlangsettings")) AND 8)=8 then response.write "selected"%>><%=yyCntNam%></option>
					<option value="16" <% if (Int(rs("adminlangsettings")) AND 16)=16 then response.write "selected"%>><%=yyPOName%></option>
					<option value="32" <% if (Int(rs("adminlangsettings")) AND 32)=32 then response.write "selected"%>><%=yyPOChoi%></option>
					<option value="64" <% if (Int(rs("adminlangsettings")) AND 64)=64 then response.write "selected"%>><%=yyOrdSta%></option>
					<option value="128" <% if (Int(rs("adminlangsettings")) AND 128)=128 then response.write "selected"%>><%=yyPayMet%></option>
					<option value="256" <% if (Int(rs("adminlangsettings")) AND 256)=256 then response.write "selected"%>><%=yyCatNam%></option>
					<option value="512" <% if (Int(rs("adminlangsettings")) AND 512)=512 then response.write "selected"%>><%=yyCatDes%></option>
					<option value="1024" <% if (Int(rs("adminlangsettings")) AND 1024)=1024 then response.write "selected"%>><%=yyDisTxt%></option>
					</select>
					</td>
			  </tr>
			  <tr>
				<td width="100%" align="center" colspan="2"><hr width="70%" /><%=yyDelUnc%></td>
			  </tr>
			  <tr>
				<td width="50%" align="right"><strong><%=yyDelAft%>: </strong></td>
				<td width="50%" align="left"><select name="deleteUncompleted" size="1">
					<option value="0"><%=yyNever%></option>
					<option value="1" <% if Int(rs("adminDelUncompleted"))=1 then response.write "selected"%>>1 <%=yyDay%></option>
					<option value="2" <% if Int(rs("adminDelUncompleted"))=2 then response.write "selected"%>>2 <%=yyDays%></option>
					<option value="3" <% if Int(rs("adminDelUncompleted"))=3 then response.write "selected"%>>3 <%=yyDays%></option>
					<option value="4" <% if Int(rs("adminDelUncompleted"))=4 then response.write "selected"%>>4 <%=yyDays%></option>
					<option value="7" <% if Int(rs("adminDelUncompleted"))=7 then response.write "selected"%>>1 <%=yyWeek%></option>
					<option value="14" <% if Int(rs("adminDelUncompleted"))=14 then response.write "selected"%>>2 <%=yyWeeks%></option>
					</select>
					</td>
			  </tr>
			  <tr>
				<td width="100%" align="center" colspan="2"><hr width="70%" /><%=yyDelCC%></td>
			  </tr>
			  <tr>
				<td width="50%" align="right"><strong><%=yyDelAft%>: </strong></td>
				<td width="50%" align="left"><select name="adminDelCC" size="1">
					<option value="0"><%=yyNever%></option>
					<option value="1" <% if Int(rs("adminDelCC"))=1 then response.write "selected"%>>1 <%=yyDay%></option>
					<option value="2" <% if Int(rs("adminDelCC"))=2 then response.write "selected"%>>2 <%=yyDays%></option>
					<option value="3" <% if Int(rs("adminDelCC"))=3 then response.write "selected"%>>3 <%=yyDays%></option>
					<option value="4" <% if Int(rs("adminDelCC"))=4 then response.write "selected"%>>4 <%=yyDays%></option>
					<option value="7" <% if Int(rs("adminDelCC"))=7 then response.write "selected"%>>1 <%=yyWeek%></option>
					<option value="14" <% if Int(rs("adminDelCC"))=14 then response.write "selected"%>>2 <%=yyWeeks%></option>
					</select>
					</td>
			  </tr>
			  <tr>
				<td width="100%" align="center" colspan="2"><hr width="70%" /><%=yyAdmTwk%><br /><font size="1"><%=yyMulSel%></font></td>
			  </tr>
			  <tr>
				<td width="50%" align="right"><strong><%=yyApTwk%>: </strong></td>
				<td width="50%" align="left"><select name="admintweaks" size="3" multiple>
					<option value="1" <% if (Int(rs("adminTweaks")) AND 1)=1 then response.write "selected"%>><%=yySmpCnt%></option>
					<option value="2" <% if (Int(rs("adminTweaks")) AND 2)=2 then response.write "selected"%>><%=yySmpOpt%></option>
					<option value="4" <% if (Int(rs("adminTweaks")) AND 4)=4 then response.write "selected"%>><%=yySmpSec%></option>
					</select>
					</td>
			  </tr>
			  <tr> 
                <td width="100%" colspan="2" align="center"><br /><input type="submit" value="<%=yySubmit%>" />&nbsp; &nbsp;<input type="reset" value="<%=yyReset%>" /><br />&nbsp;</td>
			  </tr>
            </table></td>
		  </form>
        </tr>
<%	rs.Close
end if
cnn.Close
set rs = nothing
set cnn = nothing%>
      </table>