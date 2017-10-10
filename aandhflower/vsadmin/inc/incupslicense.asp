<%
if storesessionvalue="" then storesessionvalue="virtualstore"
if Session("loggedon") <> storesessionvalue OR disallowlogin=TRUE then response.end
Response.Charset = "8859-1"
success=true
Set rs = Server.CreateObject("ADODB.RecordSet")
Set cnn=Server.CreateObject("ADODB.Connection")
cnn.open sDSN
alreadygotadmin = getadminsettings()
countryCode = origCountryCode
Function ParseUPSLicenseOutput(xmlDoc, rootNodeName, byRef thetext, byRef errormsg)
Dim noError, nodeList, e, i, j, k, l, n, t, t2
	noError = True
	errormsg = ""
	gotxml=false
	thetext=""
	Set t2 = xmlDoc.getElementsByTagName(rootNodeName).Item(0)
	' response.write Replace(Replace(xmlDoc.xml,"</","&lt;/"),"<","<br />&lt;")&"<HR>"
	for j = 0 to t2.childNodes.length - 1
		Set n = t2.childNodes.Item(j)
		if n.nodename="Response" then
			For i = 0 To n.childNodes.length - 1
				Set e = n.childNodes.Item(i)
				if e.nodeName="ResponseStatusCode" then
					noError = Int(e.firstChild.nodeValue)=1
				end if
				if e.nodeName="Error" then
					errormsg = ""
					For k = 0 To e.childNodes.length - 1
						Set t = e.childNodes.Item(k)
						Select Case t.nodeName
							Case "ErrorSeverity"
								if t.firstChild.nodeValue="Transient" then errormsg = "This is a temporary error. Please wait a few moments then refresh this page.<br />" & errormsg
							Case "ErrorDescription"
								errormsg = errormsg & t.firstChild.nodeValue
						End Select
					Next
				end if
				' response.write "The Nodename is : " & e.nodeName & ":" & e.firstChild.nodeValue & "<br />"
			Next
		elseif n.nodename="AccessLicenseNumber" then
			thetext = n.firstChild.nodeValue
		elseif n.nodename="AccessLicenseText" then
			thetext = n.firstChild.nodeValue
			if mysqlserver then rs.CursorLocation = 3
			rs.Open "SELECT * FROM admin WHERE adminID=1",cnn,1,3,&H0001
			rs.Fields("adminUPSLicense")=n.firstChild.nodeValue
			rs.Update
			rs.Close
		elseif n.nodename="UserId" then
			thetext = n.firstChild.nodeValue
		end if
	Next
	ParseUPSLicenseOutput = noError
end Function

if request.form("upsstep")="4" then
	Set docXML = Server.CreateObject("MSXML2.DOMDocument")
	sSQL = "SELECT adminUPSLicense FROM admin WHERE adminID=1"
	rs.Open sSQL,cnn,0,1
	sXML = "<?xml version=""1.0"" encoding=""ISO-8859-1""?>"
	sXML = sXML & "<AccessLicenseRequest xml:lang=""en-US""><Request><TransactionReference><CustomerContext>Ecomm Plus UPS Reg</CustomerContext><XpciVersion>1.0001</XpciVersion></TransactionReference>"
	sXML = sXML & "<RequestAction>AccessLicense</RequestAction><RequestOption>AllTools</RequestOption></Request>"
	sXML = sXML & "<CompanyName>" & Request.Form("company") & "</CompanyName>"
	sXML = sXML & "<Address><AddressLine1>" & Request.Form("address") & "</AddressLine1>"
	if Trim(Request.Form("address2"))<>"" then sXML = sXML & "<AddressLine2>" & Request.Form("address2") & "</AddressLine2>"
	sXML = sXML & "<City>" & Request.Form("city") & "</City>"
	if Trim(Request.Form("country"))="US" OR Trim(Request.Form("country"))="CA" then
		sXML = sXML & "<StateProvinceCode>" & Request.Form("usstate") & "</StateProvinceCode>"
	else
		sXML = sXML & "<StateProvinceCode>XX</StateProvinceCode>"
	end if
	if Trim(Request.Form("postcode"))<>"" then sXML = sXML & "<PostalCode>" & Request.Form("postcode") & "</PostalCode>"
	sXML = sXML & "<CountryCode>" & Request.Form("country") & "</CountryCode></Address>"
	sXML = sXML & "<PrimaryContact><Name>" & Request.Form("contact") & "</Name><Title>" & Request.Form("ctitle") & "</Title>"
	sXML = sXML & "<EMailAddress>" & Request.Form("email") & "</EMailAddress><PhoneNumber>" & Request.Form("telephone") & "</PhoneNumber></PrimaryContact>"
	sXML = sXML & "<CompanyURL>" & Request.Form("websiteurl") & "</CompanyURL>"
	if Trim(Request.Form("upsaccount"))<>"" then sXML = sXML & "<ShipperNumber>" & Request.Form("upsaccount") & "</ShipperNumber>"
	sXML = sXML & "<DeveloperLicenseNumber>BB9341E83CC05B12</DeveloperLicenseNumber>"
	sXML = sXML & "<AccessLicenseProfile><CountryCode>" & Request.Form("countryCode") & "</CountryCode><LanguageCode>" & Request.Form("languageCode") & "</LanguageCode>"
	sXML = sXML & "<AccessLicenseText>" & rs("adminUPSLicense") & "</AccessLicenseText>"
	sXML = sXML & "</AccessLicenseProfile>"
	sXML = sXML & "<OnLineTool><ToolID>RateXML</ToolID><ToolVersion>1.0</ToolVersion></OnLineTool><OnLineTool><ToolID>TrackXML</ToolID><ToolVersion>1.0</ToolVersion></OnLineTool>"
	sXML = sXML & "<ClientSoftwareProfile><SoftwareInstaller>" & Request.Form("upsrep") & "</SoftwareInstaller><SoftwareProductName>default</SoftwareProductName><SoftwareProvider>Internet Business Solutions SL</SoftwareProvider><SoftwareVersionNumber>2.5</SoftwareVersionNumber></ClientSoftwareProfile>"
	sXML = sXML & "</AccessLicenseRequest>"
	docXML.loadXML(sXML)
	rs.Close
	' response.write Replace(Replace(docXML.xml,"</","&lt;/"),"<","<br />&lt;")&"<HR>"
	' response.flush
	set objHttp = Server.CreateObject("Msxml2.ServerXMLHTTP")
	objHttp.open "POST", "https://www.ups.com/ups.app/xml/License", false
	objHttp.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
	on error resume next
	err.number=0
	objHttp.Send docXML
	on error goto 0
	If err.number <> 0 OR objHttp.status <> 200 Then
		errormsg = "Error, couldn't connect to UPS server<br />" & objHttp.status & ": (" & objHttp.statusText & ") " & err.number & ": (" & err.Description & ")"
		success = false
	Else
		saveLCID = Session.LCID
		Session.LCID = 1033
		'response.write Replace(Replace(objHttp.responseText,"</","&lt;/"),"<","<br />&lt;")&"<HR>"
		success = ParseUPSLicenseOutput(objHttp.responseXML, "AccessLicenseResponse", accessnumber, errormsg)
		Session.LCID = saveLCID
	End If
	set objHttp = nothing
	if success then
		sSQL = "UPDATE admin SET adminUPSAccess='"&accessnumber&"'"
		cnn.Execute(sSQL)
		noloops=0
		Randomize
		upperbound = "999999"
		lowerbound = "100000"
		thepw = "ecp" & Int((upperbound - lowerbound + 1) * Rnd + lowerbound)
		theuser = "ecu" & Int((upperbound - lowerbound + 1) * Rnd + lowerbound)
		do while theuser<>"" AND success AND noloops < 5
			saveuser = theuser
			sXML = "<?xml version=""1.0"" encoding=""ISO-8859-1""?>"
			sXML = sXML & "<RegistrationRequest><Request><TransactionReference><CustomerContext>Ecomm Plus UPS Reg</CustomerContext><XpciVersion>1.0001</XpciVersion></TransactionReference>"
			sXML = sXML & "<RequestAction>Register</RequestAction><RequestOption>suggest</RequestOption></Request>"
			sXML = sXML & "<UserId>"&theuser&"</UserId><Password>"&thepw&"</Password><RegistrationInformation>"
			sXML = sXML & "<UserName>" & Request.Form("contact") & "</UserName>"
			sXML = sXML & "<CompanyName>" & Request.Form("company") & "</CompanyName>"
			sXML = sXML & "<Title>" & Request.Form("ctitle") & "</Title><Address>"
			sXML = sXML & "<AddressLine1>" & Request.Form("address") & "</AddressLine1>"
			if Trim(Request.Form("address2"))<>"" then sXML = sXML & "<AddressLine2>" & Request.Form("address2") & "</AddressLine2>"
			sXML = sXML & "<City>" & Request.Form("city") & "</City>"
			if Trim(Request.Form("country"))="US" OR Trim(Request.Form("country"))="CA" then
				sXML = sXML & "<StateProvinceCode>" & Request.Form("usstate") & "</StateProvinceCode>"
			else
				sXML = sXML & "<StateProvinceCode>XX</StateProvinceCode>"
			end if
			if Trim(Request.Form("postcode"))<>"" then sXML = sXML & "<PostalCode>" & Request.Form("postcode") & "</PostalCode>"
			sXML = sXML & "<CountryCode>" & Request.Form("country") & "</CountryCode></Address>"
			sXML = sXML & "<PhoneNumber>" & Request.Form("telephone") & "</PhoneNumber>"
			sXML = sXML & "<EMailAddress>" & Request.Form("email") & "</EMailAddress>"
			'if Trim(Request.Form("upsaccount"))<>"" then sXML = sXML & "<ShipperNumber>" & Request.Form("upsaccount") & "</ShipperNumber>"
			sXML = sXML & "</RegistrationInformation></RegistrationRequest>"
			'response.write Replace(Replace(sXML,"</","&lt;/"),"<","<br />&lt;")&"<HR>"
			set objHttp = Server.CreateObject("Msxml2.ServerXMLHTTP")
			objHttp.open "POST", "https://www.ups.com/ups.app/xml/Register", false
			objHttp.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
			on error resume next
			err.number=0
			objHttp.Send sXML
			on error goto 0
			If err.number <> 0 OR objHttp.status <> 200 Then
				errormsg = "Error, couldn't connect to UPS server<br />" & objHttp.status & ": (" & objHttp.statusText & ") " & err.number & ": (" & err.Description & ")"
				success = false
			Else
				saveLCID = Session.LCID
				Session.LCID = 1033
				'response.write Replace(Replace(objHttp.responseText,"</","&lt;/"),"<","<br />&lt;")&"<HR>"
				success = ParseUPSLicenseOutput(objHttp.responseXML, "RegistrationResponse", theuser, errormsg)
				Session.LCID = saveLCID
			End If
			set objHttp = nothing
			noloops=noloops+1
		loop
	end if
	Set docXML = nothing
%>
	<form method="post" name="licform" action="admin.asp">
	  <input type="hidden" name="upsstep" value="5" />
      <table border="0" cellspacing="0" cellpadding="0" width="100%" bgcolor="" align="center">
        <tr> 
          <td width="100%">
            <table width="100%" border="0" cellspacing="0" cellpadding="2" bgcolor="">
			  <tr>
				<td rowspan="3" width="70" align="center" valign="top"><img src="../images/LOGO_S.gif" border="0" alt="UPS" /><br />&nbsp;<br />&nbsp;<br />&nbsp;<br />&nbsp;</td>
                <td width="100%" align="center"><strong><%=yyUPSWiz%> - <% if success then response.write yyRegSucc else response.write yyError %></strong><br />&nbsp;
                </td>
			  </tr>
<%	if success then
		sSQL = "UPDATE admin SET adminUPSUser='"&upsencode(saveuser, "")&"',adminUPSpw='"&upsencode(thepw, "")&"'"
		cnn.Execute(sSQL)
		Application.Lock()
		Application("getadminsettings")=""
		Application.UnLock()
%>
			  <tr> 
                <td width="100%" align="left">
				  <p><strong><%=yyRegSucc%> !</strong></p>
				  <p><%=yyUPSLi5%></p>
				  <p><%=yyUPSLi6%> <a href="http://www.ec.ups.com" target="_blank">www.ec.ups.com</a>.</p>
				  <p><%=yyUPSLi7%> <a href="adminmain.asp"><%=yyAdmMai%></a>.</p>
				  <p><%=yyUPSLi8%> <a href="http://ups.com/bussol/solutions/internetship.html" target="_blank"><%=yyClkHer%></a>.</p>
				  <p>&nbsp;</p>
				  <p align="center"><input type="submit" value="<%=yyDone%>" /></p>
				  <p>&nbsp;</p>
                </td>
			  </tr>
<%	else %>
			  <tr> 
                <td width="100%" align="center"><p><%=yySorErr%></strong></p>
				<p>&nbsp;</p>
				<p><%=errormsg%></p>
				<p>&nbsp;</p>
				<p><%=yyTryBac%> <a href="javascript:history.go(-1)"><%=yyClkHer%></a>.</p>
				<p>&nbsp;</p>
                </td>
			  </tr>
<%	end if %>
			  <tr> 
                <td colspan="2" width="100%" align="center">
				  <p><img src="../images/clearpixel.gif" width="300" height="5" alt="" /></p>
				  <p><font size="1"><%=yyUPStm%></font></p>
                </td>
			  </tr>
            </table>
          </td>
        </tr>
      </table>
	</form>
<%
elseif request.form("upsstep")="3" AND request.form("doagree")="1" then
%>
<script language="javascript" type="text/javascript">
<!--
function checkforamp(checkObj){
  checkStr = checkObj.value;
  for (i = 0;  i < checkStr.length;  i++){
	if (checkStr.charAt(i) == "&"){
	  alert("Please do not use the ampersand \"&\" character in any field.");
	  checkObj.focus();
	  return(false);
	}
  }
  return(true);
}
function formvalidator(theForm)
{
  if(theForm.contact.value == ""){
    alert("<%=yyPlsEntr%> \"<%=yyConNam%>\".");
    theForm.contact.focus();
    return (false);
  }
  if(!checkforamp(theForm.contact)) return(false);
  if(theForm.ctitle.value == ""){
    alert("<%=yyPlsEntr%> \"<%=yyTitle%>\".");
    theForm.ctitle.focus();
    return (false);
  }
  if(!checkforamp(theForm.ctitle)) return(false);
  if(theForm.company.value == ""){
    alert("<%=yyPlsEntr%> \"<%=yyComNam%>\".");
    theForm.company.focus();
    return (false);
  }
  if(!checkforamp(theForm.company)) return(false);
  if(theForm.address.value == ""){
    alert("<%=yyPlsEntr%> \"<%=yyStrAdd%>\".");
    theForm.address.focus();
    return (false);
  }
  if(!checkforamp(theForm.address)) return(false);
  if(theForm.city.value == ""){
    alert("<%=yyPlsEntr%> \"<%=yyCity%>\".");
    theForm.city.focus();
    return (false);
  }
  if(!checkforamp(theForm.city)) return(false);
  var cntry = theForm.country[theForm.country.selectedIndex].value;
  if(cntry=="US" || cntry=="CA"){
	if (theForm.usstate.selectedIndex == 0){
      alert("<%=yyPlsSel%> \"<%=yyState%>\".");
      theForm.usstate.focus();
      return (false);
	}
  }
  if(theForm.country.selectedIndex == 0){
    alert("<%=yyPlsSel%> \"<%=yyCountry%>\".");
    theForm.country.focus();
    return (false);
  }
  if(cntry!='CL' && cntry!='CO' && cntry!='CR' && cntry!='DO' && cntry!='GT' && cntry!='HK' && cntry!='IE' && cntry!='PA'){
	if (theForm.postcode.value == ""){
	  alert("<%=yyPlsEntr%> \"<%=yyPCode%>\".");
	  theForm.postcode.focus();
	  return (false);
	}
  }
  if(!checkforamp(theForm.postcode)) return(false);
  if(theForm.telephone.value == ""){
    alert("<%=yyPlsEntr%> \"<%=yyTelep%>\".");
    theForm.telephone.focus();
    return (false);
  }
  if(theForm.telephone.value.length < 10 || theForm.telephone.value.length > 14){
    alert("<%=yyValTN%>");
    theForm.telephone.focus();
    return (false);
  }
  var checkOK = "0123456789";
  var checkStr = theForm.telephone.value;
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
  if(!allValid)
  {
    alert("<%=yyOnDig%> \"<%=yyTelep%>\".");
    theForm.telephone.focus();
    return (false);
  }
  if(theForm.websiteurl.value == ""){
    alert("<%=yyPlsEntr%> \"<%=yyWebURL%>\".");
    theForm.websiteurl.focus();
    return (false);
  }
  if(!checkforamp(theForm.websiteurl)) return(false);
  var checkStr = theForm.websiteurl.value;
  var gotDot = false;
  for (i = 0;  i < checkStr.length;  i++)
  {
    ch = checkStr.charAt(i);
	if (ch == ".") gotDot = true;
  }
  if(!(gotDot))
  {
    alert("<%=yyValEnt%> \"<%=yyWebURL%>\".");
    theForm.websiteurl.focus();
    return (false);
  }
  if(theForm.email.value == ""){
    alert("<%=yyPlsEntr%> \"<%=yyEmail%>\".");
    theForm.email.focus();
    return (false);
  }
  var checkStr = theForm.email.value;
  var gotDot = false;
  var gotAt = false;
  for (i = 0;  i < checkStr.length;  i++)
  {
    ch = checkStr.charAt(i);
    if (ch == "@") gotAt = true;
	if (ch == ".") gotDot = true;
  }
  if (!(gotDot && gotAt))
  {
    alert("<%=yyValEnt%> \"<%=yyEmail%>\".");
    theForm.email.focus();
    return (false);
  }
  if(theForm.upsrep[0].checked==false && theForm.upsrep[1].checked==false){
    alert("<%=yyUPSrep%>");
    return (false);
  }
  return (true);
}
//-->
</script>
	<form method="post" name="licform" action="adminupslicense.asp" onsubmit="return formvalidator(this)">
	  <input type="hidden" name="upsstep" value="4" />
	  <input type="hidden" name="countryCode" value="<%=Request.Form("countryCode")%>" />
	  <input type="hidden" name="languageCode" value="<%=Request.Form("languageCode")%>" />
      <table border="0" cellspacing="0" cellpadding="0" width="100%" bgcolor="" align="center">
        <tr> 
          <td width="100%">
            <table width="100%" border="0" cellspacing="0" cellpadding="2" bgcolor="">
			  <tr>
				<td rowspan="18" width="70" align="center" valign="top"><img src="../images/LOGO_S.gif" border="0" alt="UPS" /><br />&nbsp;<br />&nbsp;<br />&nbsp;<br />&nbsp;</td>
                <td width="100%" align="center" colspan="2"><strong><%=yyUPSWiz%> - <%=yyStep%> 2</strong><br />&nbsp;
                </td>
			  </tr>
			  <tr> 
                <td width="40%" align="right"><strong><%=yyConNam%> : </strong></td>
				<td width="60%"><input type="text" name="contact" size="30" /></td>
			  </tr>
			  <tr> 
                <td align="right"><strong><%=yyTitle%> : </strong></td>
				<td><input type="text" name="ctitle" size="10" /></td>
			  </tr>
			  <tr> 
                <td align="right"><strong><%=yyComNam%> : </strong></td>
				<td><input type="text" name="company" size="30" /></td>
			  </tr>
			  <tr> 
                <td align="right"><strong><%=yyStrAdd%> : </strong></td>
				<td><input type="text" name="address" size="30" /></td>
			  </tr>
			  <tr> 
                <td align="right"><strong><%=yyAddr2%> : </strong></td>
				<td><input type="text" name="address2" size="30" /></td>
			  </tr>
			  <tr> 
                <td align="right"><strong><%=yyCity%> : </strong></td>
				<td><input type="text" name="city" size="30" /></td>
			  </tr>
			  <tr> 
                <td align="right"><strong><%=yyState%> <%=yyUSCan%> : </strong></td>
				<td><select name="usstate" size="1">
<option value=''><%=yyOutUS%></option>
<option value='AL'>Alabama</option>
<option value='AK'>Alaska</option>
<option value='AB'>Alberta</option>
<option value='AZ'>Arizona</option>
<option value='AR'>Arkansas</option>
<option value='BC'>British Columbia</option>
<option value='CA'>California</option>
<option value='CO'>Colorado</option>
<option value='CT'>Connecticut</option>
<option value='DE'>Delaware</option>
<option value='DC'>District Of Columbia</option>
<option value='FL'>Florida</option>
<option value='GA'>Georgia</option>
<option value='HI'>Hawaii</option>
<option value='ID'>Idaho</option>
<option value='IL'>Illinois</option>
<option value='IN'>Indiana</option>
<option value='IA'>Iowa</option>
<option value='KS'>Kansas</option>
<option value='KY'>Kentucky</option>
<option value='LA'>Louisiana</option>
<option value='ME'>Maine</option>
<option value='MB'>Manitoba</option>
<option value='MD'>Maryland</option>
<option value='MA'>Massachusetts</option>
<option value='MI'>Michigan</option>
<option value='MN'>Minnesota</option>
<option value='MS'>Mississippi</option>
<option value='MO'>Missouri</option>
<option value='MT'>Montana</option>
<option value='NE'>Nebraska</option>
<option value='NV'>Nevada</option>
<option value='NB'>New Brunswick</option>
<option value='NH'>New Hampshire</option>
<option value='NJ'>New Jersey</option>
<option value='NM'>New Mexico</option>
<option value='NY'>New York</option>
<option value='NF'>Newfoundland</option>
<option value='NC'>North Carolina</option>
<option value='ND'>North Dakota</option>
<option value='NT'>Northwest Territories</option>
<option value='NS'>Nova Scotia</option>
<option value='NU'>Nunavut</option>
<option value='OH'>Ohio</option>
<option value='OK'>Oklahoma</option>
<option value='ON'>Ontario</option>
<option value='OR'>Oregon</option>
<option value='PA'>Pennsylvania</option>
<option value='PI'>Prince Edward Island</option>
<option value='PQ'>Quebec</option>
<option value='RI'>Rhode Island</option>
<option value='SK'>Saskatchewan</option>
<option value='SC'>South Carolina</option>
<option value='SD'>South Dakota</option>
<option value='TN'>Tennessee</option>
<option value='TX'>Texas</option>
<option value='UT'>Utah</option>
<option value='VT'>Vermont</option>
<option value='VA'>Virginia</option>
<option value='WA'>Washington</option>
<option value='WV'>West Virginia</option>
<option value='WI'>Wisconsin</option>
<option value='WY'>Wyoming</option>
<option value='YT'>Yukon</option>
</select></td>
			  </tr>
			  <tr> 
                <td align="right"><strong><%=yyCountry%> : </strong></td>
				<td><select name="country" size="1">
<option value=''><%=yySelect%></option>
<%
sSQL = "SELECT countryName,countryCode FROM countries WHERE countryCode IN ('AR','AU','AT','BE','BR','CA','CL','CN','CO','CR','DK','DO','FI','FR','DE','GR','GT','HK','IN','IE','IL','IT','JP','MY','MX','NL','NZ','NO','PA','PH','PT','PR','SG','KR','ES','SE','CH','TW','TH','GB','US') ORDER BY countryName"
rs.Open sSQL,cnn,0,1
do while not rs.EOF
	response.write "<option value='"&rs("countryCode")&"'>"&rs("countryName")&"</option>"
	rs.MoveNext
loop
rs.Close
%>
				</select></td>
			  </tr>
			  <tr> 
                <td align="right"><strong><%=yyPCode%> : </strong></td>
				<td><input type="text" name="postcode" size="15" /></td>
			  </tr>
			  <tr> 
                <td align="right"><strong><%=yyTelep%> : </strong></td>
				<td><input type="text" name="telephone" size="30" /></td>
			  </tr>
			  <tr> 
                <td align="right"><strong><%=yyWebURL%> : </strong></td>
				<td><input type="text" name="websiteurl" size="30" /></td>
			  </tr>
			  <tr> 
                <td align="right"><strong><%=yyEmail%> : </strong></td>
				<td><input type="text" name="email" size="30" /></td>
			  </tr>
			  <tr> 
                <td align="right"><strong><%=yyUPSac%> : </strong></td>
				<td><input type="text" name="upsaccount" size="30" /></td>
			  </tr>
			  <tr> 
                <td align="center" colspan="2">
				  <table border="0" cellspacing="0" cellpadding="0" width="100%" bgcolor="" align="center">
        <tr> 
          <td width="70%" align="center"><%=yyUPSsr%><br /><input type="radio" name="upsrep" value="yes" /> <strong><%=yyYes%></strong> <input type="radio" name="upsrep" value="no" /> <strong><%=yyNo%></strong></td>
				   </tr></table></td>
			  </tr>
			  <tr>
                <td width="100%" align="center" colspan="2"><br />&nbsp;<input type="submit" name="agree" value="&nbsp;&nbsp;<%=yyNext%>&nbsp;&nbsp;" />&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
				<input type="button" value="<%=yyCancel%>" onclick="javascript:window.location='admin.asp';" />
                </td>
			  </tr>
			  <tr> 
                <td align="center" colspan="2"><p><font size="1"><%=yyUPSop%> <a href="http://www.ups.com/content/us/en/resources/service/account.html" target="_blank"><%=yyClkHer%></a> <%=yyUPScl%><br />
				<%=yyUPSMI%> <a href="http://www.ec.ups.com" target="_blank"><%=yyClkHer%></a>.<br />
				<%=yyUPshp%> <a href="http://ups.com/bussol/solutions/internetship.html" target="_blank"><%=yyClkHer%></a></font></p>
				</td>
			  </tr>
			  <tr> 
                <td colspan="3" width="100%" align="center">
				  <p><img src="../images/clearpixel.gif" width="300" height="5" alt="" /></p>
				  <p><font size="1"><%=yyUPStm%></font></p>
                </td>
			  </tr>
            </table>
          </td>
        </tr>
      </table>
	</form>
<%
elseif request.form("upsstep")="2" then
	languageCode="EN"
	if countryCode="AR" OR countryCode="ES" OR countryCode="MX" OR countryCode="CA" OR countryCode="DO" OR countryCode="GT" OR countryCode="CR" OR countryCode="CO" OR countryCode="PA" OR countryCode="PR" OR countryCode="CL" then
		languageCode="ES"
	elseif countryCode="AT" OR countryCode="DE" then
		languageCode="DE"
	elseif countryCode="PT" OR countryCode="BR" then
		languageCode="PT"
	elseif countryCode="FR" OR countryCode="CH" OR countryCode="BE" then
		languageCode="FR"
	elseif countryCode="CN" OR countryCode="HK" then
		languageCode="ZH"
	elseif countryCode="DK" then
		languageCode="DA"
	elseif countryCode="FI" then
		languageCode="FI"
	elseif countryCode="GR" then
		languageCode="EL"
	elseif countryCode="IN" then
		languageCode="GU"
	elseif countryCode="IL" then
		languageCode="IW"
	elseif countryCode="IT" then
		languageCode="IT"
	elseif countryCode="JP" then
		languageCode="JA"
	elseif countryCode="MY" then
		languageCode="MS"
	elseif countryCode="NL" then
		languageCode="NL"
	elseif countryCode="NO" then
		languageCode="NO"
	elseif countryCode="KR" then
		languageCode="KO"
	elseif countryCode="SE" then
		languageCode="SV"
	elseif countryCode="TH" then
		languageCode="TH"
	end if
	sXML = "<?xml version=""1.0"" encoding=""ISO-8859-1""?>"
	sXML = sXML & "<AccessLicenseAgreementRequest><Request><RequestOption>AllTools</RequestOption><TransactionReference><CustomerContext>Ecomm Plus UPS License</CustomerContext><XpciVersion>1.0001</XpciVersion></TransactionReference>"
	sXML = sXML & "<RequestAction>AccessLicense</RequestAction></Request><DeveloperLicenseNumber>BB9341E83CC05B12</DeveloperLicenseNumber>"
	sXML = sXML & "<AccessLicenseProfile><CountryCode>"&countryCode&"</CountryCode><LanguageCode>"&languageCode&"</LanguageCode></AccessLicenseProfile>"
	sXML = sXML & "<OnLineTool><ToolID>RateXML</ToolID><ToolVersion>1.0</ToolVersion></OnLineTool><OnLineTool><ToolID>TrackXML</ToolID><ToolVersion>1.0</ToolVersion></OnLineTool></AccessLicenseAgreementRequest>"

	' response.write Replace(Replace(sXML,"</","&lt;/"),"<","<br />&lt;")&"<HR>"
	set objHttp = Server.CreateObject("Msxml2.ServerXMLHTTP")
	objHttp.open "POST", "https://www.ups.com/ups.app/xml/License", false
	objHttp.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
	on error resume next
	err.number=0
	objHttp.Send sXML
	on error goto 0
	lictext = ""
	If err.number <> 0 OR objHttp.status <> 200 Then
		errormsg = "Error, couldn't connect to UPS server<br />" & objHttp.status & ": (" & objHttp.statusText & ") " & err.number & ": (" & err.Description & ")"
		success = false
	Else
		saveLCID = Session.LCID
		Session.LCID = 1033
		success = ParseUPSLicenseOutput(objHttp.responseXML, "AccessLicenseAgreementResponse", lictext, errormsg)
		Session.LCID = saveLCID
	End If
	set objHttp = nothing
%>
<script language="javascript" type="text/javascript">
<!--
var origlictext="";
function printlicense()
{
	var prnttext = '<html><body>\n';
	if(origlictext != document.licform.lictext.value){
		alert("It appears that the license text has been modified. Cannot print license.");
		return;
	}
	prnttext += document.licform.lictext.value.replace(/\n|\r\n/g,'<br />');
	prnttext += '</body></html>';
	var newwin = window.open("","printlicense",'menubar=no, scrollbars=yes, width=500, height=400, directories=no,location=no,resizable=yes,status=no,toolbar=no');
	newwin.document.open();
	newwin.document.write(prnttext);
	newwin.document.close();
	newwin.print();
}
function checkaccept(theForm)
{
  if(origlictext != document.licform.lictext.value){
	alert("It appears that the license text has been modified. Cannot proceed.");
	return (false);
  }
  if (theForm.doagree[0].checked == false)
  {
    alert("<%=yyUPSLi4%>");
    return (false);
  }
  return (true);
}
//-->
</script>
	<form method="post" name="licform" action="adminupslicense.asp" onsubmit="return checkaccept(this)">
	  <input type="hidden" name="upsstep" value="3" />
	  <input type="hidden" name="countryCode" value="<%=countryCode%>" />
	  <input type="hidden" name="languageCode" value="<%=languageCode%>" />
      <table border="0" cellspacing="0" cellpadding="0" width="100%" bgcolor="" align="center">
        <tr> 
          <td width="100%">
            <table width="100%" border="0" cellspacing="2" cellpadding="0" bgcolor="">
			  <tr>
                <td width="100%" align="center"><img src="../images/LOGO_S.gif" border="0" align="middle" alt="" />&nbsp;&nbsp;<strong><%=yyUPSWiz%> - <%=yyStep%> 1</strong><br />&nbsp;
                </td>
			  </tr>
<%	if success then %>
			  <tr> 
                <td width="100%" align="center"><textarea cols="80" rows="20" name="lictext"><%=lictext%></textarea><br /><br />
				<p><%=yyUPSTer%></p>
				<p><%=yyAgree%> <input type="radio" name="doagree" value="1" />&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<%=yyNoAgre%> <input type="radio" name="doagree" value="0" /></p>
				<p>&nbsp;</p>
                </td>
			  </tr>
<script language="javascript" type="text/javascript">
<!--
var origlictext=document.licform.lictext.value;
//-->
</script>
<%	else %>
			  <tr> 
                <td width="100%" align="center"><p><%=yySorErr%></strong></p>
				<p>&nbsp;</p>
				<p><%=errormsg%></p>
				<p>&nbsp;</p>
                </td>
			  </tr>
<%	end if %>
			  <tr> 
                <td width="100%" align="center"><% if success then %><input type="button" value="&nbsp;<%=yyPrint%>&nbsp;" onclick="javascript:printlicense();" />&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input type="submit" name="agree" value="&nbsp;&nbsp;<%=yyNext%>&nbsp;&nbsp;" />&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<% end if %>
				<input type="button" value="<%=yyCancel%>" onclick="javascript:window.location='admin.asp';" />
                </td>
			  </tr>
			  <tr> 
                <td align="center"><p><font size="1"><%=yyUPSop%> <a href="http://www.ups.com/content/us/en/resources/service/account.html" target="_blank"><%=yyClkHer%></a> <%=yyUPScl%><br />
				<%=yyUPSMI%> <a href="http://www.ec.ups.com" target="_blank"><%=yyClkHer%></a>.<br />
				<%=yyUPshp%> <a href="http://ups.com/bussol/solutions/internetship.html" target="_blank"><%=yyClkHer%></a>.</font></p>
				</td>
			  </tr>
			  <tr> 
                <td width="100%" align="center">
				  <p><img src="../images/clearpixel.gif" width="300" height="5" alt="" /></p>
				  <p><font size="1"><%=yyUPStm%></font></p>
                </td>
			  </tr>
            </table>
          </td>
        </tr>
      </table>
	</form>
<%
else %>
	<form method="post" action="adminupslicense.asp">
	  <input type="hidden" name="upsstep" value="2" />
      <table border="0" cellspacing="0" cellpadding="0" width="100%" bgcolor="" align="center">
        <tr> 
          <td width="100%">
            <table width="100%" border="0" cellspacing="0" cellpadding="2" bgcolor="">
			  <tr>
				<td rowspan="5" width="70" align="center" valign="top"><img src="../images/LOGO_S.gif" border="0" alt="" /><br />&nbsp;</td>
                <td width="100%" align="center"><strong><%=yyUPSWiz%></strong><br />&nbsp;
                </td>
			  </tr>
			  <tr> 
                <td width="100%"><ul><li><%=yyUPSLi1%><br /><br /></li>
				<li><%=yyUPSLi2%><br /><br /></li>
				<li><%=yyUPSLi3%> <%=yyNoCou%> <a href="adminmain.asp"><%=yyClkHer%></a>.<br /><br /></li>
				<li><%=yyUPSMI%> <a href="http://www.ec.ups.com" target="_blank"><%=yyClkHer%></a>.<br /><br /></li>
				<li><%=yyUPshp%> <a href="http://ups.com/bussol/solutions/internetship.html" target="_blank"><%=yyClkHer%></a>.</li>
				</ul>
				<p>&nbsp;</p>
                </td>
			  </tr>
			  <tr> 
                <td width="100%" align="center"><input type="submit" name="agree" value="&nbsp;&nbsp;<%=yyNext%>&nbsp;&nbsp;" />&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
				<input type="button" value="<%=yyCancel%>" onclick="javascript:window.location='admin.asp';" />
                </td>
			  </tr>
			  <tr> 
                <td align="center" colspan="2"><p><font size="1"><%=yyUPSop%> <a href="http://www.ups.com/content/us/en/resources/service/account.html" target="_blank"><%=yyClkHer%></a> <%=yyUPScl%><br />
				<%=yyUPSMI%> <a href="http://www.ec.ups.com" target="_blank"><%=yyClkHer%></a>.<br />
				<%=yyUPshp%> <a href="http://ups.com/bussol/solutions/internetship.html" target="_blank"><%=yyClkHer%></a>.</font></p>
				</td>
			  </tr>
			  <tr> 
                <td width="100%" align="center">
				  <p><img src="../images/clearpixel.gif" width="300" height="5" alt="" /></p>
				  <p><font size="1"><%=yyUPStm%></font></p>
                </td>
			  </tr>
            </table>
          </td>
        </tr>
      </table>
	</form>
<%
end if
cnn.Close
set rs = nothing
set cnn = nothing
%>