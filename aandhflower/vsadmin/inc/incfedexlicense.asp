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
Function ParseFedexOutput(sXML, rootNodeName, byRef thetext, byRef errormsg)
Dim noError, nodeList, e, i, j, k, l, n, t, t2
	noError = False
	errormsg = ""
	gotxml=false
	thetext=""
	set xmlDoc = Server.CreateObject("MSXML2.DOMDocument")
	xmlDoc.validateOnParse = False
	xmlDoc.loadXML (sXML)
	for k = 0 to xmlDoc.childNodes.length - 1
		Set t2 = xmlDoc.childNodes.Item(k)
		if t2.nodename=rootNodeName then
			for j = 0 to t2.childNodes.length - 1
				Set n = t2.childNodes.Item(j)
				if n.nodename="Error" then
					For i = 0 To n.childNodes.length - 1
						Set e = n.childNodes.Item(i)
						if e.nodeName="Code" then
							' No action
						elseif e.nodeName="Message" then
							errormsg = e.firstChild.nodeValue
						end if
					Next
				elseif n.nodename="MeterNumber" then
					thetext = n.firstChild.nodeValue
					noError = True
				end if
			Next
		elseif t2.nodename="Error" then
			for j = 0 to t2.childNodes.length - 1
				Set n = t2.childNodes.Item(j)
				if n.nodeName="Code" then
					' No action
				elseif n.nodeName="Message" then
					errormsg = n.firstChild.nodeValue
				end if
			Next
		end if
	next
	ParseFedexOutput = noError
end Function
if request.querystring("act")="version" then %>
	<form method="post" name="licform" action="admin.asp">
	  <input type="hidden" name="upsstep" value="5" />
      <table border="0" cellspacing="0" cellpadding="0" width="100%" bgcolor="" align="center">
        <tr> 
          <td width="100%">
            <table width="100%" border="0" cellspacing="0" cellpadding="2" bgcolor="">
			  <tr>
				<td rowspan="3" width="70" align="center" valign="top"><img src="../images/fedexsmall.gif" border="0" alt="FedEx" /><br />&nbsp;<br />&nbsp;<br />&nbsp;<br />&nbsp;</td>
                <td width="100%" align="center"><strong><%=yyFdxWiz%> - Updating FedEx® version information.</strong><br />&nbsp;
                </td>
			  </tr>
			  <tr> 
                <td width="100%" align="left">
				  <p>&nbsp;</p>
				  <p>Please wait while we update your FedEx version information.</p>
				  <p>&nbsp;</p>
				  <p>Step 1, getting location id. <span name="step1span" id="step1span"><strong>Please wait!</strong></span></p>
				  <p>&nbsp;</p>
				  <p>Step 2, updating version. <span name="step2span" id="step2span"><strong>Please wait!</strong></span></p>
				  <p>&nbsp;</p>
				  <p align="center" name="donebutton" id="donebutton" style="display:none"><input type="submit" value="<%=yyDone%>" /></p>
				  <p>&nbsp;</p>
                </td>
			  </tr>
			  <tr> 
                <td colspan="2" width="100%" align="center">
				  <p><img src="../images/clearpixel.gif" width="300" height="5" alt="" /></p>
				  <p><font size="1">FedEx® is a registered service mark of Federal Express Corporation.
FedEx logos used by permission. All rights reserved.</font></p>
                </td>
			  </tr>
            </table>
          </td>
        </tr>
      </table>
	</form>
<%	response.flush
	sSQL = "SELECT adminVersion,FedexAccountNo,FedexMeter,adminZipCode,countryCode FROM admin INNER JOIN countries ON admin.adminCountry=countries.countryID WHERE adminID=1"
	rs.Open sSQL,cnn,0,1
	if not rs.EOF then
		version = trim(rs("adminVersion")&"")
		fedexacctno = trim(rs("FedexAccountNo")&"")
		fedexmeter = trim(rs("FedexMeter")&"")
		zipcode = trim(rs("adminZipCode")&"")
		countrycode = trim(rs("countryCode")&"")
	end if
	rs.Close
	versionarray = split(version, " v", 2)
	version = versionarray(1)
	versionarray = split(version, ".")
	if int(versionarray(0)<10) then version = "0" & versionarray(0) & versionarray(1) & "0" else version = versionarray(0) & versionarray(1) & "0"
	sXML = "<?xml version=""1.0"" encoding=""UTF-8""?>"
	sXML = sXML & "<FDXZipInquiryRequest xmlns:api=""http://www.fedex.com/fsmapi"" xsi:noNamespaceSchemaLocation=""FDXSubscriptionRequest.xsd"" xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"">"
	sXML = sXML & "<RequestHeader><CustomerTransactionIdentifier>ZipRequest</CustomerTransactionIdentifier>"
	sXML = sXML & "<AccountNumber>" & fedexacctno & "</AccountNumber><MeterNumber>" & fedexmeter & "</MeterNumber><CarrierCode></CarrierCode>"
	sXML = sXML & "</RequestHeader>"
	sXML = sXML & "<DestinationPostalCode>" & zipcode & "</DestinationPostalCode>"
	sXML = sXML & "<DestinationCountryCode>" & countrycode & "</DestinationCountryCode>"
	sXML = sXML & "</FDXZipInquiryRequest>"
	' response.write Replace(Replace(sXML,"</","&lt;/"),"<","<br />&lt;")&"<HR>"
	success = callxmlfunction("https://gateway.fedex.com:443/GatewayDC", sXML, xmlres, "", "Msxml2.ServerXMLHTTP", errormsg, FALSE)
	' response.write Replace(Replace(xmlres,"</","&lt;/"),"<","<br />&lt;")&"<HR>"
	set xmlDoc = Server.CreateObject("MSXML2.DOMDocument")
	xmlDoc.validateOnParse = False
	xmlDoc.loadXML (xmlres)
	Set t2 = xmlDoc.getElementsByTagName("DestinationLocationID").Item(0)
	locationid = t2.firstChild.nodeValue
	response.write "<script language=""javascript"" type=""text/javascript"">document.getElementById('step1span').innerHTML='<strong>Completed!</strong>';</script>"
	response.flush
	sXML = "<?xml version=""1.0"" encoding=""UTF-8""?>"
	sXML = sXML & "<FDXSSPVersionCaptureRequest xmlns:api=""http://www.fedex.com/fsmapi"" xsi:noNamespaceSchemaLocation=""FDXSubscriptionRequest.xsd"" xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"">"
	sXML = sXML & "<RequestHeader><CustomerTransactionIdentifier>VersionCapture</CustomerTransactionIdentifier>"
	sXML = sXML & "<AccountNumber>" & fedexacctno & "</AccountNumber><MeterNumber>" & fedexmeter & "</MeterNumber>"
	sXML = sXML & "</RequestHeader>"
	sXML = sXML & "<LocationID>" & locationid & "</LocationID>"
	sXML = sXML & "<VendorProductID>IBTP</VendorProductID>"
	sXML = sXML & "<VendorProductPlatform>ASP</VendorProductPlatform>"
	sXML = sXML & "<VendorProductVersion>" & version & "</VendorProductVersion>"
	sXML = sXML & "</FDXSSPVersionCaptureRequest>"
	' response.write Replace(Replace(sXML,"</","&lt;/"),"<","<br />&lt;")&"<HR>"
	success = callxmlfunction("https://gateway.fedex.com:443/GatewayDC", sXML, xmlres, "", "Msxml2.ServerXMLHTTP", errormsg, FALSE)
	' response.write Replace(Replace(xmlres,"</","&lt;/"),"<","<br />&lt;")&"<HR>"
	response.write "<script language=""javascript"" type=""text/javascript"">document.getElementById('step2span').innerHTML='<strong>Completed!</strong>';document.getElementById('donebutton').style.display='block';</script>"
elseif request.form("upsstep")="3" then
	sXML = "<?xml version=""1.0"" encoding=""UTF-8""?>"
	sXML = sXML & "<FDXSubscriptionRequest xmlns:api=""http://www.fedex.com/fsmapi"" xsi:noNamespaceSchemaLocation=""FDXSubscriptionRequest.xsd"" xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"">"
	sXML = sXML & "<RequestHeader><CustomerTransactionIdentifier>Subscribe</CustomerTransactionIdentifier>"
	sXML = sXML & "<AccountNumber>" & trim(Request.Form("fedexaccount")) & "</AccountNumber>"
	sXML = sXML & "</RequestHeader><Contact>"
	sXML = sXML & "<PersonName>" & Request.Form("contact") & "</PersonName>"
	if trim(Request.Form("company"))<>"" then sXML = sXML & "<CompanyName>" & Request.Form("company") & "</CompanyName>"
	if trim(Request.Form("department"))<>"" then sXML = sXML & "<Department>" & Request.Form("department") & "</Department>"
	sXML = sXML & "<PhoneNumber>" & Request.Form("telephone") & "</PhoneNumber>"
	if trim(Request.Form("pager"))<>"" then sXML = sXML & "<PagerNumber>" & Request.Form("pager") & "</PagerNumber>"
	if trim(Request.Form("fax"))<>"" then sXML = sXML & "<FaxNumber>" & Request.Form("fax") & "</FaxNumber>"
	if trim(Request.Form("email"))<>"" then sXML = sXML & "<E-MailAddress>" & Request.Form("email") & "</E-MailAddress>"
	sXML = sXML & "</Contact><Address><Line1>" & Request.Form("address") & "</Line1>"
	if trim(Request.Form("address2"))<>"" then sXML = sXML & "<Line2>" & Request.Form("address2") & "</Line2>"
	sXML = sXML & "<City>" & Request.Form("city") & "</City>"
	if Trim(Request.Form("country"))="US" OR Trim(Request.Form("country"))="CA" then
		sXML = sXML & "<StateOrProvinceCode>" & Request.Form("usstate") & "</StateOrProvinceCode>"
	else
		sXML = sXML & "<StateOrProvinceCode></StateOrProvinceCode>"
	end if
	sXML = sXML & "<PostalCode>" & Request.Form("postcode") & "</PostalCode>"
	sXML = sXML & "<CountryCode>" & Request.Form("country") & "</CountryCode></Address>"
	sXML = sXML & "<CSPSolutionType>100</CSPSolutionType><CSPIndicator>01</CSPIndicator></FDXSubscriptionRequest>"
	success = callxmlfunction("https://gateway.fedex.com:443/GatewayDC", sXML, xmlres, "", "Msxml2.ServerXMLHTTP", errormsg, FALSE)
	if success then
		success = ParseFedexOutput(xmlres, "FDXSubscriptionReply", fedexmeter, errormsg)
	end if
%>
	<form method="post" name="licform" action="admin.asp">
	  <input type="hidden" name="upsstep" value="5" />
      <table border="0" cellspacing="0" cellpadding="0" width="100%" bgcolor="" align="center">
        <tr> 
          <td width="100%">
            <table width="100%" border="0" cellspacing="0" cellpadding="2" bgcolor="">
			  <tr>
				<td rowspan="3" width="70" align="center" valign="top"><img src="../images/fedexsmall.gif" border="0" alt="FedEx" /><br />&nbsp;<br />&nbsp;<br />&nbsp;<br />&nbsp;</td>
                <td width="100%" align="center"><strong><%=yyFdxWiz%> - <% if success then response.write yyRegSucc else response.write yyError %></strong><br />&nbsp;
                </td>
			  </tr>
<%	if success then
		sSQL = "UPDATE admin SET FedexAccountNo='"&trim(Request.Form("fedexaccount"))&"',FedexMeter='"&fedexmeter&"'"
		cnn.Execute(sSQL)
		Application.Lock()
		Application("getadminsettings")=""
		Application.UnLock()
%>
			  <tr> 
                <td width="100%" align="left">
				  <p><strong><%=yyRegSucc%> !</strong></p>
				  <p>Thank you for registering to use FedEx&reg; Rates and Tracking.</p>
				  <p>To learn more about FedEx shipping services please visit <a href="http://www.fedex.com" target="_blank">www.fedex.com</a>.</p>
				  <p>To begin using FedEx shipping calculations please don't forget to select FedEx Shipping from the <strong>Shipping Type</strong> dropdown in the page <a href="adminmain.asp"><%=yyAdmMai%></a>.</p>
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
				  <p><font size="1">FedEx® is a registered service mark of Federal Express Corporation.
FedEx logos used by permission. All rights reserved.</font></p>
                </td>
			  </tr>
            </table>
          </td>
        </tr>
      </table>
	</form>
<%
elseif request.form("upsstep")="2" then
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
  if (theForm.postcode.value == ""){
	alert("<%=yyPlsEntr%> \"<%=yyPCode%>\".");
	theForm.postcode.focus();
	return (false);
  }
  if(!checkforamp(theForm.postcode)) return(false);
  if(theForm.telephone.value == ""){
    alert("<%=yyPlsEntr%> \"<%=yyTelep%>\".");
    theForm.telephone.focus();
    return (false);
  }
  if(theForm.telephone.value.length < 6 || theForm.telephone.value.length > 16){
    alert("<%=yyValTN%>");
    theForm.telephone.focus();
    return (false);
  }
  var checkOK = "0123456789";
  var checkStr = theForm.telephone.value;
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
  if(!allValid){
    alert("<%=yyOnDig%> \"<%=yyTelep%>\".");
    theForm.telephone.focus();
    return (false);
  }
  if(!checkforamp(theForm.fedexaccount)) return(false);
  if(theForm.fedexaccount.value == ""){
    alert("<%=yyPlsEntr%> \"Fedex Account Number\".");
    theForm.fedexaccount.focus();
    return (false);
  }
  var checkOK = "0123456789";
  var checkStr = theForm.fedexaccount.value;
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
  if(!allValid){
    alert("<%=yyOnDig%> \"Fedex Account Number\".");
    theForm.fedexaccount.focus();
    return (false);
  }
  return (true);
}
//-->
</script>
	<form method="post" name="licform" action="adminfedexlicense.asp" onsubmit="return formvalidator(this)">
	  <input type="hidden" name="upsstep" value="3" />
	  <input type="hidden" name="countryCode" value="<%=Request.Form("countryCode")%>" />
	  <input type="hidden" name="languageCode" value="<%=Request.Form("languageCode")%>" />
      <table border="0" cellspacing="0" cellpadding="0" width="100%" bgcolor="" align="center">
        <tr> 
          <td width="100%">
            <table width="100%" border="0" cellspacing="0" cellpadding="2" bgcolor="">
			  <tr>
				<td rowspan="18" width="70" align="center" valign="top"><img src="../images/fedexsmall.gif" border="0" alt="FedEx" /><br />&nbsp;<br />&nbsp;<br />&nbsp;<br />&nbsp;</td>
                <td width="100%" align="center" colspan="2"><strong><%=yyFdxWiz%></strong><br />&nbsp;
                </td>
			  </tr>
			  <tr> 
                <td width="40%" align="right"><strong><font color="#FF0000">*</font><%=yyConNam%> : </strong></td>
				<td width="60%"><input type="text" name="contact" size="30" /></td>
			  </tr>
			  <tr> 
                <td align="right"><strong><%=yyComNam%> : </strong></td>
				<td><input type="text" name="company" size="30" /></td>
			  </tr>
			  <tr> 
                <td align="right"><strong>Department : </strong></td>
				<td><input type="text" name="department" size="30" /></td>
			  </tr>
			  <tr> 
                <td align="right"><strong><font color="#FF0000">*</font><%=yyStrAdd%> : </strong></td>
				<td><input type="text" name="address" size="30" /></td>
			  </tr>
			  <tr> 
                <td align="right"><strong><%=yyAddr2%> : </strong></td>
				<td><input type="text" name="address2" size="30" /></td>
			  </tr>
			  <tr> 
                <td align="right"><strong><font color="#FF0000">*</font><%=yyCity%> : </strong></td>
				<td><input type="text" name="city" size="30" /></td>
			  </tr>
			  <tr> 
                <td align="right"><strong><font color="#FF0000">*</font><%=yyState%> <%=yyUSCan%> : </strong></td>
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
                <td align="right"><strong><font color="#FF0000">*</font><%=yyCountry%> : </strong></td>
				<td><select name="country" size="1">
<option value=''><%=yySelect%></option>
<%
' sSQL = "SELECT countryName,countryCode FROM countries WHERE countryCode IN ('AR','AU','AT','BE','BR','CA','CL','CN','CO','CR','DK','DO','FI','FR','DE','GR','GT','HK','IN','IE','IL','IT','JP','MY','MX','NL','NZ','NO','PA','PH','PT','PR','SG','KR','ES','SE','CH','TW','TH','GB','US') ORDER BY countryName"
sSQL = "SELECT countryName,countryCode FROM countries WHERE countryCode IN ('US') ORDER BY countryName"
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
                <td align="right"><strong><font color="#FF0000">*</font><%=yyPCode%> : </strong></td>
				<td><input type="text" name="postcode" size="15" /></td>
			  </tr>
			  <tr> 
                <td align="right"><strong><font color="#FF0000">*</font><%=yyTelep%> : </strong></td>
				<td><input type="text" name="telephone" size="30" /></td>
			  </tr>
			  <tr> 
                <td align="right"><strong>Pager Number : </strong></td>
				<td><input type="text" name="pager" size="30" /></td>
			  </tr>
			  <tr> 
                <td align="right"><strong>Fax Number : </strong></td>
				<td><input type="text" name="fax" size="30" /></td>
			  </tr>
			  <tr> 
                <td align="right"><strong><%=yyEmail%> : </strong></td>
				<td><input type="text" name="email" size="30" /></td>
			  </tr>
			  <tr> 
                <td align="right"><strong><font color="#FF0000">*</font>Fedex Account Number : </strong></td>
				<td><input type="text" name="fedexaccount" size="30" /></td>
			  </tr>
			  <tr>
                <td width="100%" align="center" colspan="2"><br />&nbsp;<input type="submit" name="agree" value="&nbsp;&nbsp;<%=yyNext%>&nbsp;&nbsp;" />&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
				<input type="button" value="<%=yyCancel%>" onclick="javascript:window.location='admin.asp';" />
                </td>
			  </tr>
			  <tr> 
                <td colspan="2" width="100%" align="center">
				  <p><img src="../images/clearpixel.gif" width="300" height="5" alt="" /></p>
				  <p><font size="1">FedEx® is a registered service mark of Federal Express Corporation.
FedEx logos used by permission. All rights reserved.</font></p>
                </td>
			  </tr>
            </table>
          </td>
        </tr>
      </table>
	</form>
<%
else %>
	<form method="post" action="adminfedexlicense.asp">
	  <input type="hidden" name="upsstep" value="2" />
      <table border="0" cellspacing="0" cellpadding="0" width="100%" bgcolor="" align="center">
        <tr> 
          <td width="100%">
            <table width="100%" border="0" cellspacing="0" cellpadding="2" bgcolor="">
			  <tr>
				<td rowspan="5" width="70" align="center" valign="top"><img src="../images/fedexsmall.gif" border="0" alt="" /><br />&nbsp;</td>
                <td width="100%" align="center"><strong><%=yyFdxWiz%></strong><br />&nbsp;
                </td>
			  </tr>
<%	isregistered=FALSE
	sSQL = "SELECT FedexAccountNo,FedexMeter FROM admin WHERE adminID=1"
	rs.Open sSQL,cnn,0,1
	if not rs.EOF then
		if trim(rs("FedexAccountNo")&"")<>"" AND trim(rs("FedexMeter")&"")<>"" then isregistered=true
	end if
	rs.Close
	if isregistered then %>
			  <tr> 
                <td width="100%">You have already successfully completed the FedEx licensing and registration wizard. If you would like to re-register then please 
				click the "Re-register" button below. If you would just like to update your Ecommerce Plus version information with 
				FedEx then please click the "Update Version" button below.
				<p>&nbsp;</p>
                </td>
			  </tr>
			  <tr> 
                <td width="100%" align="center"><input type="submit" name="agree" value="&nbsp;&nbsp;Re-Register&nbsp;&nbsp;" />&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
				<input type="button" value="Update Version" onclick="javascript:window.location='adminfedexlicense.asp?act=version';" />
                </td>
			  </tr>
<%	else %>
			  <tr> 
                <td width="100%"><ul><li>This wizard will assist you in completing the necessary licensing and registration requirements to activate and use the FedEx&reg; Rates and Tracking services from your Ecommerce Plus Template.<br /><br /></li>
				<li>If you do not wish to use any of the functions that utilize the FedEx Rates and Tracking services, click the Cancel button and those functions will not be enabled. If, at a later time, you wish to use these services, return to this section and complete the FedEx licensing and registration process.<br /><br /></li>
				<li>For more information about FedEx services, please <a href="http://www.fedex.com" target="_blank"><%=yyClkHer%></a>.<br /><br /></li>
				</ul>
				<p>&nbsp;</p>
                </td>
			  </tr>
			  <tr> 
                <td width="100%" align="center"><input type="submit" name="agree" value="&nbsp;&nbsp;<%=yyNext%>&nbsp;&nbsp;" />&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
				<input type="button" value="<%=yyCancel%>" onclick="javascript:window.location='admin.asp';" />
                </td>
			  </tr>
<%	end if %>
			  <tr> 
                <td align="center" colspan="2"><p><font size="1"><br />To open a FedEx account, please <a href="https://www.fedex.com/us/OADR/index.html?link=4" target="_blank"><strong><%=yyClkHer%></strong></a><br /></p></td>
			  </tr>
			  <tr> 
                <td width="100%" align="center">
				  <p><img src="../images/clearpixel.gif" width="300" height="5" alt="" /></p>
				  <p><font size="1">FedEx® is a registered service mark of Federal Express Corporation.
FedEx logos used by permission. All rights reserved.</font></p>
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