<%
if digidownloadsecret="" then digidownloadsecret="this is some secret text"
Sub order_success(sorderid,sEmail,sendstoreemail)
	call do_order_success(sorderid,sEmail,sendstoreemail,TRUE,TRUE,TRUE,TRUE)
End sub
Sub do_order_success(sorderid,sEmail,sendstoreemail,doshowhtml,sendcustemail,sendaffilemail,sendmanufemail)
Dim custEmail,ordAddInfo,affilID,dropShippers()
Redim dropShippers(2,10)
if htmlemails=true then emlNl = "<br />" else emlNl=vbCrLf
affilID = ""
ordID = sorderid
hasdownload=FALSE
sSQL = "SELECT ordID,ordName,ordAddress,ordAddress2,ordCity,ordState,ordZip,ordCountry,ordEmail,ordPhone,ordShipName,ordShipAddress,ordShipAddress2,ordShipCity,ordShipState,ordShipZip,ordShipCountry,ordPayProvider,ordAuthNumber,ordTotal,ordDate,ordStateTax,ordCountryTax,ordHSTTax,ordHandling,ordShipping,ordAffiliate,ordShipType,ordDiscount,ordDiscountText,ordComLoc,ordExtra1,ordExtra2,ordExtra3,ordSessionID,payProvID,ordAddInfo FROM orders INNER JOIN payprovider ON payprovider.payProvID=orders.ordPayProvider WHERE ordAuthNumber<>'' AND ordID="&replace(sorderid,"'","")
rs.Open sSQL,cnn,0,1
if NOT rs.EOF then
	orderText = ""
	saveHeader = ""
	success=true
	ordAuthNumber = rs("ordAuthNumber")
	ordSessionID = rs("ordSessionID")
	payprovid = rs("payProvID")
	ordName = rs("ordName")
	if emailheader<>"" then saveHeader = emailheader
	execute("emailheader = emailheader" & payprovid)
	if emailheader<>"" then saveHeader = saveHeader & emailheader
	saveHeader = replace(saveHeader, "%ordername%", ordName)
	orderText = orderText & xxOrdId & ": " & rs("ordID") & emlNl
	if thereference<>"" then orderText = orderText & "Transaction Ref" & ": " & thereference & emlNl
	orderText = orderText & xxCusDet & ": " & emlNl
	if Trim(extraorderfield1)<>"" then orderText = orderText & extraorderfield1 & ": " & rs("ordExtra1") & emlNl
	orderText = orderText & ordName & emlNl
	orderText = orderText & rs("ordAddress") & emlNl
	if useaddressline2=TRUE AND trim(rs("ordAddress2"))<>""then orderText = orderText & rs("ordAddress2") & emlNl
	orderText = orderText & rs("ordCity") & ", " & rs("ordState") & emlNl
	orderText = orderText & rs("ordZip") & emlNl
	orderText = orderText & rs("ordCountry") & emlNl
	orderText = orderText & xxEmail & ": " & rs("ordEmail") & emlNl
	custEmail = rs("ordEmail")
	orderText = orderText & xxPhone & ": " & rs("ordPhone") & emlNl
	if Trim(extraorderfield2)<>"" then orderText = orderText & extraorderfield2 & ": " & rs("ordExtra2") & emlNl
	if Trim(rs("ordShipName")) <> "" OR Trim(rs("ordShipAddress")) <> "" then
		orderText = orderText & xxShpDet & ": " & emlNl
		if Trim(extraorderfield3)<>"" AND trim(rs("ordExtra3")&"")<>"" then orderText = orderText & extraorderfield3 & ": " & rs("ordExtra3") & emlNl
		orderText = orderText & rs("ordShipName") & emlNl
		orderText = orderText & rs("ordShipAddress") & emlNl
		if useaddressline2=TRUE AND trim(rs("ordShipAddress2"))<>"" then orderText = orderText & rs("ordShipAddress2") & emlNl
		orderText = orderText & rs("ordShipCity") & ", " & rs("ordShipState") & emlNl
		orderText = orderText & rs("ordShipZip") & emlNl
		orderText = orderText & rs("ordShipCountry") & emlNl
	end if
	ordShipType = rs("ordShipType")
	if ordShipType <> "" then
		orderText = orderText & emlNl & xxShpMet & ": " & ordShipType
		if (rs("ordComLoc") AND 2)=2 then orderText = orderText & xxWtIns
		orderText = orderText & emlNl
		if (rs("ordComLoc") AND 1)=1 then orderText = orderText & xxCerCLo & emlNl
		if (rs("ordComLoc") AND 4)=4 then orderText = orderText & xxSatDeR & emlNl
	end if
	ordAddInfo = Trim(rs("ordAddInfo"))
	if ordAddInfo <> "" then
		orderText = orderText & emlNl & xxAddInf & ": " & emlNl
		orderText = orderText & ordAddInfo & emlNl
	end if
	ordTotal = rs("ordTotal")
	ordDate = rs("ordDate")
	ordStateTax = rs("ordStateTax")
	ordDiscount = rs("ordDiscount")
	ordDiscountText = rs("ordDiscountText")
	ordCountryTax = rs("ordCountryTax")
	ordHSTTax = rs("ordHSTTax")
	ordShipping = rs("ordShipping")
	ordHandling = rs("ordHandling")
	affilID = Trim(rs("ordAffiliate"))
else
	orderText = "Cannot find customer details for order id " & sorderid & emlNl
end if
rs.Close
saveCustomerDetails=orderText
orderText = saveHeader & "%digidownloadplaceholder%" & orderText
sSQL = "SELECT cartProdId,cartProdName,cartProdPrice,cartQuantity,cartID,pDropship"&IIfVr(digidownloads=TRUE,",pDownload","")&" FROM cart INNER JOIN products ON cart.cartProdId=products.pID WHERE cartOrderID="&replace(sorderid,"'","")
rs.Open sSQL,cnn,0,1
if NOT rs.EOF then
	do while not rs.EOF
		localhasdownload=FALSE
		if digidownloads=TRUE then
			if trim(rs("pDownload")&"")<>"" then localhasdownload=TRUE
		end if
		saveCartItems = "--------------------------" & emlNl
		saveCartItems = saveCartItems & xxPrId & ": " & rs("cartProdId") & emlNl
		saveCartItems = saveCartItems & xxPrNm & ": " & rs("cartProdName") & emlNl
		saveCartItems = saveCartItems & xxQuant & ": " & rs("cartQuantity") & emlNl
		orderText = orderText & saveCartItems
		theoptions = ""
		theoptionspricediff=0
		sSQL = "SELECT coOptGroup,coCartOption,coPriceDiff,optRegExp FROM cartoptions INNER JOIN options ON cartoptions.coOptID=options.optID WHERE coCartID="&rs("cartID") & " ORDER BY coID"
		rs2.Open sSQL,cnn,0,1
		do while NOT rs2.EOF
			theoptionspricediff = theoptionspricediff + rs2("coPriceDiff")
			optionline = IIfVr(htmlemails=true,"&nbsp;&nbsp;&nbsp;&nbsp;>&nbsp;","> > > ") & rs2("coOptGroup") & " : " & replace(rs2("coCartOption")&"", vbCrLf, emlNl)
			theoptions = theoptions & optionline
			saveCartItems = saveCartItems & optionline & emlNl
			if rs2("coPriceDiff")=0 OR hideoptpricediffs=TRUE then
				theoptions = theoptions & emlNl
			else
				theoptions = theoptions & " ("
				if rs2("coPriceDiff") > 0 then theoptions = theoptions & "+"
				theoptions = theoptions & FormatEmailEuroCurrency(rs2("coPriceDiff")) & ")" & emlNl
			end if
			if rs2("optRegExp") = "!!" then localhasdownload=FALSE
			rs2.MoveNext
		loop
		rs2.Close
		orderText = orderText & xxUnitPr & ": " & IIfVr(hideoptpricediffs=TRUE,FormatEmailEuroCurrency(rs("cartProdPrice")+theoptionspricediff),FormatEmailEuroCurrency(rs("cartProdPrice"))) & emlNl
		orderText = orderText & theoptions
		if rs("pDropship")<>0 then
			index=0
			do while TRUE
				if index>=UBOUND(dropShippers,2) then Redim Preserve dropShippers(2,index+10)
				if dropShippers(0, index)="" OR dropShippers(0, index)=rs("pDropship") then exit do
				index=index+1
			loop
			dropShippers(0, index)=rs("pDropship")
			dropShippers(1, index)=dropShippers(1, index) & saveCartItems
		end if
		if localhasdownload=TRUE then hasdownload=TRUE
		rs.MoveNext
	loop
	orderText = orderText & "--------------------------" & emlNl

	orderText = orderText & xxOrdTot & " : " & FormatEmailEuroCurrency(ordTotal) & emlNl
	if combineshippinghandling=TRUE then
		orderText = orderText & xxShipHa & " : " & FormatEmailEuroCurrency(ordShipping + ordHandling) & emlNl
	else
		if shipType<>0 then orderText = orderText & xxShippg & " : " & FormatEmailEuroCurrency(ordShipping) & emlNl
		if cDbl(ordHandling)<>0.0 then orderText = orderText & xxHndlg & " : " & FormatEmailEuroCurrency(ordHandling) & emlNl
	end if
	if cDbl(ordDiscount)<>0.0 then orderText = orderText & xxDscnts & " : " & FormatEmailEuroCurrency(ordDiscount) & emlNl
	if cDbl(ordStateTax)<>0.0 then orderText = orderText & xxStaTax & " : " & FormatEmailEuroCurrency(ordStateTax) & emlNl
	if cDbl(ordCountryTax)<>0.0 then orderText = orderText & xxCntTax & " : " & FormatEmailEuroCurrency(ordCountryTax) & emlNl
	if cDbl(ordHSTTax)<>0.0 then orderText = orderText & xxHST & " : " & FormatEmailEuroCurrency(ordHSTTax) & emlNl
	ordGrandTotal = (ordTotal+ordStateTax+ordCountryTax+ordHSTTax+ordShipping+ordHandling)-ordDiscount
	orderText = orderText & xxGndTot & " : " & FormatEmailEuroCurrency(ordGrandTotal) & emlNl

	execute("emailheader = emailfooter" & payprovid)
	if emailheader<>"" then orderText = orderText & emailheader
	if emailfooter<>"" then orderText = orderText & emailfooter
else
	orderText = orderText & "Cannot find order details for order id " & sorderid & emlNl
end if
rs.Close
if hasdownload=TRUE AND digidownloademail<>"" then
	fingerprint = HMAC(digidownloadsecret, sorderid & ordAuthNumber & ordSessionID)
	fingerprint = Left(fingerprint, 14)
	digidownloademail = replace(digidownloademail,"%orderid%",ordID)
	digidownloademail = replace(digidownloademail,"%password%",fingerprint)
	digidownloademail = replace(digidownloademail,"%nl%",emlNl)
	orderEmailText = replace(orderText,"%digidownloadplaceholder%",digidownloademail)
else
	orderEmailText = replace(orderText,"%digidownloadplaceholder%","")
end if
orderText = replace(orderText,"%digidownloadplaceholder%","")
if sendstoreemail then
	Call DoSendEmailEO(sEmail,sEmail,"",replace(xxOrdStr, "%orderid%", sorderid),orderEmailText,emailObject,themailhost,theuser,thepass)
end if
' And one for the customer
if sendcustemail then
	Call DoSendEmailEO(custEmail,sEmail,"",replace(xxTnxOrd, "%ordername%", ordName),xxTouSoo & emlNl & emlNl & orderEmailText,emailObject,themailhost,theuser,thepass)
end if
' Drop Shippers / Manufacturers
if sendmanufemail then
	for index=0 to UBOUND(dropShippers,2)
		if dropShippers(0, index)="" then exit for
		if dropshipsubject="" then dropshipsubject="We have received the following order"
		sSQL = "SELECT dsEmail,dsAction FROM dropshipper WHERE dsID="&dropShippers(0, index)
		rs.Open sSQL,cnn,0,1
		if NOT rs.EOF then
			if (rs("dsAction") AND 1)=1 OR sendmanufemail=2 then
				saveHeader = ""
				saveFooter = ""
				if dropshipheader<>"" then saveHeader = dropshipheader
				execute("emailheader = dropshipheader" & payprovid)
				if emailheader<>"" then saveHeader = saveHeader & emailheader
				execute("saveFooter = dropshipfooter" & payprovid)
				if dropshipfooter<>"" then saveFooter = saveFooter & dropshipfooter
				Call DoSendEmailEO(Trim(rs("dsEmail")),sEmail,"",dropshipsubject,saveHeader & saveCustomerDetails & dropShippers(1, index) & saveFooter,emailObject,themailhost,theuser,thepass)
			end if
		end if
		rs.Close
	next
end if
if sendaffilemail then
	if affilID<>"" then
		sSQL = "SELECT affilEmail,affilInform FROM affiliates WHERE affilID='"&replace(affilID,"'","")&"'"
		rs.Open sSQL,cnn,0,1
		if NOT rs.EOF then
			if Int(rs("affilInform"))=1 then
				affiltext = xxAff1 & " "&FormatEmailEuroCurrency(ordTotal-ordDiscount)&"."&emlNl&emlNl&xxAff2&emlNl&emlNl&xxThnks&emlNl
				Call DoSendEmailEO(Trim(rs("affilEmail")),sEmail,"",xxAff3,affiltext,emailObject,themailhost,theuser,thepass)
			end if
		end if
		rs.Close
	end if
end if
if doshowhtml then
%>
<script language="javascript" type="text/javascript">
<!--
function doprintcontent()
{
	var prnttext = '<html><body>\n';
	prnttext += document.getElementById('printcontent').innerHTML;
	prnttext += '</body></html>';
	var newwin = window.open("","printit",'menubar=no, scrollbars=yes, width=500, height=400, directories=no,location=no,resizable=yes,status=no,toolbar=no');
	newwin.document.open();
	newwin.document.write(prnttext);
	newwin.document.close();
	newwin.print();
}
//-->
</script>
      <table border="0" cellspacing="<%=maintablespacing%>" cellpadding="<%=maintablepadding%>" width="<%=maintablewidth%>" bgcolor="<%=maintablebg%>" align="center">
        <tr>
          <td width="100%">
            <table width="<%=innertablewidth%>" border="0" cellspacing="<%=innertablespacing%>" cellpadding="<%=innertablepadding%>" bgcolor="<%=innertablebg%>">
			  <tr> 
                <td width="100%" align="center"><%=xxThkYou%>
                </td>
			  </tr>
	<%	if digidownloads<>true then %>
			  <tr> 
                <td width="100%" align="left">
				  <span name="printcontent" id="printcontent">
					<%=Replace(orderText,vbCrLf,"<br />")%>
				  </span>
                </td>
			  </tr>
			  <tr> 
                <td width="100%" align="center"><br />
				<% if xxRecEml<>"" then response.write xxRecEml&"<br /><br />"%>
				<input type="button" value="&nbsp;<%=xxCntShp%>&nbsp;" onclick="document.location='<%=storeUrl%>';" />
				<input type="button" value="&nbsp;<%=xxPrint%>&nbsp;" onclick="doprintcontent();" /><br />
				<img src="images/clearpixel.gif" width="300" height="3" alt="" />
                </td>
			  </tr>
	<%	end if %>
			</table>
		  </td>
        </tr>
      </table>
<%
end if
End sub
Sub DoSendEmail(seTo,seFrom,seSubject,seBody)
	Set rsSE = Server.CreateObject("ADODB.RecordSet")
	sSQL="SELECT emailObject,smtpserver,emailUser,emailPass FROM admin WHERE adminID=1"
	rsSE.Open sSQL,cnn,0,1
	emailObject = rsSE("emailObject")
	themailhost = Trim(rsSE("smtpserver")&"")
	theuser = Trim(rsSE("emailUser")&"")
	thepass = Trim(rsSE("emailPass")&"")
	rsSE.Close
	Call DoSendEmailEO(seTo,seFrom,"",seSubject,seBody,emailObject,themailhost,theuser,thepass)
	set rsSE = nothing
End Sub

Sub DoSendEmailEO(seTo,seFrom,seReplyTo,seSubject,seBody,emailObject,emailhost,username,password)
	seReplyTo = Trim(seReplyTo)
	on error resume next
	if emailObject=0 then
		Set EmailObj = Server.CreateObject("CDONTS.NewMail")
		EmailObj.MailFormat = 0
		if htmlemails=true then EmailObj.BodyFormat=0
		EmailObj.To = seTo
		EmailObj.From = seFrom
		if seReplyTo<>"" then EmailObj.Value("Reply-To") = seReplyTo
		EmailObj.Subject = seSubject
		EmailObj.Body = seBody
		EmailObj.Send
	elseif emailObject=1 then
		Set EmailObj = Server.CreateObject("CDO.Message")
		Set iConf = CreateObject("CDO.Configuration")
		if NOT (emailhost = "your.mailserver.com" OR emailhost = "") then
			Set Flds = iConf.Fields
			Flds.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
			Flds.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = emailhost
			if username<>"" AND password<>"" then
				Flds.Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
				Flds.Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = username
				Flds.Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = password
			end if
			Flds.Update
			EmailObj.Configuration = iConf
		else
			Set Flds = iConf.Fields
			if username<>"" AND password<>"" then
				Flds.Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
				Flds.Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = username
				Flds.Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = password
			end if
			Flds.Update
			EmailObj.Configuration = iConf
		end if
		EmailObj.From = Chr(34) & seFrom & Chr(34) & Chr(60) & seFrom & Chr(62)
		if seReplyTo<>"" then
			EmailObj.ReplyTo = Chr(34) & seReplyTo & Chr(34) & Chr(60) & seReplyTo & Chr(62)
		else
			EmailObj.ReplyTo = Chr(34) & seFrom & Chr(34) & Chr(60) & seFrom & Chr(62)
		end if
		EmailObj.Subject = seSubject
		EmailObj.Fields.Update
		if htmlemails=true then
			EmailObj.HTMLBody = seBody
			if emailencoding<> "iso-8859-1" then
				EmailObj.HTMLBodyPart.Charset = emailencoding
				EmailObj.TextBodyPart.Charset = emailencoding
				EmailObj.BodyPart.Charset = emailencoding
			end if
		else
			EmailObj.TextBody = seBody
			if emailencoding<> "iso-8859-1" then
				EmailObj.HTMLBodyPart.Charset = emailencoding
				EmailObj.TextBodyPart.Charset = emailencoding
				EmailObj.BodyPart.Charset = emailencoding
			end if
		end if
		EmailObj.To = Chr(34) & seTo & Chr(34) & " <" & seTo & ">"
		EmailObj.Send
	elseif emailObject=2 then
		Set EmailObj = Server.CreateObject("Persits.MailSender")
		if username<>"" AND password<>"" then
			EmailObj.Username = username
			EmailObj.Password = password
		end if
		EmailObj.Host = emailhost
		if htmlemails=true then EmailObj.IsHTML = true
		EmailObj.AddAddress seTo
		EmailObj.From = seFrom
		EmailObj.FromName = seFrom
		if seReplyTo<>"" then
			EmailObj.AddReplyTo seReplyTo,seReplyTo
		end if
		EmailObj.Subject = seSubject
		if emailencoding<> "iso-8859-1" then
			EmailObj.Charset = emailencoding
		end if
		EmailObj.Body = seBody
		if emailencoding<> "iso-8859-1" then
			EmailObj.ContentTransferEncoding = "Quoted-Printable"
		end if
		EmailObj.Send
	elseif emailObject=3 then
		Set EmailObj = Server.CreateObject("SMTPsvg.Mailer")
		if htmlemails=true then EmailObj.ContentType = "text/html"
		EmailObj.RemoteHost = emailhost
		EmailObj.AddRecipient seTo, seTo
		EmailObj.FromAddress = seFrom
		if seReplyTo<>"" then EmailObj.ReplyTo = seReplyTo
		EmailObj.Subject = seSubject
		EmailObj.BodyText = seBody
		EmailObj.SendMail
	elseif emailObject=4 then
		Set EmailObj = Server.CreateObject("JMail.SMTPMail")
		if htmlemails=true then EmailObj.ContentType = "text/html"
		EmailObj.silent = true
		EmailObj.Logging = true
		EmailObj.ServerAddress = emailhost
		EmailObj.AddRecipient seTo
		EmailObj.Sender = seFrom
		if seReplyTo<>"" then EmailObj.ReplyTo = seReplyTo
		EmailObj.Subject = seSubject
		EmailObj.Body = seBody
		EmailObj.Execute
	elseif emailObject=5 then
		Set EmailObj = Server.CreateObject("SoftArtisans.SMTPMail")
		if username<>"" AND password<>"" then
			EmailObj.UserName = username
			EmailObj.Password = password
		end if
		if htmlemails=true then EmailObj.ContentType = "text/html"
		EmailObj.RemoteHost = emailhost
		EmailObj.AddRecipient seTo , seTo
		EmailObj.FromAddress = seFrom
		if seReplyTo<>"" then EmailObj.ReplyTo = seReplyTo
		EmailObj.Subject = seSubject
		EmailObj.BodyText = seBody
		if NOT EmailObj.SendMail then Response.write "<br /> " & EmailObj.Response
	elseif emailObject=6 then
		Set EmailObj = Server.CreateObject("JMail.Message")
		if htmlemails=true then EmailObj.ContentType = "text/html"
		EmailObj.silent = true
		EmailObj.Logging = true
		EmailObj.AddRecipient seTo
		EmailObj.From = seFrom
		if seReplyTo<>"" then EmailObj.ReplyTo = seReplyTo
		EmailObj.Subject = seSubject
		if htmlemails=true then EmailObj.HTMLBody = seBody else EmailObj.Body = seBody
		EmailObj.Send(emailhost)
	end if
	Set EmailObj = nothing
	on error goto 0
End Sub
%>
