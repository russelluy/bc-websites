<%
sub sortshippingarray()
	for ssaindex2=0 to UBOUND(intShipping,2)
		intShipping(2,ssaindex2) = cDbl(intShipping(2,ssaindex2))
		for ssaindex=1 to UBOUND(intShipping,2)
			if intShipping(3,ssaindex) AND cDbl(intShipping(2,ssaindex))<cDbl(intShipping(2,ssaindex-1)) then
				tt0 = intShipping(0,ssaindex)
				tt1 = intShipping(1,ssaindex)
				tt2 = intShipping(2,ssaindex)
				tt3 = intShipping(3,ssaindex)
				tt4 = intShipping(4,ssaindex)
				intShipping(0,ssaindex) = intShipping(0,ssaindex-1)
				intShipping(1,ssaindex) = intShipping(1,ssaindex-1)
				intShipping(2,ssaindex) = intShipping(2,ssaindex-1)
				intShipping(3,ssaindex) = intShipping(3,ssaindex-1)
				intShipping(4,ssaindex) = intShipping(4,ssaindex-1)
				intShipping(0,ssaindex-1) = tt0
				intShipping(1,ssaindex-1) = tt1
				intShipping(2,ssaindex-1) = tt2
				intShipping(3,ssaindex-1) = tt3
				intShipping(4,ssaindex-1) = tt4
			end if
		next
	next
'	for ssaindex=0 to UBOUND(intShipping,2)
'		response.write intShipping(0,ssaindex) & ":" & intShipping(1,ssaindex) & ":" & intShipping(2,ssaindex) & ":" & intShipping(3,ssaindex) & ":" & intShipping(4,ssaindex) & "<br>"
'	next
end sub
Function ParseUSPSXMLOutput(sXML, international, byRef totalCost, byRef errormsg, byRef intShipping)
Dim noError, nodeList, packCost, xmlDoc, e, i, j, k, l, n, t, t2, s2
	noError = True
	totalCost = 0
	packCost = 0
	errormsg = ""
	gotxml=false
	on error resume next
	err.number=0
	set xmlDoc = Server.CreateObject("MSXML2.DOMDocument")
	if err.number=0 then gotxml=true
	if NOT gotxml then
		err.number=0
		set xmlDoc = Server.CreateObject("MSXML.DOMDocument")
		if err.number=0 then gotxml=true
	end if
	on error goto 0
	xmlDoc.validateOnParse = False
	xmlDoc.loadXML (sXML)
	If xmlDoc.documentElement.nodeName = "Error" then 'Top-level Error
		noError = False
		Set nodeList = xmlDoc.getElementsByTagName("Error")
		Set n = nodeList.Item(0)
		For i = 0 To n.childNodes.length - 1
			Set e = n.childNodes.Item(i)
			Select Case e.nodeName
				Case "Source"
				Case "Number"
				Case "Description"
					errormsg = e.firstChild.nodeValue
				Case "HelpFile"
				Case "HelpContext"
			End Select
		Next
	Else 'no Top-level Error
		Set nodeList = xmlDoc.getElementsByTagName("Package")
		For i = 0 To nodeList.length - 1
			Set n = nodeList.Item(i)
			tmpArr = Split(n.getAttribute("ID"),"x")
			quantity = Int(tmpArr(1))
			For j = 0 To n.childNodes.length - 1
				Set e = n.childNodes.Item(j)
				If e.nodeName = "Error" Then 'Lower-level error
					For k = 0 To e.childNodes.length - 1
						Set t = e.childNodes.Item(k)
						Select Case t.nodeName
							Case "Description"
								if debugmode=TRUE then response.write "USPS warning: " & t.firstChild.nodeValue & "<br>"
						End Select
					Next
				else
					Select Case e.nodeName
						Case "Postage"
							if international = "" then
								l = 0
								do while (intShipping(0, l) <> thisService AND intShipping(0, l) <> "")
									l = l + 1
								loop
								intShipping(0, l) = thisService
								if thisService="PARCEL" then
									intShipping(1, l) = "2-7 " & xxDays
								elseif thisService="EXPRESS" then
									intShipping(1, l) = "Overnight to most areas"
								elseif thisService="PRIORITY" then
									intShipping(1, l) = "1-2 " & xxDays
								elseif thisService="BPM" then
									intShipping(1, l) = "2-7 " & xxDays
								elseif thisService="Media" then
									intShipping(1, l) = "2-7 " & xxDays
								elseif thisService="FIRST CLASS" then
									intShipping(1, l) = "1-3 " & xxDays
								end if
								intShipping(2, l) = intShipping(2, l) + (e.firstChild.nodeValue * quantity)
								intShipping(3, l) = intShipping(3, l) + 1
							end if
						Case "Service"
							if international <> "" then
								Set t2 = e.getElementsByTagName("SvcDescription")
								Set s2 = t2.Item(0)
								l = 0
								do while (intShipping(0, l) <> s2.firstChild.nodeValue AND intShipping(0, l) <> "")
									l = l + 1
								loop
								intShipping(0, l) = s2.firstChild.nodeValue
								Set t2 = e.getElementsByTagName("SvcCommitments")
								Set s2 = t2.Item(0)
								intShipping(1, l) = s2.firstChild.nodeValue
								Set t2 = e.getElementsByTagName("Postage")
								Set s2 = t2.Item(0)
								intShipping(2, l) = intShipping(2, l) + (s2.firstChild.nodeValue * quantity)
								intShipping(3, l) = intShipping(3, l) + 1
							else
								thisService = e.firstChild.nodeValue
							end if
					End Select
				End If
			Next
			totalCost = totalCost + packCost
			packCost = 0
		Next
	End If
	set xmlDoc = nothing
	ParseUSPSXMLOutput = noError
end Function
Function checkUPSShippingMeth(method, byRef discountsApply, byRef showAs)
	retval = false
	for xx=0 to UBOUND(uspsmethods,2)
		if method=uspsmethods(0,xx) then
			retval=true
			discountsApply = uspsmethods(1,xx)
			showAs = uspsmethods(2,xx)
			exit for
		end if
	next
	checkUPSShippingMeth = retval
End Function
Function ParseUPSXMLOutput(xmlDoc, international, byRef totalCost, byRef errormsg, byRef errorcode, byRef intShipping)
Dim noError, nodeList, e, i, j, k, l, n, t, t2, indexus
	noError = True
	totalCost = 0
	indexus = 0
	l = 0
	errormsg = ""
	Set t2 = xmlDoc.getElementsByTagName("RatingServiceSelectionResponse").Item(0)
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
							Case "ErrorCode"
								errorcode = t.firstChild.nodeValue
							Case "ErrorSeverity"
								if t.firstChild.nodeValue="Transient" then errormsg = "This is a temporary error. Please wait a few moments then refresh this page.<br />" & errormsg
							Case "ErrorDescription"
								errormsg = errormsg & t.firstChild.nodeValue
						End Select
					Next
				end if
				' response.write "The Nodename is : " & e.nodeName & ":" & e.firstChild.nodeValue & "<br />"
			Next
		elseif n.nodename="RatedShipment" then
			wantthismethod=true
			For i = 0 To n.childNodes.length - 1
				Set e = n.childNodes.Item(i)
				Select Case e.nodeName
					Case "Service"
						For k = 0 To e.childNodes.length - 1
							Set t = e.childNodes.Item(k)
							if t.nodeName = "Code" then
								Select Case cStr(t.firstChild.nodeValue)
									Case "01"
										intShipping(0, l) = "UPS Next Day Air&reg;"
									Case "02"
										intShipping(0, l) = "UPS 2nd Day Air&reg;"
									Case "03"
										intShipping(0, l) = "UPS Ground"
									Case "07"
										intShipping(0, l) = "UPS Worldwide Express"
									Case "08"
										intShipping(0, l) = "UPS Worldwide Expedited"
									Case "11"
										intShipping(0, l) = "UPS Standard"
									Case "12"
										intShipping(0, l) = "UPS 3 Day Select&reg;"
									Case "13"
										intShipping(0, l) = "UPS Next Day Air Saver&reg;"
									Case "14"
										intShipping(0, l) = "UPS Next Day Air&reg; Early A.M.&reg;"
									Case "54"
										intShipping(0, l) = "UPS Worldwide Express Plus"
									Case "59"
										intShipping(0, l) = "UPS 2nd Day Air A.M.&reg;"
									Case "65"
										intShipping(0, l) = "UPS Express Saver"
								End Select
								wantthismethod = checkUPSShippingMeth(t.firstChild.nodeValue, discntsApp, notUsed)
								intShipping(4, l) = discntsApp
							end if
						Next
					Case "TotalCharges"
						For k = 0 To e.childNodes.length - 1
							Set t = e.childNodes.Item(k)
							if t.nodeName = "MonetaryValue" then intShipping(2, l) = cDbl(t.firstChild.nodeValue)
						Next
					Case "GuaranteedDaysToDelivery"
						if e.childNodes.length > 0 then
							if e.firstChild.nodeValue="1" then
								intShipping(1, l) = "1 " & xxDay & intShipping(1, l)
							else
								intShipping(1, l) = e.firstChild.nodeValue & " " & xxDays & intShipping(1, l)
							end if
						end if
					Case "ScheduledDeliveryTime"
						if e.childNodes.length > 0 then intShipping(1, l) = intShipping(1, l) & " by " & e.firstChild.nodeValue
				End select
			Next
			if wantthismethod=true then 
				intShipping(3, l) = true
				l = l + 1
			else
				intShipping(1, l) = ""
			end if
			wantthismethod=true
			' response.write "The RatedShipment is : " & n.nodeName & ":" & n.firstChild.nodeValue & "<br />"
		end if
	Next
	ParseUPSXMLOutput = noError
end Function
Function ParseCanadaPostXMLOutput(xmlDoc, international, byRef totalCost, byRef errormsg, byRef errorcode, byRef intShipping)
Dim noError, nodeList, e, i, j, k, l, n, t, t2, indexus
	noError = True
	totalCost = 0
	indexus = 0
	l = 0
	errormsg = ""
	Set t2 = xmlDoc.getElementsByTagName("eparcel").Item(0)
	for j = 0 to t2.childNodes.length - 1
		Set n = t2.childNodes.Item(j)
		if n.nodename="error" then
			noError = false
			For i = 0 To n.childNodes.length - 1
				Set e = n.childNodes.Item(i)
				if e.nodeName="statusMessage" then
					errormsg = errormsg & e.firstChild.nodeValue
				elseif e.nodeName="statusCode" then
					errorcode = e.firstChild.nodeValue
				end if
			Next
		elseif n.nodename="ratesAndServicesResponse" then
			For i = 0 To n.childNodes.length - 1
				Set e = n.childNodes.Item(i)
				if e.nodeName="product" then
					wantthismethod = checkUPSShippingMeth(e.getAttribute("id"), discntsApp, notUsed)
					intShipping(4, l) = discntsApp
					wantthismethod=true
					For k = 0 To e.childNodes.length - 1
						Set t = e.childNodes.Item(k)
						Select Case t.nodeName
							Case "name"
								intShipping(0, l) = t.firstChild.nodeValue
							Case "rate"
								intShipping(2, l) = cDbl(t.firstChild.nodeValue)
							Case "deliveryDate"
								if IsDate(t.firstChild.nodeValue) then
									numdays = DateValue(t.firstChild.nodeValue) - Date()
									intShipping(1, l) = numdays & " " & IIfVr(numdays<2,xxDay,xxDays) & intShipping(1, l)
								else
									intShipping(1, l) = t.firstChild.nodeValue & intShipping(1, l)
								end if
							Case "nextDayAM"
								if t.firstChild.nodeValue="true" then intShipping(1, l) = intShipping(1, l) & " AM"
						End select
					Next
					if wantthismethod=true then 
						intShipping(3, l) = true
						l = l + 1
					else
						intShipping(1, l) = ""
					end if
					wantthismethod=true
				end if
			next
		end if
	Next
	ParseCanadaPostXMLOutput = noError
end Function
function addUSPSDomestic(id,service,orig,dest,iWeight,quantity,container,size,machinable)
	Dim sXML
	sXML = ""
	pounds = Int(iWeight)
	ounces = round((iWeight-pounds)*16)
	if pounds=0 AND ounces=0 then ounces=1
	if IsArray(uspsmethods) then
		for indexus=0 TO UBOUND(uspsmethods,2)
			sXML = sXML & "<Package ID="""&uspsmethods(0,indexus)&id&"x"&quantity&""">"
			sXML = sXML & "<Service>"&uspsmethods(0,indexus)&"</Service>"
			sXML = sXML & "<ZipOrigination>"&orig&"</ZipOrigination><ZipDestination>"&left(dest, 5)&"</ZipDestination>"
			sXML = sXML & "<Pounds>"&pounds&"</Pounds><Ounces>"&ounces&"</Ounces>"
			sXML = sXML & "<Container>"&container&"</Container><Size>"&size&"</Size>"
			sXML = sXML & "<Machinable>"&machinable&"</Machinable></Package>"
		next
	end if
	addUSPSDomestic = sXML
end function
function addUSPSInternational(id,iWeight,quantity,mailtype,country)
	Dim sXML
	pounds = Int(iWeight)
	ounces = round((iWeight-pounds)*16)
	if pounds=0 AND ounces=0 then ounces=1
	sXML = "<Package ID="""&id&"x"&quantity&"""><Pounds>"&pounds&"</Pounds><Ounces>"&ounces&"</Ounces><MailType>"&mailtype&"</MailType><Country>"&country&"</Country>"
	addUSPSInternational = sXML & "</Package>"
end function
function addUPSInternational(iWeight,adminUnits,packTypeCode,country,packcost,dimens)
	Dim sXML
	if iWeight < 0.1 then iWeight=0.1
	sXML = "<Package><PackagingType><Code>"&packTypeCode&"</Code><Description>Package</Description></PackagingType>"
	if oversize<>0 then sXML = sXML & "<OversizePackage>" & oversize & "</OversizePackage>"
	oversize = 0
	if dimens(0) > 0 AND dimens(1) > 0 AND dimens(2) > 0 then sXML = sXML & "<Dimensions><Length>" & vsround(dimens(0),0) & "</Length><Width>" & vsround(dimens(1),0) & "</Width><Height>" & vsround(dimens(2),0) & "</Height><UnitOfMeasurement><Code>"&IIfVr((adminUnits AND 12)=4,"IN","CM")&"</Code></UnitOfMeasurement></Dimensions>"
	dimens(0)=0 : dimens(1)=0 : dimens(2)=0
	sXML = sXML & "<Description>Rate Shopping</Description><PackageWeight><UnitOfMeasurement><Code>"&IIfVr((adminUnits AND 1)=1,"LBS","KGS")&"</Code></UnitOfMeasurement><Weight>"&iWeight&"</Weight></PackageWeight><PackageServiceOptions>"
	if abs(addshippinginsurance)=1 OR (abs(addshippinginsurance)=2 AND wantinsurancepost="Y") then
		if packcost > 50000 then packcost=50000
		sXML = sXML & "<InsuredValue><CurrencyCode>" & countryCurrency & "</CurrencyCode><MonetaryValue>" & FormatNumber(packcost,2,-1,0,0) & "</MonetaryValue></InsuredValue>"
	end if
	if payproviderpost<>"" then
		if int(payproviderpost)=codpaymentprovider then sXML = sXML & "<COD><CODFundsCode>0</CODFundsCode><CODCode>3</CODCode><CODAmount><CurrencyCode>"&countryCurrency&"</CurrencyCode><MonetaryValue>" & FormatNumber(packcost,2,-1,0,0) & "</MonetaryValue></CODAmount></COD>"
	end if
	if signatureoption="indirect" then
		sXML = sXML & "<DeliveryConfirmation><DCISType>1</DCISType></DeliveryConfirmation>"
	elseif signatureoption="direct" then
		sXML = sXML & "<DeliveryConfirmation><DCISType>2</DCISType></DeliveryConfirmation>"
	elseif signatureoption="adult" then
		sXML = sXML & "<DeliveryConfirmation><DCISType>3</DCISType></DeliveryConfirmation>"
	end if
	addUPSInternational = sXML & "</PackageServiceOptions></Package>"
end function
function addCanadaPostPackage(iWeight,adminUnits,packTypeCode,country,packcost,dimens)
	if iWeight < 0.1 then iWeight=0.1
	if packtogether then thesize = 1 else thesize = 19
	if dimens(0)=0 then dimens(0) = thesize
	if dimens(1)=0 then dimens(1) = thesize
	if dimens(2)=0 then dimens(2) = thesize
	addCanadaPostPackage = "<item><quantity> 1 </quantity><weight> "&iWeight&" </weight><length> "&dimens(0)&" </length><width> "&dimens(1)&" </width><height> "&dimens(2)&" </height><description> Goods for shipping rates selection </description></item>"
end function
function addFedexPackage(iWeight,packages,packcost,dimens)
	Session.LCID = 1033
	tmpXML = ""
	if iWeight < 0.1 then iWeight=0.1
	if dimens(0) > 0 AND dimens(1) > 0 AND dimens(2) > 0 then tmpXML = "<Dimensions><Length>" & vsround(dimens(0),0) & "</Length><Width>" & vsround(dimens(1),0) & "</Width><Height>" & vsround(dimens(2),0) & "</Height><Units>"&IIfVr((adminUnits AND 12)=4,"IN","CM")&"</Units></Dimensions>"
	dimens(0)=0 : dimens(1)=0 : dimens(2)=0
	addFedexPackage = tmpXML & "<DeclaredValue>" & packcost & "</DeclaredValue><PackageCount>"&packages&"</PackageCount><Weight>"&formatnumber(iWeight,1,-1,0,0)&"</Weight>"
	Session.LCID = saveLCID
end function
function USPSCalculate(sXML,international,byRef totalCost, byRef errormsg, byRef intShipping)
	Dim objHttp, i
	if destZip="" then
		errormsg=xxPlsZip
		USPSCalculate=FALSE
	else
		set objHttp = Server.CreateObject("Msxml2.ServerXMLHTTP")
		objHttp.open "POST", "http://production.shippingapis.com/ShippingAPI.dll", false
		on error resume next
		err.number=0
		objHttp.Send "API="&international&"Rate&XML=" & Server.URLEncode(sXML)
		on error goto 0
		If err.number <> 0 OR objHttp.status <> 200 Then
			errormsg = "Error, couldn't connect to USPS server"
			USPSCalculate = false
		Else
			saveLCID = Session.LCID
			Session.LCID = 1033
			USPSCalculate = ParseUSPSXMLOutput(objHttp.responseText, international, totalCost, errormsg, intShipping)
			sortshippingarray()
			Session.LCID = saveLCID
		End If
		set objHttp = nothing
	end if
end function
function UPSCalculate(sXML,international,byRef totalCost, byRef errormsg, byRef intShipping)
	Dim objHttp, i
	if destZip="" then
		errormsg=xxPlsZip
		UPSCalculate=FALSE
	else
		set objHttp = Server.CreateObject("Msxml2.ServerXMLHTTP")
		objHttp.open "POST", "https://www.ups.com/ups.app/xml/Rate", false
		objHttp.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		on error resume next
		err.number=0
		objHttp.Send sXML
		on error goto 0
		If err.number <> 0 OR objHttp.status <> 200 Then
			errormsg = "Error, couldn't connect to UPS server"
			UPSCalculate = false
		Else
			saveLCID = Session.LCID
			Session.LCID = 1033
			UPSCalculate = ParseUPSXMLOutput(objHttp.responseXML, international, totalCost, errormsg, errorcode, intShipping)
			sortshippingarray()
			if errorcode = 111210 then errormsg = "The destination zip / postal code is invalid."
			Session.LCID = saveLCID
		End If
		set objHttp = nothing
	end if
end function
function CanadaPostCalculate(sXML,international,byRef totalCost, byRef errormsg, byRef intShipping)
	Dim objHttp, i
	if destZip="" then
		errormsg=xxPlsZip
		CanadaPostCalculate=FALSE
	else
		set objHttp = Server.CreateObject("Msxml2.ServerXMLHTTP")
		if canadaposttest = TRUE then
			objHttp.open "POST", "http://206.191.4.228:30000", false
		else
			objHttp.open "POST", "http://216.191.36.73:30000", false
		end if
		objHttp.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		on error resume next
		err.number=0
		objHttp.Send sXML
		on error goto 0
		If err.number <> 0 OR objHttp.status <> 200 Then
			errormsg = "Error, couldn't connect to CanadaPost server"
			CanadaPostCalculate = false
		Else
			saveLCID = Session.LCID
			Session.LCID = 1033
			' response.write Replace(Replace(objHttp.responseText,"</","&lt;/"),"<","<br />&lt;")&"<HR>"
			CanadaPostCalculate = ParseCanadaPostXMLOutput(objHttp.responseXML, international, totalCost, errormsg, errorcode, intShipping)
			sortshippingarray()
			Session.LCID = saveLCID
		End If
		set objHttp = nothing
	end if
end function
Function parsefedexXMLoutput(sXML, international, byRef errormsg, byRef errorcode, byRef intShipping)
	noError = True
	errormsg = ""
	l = 0
	set xmlDoc = Server.CreateObject("MSXML2.DOMDocument")
	xmlDoc.validateOnParse = False
	xmlDoc.loadXML (sXML)
	Set t2 = xmlDoc.getElementsByTagName("FDXRateAvailableServicesReply").Item(0)
	for j = 0 to t2.childNodes.length - 1
		Set n = t2.childNodes.Item(j)
		if n.nodename="Error" then
			noError = false
			For i = 0 To n.childNodes.length - 1
				Set e = n.childNodes.Item(i)
				if e.nodeName="Message" then
					errormsg = errormsg & e.firstChild.nodeValue
				elseif e.nodeName="Code" then
					errorcode = e.firstChild.nodeValue
				end if
			Next
		elseif n.nodename="Entry" then
			wantthismethod=FALSE
			For i = 0 To n.childNodes.length - 1
				Set e = n.childNodes.Item(i)
				if e.nodeName="Service" then
					wantthismethod = checkUPSShippingMeth(e.firstChild.nodeValue, discntsApp, showAs)
					if wantthismethod then
						intShipping(0, l) = showAs
						intShipping(4, l) = discntsApp
					end if
				elseif e.nodeName="EstimatedCharges" then
					For k9 = 0 To e.childNodes.length - 1
						Set f9 = e.childNodes.Item(k9)
						if f9.nodeName="DiscountedCharges" then
						intShipping(2, l) = 0
							For m = 0 To f9.childNodes.length - 1
								Set g9 = f9.childNodes.Item(m)
								if g9.nodeName="NetCharge" then
									intShipping(2, l) = intShipping(2, l) + cDbl(g9.firstChild.nodeValue)
								elseif g9.nodeName="TotalDiscount" then
									if uselistshippingrates=TRUE then intShipping(2, l) = intShipping(2, l) + cDbl(g9.firstChild.nodeValue)
								end if
							next
						end if
					next
				elseif e.nodeName="DeliveryDate" then
					numdays = DateValue(e.firstChild.nodeValue) - Date()
					if numdays < 1 then numdays = 1
					intShipping(1, l) = numdays & " " & IIfVr(numdays<2,xxDay,xxDays)
				end if
			next
			if wantthismethod then
				intShipping(3, l) = TRUE
				l = l + 1
			end if
		end if
	Next
	parsefedexXMLoutput = noError
end Function
function fedexcalculate(sXML,international, byRef errormsg, byRef intShipping)
	if destZip="" then
		errormsg=xxPlsZip
		fedexcalculate=FALSE
	else
		Session.LCID = 1033
		if payproviderpost<>"" then
			if int(payproviderpost)=codpaymentprovider then sXML = replace(sXML, "XXXFILLCODAMTHEREYYY", FormatNumber(totalgoods,2,-1,0,0), 1)
		end if
		' response.write Replace(Replace(sXML,"</","&lt;/"),"<","<br />&lt;")&"<HR>"
		success = callxmlfunction("https://gateway.fedex.com:443/GatewayDC", sXML, xmlres, "", "Msxml2.ServerXMLHTTP", errormsg, FALSE)
		' response.write Replace(Replace(xmlres,"</","&lt;/"),"<","<br />&lt;")&"<HR>"
		if success then
			success = parsefedexXMLoutput(xmlres, international, errormsg, errorcode, intShipping)
		end if
		if success then sortshippingarray()
		fedexcalculate = success
		Session.LCID = saveLCID
	end if
end function
%>
