<%@LANGUAGE="VBScript"%>
<!--#include file="inc/md5.asp"-->
<!--#include file="db_conn_open.asp"-->
<!--#include file="includes.asp"-->
<!--#include file="inc/incemail.asp"-->
<!--#include file="inc/languagefile.asp"-->
<%
if debugmode then thetime1 = timer()
enableclientlogin=FALSE
if dateadjust="" then dateadjust=0
Dim cacheaddress(20,4) ' post code / country code / address id / actual results.
maxcacheid=0
%>
<!--#include file="inc/incfunctions.asp"-->
<%
Dim str,Txn_id,Payment_status,objHttp
sub writeresultstructure()
	responsexml2 = "<result"& IIfVr(noshipping, "", " shipping-name="""&shipmethods(0, gindex3)&"""")&" address-id="""&addressid&""">"
	if NOT noshipping then responsexml2 = responsexml2 & "<shipping-rate currency="""&countryCurrency&""">"&vsround((shipping+handling)-freeshipamnt,2)&"</shipping-rate>"
	responsexml2 = responsexml2 & "<shippable>true</shippable>"
	if numcpncodes > 0 then
		responsexml2 = responsexml2 & "<merchant-code-results>"
		for gindex5=0 to numcpncodes-1
			if cpncodes(gindex5)=cpncode then
				responsexml2 = responsexml2 & "<coupon-result><valid>"&IIfVr(gotcpncode, "true", "false")&"</valid>"
				if totaldiscounts>0 then responsexml2 = responsexml2 & "<calculated-amount currency="""&countryCurrency&""">"&vsround(appliedcouponamount,2)&"</calculated-amount>"
				responsexml2 = responsexml2 & "<code>"&cpncode&"</code>"
				if cpnmessage<>"" then
					responsexml2 = responsexml2 & "<message>"&replace(appliedcouponname,"<br />",vbCrLf)&"</message>"
				end if
				responsexml2 = responsexml2 & "</coupon-result>"
			else
				responsexml2 = responsexml2 & "<coupon-result><valid>false</valid>"
				responsexml2 = responsexml2 & "<code>"&cpncodes(gindex5)&"</code>"
				responsexml2 = responsexml2 & "<message>This coupon is not valid in conjunction with other coupons.</message>"
				responsexml2 = responsexml2 & "</coupon-result>"
			end if
		next
		responsexml2 = responsexml2 & "</merchant-code-results>"
	end if
	responsexml2 = responsexml2 & "<total-tax currency="""&countryCurrency&""">"&vsround(stateTax+countryTax, 2)&"</total-tax>"
	responsexml2 = responsexml2 & "</result>"
	responsexml = responsexml & responsexml2
	cacheaddress(maxcacheid,3) = cacheaddress(maxcacheid,3) & responsexml2
end sub
Sub release_stock(smOrdId)
	if stockManage <> 0 then
		sSQL="SELECT cartID,cartProdID,cartQuantity,pStockByOpts FROM cart INNER JOIN products ON cart.cartProdID=products.pID WHERE cartCompleted=1 AND cartOrderID=" & smOrdId
		rsl.Open sSQL,cnn,0,1
		do while NOT rsl.EOF
			if cint(rsl("pStockByOpts")) <> 0 then
				sSQL = "SELECT coOptID FROM cartoptions INNER JOIN (options INNER JOIN optiongroup ON options.optGroup=optiongroup.optGrpID) ON cartoptions.coOptID=options.optID WHERE optType IN (-2,-1,1,2) AND coCartID=" & rsl("cartID")
				rs.Open sSQL,cnn,0,1
				do while NOT rs.EOF
					sSQL = "UPDATE options SET optStock=optStock+"&rsl("cartQuantity")&" WHERE optID="&rs("coOptID")
					cnn.Execute(sSQL)
					rs.MoveNext
				loop
				rs.Close
			else
				sSQL = "UPDATE products SET pInStock=pInStock+"&rsl("cartQuantity")&" WHERE pID='"&rsl("cartProdID")&"'"
				cnn.Execute(sSQL)
			end if
			rsl.MoveNext
		loop
		rsl.Close
	end if
End Sub
	'**************************************
    ' Name: ANSI to Unicode
    ' Description:Converts from ANSI to Unic
    '     ode very fast. Inspired by code found in
    '     UltraFastAspUpload by Cakkie (on PSC). T
    '     his should work slightly faster then Cak
    '     kies due to how some of the code has bee
    '     n arranged.
    ' By: Lewis E. Moten III
    '
    ' This code is copyrighted and has
    ' limited warranties. Please see http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=7266&lngWId=4
	' for details.
    '**************************************
	function ANSIToUnicode(ByRef pbinBinaryData)
    	Dim lbinData	' Binary Data (ANSI)
    	Dim llngLength	' Length of binary data (byte count)
    	Dim lobjRs		' Recordset
    	Dim lstrData	' Unicode Data
    	' VarType Reference
    	'8 = Integer (this is expected var type)
    	'17 = Byte Subtype
    	' 8192 = Array
    	' 8209 = Byte Subtype + Array
    	Set lobjRs = Server.CreateObject("ADODB.Recordset")
    	if VarType(pbinBinaryData) = 8 Then
    		' Convert integers(4 bytes) To Byte Subtype Array (1 byte)
    		llngLength = LenB(pbinBinaryData)
    		if llngLength = 0 Then
    			lbinData = ChrB(0)
    		Else
    			Call lobjRs.Fields.Append("BinaryData", adLongVarBinary, llngLength)
    			Call lobjRs.Open()
    			Call lobjRs.AddNew()
    			Call lobjRs.Fields("BinaryData").AppendChunk(pbinBinaryData & ChrB(0)) ' + Null terminator
    			Call lobjRs.Update()
    			lbinData = lobjRs.Fields("BinaryData").GetChunk(llngLength)
    			Call lobjRs.Close()
    		End if
    	Else
    		lbinData = pbinBinaryData
    	End if
    	' Do REAL conversion now!	
    	llngLength = LenB(lbinData)
    	if llngLength = 0 Then
    		lstrData = ""
    	Else
    		Call lobjRs.Fields.Append("BinaryData", 201, llngLength)
    		Call lobjRs.Open()
    		Call lobjRs.AddNew()
    		Call lobjRs.Fields("BinaryData").AppendChunk(lbinData)
    		Call lobjRs.Update()
    		lstrData = lobjRs.Fields("BinaryData").Value
    		Call lobjRs.Close()
    	End if
    				
    	Set lobjRs = Nothing
    	ANSIToUnicode = lstrData
    End function
	Const sBASE_64_CHARACTERS = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"
    ' --------------------------------------
    '     ---------------------------------------
    function Base64decode(ByVal asContents)
    	Dim lsResult
    	Dim lnPosition
    	Dim lsGroup64, lsGroupBinary
    	Dim Char1, Char2, Char3, Char4
    	Dim Byte1, Byte2, Byte3
    	'if Len(asContents) Mod 4 > 0 Then asContents = asContents & String(4 - (Len(asContents) Mod 4), " ")
    	lsResult = ""
    	
    	For lnPosition = 1 To Len(asContents) Step 4
    		lsGroupBinary = ""
    		lsGroup64 = Mid(asContents, lnPosition, 4)
    		Char1 = INSTR(sBASE_64_CHARACTERS, Mid(lsGroup64, 1, 1)) - 1
    		Char2 = INSTR(sBASE_64_CHARACTERS, Mid(lsGroup64, 2, 1)) - 1
    		Char3 = INSTR(sBASE_64_CHARACTERS, Mid(lsGroup64, 3, 1)) - 1
    		Char4 = INSTR(sBASE_64_CHARACTERS, Mid(lsGroup64, 4, 1)) - 1
    		Byte1 = Chr(((Char2 And 48) \ 16) Or (Char1 * 4) And &HFF)
    		if char2>=0 AND char3>=0 then Byte2 = lsGroupBinary & Chr(((Char3 And 60) \ 4) Or (Char2 * 16) And &HFF) else Byte2=""
    		if char3>=0 AND char4>=0 then Byte3 = Chr((((Char3 And 3) * 64) And &HFF) Or (Char4 And 63)) else Byte3=""
    		lsGroupBinary = Byte1 & Byte2 & Byte3
    		
    		lsResult = lsResult + lsGroupBinary
    	Next
    	Base64decode = lsResult
    End function
Set rs = Server.CreateObject("ADODB.RecordSet")
Set rs2 = Server.CreateObject("ADODB.RecordSet")
Set cnn=Server.CreateObject("ADODB.Connection")
b64pad="="
cnn.open sDSN
alreadygotadmin = getadminsettings()
success = getpayprovdetails(20,googledata1,googledata2,googledata3,googledemomode,ppmethod)
Dim biData
biData = Request.BinaryRead(Request.TotalBytes)
xmlResponse = ANSIToUnicode(biData)

xmlResponse2 = "<?xml version=""1.0"" encoding=""UTF-8""?>" & _
"<merchant-calculation-callback xmlns=""http://checkout.google.com/schema/2"" serial-number=""9d761c2b-2b81-42df-8716-74432545cbb8"">" & _
"  <shopping-cart>" & _
"    <merchant-private-data>" & _
"      <privateitems>" & _
"        <sessionid>158938536</sessionid>" & _
"        <partner />" & _
"        <clientuser />" & _
"      </privateitems>" & _
"    </merchant-private-data>" & _
"    <items>" & _
"      <item>" & _
"        <item-name>Digital Camera</item-name>" & _
"        <item-description>GreatBrand™ Digital Camera</item-description>" & _
"        <quantity>1</quantity>" & _
"        <unit-price currency=""USD"">400.0</unit-price>" & _
"        <merchant-private-item-data>" & _
"          <product-id>digitalcamera</product-id>" & _
"        </merchant-private-item-data>" & _
"      </item>" & _
"    </items>" & _
"  </shopping-cart>" & _
"  <buyer-id>752681359912242</buyer-id>" & _
"  <calculate>" & _
"    <shipping>" & _
"      <method name=""FedEx Ground"" />" & _
"      <method name=""FedEx 1Day Freight"" />" & _
"      <method name=""FedEx 2Day Freight"" />" & _
"      <method name=""FedEx 2Day"" />" & _
"      <method name=""FedEx 3Day Freight"" />" & _
"      <method name=""FedEx Europe First - Int'l Priority"" />" & _
"      <method name=""FedEx Express Saver"" />" & _
"      <method name=""FedEx First Overnight"" />" & _
"      <method name=""FedEx Home Delivery"" />" & _
"      <method name=""FedEx International Economy Freight"" />" & _
"      <method name=""FedEx International Economy"" />" & _
"      <method name=""FedEx International Next Flight "" />" & _
"      <method name=""FedEx International Priority Freight"" />" & _
"      <method name=""FedEx International Priority "" />" & _
"      <method name=""FedEx Priority Overnight"" />" & _
"      <method name=""FedEx Standard Overnight"" />" & _
"    </shipping>" & _
"    <addresses>" & _
"      <anonymous-address id=""228145593116372"">" & _
"        <country-code>US</country-code>" & _
"        <city>San Jose</city>" & _
"        <region>CA</region>" & _
"        <postal-code>95129</postal-code>" & _
"      </anonymous-address>" & _
"      <anonymous-address id=""639235019263105"">" & _
"        <country-code>US</country-code>" & _
"        <city>JACKSON</city>" & _
"        <region>MS</region>" & _
"        <postal-code>39201</postal-code>" & _
"      </anonymous-address>" & _
"      <anonymous-address id=""770975562919763"">" & _
"        <country-code>US</country-code>" & _
"        <city>JACKSON</city>" & _
"        <region>MS</region>" & _
"        <postal-code>39201</postal-code>" & _
"      </anonymous-address>" & _
"    </addresses>" & _
"    <merchant-code-strings />" & _
"    <tax>true</tax>" & _
"  </calculate>" & _
"  <buyer-language>en_US</buyer-language>" & _
"</merchant-calculation-callback>"

standalonetestmode=false

if standalonetestmode then xmlResponse=xmlResponse2

if disablebasicauth=TRUE then
	' Do Nothing
elseif success then
	http_auth = request.servervariables("HTTP_AUTHORIZATION")
	if http_auth="" then http_auth = request.servervariables("HTTP_AUTHENTICATION")
	if left(http_auth, 6)="Basic " then
		http_auth = right(http_auth, len(http_auth)-6)
		http_auth = Base64decode(http_auth)
		if InStr(http_auth, ":")=0 then
			success=FALSE
		else
			auth_split = split(http_auth,":")
			if googledata1<>auth_split(0) OR googledata2<>auth_split(1) then success=FALSE
		end if
	else
		success=FALSE
	end if
end if
if standalonetestmode then success=TRUE
if NOT success then
	response.clear
	response.status = "401 Unauthorized"
	response.write "<html><head><title>401 Unauthorized</title><body>"
	response.write "I'm sorry, you are not authorized to view this page.<br>"
	response.write "</body></html>"
else
	set gcXmlDoc = Server.CreateObject("MSXML2.DOMDocument")
	gcXmlDoc.validateOnParse = False
	gcXmlDoc.loadXML (xmlResponse)
	thismessage = gcXmlDoc.documentElement.tagName
	Select Case thismessage
	Case "merchant-calculation-callback"
		Dim shipmethods(), cpncodes()
		cartisincluded=TRUE
		cpncode=""
		ordPayProvider=20
		' response.clear
		if standalonetestmode then response.write "<html><body>"
		responsexml = "<?xml version=""1.0"" encoding=""UTF-8""?>"
		responsexml = responsexml & "<merchant-calculation-results xmlns=""http://checkout.google.com/schema/2"">"
		responsexml = responsexml & "<results>"
		Set obj1 = gcXmlDoc.getElementsByTagName("sessionid").Item(0)
		thesessionid = replace(obj1.firstChild.nodeValue,"'","")
		Set obj1 = gcXmlDoc.getElementsByTagName("clientuser").Item(0)
		if obj1.hasChildNodes then clientuser=obj1.firstChild.nodeValue else clientuser=""
		Session("clientUser")=""
		if clientuser<>"" then
			sSQL = "SELECT clientUser,clientActions,clientLoginLevel,clientPercentDiscount FROM clientlogin WHERE clientUser='"&replace(clientuser,"'","")&"'"
			rs.Open sSQL,cnn,0,1
			if NOT rs.EOF then
				Session("clientUser")=rs("clientUser")
				Session("clientActions")=rs("clientActions")
				Session("clientLoginLevel")=rs("clientLoginLevel")
				Session("clientPercentDiscount")=(100.0-cDbl(rs("clientPercentDiscount")))/100.0
			end if
			rs.Close
		end if
%><!--#include file="inc/inccart.asp"--><%
		for gindex=0 to gcXmlDoc.documentElement.childNodes.length-1
			set obj1 = gcXmlDoc.documentElement.childNodes.Item(gindex)
			select case obj1.nodeName
			case "calculate"
				redim shipmethods(2,10)
				redim cpncodes(10)
				numshipmethods=0
				numcpncodes=0
				usestateabbrev=TRUE
				savehandling=handling
				cpnmessage = "<br />"
				set obj2 = obj1.getElementsByTagName("merchant-code-strings")
				if obj2.length > 0 then
					set obj3 = obj2.item(0).childNodes
					for gindex3=0 to obj3.length-1
						if cpncode="" then cpncode = obj3.Item(gindex3).getAttribute("code") ' Because they arrive in NON reverse order
						cpncodes(numcpncodes) = obj3.Item(gindex3).getAttribute("code")
						numcpncodes=numcpncodes+1
					next
				end if
				set obj2 = obj1.getElementsByTagName("shipping").item(0)
				for gindex2=0 to obj2.childNodes.length-1
					Set obj3 = obj2.childNodes.Item(gindex2)
					if obj3.nodeName="method" then
						shipMethod = obj3.getAttribute("name")
						shipmethods(0, numshipmethods)=shipMethod
						shipmethods(1, numshipmethods)=FALSE
						numshipmethods=numshipmethods+1
						if numshipmethods >= UBOUND(shipmethods) then redim preserve shipmethods(2, UBOUND(shipmethods, 2) + 10)
					end if
				next
				set obj2 = obj1.getElementsByTagName("addresses").item(0)
				for gindex2=0 to obj2.childNodes.length-1
					stateTaxRate=0
					Set obj3 = obj2.childNodes.Item(gindex2)
					if obj3.nodeName="anonymous-address" then
						numshipoptions=0
						totShipOptions=0
						freeshippingapplied=FALSE
						noshipping=(shipType=0)
						totaldiscounts=0
						gotcpncode=FALSE
						cpnmessage = "<br />"
						iTotItems = 0
						destinationsupported=TRUE
						addressid = obj3.getAttribute("id")
						for gindex3=0 to obj3.childNodes.length-1
							Set obj4 = obj3.childNodes.Item(gindex3)
							select case obj4.nodeName
							case "country-code"
								shipCountryCode=obj4.firstChild.nodeValue
							case "region"
								shipstate=obj4.firstChild.nodeValue
							case "postal-code"
								destZip=obj4.firstChild.nodeValue
							end select
						next
						' Firstly check in the cache
						foundincache=-1
						for gindex3=0 to maxcacheid-1
							if cacheaddress(gindex3,0)=destZip AND cacheaddress(gindex3,1)=shipCountryCode then foundincache=gindex3
						next
						if foundincache >= 0 then
							responsexml = responsexml & replace(cacheaddress(foundincache,3), cacheaddress(foundincache,2), addressid)
						else
							cacheaddress(maxcacheid,0) = destZip
							cacheaddress(maxcacheid,1) = shipCountryCode
							cacheaddress(maxcacheid,2) = addressid
							cacheaddress(maxcacheid,3) = ""
							sSQL = "SELECT countryID,countryName,countryTax,countryCode,countryFreeShip,countryOrder,countryEnabled FROM countries WHERE countryCode='"&shipCountryCode&"'"
							rs.Open sSQL,cnn,0,1
							if NOT rs.EOF then
								if trim(Session("clientUser")) <> "" AND (Session("clientActions") AND 2)=2 then countryTaxRate=0 else countryTaxRate = rs("countryTax")
								shipCountryID = rs("countryID")
								shipCountryCode = rs("countryCode")
								freeshipapplies = (rs("countryFreeShip")=1)
								shiphomecountry = (rs("countryOrder")=2)
								shipcountry = rs("countryName")
								if rs("countryEnabled")=0 then destinationsupported=FALSE
							end if
							rs.Close
							if shiphomecountry then
								sSQL = "SELECT stateTax,stateAbbrev,stateFreeShip,stateEnabled FROM states WHERE stateAbbrev='"&replace(shipstate,"'","''")&"'"
								rs.Open sSQL,cnn,0,1
								if NOT rs.EOF then
									stateTaxRate=rs("stateTax")
									shipStateAbbrev=rs("stateAbbrev")
									freeshipapplies=(freeshipapplies AND (rs("stateFreeShip")=1))
									if rs("stateEnabled")=0 then destinationsupported=FALSE
								end if
								rs.Close
							end if
							if NOT destinationsupported then
								for gindex3=0 to numshipmethods-1
									if shipmethods(1, gindex3)<>TRUE then
										responsexml = responsexml & "<result"& IIfVr(noshipping, "", " shipping-name="""&shipmethods(0, gindex3)&"""")&" address-id="""&addressid&"""><shipping-rate currency="""&countryCurrency&""">0.00</shipping-rate><shippable>false</shippable><total-tax currency="""&countryCurrency&""">0.00</total-tax></result>"
									end if
								next
							else
								initshippingmethods()
								totalgoods=0
								alldata=""
								if mysqlserver=true then
									sSQL = "SELECT cartID,cartProdID,cartProdName,cartProdPrice,cartQuantity,pWeight,pShipping,pShipping2,pExemptions,pSection,topSection,pDims,pTax,"&getlangid("pDescription",2)&" FROM cart INNER JOIN products ON cart.cartProdID=products.pID LEFT OUTER JOIN sections ON products.pSection=sections.sectionID WHERE cartCompleted=0 AND cartSessionID="&thesessionid
								else
									sSQL = "SELECT cartID,cartProdID,cartProdName,cartProdPrice,cartQuantity,pWeight,pShipping,pShipping2,pExemptions,pSection,topSection,pDims,pTax,"&getlangid("pDescription",2)&" FROM cart INNER JOIN (products LEFT OUTER JOIN sections ON products.pSection=sections.sectionID) ON cart.cartProdID=products.pID WHERE cartCompleted=0 AND cartSessionID="&thesessionid
								end if
								if standalonetestmode then response.write sSQL & "<br>"
								rs.Open sSQL,cnn,0,1
								if NOT (rs.EOF OR rs.BOF) then alldata=rs.getrows
								rs.Close
								if isarray(alldata) then
									for index=0 to UBOUND(alldata,2)
										sSQL = "SELECT SUM(coPriceDiff) AS coPrDff FROM cartoptions WHERE coCartID="&alldata(0,index)
										rs.Open sSQL,cnn,0,1
										if NOT rs.EOF then
											if NOT IsNull(rs("coPrDff")) then alldata(3,index)=cDbl(alldata(3,index))+cDbl(rs("coPrDff"))
										end if
										rs.Close
										sSQL = "SELECT SUM(coWeightDiff) AS coWghtDff FROM cartoptions WHERE coCartID="&alldata(0,index)
										rs.Open sSQL,cnn,0,1
										if NOT rs.EOF then
											if NOT IsNull(rs("coWghtDff")) then alldata(5,index)=cDbl(alldata(5,index))+cDbl(rs("coWghtDff"))
										end if
										rs.Close
										runTot=(alldata(3,index)*Int(alldata(4,index)))
										totalquantity = totalquantity + alldata(4,index)
										totalgoods=totalgoods+runTot
										thistopcat=0
										if trim(Session("clientUser"))<>"" then alldata(8,index) = (alldata(8,index) OR Session("clientActions"))
										if (shipType=2 OR shipType=3 OR shipType=4 OR shipType=6 OR shipType=7) AND cDbl(alldata(5,index))<=0.0 then alldata(8,index) = (alldata(8,index) OR 4)
										if (alldata(8,index) AND 1)=1 then statetaxfree = statetaxfree + runTot
										if perproducttaxrate=TRUE then
											if isnull(alldata(12,index)) then alldata(12,index)=countryTaxRate
											if (alldata(8,index) AND 2)<>2 then countryTax = countryTax + ((alldata(12,index) * runTot) / 100.0)
										else
											if (alldata(8,index) AND 2)=2 then countrytaxfree = countrytaxfree + runTot
										end if
										if (alldata(8,index) AND 4)=4 then shipfreegoods = shipfreegoods + runTot
										call addproducttoshipping(alldata, index)
									next
								else
									errormsg = "Error, couldn't find cart"
									success = FALSE
								end if
								call calculatediscounts(totalgoods, false, cpncode)
								if totaldiscounts > totalgoods then totaldiscounts = totalgoods
								if success AND calculateshipping() then
									freeshipamnt=0
									insuranceandtaxaddedtoshipping()
									calculateshippingdiscounts(false)
									freeshipamnt=0
									cpnmessage = Right(cpnmessage,Len(cpnmessage)-6)
									if numshipmethods=0 then
										noshipping=TRUE
										handling=savehandling
										shipping=0
										calculatetaxandhandling()
										writeresultstructure()
									else
										if shipType=1 then
											for gindex3=0 to numshipmethods-1
												if xmlencodecharref(xxShipHa) = shipmethods(0, gindex3) then
													handling=savehandling
													if freeshippingapplied then shipping=0
													freeshipamnt=0
													calculatetaxandhandling()
													writeresultstructure()
													shipmethods(1, gindex3)=TRUE
												end if
											next
										elseif (shipType=2 OR shipType=3 OR shipType=4 OR shipType=5 OR shipType=6 OR shipType=7) then
											' response.write "numshipmethods: " & numshipmethods & ", numshipoptions: " & numshipoptions & "<br>"
											if shipType=2 OR shipType=5 then totShipOptions=numshipoptions+1 else totShipOptions=UBOUND(intShipping,2)
											for gindex4=0 to totShipOptions-1
												for gindex3=0 to numshipmethods-1
													' response.write "matching: " & intShipping(0,gindex4) & " : " & shipmethods(gindex3) & "<br>"
													if shipType=3 then
														if iTotItems=intShipping(3,gindex4) then
															if xmlencodecharref(intShipping(5,gindex4)&"") = shipmethods(0, gindex3) then
																handling=savehandling
																if freeshippingapplied AND intShipping(4,gindex4) <> 0 then shipping=0 else shipping=intShipping(2,gindex4)
																calculatetaxandhandling()
																writeresultstructure()
																shipmethods(1, gindex3)=TRUE
															end if
														end if
													elseif shipType=4 OR shipType=6 OR shipType=7 then
														if intShipping(3,gindex4)=TRUE then
															if xmlencodecharref(intShipping(0,gindex4)&"") = shipmethods(0, gindex3) then
																handling=savehandling
																if freeshippingapplied AND intShipping(4,gindex4) <> 0 then shipping=0 else shipping=intShipping(2,gindex4)
																calculatetaxandhandling()
																writeresultstructure()
																shipmethods(1, gindex3)=TRUE
															end if
														end if
													else
														if xmlencodecharref(intShipping(0,gindex4)&"") = shipmethods(0, gindex3) then
															handling=savehandling
															if freeshippingapplied AND intShipping(4,gindex4) <> 0 then shipping=0 else shipping=intShipping(2,gindex4)
															calculatetaxandhandling()
															writeresultstructure()
															shipmethods(1, gindex3)=TRUE
														end if
													end if
												next
											next
										elseif shipType=0 then
											handling=savehandling
											shipping=0
											calculatetaxandhandling()
											writeresultstructure()
										end if
										if willpickuptext<>"" then
											noshipping=FALSE
											for gindex3=0 to numshipmethods-1
												if xmlencodecharref(willpickuptext) = shipmethods(0, gindex3) then
													if willpickupcost="" then shipping=0 else shipping=willpickupcost
													handling=savehandling
													freeshipamnt=0
													calculatetaxandhandling()
													writeresultstructure()
													shipmethods(1, gindex3)=TRUE
												end if
											next
										end if
										for gindex3=0 to numshipmethods-1
											if shipmethods(1, gindex3)<>TRUE then
												responsexml2 = "<result"& IIfVr(noshipping, "", " shipping-name="""&shipmethods(0, gindex3)&"""")&" address-id="""&addressid&"""><shipping-rate currency="""&countryCurrency&""">0.00</shipping-rate><shippable>false</shippable><total-tax currency="""&countryCurrency&""">0.00</total-tax></result>"
												responsexml = responsexml & responsexml2
												cacheaddress(maxcacheid,3) = cacheaddress(maxcacheid,3) & responsexml2
											end if
											shipmethods(1, gindex3)=FALSE
										next
									end if
								else
									responsexml = responsexml & "<error-message>" & errormsg & "</error-message>"
								end if
								maxcacheid = maxcacheid+1
							end if
						end if
					end if
				next
			end select
		next
		responsexml = responsexml & "</results></merchant-calculation-results>"
		if standalonetestmode then
			response.write "<HR>" & Replace(Replace(responsexml,"</","&lt;/"),"<","<br />&lt;")
		else
			response.clear
			response.write responsexml
		end if
	Case "new-order-notification"
		sub get_google_address(xmlobj,ByRef gEmail,ByRef gName,ByRef gAddress,ByRef gAddress2,ByRef gCity,ByRef gState,ByRef gZip,ByRef gCountry,ByRef gPhone)
			for index2=0 to xmlobj.childNodes.length-1
				Set t = xmlobj.childNodes.Item(index2)
				Select Case t.nodeName
				case "email"
					gEmail=t.firstChild.nodeValue
				case "contact-name"
					gName=t.firstChild.nodeValue
				case "address1"
					gAddress=t.firstChild.nodeValue
				case "address2"
					if t.hasChildNodes then gAddress2=t.firstChild.nodeValue else gAddress2=""
				case "city"
					if t.hasChildNodes then gCity=t.firstChild.nodeValue else gCity=""
				case "region"
					if t.hasChildNodes then gState=t.firstChild.nodeValue else gState=""
				case "postal-code"
					gZip=t.firstChild.nodeValue
				case "country-code"
					gCountry=t.firstChild.nodeValue
				case "phone"
					if t.hasChildNodes then gPhone=t.firstChild.nodeValue else gPhone=""
				end select
			next
		end sub
		totaldiscounts=0
		stateTax=0
		countryTax=0
		totalgoods=0
		handling=0
		shipping=0
		freeshipamnt=0
		cpnmessage=""
		ordComLoc=0
		ordAddInfo=""
		ordAffiliate=""
		ordExtra1=""
		ordExtra2=""
		ordExtra3=""
		for index=0 to gcXmlDoc.documentElement.childNodes.length-1
			set obj1 = gcXmlDoc.documentElement.childNodes.Item(index)
			select case obj1.nodeName
			case "google-order-number"
				ordAuthNumber=obj1.firstChild.nodeValue
			case "order-total"
				ordTotal=obj1.firstChild.nodeValue
			case "shopping-cart"
				thesessionid = obj1.getElementsByTagName("sessionid").item(0).firstChild.nodeValue
				set obj2 = obj1.getElementsByTagName("partner")
				if obj2.length > 0 then
					if obj2.item(0).hasChildNodes then ordAffiliate = trim(obj2.item(0).firstChild.nodeValue&"")
				end if
				set lineitems = obj1.getElementsByTagName("items").item(0)
				for index2 = 0 to lineitems.childNodes.length - 1
					Set obj2 = lineitems.childNodes.Item(index2)
					Set obj3 = obj2.getElementsByTagName("discountflag")
					if obj3.length > 0 then
						if obj3.item(0).firstChild.nodeValue="true" then
							set obj3 = obj2.getElementsByTagName("unit-price")
							if obj3.length > 0 then
								totaldiscounts = totaldiscounts + (0 - obj3.item(0).firstChild.nodeValue)
							end if
							set obj3 = obj2.getElementsByTagName("item-description")
							if obj3.length > 0 then
								cpnmessage = replace(obj3.item(0).firstChild.nodeValue, " - ", "<br />") & "<br />" & cpnmessage
							end if
						end if
					end if
				next
			case "total-tax"
				countryTax=obj1.firstChild.nodeValue
			case "order-adjustment"
				set obj2 = obj1.getElementsByTagName("coupon-adjustment")
				if obj2.length > 0 then
					set obj3 = obj2.Item(0).getElementsByTagName("applied-amount")
					if obj3.length > 0 then
						totaldiscounts = totaldiscounts + obj3.item(0).firstChild.nodeValue
					end if
					set obj3 = obj2.Item(0).getElementsByTagName("message")
					if obj3.length > 0 then
						cpnmessage = obj3.item(0).firstChild.nodeValue & "<br />" & cpnmessage
					end if
				end if
				set obj2 = obj1.getElementsByTagName("shipping-name")
				if obj2.length > 0 then
					shipMethod = obj2.item(0).firstChild.nodeValue
				end if
				set obj2 = obj1.getElementsByTagName("shipping-cost")
				if obj2.length > 0 then
					shipping = obj2.item(0).firstChild.nodeValue
				end if
				set obj2 = obj1.getElementsByTagName("total-tax")
				if obj2.length > 0 then
					countryTax = countryTax + obj2.item(0).firstChild.nodeValue
				end if
			case "buyer-billing-address"
				call get_google_address(obj1,ordEmail,ordName,ordAddress,ordAddress2,ordCity,ordState,ordZip,ordCountry,ordPhone)
			case "buyer-shipping-address"
				call get_google_address(obj1,dummyEmail,ordShipName,ordShipAddress,ordShipAddress2,ordShipCity,ordShipState,ordShipZip,ordShipCountry,ordShipPhone)
			end select
		next
		
		sSQL = "SELECT cartID FROM cart WHERE cartCompleted=0 AND cartSessionID=" & replace(thesessionid, "'", "")
		rs.Open sSQL,cnn,0,1
		success = (NOT rs.EOF)
		rs.Close
		if success then
			totalgoods = (ordTotal - (stateTax+countryTax+shipping+handling)) + totaldiscounts
			sSQL = "SELECT ordID FROM orders WHERE ordSessionID="&replace(thesessionid,"'","")&" AND ordAuthNumber=''"
			rs.Open sSQL,cnn,0,1
			if NOT rs.EOF then orderid=rs("ordID") else orderid=""
			rs.Close
			if orderid="" then
				rs.Open "orders",cnn,1,3,&H0002
				rs.AddNew
			else
				if mysqlserver then rs.CursorLocation = 3
				rs.Open "SELECT * FROM orders WHERE ordID="&orderid,cnn,1,3,&H0001
			end if
			if ordShipName="" AND ordShipAddress="" AND ordShipAddress2="" AND ordShipCity="" then ordShipCountry=""
			rs.Fields("ordSessionID")	= thesessionid
			rs.Fields("ordName")		= ordName
			rs.Fields("ordAddress")		= ordAddress
			rs.Fields("ordAddress2")	= ordAddress2
			rs.Fields("ordCity")		= ordCity
			rs.Fields("ordState")		= ordState
			rs.Fields("ordZip")			= ordZip
			rs.Fields("ordCountry")		= ordCountry
			rs.Fields("ordEmail")		= ordEmail
			rs.Fields("ordPhone")		= ordPhone
			rs.Fields("ordShipName")	= ordShipName
			rs.Fields("ordShipAddress")	= ordShipAddress
			rs.Fields("ordShipAddress2")= ordShipAddress2
			rs.Fields("ordShipCity")	= ordShipCity
			rs.Fields("ordShipState")	= ordShipState
			rs.Fields("ordShipZip")		= ordShipZip
			rs.Fields("ordShipCountry")	= ordShipCountry
			rs.Fields("ordPayProvider") = 20
			rs.Fields("ordAuthNumber")	= ordAuthNumber
			rs.Fields("ordShipping")	= shipping - freeshipamnt
			if usehst=true then
				rs.Fields("ordHSTTax")		= stateTax + countryTax
				rs.Fields("ordStateTax")	= 0
				rs.Fields("ordCountryTax")	= 0
			else
				rs.Fields("ordHSTTax")		= 0
				rs.Fields("ordStateTax")	= stateTax
				rs.Fields("ordCountryTax")	= countryTax
			end if
			rs.Fields("ordHandling")	= handling
			rs.Fields("ordShipType")	= shipMethod
			if adminIntShipping<>0 AND ordShipCountry<>origCountryCode then
				rs.Fields("ordShipCarrier")	= adminIntShipping
			else
				rs.Fields("ordShipCarrier")	= shipType
			end if
			rs.Fields("ordTotal")		= totalgoods
			rs.Fields("ordDate")		= DateAdd("h",dateadjust,Now())
			rs.Fields("ordStatus")		= 2
			rs.Fields("ordStatusDate")	= DateAdd("h",dateadjust,Now())
			rs.Fields("ordIP")			= ""
			rs.Fields("ordComLoc")		= ordComLoc
			rs.Fields("ordAffiliate")	= ordAffiliate
			rs.Fields("ordAddInfo")		= ordAddInfo
			rs.Fields("ordDiscount")	= totaldiscounts
			rs.Fields("ordDiscountText")= Left(cpnmessage,255)
			rs.Update
			if mysqlserver=true then
				if orderid="" then
					rs.Close
					rs.Open "SELECT LAST_INSERT_ID() AS lstIns",cnn,0,1
					orderid = rs("lstIns")
				end if
			else
				orderid = rs.Fields("ordID")
			end if
			rs.Close
			sSQL="UPDATE cart SET cartOrderID="&orderid&",cartCompleted=2 WHERE cartCompleted=0 AND cartSessionID="&replace(thesessionid,"'","")
			cnn.Execute(sSQL)
			set objHttp = Server.CreateObject("WinHTTP.WinHTTPRequest.5.1")
			theurl="https://"&IIfVr(googledemomode, "sandbox", "checkout")&".google.com/cws/v2/Merchant/"&googledata1&"/request"
			objHttp.open "POST", theurl, false
			objHttp.setRequestHeader "Authorization", "Basic " & vrbase64_encrypt(googledata1&":"&googledata2)
			objHttp.setRequestHeader "Content-Type", "application/xml"
			objHttp.setRequestHeader "Accept", "application/xml"
			objHttp.Send "<?xml version=""1.0"" encoding=""UTF-8""?>" & _
				"<add-merchant-order-number xmlns=""http://checkout.google.com/schema/2"" google-order-number=""" & ordAuthNumber & """><merchant-order-number>" & orderid & "</merchant-order-number></add-merchant-order-number>"
		end if
		Response.write "<?xml version=""1.0"" encoding=""UTF-8""?><notification-acknowledgment xmlns=""http://checkout.google.com/schema/2""/>"
	Case "order-state-change-notification"
		Set obj1 = gcXmlDoc.getElementsByTagName("google-order-number").Item(0)
		ordnumber = replace(obj1.firstChild.nodeValue,"'","")
		sSQL = "SELECT ordID FROM orders WHERE ordAuthNumber='"&replace(ordnumber,"'","")&"' AND ordPayProvider=20"
		rs.Open sSQL,cnn,0,1
		if NOT rs.EOF then ordID=rs("ordID") else ordID=""
		rs.Close
		Set obj1 = gcXmlDoc.getElementsByTagName("new-financial-order-state").Item(0)
		financialstate = replace(obj1.firstChild.nodeValue,"'","")
		Set obj1 = gcXmlDoc.getElementsByTagName("previous-financial-order-state").Item(0)
		oldfinancialstate = replace(obj1.firstChild.nodeValue,"'","")
		Set obj1 = gcXmlDoc.getElementsByTagName("new-fulfillment-order-state").Item(0)
		fulfillmentstate = replace(obj1.firstChild.nodeValue,"'","")
		Set obj1 = gcXmlDoc.getElementsByTagName("previous-fulfillment-order-state").Item(0)
		oldfulfillmentstate = replace(obj1.firstChild.nodeValue,"'","")
		if ordID<>"" then
			if oldfinancialstate<>financialstate then
				select case financialstate
				case "CHARGEABLE"
					do_stock_management(ordID)
					cnn.Execute("UPDATE cart SET cartCompleted=1 WHERE cartOrderID="&ordID)
					cnn.Execute("UPDATE orders SET ordStatus=3,ordStatusDate=" & datedelim & VSUSDateTime(DateAdd("h",dateadjust,Now())) & datedelim & " WHERE ordID="&ordID)
					Call do_order_success(ordID,emailAddr,sendEmail,FALSE,TRUE,TRUE,TRUE)
				case "CHARGING"
				case "CHARGED"
					do_stock_management(ordID)
					cnn.Execute("UPDATE cart SET cartCompleted=1 WHERE cartOrderID="&ordID)
					cnn.Execute("UPDATE orders SET ordStatus=4,ordStatusDate=" & datedelim & VSUSDateTime(DateAdd("h",dateadjust,Now())) & datedelim & " WHERE ordID="&ordID)
					Call do_order_success(ordID,emailAddr,sendEmail,FALSE,TRUE,TRUE,TRUE)
				case "PAYMENT_DECLINED"
					cnn.Execute("UPDATE orders SET ordStatus=2,ordStatusDate=" & datedelim & VSUSDateTime(DateAdd("h",dateadjust,Now())) & datedelim & " WHERE ordID="&ordID)
				case "CANCELLED"
					release_stock(ordID)
					cnn.Execute("UPDATE orders SET ordStatus=0,ordStatusDate=" & datedelim & VSUSDateTime(DateAdd("h",dateadjust,Now())) & datedelim & " WHERE ordID="&ordID)
				case "CANCELLED_BY_GOOGLE"
					release_stock(ordID)
					sSQL = "SELECT ordStatusInfo FROM orders WHERE ordID="&ordID
					rs.Open sSQL,cnn,0,1
					if NOT rs.EOF then currstatusinfo = rs("ordStatusInfo") else currstatusinfo = ""
					rs.Close
					cnn.Execute("UPDATE orders SET ordStatus=0,ordStatusDate=" & datedelim & VSUSDateTime(DateAdd("h",dateadjust,Now())) & datedelim & ",ordStatusInfo='"&replace("Cancelled By Google." & vbCrLf & currstatusinfo,"'","''")&"' WHERE ordID="&ordID)
				end select
			end if
			if oldfulfillmentstate<>fulfillmentstate then
				if googledeliveredstate="" then googledeliveredstate=5
				select case fulfillmentstate
				case "DELIVERED"
					cnn.Execute("UPDATE orders SET ordStatus="&googledeliveredstate&",ordStatusDate=" & datedelim & VSUSDateTime(DateAdd("h",dateadjust,Now())) & datedelim & " WHERE ordID="&ordID)
				end select
			end if
		end if
		Response.write "<?xml version=""1.0"" encoding=""UTF-8""?><notification-acknowledgment xmlns=""http://checkout.google.com/schema/2""/>"
	Case "charge-amount-notification"
		' Test
		Response.write "<?xml version=""1.0"" encoding=""UTF-8""?><notification-acknowledgment xmlns=""http://checkout.google.com/schema/2""/>"
	Case "chargeback-amount-notification"
		session.LCID = 1033
		success=TRUE
		amount=0
		Set obj1 = gcXmlDoc.getElementsByTagName("google-order-number").Item(0)
		ordnumber = replace(obj1.firstChild.nodeValue,"'","")
		sSQL = "SELECT ordID,ordShipping,ordStateTax,ordCountryTax,ordHandling,ordTotal,ordDiscount,ordAuthNumber FROM orders WHERE ordAuthNumber='"&replace(ordnumber,"'","")&"' AND ordPayProvider=20"
		rs.Open sSQL,cnn,0,1
		if NOT rs.EOF then
			ordID = rs("ordID")
			amount = cDbl((rs("ordShipping")+rs("ordStateTax")+rs("ordCountryTax")+rs("ordTotal")+rs("ordHandling"))-rs("ordDiscount"))
		else
			success = FALSE
		end if
		rs.Close
		Set obj1 = gcXmlDoc.getElementsByTagName("total-chargeback-amount").Item(0)
		refundamount = cDbl(replace(obj1.firstChild.nodeValue,"'",""))
		if success AND amount <= refundamount then
			release_stock(ordID)
			cnn.Execute("UPDATE orders SET ordStatus=0,ordStatusDate=" & datedelim & VSUSDateTime(DateAdd("h",dateadjust,Now())) & datedelim & " WHERE ordID="&ordID)
		end if
		Response.write "<?xml version=""1.0"" encoding=""UTF-8""?><notification-acknowledgment xmlns=""http://checkout.google.com/schema/2""/>"
	Case "refund-amount-notification"
		session.LCID = 1033
		success=TRUE
		amount=0
		ordID=0
		Set obj1 = gcXmlDoc.getElementsByTagName("google-order-number").Item(0)
		ordnumber = replace(obj1.firstChild.nodeValue,"'","")
		sSQL = "SELECT ordID,ordShipping,ordStateTax,ordCountryTax,ordHandling,ordTotal,ordDiscount,ordAuthNumber FROM orders WHERE ordAuthNumber='"&replace(ordnumber,"'","")&"' AND ordPayProvider=20"
		rs.Open sSQL,cnn,0,1
		if NOT rs.EOF then
			ordID = rs("ordID")
			amount = cDbl((rs("ordShipping")+rs("ordStateTax")+rs("ordCountryTax")+rs("ordTotal")+rs("ordHandling"))-rs("ordDiscount"))
		else
			success = FALSE
		end if
		rs.Close
		Set obj1 = gcXmlDoc.getElementsByTagName("total-refund-amount").Item(0)
		refundamount = cDbl(replace(obj1.firstChild.nodeValue,"'",""))
		if success AND amount <= refundamount then
			release_stock(ordID)
			cnn.Execute("UPDATE orders SET ordStatus=0,ordStatusDate=" & datedelim & VSUSDateTime(DateAdd("h",dateadjust,Now())) & datedelim & " WHERE ordID="&ordID)
		end if
		Response.write "<?xml version=""1.0"" encoding=""UTF-8""?><notification-acknowledgment xmlns=""http://checkout.google.com/schema/2""/>"
	Case "risk-information-notification"
		ipaddress = ""
		avs = ""
		cvv = ""
		iseligable = ""
		partialcc = ""
		acctage = 0
		Set obj1 = gcXmlDoc.getElementsByTagName("google-order-number").Item(0)
		ordnumber = replace(obj1.firstChild.nodeValue,"'","")
		set obj2 = gcXmlDoc.getElementsByTagName("risk-information")
		if obj2.length > 0 then
			set obj3 = obj2.Item(0).getElementsByTagName("ip-address")
			if obj3.length > 0 then
				ipaddress = obj3.item(0).firstChild.nodeValue
			end if
			set obj3 = obj2.Item(0).getElementsByTagName("avs-response")
			if obj3.length > 0 then
				avs = obj3.item(0).firstChild.nodeValue
			end if
			set obj3 = obj2.Item(0).getElementsByTagName("cvn-response")
			if obj3.length > 0 then
				cvv = obj3.item(0).firstChild.nodeValue
			end if
			set obj3 = obj2.Item(0).getElementsByTagName("buyer-account-age")
			if obj3.length > 0 then
				acctage = obj3.item(0).firstChild.nodeValue
			end if
			set obj3 = obj2.Item(0).getElementsByTagName("partial-cc-number")
			if obj3.length > 0 then
				partialcc = obj3.item(0).firstChild.nodeValue
			end if
			set obj3 = obj2.Item(0).getElementsByTagName("eligible-for-protection")
			if obj3.length > 0 then
				iseligable = obj3.item(0).firstChild.nodeValue
				if iseligable="false" then iseligable=xxNo else iseligable=xxYes
			end if
		end if
		if ordnumber<>"" then
			sSQL = "UPDATE orders SET ordIP='"&replace(ipaddress,"'","")&"',ordAVS='"&replace(avs,"'","")&"/"&iseligable&"',ordCVV='"&replace(cvv,"'","")&"/"&acctage&"',ordCNum='"&partialcc&"' WHERE ordAuthNumber='"&ordnumber&"' AND ordPayProvider=20"
			cnn.Execute(sSQL)
		end if
		Response.write "<?xml version=""1.0"" encoding=""UTF-8""?><notification-acknowledgment xmlns=""http://checkout.google.com/schema/2""/>"
	Case "request-received"
		' Test
		Response.write "<?xml version=""1.0"" encoding=""UTF-8""?><notification-acknowledgment xmlns=""http://checkout.google.com/schema/2""/>"
	Case "error"
		' Test
		Response.write "<?xml version=""1.0"" encoding=""UTF-8""?><notification-acknowledgment xmlns=""http://checkout.google.com/schema/2""/>"
	Case "diagnosis"
		' Test
		Response.write "<?xml version=""1.0"" encoding=""UTF-8""?><notification-acknowledgment xmlns=""http://checkout.google.com/schema/2""/>"
	Case Else ' None of the above: message is not recognized.
	end select
end if

if debugmode=TRUE then
	htmlemails=false
	if htmlemails=true then emlNl = "<br />" else emlNl=vbCrLf
	
	emailtxt = "ThisMessage: " & thismessage & emlNl & "Response: " & xmlResponse & emlNl & emlNl
	emailtxt = emailtxt & "responsexml:" & responsexml & emlNl
	emailtxt = emailtxt & "<p>Callback took " & timer()-thetime1 & " seconds.</p>" & emlNl
	Call DoSendEmailEO(emailAddr,emailAddr,"","gcallback.asp debug",emailtxt,emailObject,themailhost,theuser,thepass)
end if
set objHttp = nothing
%>