<!--#include file="md5.asp"-->
<%
Dim sSQL,rs,alldata,quantity,grandtotal,netnav,bExists,cartID,cartEmpty,index,index2,rowcounter,objItem,totShipOptions,cpnmessage,totaldiscounts,numhomecountries,nonhomecountries,blockmultipurchase,multipurchaseblockmessage
Dim demomode,data1,data2,success,errormsg,shipping,totalgoods,orderid,sXML,destZip,allzones,stateTax,stateTaxRate,countryID,somethingToShip,taxfreegoods,uspsmethods,freeshipamnt,pzFSA
Dim iTotItems,international,checkIntOptions,shipMethod,shipArr,shipcountry,intShipping(5,20),havematch,dHighest(10),dHighWeight,dTotalWeight,dTotalWeightOz,thePQuantity,thePWeight
cartEmpty=False
isInStock=true
outofstockreason=0
if dateadjust="" then dateadjust=0
WSP = ""
OWSP = ""
nodiscounts=false
success=True : usehst=false : checkIntOptions=False : alldata = "" : shipMethod = "" : shipping = 0
iTotItems = 0 : iWeight = 0 : stateTaxRate=0 : countryTax=0 : stateTax=0
appliedcouponname="" : ordAVS="" : ordCVV="" : stateAbbrev="" : international = "" : thePQuantity = 0 : thePWeight = 0
appliedcouponamount = 0 : totalquantity = 0 : statetaxfree = 0 : countrytaxfree = 0 : shipfreegoods = 0 : totalgoods = 0
somethingToShip = false : freeshippingapplied = false : freeshipamnt = 0 : rowcounter = 0
gotcpncode=false : isstandardship = false : numshipoptions=0 : homecountry = false : totalshipitems = 0
if cartisincluded<>TRUE then
	cpncode = Trim(replace(request.form("cpncode"),"'",""))
	payerid = request.form("payerid")
	token = request("token")
	if trim(Request.form("sessionid"))<>"" then thesessionid=replace(trim(Request.form("sessionid")),"'","") else thesessionid=Session.SessionID
	if NOT isnumeric(thesessionid) then thesessionid=-1
	theid = Replace(Trim(Request.Form("id")),"'","")
	checkoutmode=request.form("mode")
	shippingpost=trim(request.form("shipping"))
	commerciallocpost = Request.Form("commercialloc")
	wantinsurancepost = trim(request.form("wantinsurance"))
	payproviderpost = trim(request.form("payprovider"))
end if
paypalexpress=FALSE
ppexpresscancel=FALSE
Set rs = Server.CreateObject("ADODB.RecordSet")
Set rs2 = Server.CreateObject("ADODB.RecordSet")
Set rs3 = Server.CreateObject("ADODB.RecordSet")
Set cnn=Server.CreateObject("ADODB.Connection")
cnn.open sDSN
alreadygotadmin = getadminsettings()
countryTax=0 ' At present both countryTaxRate and countryTax are set in incfunctions
origShipType=shipType
if cartisincluded<>TRUE then
	if (alternateratesups<>"" OR alternateratesusps<>"" OR alternateratesweightbased<>"" OR alternateratescanadapost<>"" OR alternateratesfedex<>"") then alternaterates = TRUE else alternaterates = FALSE
	if request.form("altrates")<>"" then
		altrate=int(request.form("altrates"))
		if alternateratesups<>"" AND altrate=4 then shipType=4
		if alternateratesusps<>"" AND altrate=3 then shipType=3
		if alternateratesweightbased<>"" AND altrate=2 then shipType=2
		if alternateratescanadapost<>"" AND altrate=6 then shipType=6
		if alternateratesfedex<>"" AND altrate=7 then shipType=7
	end if
	ordPayProvider = replace(payproviderpost,"'","")
end if
if ordPayProvider<>"" then execute("handling = handling + handlingcharge" & ordPayProvider & " : handlingchargepercent = handlingchargepercent" & ordPayProvider)
if Session("clientUser")<>"" then
	if (Session("clientActions") AND 8) = 8 then
		WSP = "pWholesalePrice AS "
		if wholesaleoptionpricediff=TRUE then OWSP = "optWholesalePriceDiff AS "
		if nowholesalediscounts=true then nodiscounts=true
	end if
	if (Session("clientActions") AND 16) = 16 then
		Session.LCID = 1033
		WSP = Session("clientPercentDiscount") & "*pPrice AS "
		if wholesaleoptionpricediff=TRUE then OWSP = Session("clientPercentDiscount") & "*optPriceDiff AS "
		if nowholesalediscounts=true then nodiscounts=true
		Session.LCID = saveLCID
	end if
end if
if Session("couponapply")<>"" then
	cnn.Execute("UPDATE coupons SET cpnNumAvail=cpnNumAvail+1 WHERE cpnID IN (0" & Session("couponapply")&")")
	Session("couponapply")=""
end if
Function show_states(tstate)
	Dim foundmatch
	foundmatch=false
	if xxOutState<>"" then response.write "<option value=''>"&xxOutState&"</option>"
	if IsArray(allstates) then
		for rowcounter=0 to UBOUND(allstates,2)
			response.write "<option value="""&Replace(IIfVr(usestateabbrev=TRUE,allstates(1,rowcounter),allstates(0,rowcounter)),"""","&quot;")&""""
			if tstate=allstates(0,rowcounter) OR tstate=allstates(1,rowcounter) then
				response.write " selected"
				foundmatch=true
			end if
			response.write ">"&allstates(0,rowcounter)&"</option>"&vbCrLf
		next
	end if
	show_states=foundmatch
End Function
Sub show_countries(tcountry)
	if IsArray(allcountries) then
		for rowcounter=0 to UBOUND(allcountries,2)
			response.write "<option value="""&Replace(allcountries(0,rowcounter),"""","&quot;")&""""
			if tcountry=allcountries(0,rowcounter) then response.write " selected"
			response.write ">"&allcountries(2,rowcounter)&"</option>"&vbCrLf
		next
	end if
End Sub
function checkuserblock(thepayprov)
	multipurchaseblocked=FALSE
	if multipurchaseblockmessage="" then multipurchaseblockmessage="I'm sorry. We are experiencing temporary difficulties at the moment. Please try your purchase again later."
	if thepayprov<>"7" AND thepayprov <> "13" then
		theip = trim(replace(left(request.servervariables("REMOTE_ADDR"), 48), "'", ""))
		if theip = "" then theip = "none"
		if blockmultipurchase<>"" then
			cnn.Execute("DELETE FROM multibuyblock WHERE lastaccess<" & datedelim & VSUSDateTime(Now()-1) & datedelim)
			sSQL = "SELECT ssdenyid,sstimesaccess FROM multibuyblock WHERE ssdenyip = '" & theip & "'"
			rs.Open sSQL,cnn,0,1
			if NOT rs.EOF then
				cnn.Execute("UPDATE multibuyblock SET sstimesaccess=sstimesaccess+1,lastaccess=" & datedelim & VSUSDateTime(Now()) & datedelim & " WHERE ssdenyid=" & rs("ssdenyid"))
				if rs("sstimesaccess") >= blockmultipurchase then multipurchaseblocked=TRUE
			else
				cnn.Execute("INSERT INTO multibuyblock (ssdenyip,lastaccess) VALUES ('" & theip & "'," & datedelim & VSUSDateTime(Now()) & datedelim & ")")
			end if
			rs.Close
		end if
		if theip = "none" then
			sSQL = "SELECT TOP 1 dcid FROM ipblocking"
		else
			sSQL = "SELECT dcid FROM ipblocking WHERE (dcip1=" & ip2long(theip) & " AND dcip2=0) OR (dcip1 <= " & ip2long(theip) & " AND " & ip2long(theip) & " <= dcip2 AND dcip2 <> 0)"
		end if
		rs.Open sSQL,cnn,0,1
		if NOT rs.EOF then multipurchaseblocked = TRUE
		rs.Close
	end if
	checkuserblock = multipurchaseblocked
end function
sub checkpricebreaks(cpbpid,origprice)
	newprice=""
	sSQL = "SELECT SUM(cartQuantity) AS totquant FROM cart WHERE cartCompleted=0 AND cartSessionID="&Session.SessionID&" AND cartProdID='"&replace(cpbpid,"'","''")&"'"
	rs2.Open sSQL,cnn,0,1
	if IsNull(rs2("totquant")) then thetotquant=0 else thetotquant = rs2("totquant")
	rs2.Close
	sSQL="SELECT "&WSP&"pPrice FROM pricebreaks WHERE "&thetotquant&">=pbQuantity AND pbProdID='"&replace(cpbpid,"'","''")&"' ORDER BY " & IIfVr(WSP="","pPrice",replace(WSP," AS ",""))
	rs2.Open sSQL,cnn,0,1
	if NOT rs2.EOF then
		thepricebreak = rs2("pPrice")
	else
		thepricebreak = origprice
	end if
	rs2.Close
	Session.LCID = 1033
	sSQL = "UPDATE cart SET cartProdPrice="&FormatNumber(thepricebreak,4,-1,0,0)&" WHERE cartCompleted=0 AND cartSessionID="&Session.SessionID&" AND cartProdID='"&replace(cpbpid,"'","''")&"'"
	Session.LCID = saveLCID
	cnn.Execute(sSQL)
end sub
function multShipWeight(theweight, themul)
	multShipWeight = (theweight*themul)/100.0
end function
sub subtaxesfordiscounts(theExemptions, discAmount)
	if (theExemptions AND 1)=1 then statetaxfree = statetaxfree - discAmount
	if (theExemptions AND 2)=2 then countrytaxfree = countrytaxfree - discAmount
	if (theExemptions AND 4)=4 then shipfreegoods = shipfreegoods - discAmount
end sub
sub addadiscount(resset, groupdiscount, dscamount, subcpns, cdcpncode, statetaxhandback, countrytaxhandback, theexemptions, thetax)
	totaldiscounts = totaldiscounts + dscamount
	if groupdiscount then
		statetaxfree = statetaxfree - (dscamount * statetaxhandback)
		countrytaxfree = countrytaxfree - (dscamount * countrytaxhandback)
	else
		call subtaxesfordiscounts(theexemptions, dscamount)
		if perproducttaxrate then countryTax = countryTax - ((dscamount * thetax) / 100.0)
	end if
	if InStr(cpnmessage,"<br />" & resset("cpnName") & "<br />")=0 then cpnmessage = cpnmessage & resset("cpnName") & "<br />"
	if subcpns then
		Set theres = cnn.Execute("SELECT cpnID FROM coupons WHERE cpnNumAvail>0 AND cpnNumAvail<30000000 AND cpnID=" & resset("cpnID"))
		if NOT theres.EOF then Session("couponapply") = Session("couponapply") & "," & resset("cpnID")
		cnn.Execute("UPDATE coupons SET cpnNumAvail=cpnNumAvail-1 WHERE cpnNumAvail>0 AND cpnNumAvail<30000000 AND cpnID=" & resset("cpnID"))
	end if
	if cdcpncode<>"" AND LCase(Trim(resset("cpnNumber")))=LCase(cdcpncode) then gotcpncode=true : appliedcouponname = resset("cpnName") : appliedcouponamount = dscamount
end sub
function timesapply(taquant,tathresh,tamaxquant,tamaxthresh,taquantrepeat,tathreshrepeat)
	if taquantrepeat=0 AND tathreshrepeat=0 then
		tatimesapply = 1.0
	elseif tamaxquant=0 then
		tatimesapply = Int((tathresh-tamaxthresh) / tathreshrepeat)+1
	elseif tamaxthresh=0 then
		tatimesapply = Int((taquant-tamaxquant) / taquantrepeat)+1
	else
		ta1 = Int((taquant-tamaxquant) / taquantrepeat)+1
		ta2 = Int((tathresh-tamaxthresh) / tathreshrepeat)+1
		if ta2 < ta1 then tatimesapply = ta2 else tatimesapply = ta1
	end if
	timesapply = tatimesapply
end function
sub calculatediscounts(cdgndtot, subcpns, cdcpncode)
	totaldiscounts = 0
	cpnmessage = "<br />"
	cdtotquant = 0
	if cdgndtot=0 then
		statetaxhandback = 0.0
		countrytaxhandback = 0.0
	else
		statetaxhandback = 1.0 - ((cdgndtot - statetaxfree) / cdgndtot)
		countrytaxhandback = 1.0 - ((cdgndtot - countrytaxfree) / cdgndtot)
	end if
	if NOT nodiscounts then
		Session.LCID = 1033
		cdalldata = ""
		sSQL = "SELECT cartProdID,SUM(cartProdPrice*cartQuantity),SUM(cartQuantity),pSection,COUNT(cartProdID),pExemptions,pTax FROM products INNER JOIN cart ON cart.cartProdID=products.pID WHERE cartCompleted=0 AND cartSessionID="&thesessionid&" GROUP BY cartProdID,pSection,pExemptions,pTax"
		rs2.Open sSQL,cnn,0,1
		if NOT (rs2.EOF OR rs2.BOF) then cdalldata=rs2.getrows
		rs2.Close
		if IsArray(cdalldata) then
			For index=0 to UBOUND(cdalldata,2)
				sSQL = "SELECT SUM(coPriceDiff*cartQuantity) AS totOpts FROM cart LEFT OUTER JOIN cartoptions ON cart.cartID=cartoptions.coCartID WHERE cartCompleted=0 AND cartSessionID="&thesessionid&" AND cartProdID='" & replace(cdalldata(0,index), "'", "''") & "'"
				rs2.Open sSQL,cnn,0,1
				if NOT IsNull(rs2("totOpts")) then cdalldata(1,index) = cdalldata(1,index) + rs2("totOpts")
				rs2.Close
				cdtotquant = cdtotquant + cdalldata(2,index)
				topcpnids = cdalldata(3,index)
				thetopts = cdalldata(3,index)
				if isnull(cdalldata(6,index)) then cdalldata(6,index) = countryTaxRate
				if NOT IsNull(thetopts) then
					for cpnindex=0 to 10
						if thetopts=0 then
							exit for
						else
							sSQL = "SELECT topSection FROM sections WHERE sectionID=" & thetopts
							rs.Open sSQL,cnn,0,1
							if NOT rs.EOF then
								thetopts = rs("topSection")
								topcpnids = topcpnids & "," & thetopts
							else
								rs.Close
								exit for
							end if
							rs.Close
						end if
					next
				end if
				tdt = Date()
				sSQL = "SELECT DISTINCT cpnID,cpnDiscount,cpnType,cpnNumber,cpnName,cpnThreshold,cpnQuantity,cpnThresholdRepeat,cpnQuantityRepeat FROM coupons LEFT OUTER JOIN cpnassign ON coupons.cpnID=cpnassign.cpaCpnID WHERE cpnNumAvail>0 AND cpnEndDate>="&datedelim&VSUSDate(tdt)&datedelim&" AND (cpnIsCoupon=0"
				if cdcpncode<>"" then sSQL = sSQL & " OR (cpnIsCoupon=1 AND cpnNumber='"&cdcpncode&"')"
				sSQL = sSQL & ") AND cpnThreshold<="&cdalldata(1,index)&" AND (cpnThresholdMax>"&cdalldata(1,index)&" OR cpnThresholdMax=0) AND cpnQuantity<="&cdalldata(2,index)&" AND (cpnQuantityMax>"&cdalldata(2,index)&" OR cpnQuantityMax=0) AND (cpnSitewide=0 OR cpnSitewide=2) AND "
				sSQL = sSQL & "(cpnSitewide=2 OR (cpaType=2 AND cpaAssignment='"&cdalldata(0,index)&"') "
				sSQL = sSQL & "OR (cpaType=1 AND cpaAssignment IN ('"&Replace(topcpnids,",","','")&"')))"
				rs2.Open sSQL,cnn,0,1
				do while NOT rs2.EOF
					if rs2("cpnType")=1 then ' Flat Rate Discount
						thedisc = cDbl(rs2("cpnDiscount")) * timesapply(cdalldata(2,index),cdalldata(1,index),rs2("cpnQuantity"),rs2("cpnThreshold"),rs2("cpnQuantityRepeat"),rs2("cpnThresholdRepeat"))
						if cdalldata(1,index) < thedisc then thedisc = cdalldata(1,index)
						call addadiscount(rs2, false, thedisc, subcpns, cdcpncode, statetaxhandback, countrytaxhandback, cdalldata(5,index), cdalldata(6,index))
					elseif rs2("cpnType")=2 then ' Percentage Discount
						call addadiscount(rs2, false, ((cDbl(rs2("cpnDiscount")) * cDbl(cdalldata(1,index))) / 100.0), subcpns, cdcpncode, statetaxhandback, countrytaxhandback, cdalldata(5,index), cdalldata(6,index))
					end if
					rs2.MoveNext
				loop
				rs2.Close
			Next
		end if
		tdt = Date()
		sSQL = "SELECT DISTINCT cpnID,cpnDiscount,cpnType,cpnNumber,cpnName,cpnSitewide,cpnThreshold,cpnThresholdMax,cpnQuantity,cpnQuantityMax,cpnThresholdRepeat,cpnQuantityRepeat FROM coupons WHERE cpnNumAvail>0 AND cpnEndDate>="&datedelim&VSUSDate(tdt)&datedelim&" AND (cpnIsCoupon=0"
		if cdcpncode<>"" then sSQL = sSQL & " OR (cpnIsCoupon=1 AND cpnNumber='"&cdcpncode&"')"
		sSQL = sSQL & ") AND cpnThreshold<="&cdgndtot&" AND cpnQuantity<="&cdtotquant&" AND (cpnSitewide=1 OR cpnSitewide=3) AND (cpnType=1 OR cpnType=2)"
		rs.Open sSQL,cnn,0,1
		do while NOT rs.EOF
			totquant = 0
			totprice = 0
			if rs("cpnSitewide")=3 then
				sSQL = "SELECT cpaAssignment FROM cpnassign WHERE cpaType=1 AND cpacpnID=" & rs("cpnID")
				rs2.Open sSQL,cnn,0,1
				secids = ""
				addcomma = ""
				do while NOT rs2.EOF
					secids = secids & addcomma & rs2("cpaAssignment")
					addcomma = ","
					rs2.MoveNext
				loop
				rs2.Close
				if NOT (secids = "") then
					secids = getsectionids(secids, false)
					sSQL = "SELECT SUM(cartProdPrice*cartQuantity) AS totPrice,SUM(cartQuantity) AS totQuant FROM products INNER JOIN cart ON cart.cartProdID=products.pID WHERE cartCompleted=0 AND cartSessionID="&thesessionid&" AND products.pSection IN (" & secids & ")"
					rs2.Open sSQL,cnn,0,1
						if IsNull(rs2("totPrice")) then totprice = 0 else totprice = rs2("totPrice")
						if IsNull(rs2("totQuant")) then totquant = 0 else totquant = rs2("totQuant")
					rs2.Close
					if mysqlserver=true then
						sSQL = "SELECT SUM(coPriceDiff*cartQuantity) AS optPrDiff FROM products INNER JOIN cart ON cart.cartProdID=products.pID LEFT OUTER JOIN cartoptions ON cart.cartID=cartoptions.coCartID WHERE cartCompleted=0 AND cartSessionID="&thesessionid&" AND products.pSection IN (" & secids & ")"
					else
						sSQL = "SELECT SUM(coPriceDiff*cartQuantity) AS optPrDiff FROM products INNER JOIN (cart LEFT OUTER JOIN cartoptions ON cart.cartID=cartoptions.coCartID) ON cart.cartProdID=products.pID WHERE cartCompleted=0 AND cartSessionID="&thesessionid&" AND products.pSection IN (" & secids & ")"
					end if
					rs2.Open sSQL,cnn,0,1
						if NOT IsNull(rs2("optPrDiff")) then totprice = totprice + rs2("optPrDiff")
					rs2.Close
				end if
			else
				totquant = cdtotquant
				totprice = cdgndtot
			end if
			if totquant > 0 AND rs("cpnThreshold") <= totprice AND (rs("cpnThresholdMax") > totprice OR rs("cpnThresholdMax")=0) AND rs("cpnQuantity") <= totquant AND (rs("cpnQuantityMax") > totquant OR rs("cpnQuantityMax")=0) then
				if rs("cpnType")=1 then ' Flat Rate Discount
					thedisc = cDbl(rs("cpnDiscount")) * timesapply(totquant,totprice,rs("cpnQuantity"),rs("cpnThreshold"),rs("cpnQuantityRepeat"),rs("cpnThresholdRepeat"))
					if totprice < thedisc then thedisc = totprice
				elseif rs("cpnType")=2 then ' Percentage Discount
					thedisc = ((cDbl(rs("cpnDiscount")) * cDbl(totprice)) / 100.0)
				end if
				call addadiscount(rs, true, thedisc, subcpns, cdcpncode, statetaxhandback, countrytaxhandback, 3, 0)
				if perproducttaxrate AND cdgndtot > 0 then
					if IsArray(cdalldata) then
						for index=0 to UBOUND(cdalldata,2)
							if rs("cpnType")=1 then ' Flat Rate Discount
								applicdisc = thedisc / (cdtotquant / cdalldata(2,index))
							elseif rs("cpnType")=2 then ' Percentage Discount
								applicdisc = thedisc / (cdgndtot / cdalldata(1,index))
							end if
							if (cdalldata(5,index) AND 2)<>2 then countryTax = countryTax - ((applicdisc * cdalldata(6,index)) / 100.0)
						next
					end if
				end if
			end if
			rs.MoveNext
		loop
		rs.Close
		Session.LCID = saveLCID
	end if
	if statetaxfree < 0 then statetaxfree = 0
	if countrytaxfree < 0 then countrytaxfree = 0
	totaldiscounts = vsround(totaldiscounts, 2)
end sub
sub calculateshippingdiscounts(subcpns)
	freeshipamnt = 0
	if NOT nodiscounts then
		Session.LCID = 1033
		tdt = Date()
		sSQL = "SELECT cpnID,cpnName,cpnNumber,cpnDiscount,cpnThreshold,cpnCntry FROM coupons WHERE cpnType=0 AND cpnSitewide=1 AND cpnNumAvail>0 AND cpnThreshold<="&totalgoods&" AND (cpnThresholdMax>"&totalgoods&" OR cpnThresholdMax=0) AND cpnQuantity<="&totalquantity&" AND (cpnQuantityMax>"&totalquantity&" OR cpnQuantityMax=0) AND cpnEndDate>="&datedelim&VSUSDate(tdt)&datedelim&" AND (cpnIsCoupon=0 OR (cpnIsCoupon=1 AND cpnNumber='"&cpncode&"'))"
		rs.Open sSQL,cnn,0,1
		do while NOT rs.EOF
			if freeshipapplies OR Int(rs("cpnCntry"))=0 then
				if cpncode<>"" AND LCase(Trim(rs("cpnNumber")))=LCase(cpncode) then gotcpncode=true : appliedcouponname = rs("cpnName")
				if isstandardship then
					if InStr(cpnmessage,"<br />" & rs("cpnName") & "<br />")=0 then cpnmessage = cpnmessage & rs("cpnName") & "<br />"
					freeshipamnt = shipping
					if subcpns then
						Set theres = cnn.Execute("SELECT cpnID FROM coupons WHERE cpnNumAvail>0 AND cpnNumAvail<30000000 AND cpnID=" & rs("cpnID"))
						if NOT theres.EOF then Session("couponapply") = Session("couponapply") & "," & rs("cpnID")
						cnn.Execute("UPDATE coupons SET cpnNumAvail=cpnNumAvail-1 WHERE cpnNumAvail>0 AND cpnNumAvail<30000000 AND cpnID=" & rs("cpnID"))
					end if
				end if
				freeshippingapplied = true
			end if
			rs.MoveNext
		loop
		rs.Close
		Session.LCID = saveLCID
	end if
	if freeshipamnt > shipping then freeshipamnt = shipping
end sub
sub initshippingmethods()
	for i=0 to UBOUND(intShipping,2)
		intShipping(0,i)="" ' Name
		intShipping(1,i)="" ' Delivery
		intShipping(2,i)=0 ' Cost
		intShipping(3,i)=false ' Used
		intShipping(4,i)=0 ' FSA
		intShipping(5,i)="" ' Name to match (USPS)
	next
	if shipcountry <> origCountry then
		international = "Intl"
		willpickuptext = ""
		if adminIntShipping<>0 then
			if cartisincluded=TRUE then
				shipType=adminIntShipping
			elseif request.form("altrates")="" then
				shipType=adminIntShipping
			end if
		end if
	end if
	if shipType=2 OR shipType=5 then ' Weight / Price based shipping
		allzones=""
		zoneid=0
		if splitUSZones AND shiphomecountry then
			sSQL = "SELECT pzID,pzMultiShipping,pzFSA,pzMethodName1,pzMethodName2,pzMethodName3,pzMethodName4,pzMethodName5 FROM states INNER JOIN postalzones ON postalzones.pzID=states.stateZone WHERE "&IIfVr(usestateabbrev=TRUE,"stateAbbrev","stateName")&"='"&Replace(shipstate,"'","''")&"'"
		else
			sSQL = "SELECT pzID,pzMultiShipping,pzFSA,pzMethodName1,pzMethodName2,pzMethodName3,pzMethodName4,pzMethodName5 FROM countries INNER JOIN postalzones ON postalzones.pzID=countries.countryZone WHERE countryName='"&Replace(shipcountry,"'","''")&"'"
		end if
		rs.Open sSQL,cnn,0,1
		if NOT (rs.EOF OR rs.BOF) then
			zoneid=rs("pzID")
			numshipoptions=rs("pzMultiShipping")
			pzFSA = rs("pzFSA")
			for index3=0 to numshipoptions
				intShipping(0,index3)=rs("pzMethodName"&(index3+1))
				intShipping(2,index3)=0
				intShipping(3,index3)=TRUE
				intShipping(4,index3)=IIfVr((rs("pzFSA") AND (2 ^ index3))<>0, 1, 0)
			next
		else
			success=false
			if splitUSZones AND shiphomecountry AND shipstate="" then errormsg = xxPlsSta else errormsg = "Country / state shipping zone is unassigned."
		end if
		rs.Close
		sSQL = "SELECT zcWeight,zcRate,zcRate2,zcRate3,zcRate4,zcRate5,zcRatePC,zcRatePC2,zcRatePC3,zcRatePC4,zcRatePC5 FROM zonecharges WHERE zcZone="&zoneid&" ORDER BY zcWeight"
		rs.Open sSQL,cnn,0,1
		if NOT (rs.EOF OR rs.BOF) then allzones=rs.getrows
		rs.Close
	elseif shipType=3 OR shipType=4 OR shipType=6 OR shipType=7 then ' USPS / UPS / Canada Post / Fedex
		if shipType=3 then
			sSQL = "SELECT uspsMethod,uspsFSA,uspsShowAs FROM uspsmethods WHERE uspsID<100 AND uspsUseMethod=1 AND uspsLocal="
			if international="" then sSQL=sSQL&"1" else sSQL=sSQL&"0"
		elseif shipType=4 then
			shipinsuranceamt=""
			sSQL = "SELECT uspsMethod,uspsFSA,uspsShowAs FROM uspsmethods WHERE uspsID>100 AND uspsID<200 AND uspsUseMethod=1"
		elseif shipType=6 then
			sSQL = "SELECT uspsMethod,uspsFSA,uspsShowAs FROM uspsmethods WHERE uspsID>200 AND uspsID<300 AND uspsUseMethod=1"
		elseif shipType=7 then
			sSQL = "SELECT uspsMethod,uspsFSA,uspsShowAs,uspsLocal FROM uspsmethods WHERE uspsID>300 AND uspsID<400 AND uspsUseMethod=1"
			if international="" then sSQL = sSQL & " AND uspsMethod<>" & IIfVr(commerciallocpost="Y", "'GROUNDHOMEDELIVERY'", "'FEDEXGROUND'")
		end if
		rs.Open sSQL,cnn,0,1
		if NOT rs.EOF then
			uspsmethods=rs.GetRows()
		else
			success=false
			errormsg = "Admin Error: " & xxNoMeth
		end if
		rs.Close
	end if
	if shipType=3 then
		sXML = "<"&international&"RateRequest USERID="""&uspsUser&""" PASSWORD="""&uspsPw&""">"
	elseif shipType=4 then
		sXML = "<?xml version=""1.0""?><AccessRequest xml:lang=""en-US""><AccessLicenseNumber>"&upsAccess&"</AccessLicenseNumber><UserId>"&upsUser&"</UserId><Password>"&upsPw&"</Password></AccessRequest><?xml version=""1.0""?>"
		sXML = sXML & "<RatingServiceSelectionRequest xml:lang=""en-US""><Request><TransactionReference><CustomerContext>Rating and Service</CustomerContext><XpciVersion>1.0001</XpciVersion></TransactionReference>"
		sXML = sXML & "<RequestAction>Rate</RequestAction><RequestOption>shop</RequestOption></Request>"
		if upspickuptype<>"" then sXML = sXML & "<PickupType><Code>"&upspickuptype&"</Code></PickupType>"
		sXML = sXML & "<Shipment><Shipper><Address><PostalCode>"&origZip&"</PostalCode><CountryCode>"&origCountryCode&"</CountryCode></Address></Shipper>"
		sXML = sXML & "<ShipTo><Address><PostalCode>"&destZip&"</PostalCode><CountryCode>"&shipCountryCode&"</CountryCode>" & IIfVr(commerciallocpost<>"Y", "<ResidentialAddress/>", "") & "</Address></ShipTo>"
		'sXML = sXML & "<Service><Code>11</Code></Service>"
	elseif shipType=6 then
		sXML = " <?xml version=""1.0"" ?> <eparcel><language> en </language><ratesAndServicesRequest><merchantCPCID> "&adminCanPostUser&" </merchantCPCID><fromPostalCode> "&origZip&" </fromPostalCode><lineItems>"
	elseif shipType=7 then ' FedEx
		if packaging<>"" then packaging="FEDEX" & UCase(packaging) else packaging="YOURPACKAGING"
		sXML = "<?xml version=""1.0"" encoding=""UTF-8"" ?>" & _
			"<FDXRateAvailableServicesRequest xmlns:api=""http://www.fedex.com/fsmapi"" xmlns:xsi=""http://www.w3.org/2001/XMLSchemainstance"" xsi:noNamespaceSchemaLocation=""FDXRateAvailableServicesRequest.xsd""><RequestHeader>" & _
			"<CustomerTransactionIdentifier>ecommerceplusrate</CustomerTransactionIdentifier><AccountNumber>"&fedexaccount&"</AccountNumber><MeterNumber>"&fedexmeter&"</MeterNumber><CarrierCode></CarrierCode></RequestHeader>" & _
			"<DropoffType>REGULARPICKUP</DropoffType><Packaging>"&packaging&"</Packaging>" & _
			"<WeightUnits>"&IIfVr((adminUnits AND 1)=1,"LBS","KGS")&"</WeightUnits><OriginAddress>"
		if origCountryCode="US" OR origCountryCode="CA" then sXML = sXML & "<StateOrProvinceCode>"&originstatecode&"</StateOrProvinceCode>"
		sXML = sXML & "<PostalCode>"&origZip&"</PostalCode><CountryCode>"&origCountryCode&"</CountryCode></OriginAddress><DestinationAddress>"
		if shipCountryCode="US" OR shipCountryCode="CA" then sXML = sXML & "<StateOrProvinceCode>"&shipStateAbbrev&"</StateOrProvinceCode>"
		sXML = sXML & "<PostalCode>"&destZip&"</PostalCode><CountryCode>"&shipCountryCode&"</CountryCode></DestinationAddress>" & _
			"<Payment><PayorType>SENDER</PayorType></Payment><SpecialServices>"
		sXML = sXML & "<ResidentialDelivery>" & IIfVr(commerciallocpost="Y","false","true") & "</ResidentialDelivery>"
		if saturdaydelivery="Y" then sXML = sXML & "<SaturdayDelivery>true</SaturdayDelivery>"
		if saturdaypickup=TRUE then sXML = sXML & "<SaturdayPickup>true</SaturdayPickup>"
		if insidedelivery="Y" then sXML = sXML & "<InsideDelivery>true</InsideDelivery>"
		if insidepickup=TRUE then sXML = sXML & "<InsidePickup>true</InsidePickup>"
		if payproviderpost<>"" then
			if int(payproviderpost)=codpaymentprovider then sXML = sXML & "<COD><CollectionAmount>XXXFILLCODAMTHEREYYY</CollectionAmount><CollectionType>ANY</CollectionType></COD>"
		end if
		if signaturerelease="Y" AND allowsignaturerelease=TRUE then
		elseif signatureoption="indirect" then
			sXML = sXML & "<SignatureOption>INDIRECT</SignatureOption>"
		elseif signatureoption="direct" then
			sXML = sXML & "<SignatureOption>DIRECT</SignatureOption>"
		elseif signatureoption="adult" then
			sXML = sXML & "<SignatureOption>ADULT</SignatureOption>"
		elseif signatureoption="none" then
			sXML = sXML & "<SignatureOption>NONE</SignatureOption>"
		end if
		sXML = sXML & "</SpecialServices>"
		if homedelivery<>"" then sXML = sXML & "<HomeDelivery><Type>"&homedelivery&"</Type></HomeDelivery>"
	end if
end sub
totalpackdims = Array(0,0,0,0) ' len : wid : hei : vol used
sub addpackagedimensions(dimens)
	Session.LCID = 1033
	if (adminUnits AND 12)<>0 then
		origdimens = totalpackdims
		' response.write "adding package dimensions " & dimens & "<br>"
		proddims = split(dimens&"", "x")
		if UBOUND(proddims)>=0 then if proddims(0)<>"" then thelength = cDbl(proddims(0))
		if UBOUND(proddims)>=1 then if proddims(1)<>"" then thewidth = cDbl(proddims(1))
		if UBOUND(proddims)>=2 then if proddims(2)<>"" then theheight =  cDbl(proddims(2))
		if thelength<>"" AND thewidth<>"" AND theheight<>"" then
			objvol = thelength * thewidth * theheight
			if thelength > totalpackdims(0) then totalpackdims(0) = thelength
			if thewidth > totalpackdims(1) then totalpackdims(1) = thewidth
			if theheight > totalpackdims(2) then totalpackdims(2) = theheight
			if objvol + totalpackdims(3) > totalpackdims(0) * totalpackdims(1) * totalpackdims(2) then totalpackdims(2) = totalpackdims(2) + IIfVr(origdimens(2) > 0 AND origdimens(2) < theheight, origdimens(2),theheight)
			if objvol + totalpackdims(3) > totalpackdims(0) * totalpackdims(1) * totalpackdims(2) then totalpackdims(1) = totalpackdims(1) + IIfVr(origdimens(1) > 0 AND origdimens(1) < thewidth, origdimens(1),thewidth)
			if objvol + totalpackdims(3) > totalpackdims(0) * totalpackdims(1) * totalpackdims(2) then totalpackdims(0) = totalpackdims(0) + IIfVr(origdimens(0) > 0 AND origdimens(0) < thelength, origdimens(0),thelength)
			totalpackdims(3) = totalpackdims(3) + objvol
			if totalpackdims(2) > totalpackdims(1) then apdtemp = totalpackdims(1) : totalpackdims(1) = totalpackdims(2) : totalpackdims(2) = apdtemp
			if totalpackdims(1) > totalpackdims(0) then apdtemp = totalpackdims(0) : totalpackdims(0) = totalpackdims(1) : totalpackdims(1) = apdtemp
			if totalpackdims(2) > totalpackdims(1) then apdtemp = totalpackdims(1) : totalpackdims(1) = totalpackdims(2) : totalpackdims(2) = apdtemp
		end if
	end if
	' response.write "Bin is : " & totalpackdims(0)&":"& totalpackdims(1)&":"& totalpackdims(2)&" = " & (totalpackdims(0)*totalpackdims(1)*totalpackdims(2)) & "<br>"
	Session.LCID = saveLCID
end sub
sub addproducttoshipping(apsrs, prodindex)
	call addpackagedimensions(apsrs(11,prodindex))
	if packtogether then iTotItems = 1 else iTotItems = iTotItems + 1
	shipThisProd=true
	if (apsrs(8,prodindex) AND 4)=4 then ' No Shipping on this product
		if NOT packtogether then iTotItems = iTotItems - Int(apsrs(4,prodindex))
		shipThisProd=false
	end if
	if shipType=1 then ' Flat rate shipping
		if shipThisProd then shipping = shipping + apsrs(6,prodindex) + (apsrs(7,prodindex) * (apsrs(4,prodindex)-1))
	elseif (shipType=2 OR shipType=5) AND shippingpost="" then ' Weight / Price based shipping
		havematch=false
		for index3=0 to numshipoptions
			dHighest(index3)=0
		next
		if IsArray(allzones) then
			if shipThisProd then
				somethingToShip=true
				if shipType=2 then tmpweight = cDbl(apsrs(5,prodindex)) else tmpweight = cDbl(apsrs(3,prodindex))
				if packtogether then
					thePWeight = thePWeight + (cDbl(apsrs(4,prodindex))*tmpweight)
					thePQuantity = 1
				else
					thePWeight = tmpweight
					thePQuantity = cDbl(apsrs(4,prodindex))
				end if
			end if
			if ((NOT packtogether AND shipThisProd) OR (packtogether AND prodindex=UBOUND(apsrs,2))) AND somethingToShip then ' Only calculate pack together when we have the total
				for index2=0 to UBOUND(allzones,2)
					if allzones(0,index2)>=thePWeight then
						havematch=true
						for index3=0 to numshipoptions
							if cint(allzones(6+index3,index2))<>0 then ' by percentage
								intShipping(2,index3)=intShipping(2,index3)+((cDbl(allzones(1+index3,index2))*thePQuantity*thePWeight)/100.0)
							else
								intShipping(2,index3)=intShipping(2,index3)+(cDbl(allzones(1+index3,index2))*thePQuantity)
							end if
							if cDbl(allzones(1+index3,index2))=-99999.0 then intShipping(3,index3)=FALSE
						next
						exit for
					end if
					dHighWeight=allzones(0,index2)
					for index3=0 to numshipoptions
						if cint(allzones(6+index3,index2))<>0 then ' by percentage
							dHighest(index3)=(allzones(1+index3,index2)*dHighWeight)/100.0
						else
							dHighest(index3)=allzones(1+index3,index2)
						end if
					next
				next
				if NOT havematch then
					for index3=0 to numshipoptions
						intShipping(2,index3) = intShipping(2,index3) + dHighest(index3)
						if dHighest(index3)=-99999.0 then intShipping(3,index3)=FALSE
					next
					if allzones(0,0) < 0 then
						dHighWeight = thePWeight - dHighWeight
						do while dHighWeight > 0
							for index3=0 to numshipoptions
								intShipping(2,index3) = intShipping(2,index3) + (cDbl(allzones(1+index3,0))*thePQuantity)
							next
							dHighWeight = vsround(dHighWeight + allzones(0,0),4)
						loop
					end if
				end if
				for index3=numshipoptions to 0 step-1
					if intShipping(3,index3)=FALSE then
						for index4=index3+1 to numshipoptions
							intShipping(0,index4-1)=intShipping(0,index4)
							intShipping(2,index4-1)=intShipping(2,index4)
							intShipping(3,index4-1)=intShipping(3,index4)
						next
						numshipoptions = numshipoptions-1
					end if
				next
			end if
		end if
	elseif shipType=3 AND shippingpost="" then ' USPS Shipping
		if packtogether then
			if shipThisProd then
				somethingToShip=true
				iWeight = iWeight + (cDbl(apsrs(5,prodindex)) * Int(apsrs(4,prodindex)))
			end if
			if prodindex = UBOUND(apsrs,2) AND somethingToShip then
				numpacks=1
				if splitpackat<>"" then
					if iWeight > splitpackat then numpacks=-Int(-(iWeight/splitpackat))
				end if
				if numpacks > 1 then
					if international <> "" then
						sXML = sXML & addUSPSInternational(rowcounter,splitpackat,numpacks-1,"Package",shipcountry)
					else
						sXML = sXML & addUSPSDomestic(rowcounter,"Parcel",origZip,destZip,splitpackat,numpacks-1,"None","REGULAR","True")
					end if
					iTotItems = iTotItems + 1
					iWeight = iWeight - (splitpackat*(numpacks-1))
					rowcounter = rowcounter + 1
				end if
				if international <> "" then
					sXML = sXML & addUSPSInternational(rowcounter,iWeight,1,"Package",shipcountry)
				else
					sXML = sXML & addUSPSDomestic(rowcounter,"Parcel",origZip,destZip,iWeight,1,"None","REGULAR","True")
				end if
				rowcounter = rowcounter + 1
			end if
		else
			if shipThisProd then
				somethingToShip=true
				iWeight=apsrs(5,prodindex)
				numpacks=1
				if splitpackat<>"" then
					if iWeight > splitpackat then numpacks=-Int(-(iWeight/splitpackat))
				end if
				if numpacks > 1 then
					if international <> "" then
						sXML = sXML & addUSPSInternational(rowcounter,splitpackat,apsrs(4,prodindex)*(numpacks-1),"Package",shipcountry)
					else
						sXML = sXML & addUSPSDomestic(rowcounter,"Parcel",origZip,destZip,splitpackat,apsrs(4,prodindex)*(numpacks-1),"None","REGULAR","True")
					end if
					iTotItems = iTotItems + 1
					iWeight = iWeight - (splitpackat*(numpacks-1))
					rowcounter = rowcounter + 1
				end if
				if international <> "" then
					sXML = sXML & addUSPSInternational(rowcounter,iWeight,apsrs(4,prodindex),"Package",shipcountry)
				else
					sXML = sXML & addUSPSDomestic(rowcounter,"Parcel",origZip,destZip,iWeight,apsrs(4,prodindex),"None","REGULAR","True")
				end if
				rowcounter = rowcounter + 1
			end if
		end if
	elseif (shipType=4 OR shipType=6) AND shippingpost="" then ' UPS Shipping OR Canada Post
		Session.LCID = 1033
		if packaging<>"" then
			if packaging="envelope" then packaging="01"
			if packaging="pak" then packaging="04"
			if packaging="box" then packaging="21"
			if packaging="tube" then packaging="03"
			if packaging="10kgbox" then packaging="25"
			if packaging="25kgbox" then packaging="24"
		elseif upspacktype<>"" then
			packaging=upspacktype
		else
			packaging="02"
		end if
		if packtogether then
			if shipThisProd then
				somethingToShip=true
				iWeight = iWeight + (cDbl(apsrs(5,prodindex)) * Int(apsrs(4,prodindex)))
			end if
			if prodindex = UBOUND(apsrs,2) AND somethingToShip then
				numpacks=1
				if splitpackat<>"" then
					if iWeight > splitpackat then numpacks=-Int(-(iWeight/splitpackat))
				end if
				for index3 = 1 to numpacks
					if shipType=4 then
						sXML = sXML & addUPSInternational(iWeight / numpacks,adminUnits,packaging,shipCountryCode,totalgoods-shipfreegoods,totalpackdims)
					else
						sXML = sXML & addCanadaPostPackage(iWeight / numpacks,adminUnits,packaging,shipCountryCode,totalgoods-shipfreegoods,totalpackdims)
					end if
				next
			end if
		else
			if shipThisProd then
				somethingToShip=true
				iWeight=apsrs(5,prodindex)
				numpacks=1
				if splitpackat<>"" then
					if iWeight > splitpackat then numpacks=-Int(-(iWeight/splitpackat))
				end if
				for index2=0 to Int(apsrs(4,prodindex))-1
					for index3 = 1 to numpacks
						if shipType=4 then
							sXML = sXML & addUPSInternational(iWeight / numpacks,adminUnits,packaging,shipCountryCode,apsrs(3,prodindex),totalpackdims)
						else
							sXML = sXML & addCanadaPostPackage(iWeight / numpacks,adminUnits,packaging,shipCountryCode,apsrs(3,prodindex),totalpackdims)
						end if
					next
				next
			end if
		end if
		Session.LCID = saveLCID
	elseif shipType=7 AND shippingpost="" then ' FedEx
		Session.LCID = 1033
		if packtogether then
			totalshipitems=1
			if shipThisProd then
				somethingToShip=true
				iWeight = iWeight + (cDbl(apsrs(5,prodindex)) * Int(apsrs(4,prodindex)))
			end if
		else
			if shipThisProd then
				somethingToShip=true
				iWeight = iWeight + (cDbl(apsrs(5,prodindex)) * Int(apsrs(4,prodindex)))
				if splitpackat<>"" then
					if cDbl(apsrs(5,prodindex)) > splitpackat then totalshipitems=totalshipitems + (-Int(-(cDbl(apsrs(5,prodindex))/splitpackat)) * Int(apsrs(4,prodindex))) else totalshipitems=totalshipitems + Int(apsrs(4,prodindex))
				else
					totalshipitems=totalshipitems + Int(apsrs(4,prodindex))
				end if
			end if
		end if
		if prodindex = UBOUND(apsrs,2) AND somethingToShip then
			if packtogether AND splitpackat<>"" then
				if iWeight > splitpackat then totalshipitems = (-Int(-(iWeight/splitpackat)))
			end if
			sXML = sXML & addFedexPackage(iWeight,totalshipitems,totalgoods-shipfreegoods,totalpackdims)
		end if
		Session.LCID = saveLCID
	end if
end sub
function calculateshipping()
	if shipType=1 then
		isstandardship = true
	elseif (shipType=2 OR shipType=5) AND (somethingToShip OR willpickuptext<>"") then
		checkIntOptions = (shippingpost="")
		if IsArray(allzones) AND numshipoptions>=0 then
			shipping = intShipping(2,0)
			shipMethod = intShipping(0,0)
			isstandardship = ((pzFSA AND 1) = 1)
			if numshipoptions = 0 AND willpickuptext="" then checkIntOptions = FALSE
		else
			if willpickuptext<>"" then
				if willpickupcost<>"" then shipping = willpickupcost
				shipMethod = willpickuptext
			else
				success = FALSE
				errormsg=xxNoMeth
				checkIntOptions = false
			end if
		end if
	elseif shipType=3 AND somethingToShip then
		checkIntOptions = (shippingpost="")
		if shippingpost="" then
			sXML = sXML & "</"&international&"RateRequest>"
			success = USPSCalculate(sXML,international,shipping, errormsg, intShipping)
			if left(errormsg, 30)="Warning - Bound Printed Matter" then success=true
			if success AND checkIntOptions then ' Look for a single valid shipping option
				totShipOptions = 0
				for index=0 to UBOUND(intShipping,2)
					if iTotItems=intShipping(3,index) then
						for index2=0 to UBOUND(uspsmethods,2)
							if replace(lcase(intShipping(0,index)),"-"," ") = replace(lcase(uspsmethods(0,index2)),"-"," ") then
								if totShipOptions=0 then
									shipping = intShipping(2,index)
									shipMethod = Trim(uspsmethods(2,index2))
									isstandardship = Int(uspsmethods(1,index2))
								end if
								intShipping(5,index)=uspsmethods(2,index2)
								totShipOptions = totShipOptions + 1
							end if
						next
					end if
				next
				if totShipOptions=1 then
					checkIntOptions=False
				elseif totShipOptions=0 AND willpickuptext="" then
					checkIntOptions=False
					success=False
					errormsg=xxNoMeth
				end if
				if willpickuptext<>"" then checkIntOptions = True
			end if
		end if
	elseif shipType=4 AND somethingToShip then
		checkIntOptions = (shippingpost="")
		if shippingpost="" then
			sXML = sXML & "<ShipmentServiceOptions>" & IIfVr(saturdaydelivery="Y","<SaturdayDelivery/>","") & IIfVr(saturdaypickup=TRUE,"<SaturdayPickup/>","") & "</ShipmentServiceOptions></Shipment></RatingServiceSelectionRequest>"
			if Trim(upsUser)<>"" AND Trim(upsPw)<>"" then
				success = UPSCalculate(sXML,international,shipping, errormsg, intShipping)
			else
				success = false
				errormsg = "You must register with UPS by logging on to your online admin section and clicking the &quot;Register with UPS&quot; link before you can use the UPS OnLine&reg; Shipping Rates and Services Selection"
			end if
		end if
	elseif shipType=6 AND somethingToShip then
		checkIntOptions = (shippingpost="")
		if shippingpost="" then
			sXML = sXML & " </lineItems><city> </city> "
			if shipstate<>"" then
				sXML = sXML & "<provOrState> "&shipstate&" </provOrState>"
			else
				if shipCountryCode="US" OR shipCountryCode="CA" then
					thestate = IIfVr(trim(Request.form("sname")) <> "" OR trim(Request.form("saddress")) <> "", trim(request.form("sstate2")), trim(request.form("state2")))
					if thestate="" then thestate=IIfVr(shipCountryCode="US","CA","QC")
					sXML = sXML & "<provOrState> "&thestate&" </provOrState>"
				else
					sXML = sXML & "<provOrState> </provOrState>"
				end if
			end if
			sXML = sXML & "<country>"&shipCountryCode&"</country><postalCode>"&destZip&"</postalCode></ratesAndServicesRequest></eparcel>"
			success = CanadaPostCalculate(sXML,international,shipping, errormsg, intShipping)
		end if
	elseif shipType=7 AND somethingToShip then
		checkIntOptions = (shippingpost="")
		if shippingpost="" then
			sXML = sXML & "</FDXRateAvailableServicesRequest>"
			success = fedexcalculate(sXML,international, errormsg, intShipping)
		end if
	end if
	if success AND shippingpost="" AND somethingToShip AND (shipType=4 OR shipType=6 OR shipType=7) then
		totShipOptions = 0
		for index=0 to UBOUND(intShipping,2)
			if intShipping(3,index)=true then
				totShipOptions = totShipOptions + 1
				if index=0 then
					shipping = intShipping(2,index)
					shipMethod = intShipping(0,index)
					isstandardship = intShipping(4,index)
				end if
			end if
		next
		if totShipOptions=1 then
			checkIntOptions=False
		elseif totShipOptions=0 AND willpickuptext="" then
			checkIntOptions=False
			success=False
			errormsg=xxNoMeth
		end if
		if willpickuptext<>"" then checkIntOptions = True
	end if
	calculateshipping = success
end function
sub insuranceandtaxaddedtoshipping()
	if IsNumeric(shipinsuranceamt) AND shippingpost="" AND somethingToShip then
		if (wantinsurance="Y" AND addshippinginsurance=2) OR addshippinginsurance=1 then
			for index3=0 to UBOUND(intShipping,2)
				intShipping(2,index3) = intShipping(2,index3) + ((cDbl(totalgoods)*cDbl(shipinsuranceamt))/100.0)
			next
			shipping = shipping + ((cDbl(totalgoods)*cDbl(shipinsuranceamt))/100.0)
		elseif (wantinsurance="Y" AND addshippinginsurance=-2) OR addshippinginsurance=-1 then
			for index3=0 to UBOUND(intShipping,2)
				intShipping(2,index3) = intShipping(2,index3) + shipinsuranceamt
			next
			shipping = shipping + shipinsuranceamt
		end if
	end if
	if taxShipping=1 AND shippingpost="" then
		for index3=0 to UBOUND(intShipping,2)
			intShipping(2,index3) = intShipping(2,index3) + (cDbl(intShipping(2,index3))*(cDbl(stateTaxRate)+cDbl(countryTaxRate)))/100.0
		next
		shipping = shipping + (cDbl(shipping)*(cDbl(stateTaxRate)+cDbl(countryTaxRate)))/100.0
	end if
end sub
sub calculatetaxandhandling()
	if handlingchargepercent<>"" then handling = handling + (((totalgoods + shipping + handling) - (totaldiscounts + freeshipamnt)) * handlingchargepercent / 100.0)
	if taxHandling=1 then handling = handling + (cDbl(handling)*(cDbl(stateTaxRate)+cDbl(countryTaxRate)))/100.0
	if canadataxsystem=true AND shipCountryID=2 AND (shipStateAbbrev="NB" OR shipStateAbbrev="NF" OR shipStateAbbrev="NS") then usehst=true else usehst=false
	if canadataxsystem=true AND shipCountryID=2 AND (shipStateAbbrev="PE" OR shipStateAbbrev="QC") then
		statetaxable = 0
		countrytaxable = 0
		if taxShipping=2 AND (shipping - freeshipamnt > 0) then
			if proratashippingtax=TRUE then
				if totalgoods > 0 then statetaxable = statetaxable + (((cDbl(totalgoods)-(cDbl(totaldiscounts)+cDbl(statetaxfree))) / totalgoods) * (cDbl(shipping)-cDbl(freeshipamnt)))
			else
				statetaxable = statetaxable + (cDbl(shipping)-cDbl(freeshipamnt))
			end if
			countrytaxable = countrytaxable + (cDbl(shipping)-cDbl(freeshipamnt))
		end if
		if taxHandling=2 then
			statetaxable = statetaxable + cDbl(handling)
			countrytaxable = countrytaxable + cDbl(handling)
		end if
		if totalgoods>0 then
			statetaxable = statetaxable + (cDbl(totalgoods)-(cDbl(totaldiscounts)+cDbl(statetaxfree)))
			countrytaxable = countrytaxable + (cDbl(totalgoods)-(cDbl(totaldiscounts)+cDbl(countrytaxfree)))
		end if
		countryTax = countrytaxable*cDbl(countryTaxRate)/100.0
		stateTax = (statetaxable+cDbl(countryTax))*cDbl(stateTaxRate)/100.0
	else
		if totalgoods>0 then
			stateTax = ((cDbl(totalgoods)-(cDbl(totaldiscounts)+cDbl(statetaxfree)))*cDbl(stateTaxRate)/100.0)
			if perproducttaxrate<>TRUE then countryTax = ((cDbl(totalgoods)-(cDbl(totaldiscounts)+cDbl(countrytaxfree)))*cDbl(countryTaxRate)/100.0)
		end if
		if taxShipping=2 AND (shipping - freeshipamnt > 0) then
			if proratashippingtax=TRUE then
				if totalgoods>0 then stateTax = stateTax + (((cDbl(totalgoods)-(cDbl(totaldiscounts)+cDbl(statetaxfree))) / totalgoods) * (cDbl(shipping)-cDbl(freeshipamnt))*(cDbl(stateTaxRate)/100.0))
			else
				stateTax = stateTax + (cDbl(shipping)-cDbl(freeshipamnt))*(cDbl(stateTaxRate)/100.0)
			end if
			countryTax = countryTax + (cDbl(shipping)-cDbl(freeshipamnt))*(cDbl(countryTaxRate)/100.0)
		end if
		if taxHandling=2 then
			stateTax = stateTax + cDbl(handling)*(cDbl(stateTaxRate)/100.0)
			countryTax = countryTax + cDbl(handling)*(cDbl(countryTaxRate)/100.0)
		end if
	end if
	if stateTax < 0 then stateTax = 0
	if countryTax < 0 then countryTax = 0
end sub
if stockManage<>0 then
	tdt = DateAdd("h",dateadjust-stockManage,now())
	sSQL = "SELECT cartOrderID,cartID FROM cart WHERE (cartCompleted=0 AND cartOrderID=0 AND cartDateAdded<" & datedelim & VSUSDateTime(tdt) & datedelim & ")"
	if delAfter<>0 then
		tdt = Date()-delAfter
		sSQL = sSQL & " OR (cartCompleted=0 AND cartDateAdded<"&datedelim & VSUSDate(tdt) & datedelim & ")"
	end if
	rs.Open sSQL,cnn,0,1
	if NOT rs.EOF then
		addcomma=""
		delstr=""
		delcart=""
		do while NOT rs.EOF
			delcart = delcart & addcomma & rs("cartOrderID")
			delstr = delstr & addcomma & rs("cartID")
			addcomma=","
			rs.MoveNext
		loop
		if delAfter<>0 then cnn.Execute("DELETE FROM orders WHERE ordID IN ("&delcart&")")
		cnn.Execute("DELETE FROM cart WHERE cartID IN ("&delstr&")")
		cnn.Execute("DELETE FROM cartoptions WHERE coCartID IN ("&delstr&")")
	end if
	rs.Close
end if
if request.querystring("token") <> "" then
	call getpayprovdetails(18,username,data2pwd,data2hash,demomode,ppmethod)
	sXML = ppsoapheader(username, data2pwd, data2hash) & _
		"<soap:Body><GetExpressCheckoutDetailsReq xmlns=""urn:ebay:api:PayPalAPI""><GetExpressCheckoutDetailsRequest><Version xmlns=""urn:ebay:apis:eBLBaseComponents"">1.00</Version>" & _
		"  <Token>" & request.querystring("token") & "</Token>" & _
		"</GetExpressCheckoutDetailsRequest></GetExpressCheckoutDetailsReq></soap:Body></soap:Envelope>"
	if demomode then sandbox = ".sandbox" else sandbox = ""
	if callxmlfunction("https://api-aa" & IIfVr(sandbox="" AND data2hash<>"", "-3t", "") & sandbox & ".paypal.com/2.0/", sXML, res, IIfVr(data2hash<>"","",username), "WinHTTP.WinHTTPRequest.5.1", errormsg, FALSE) then
		countryid=0
		success = FALSE
		ordPayProvider = "19"
		ordEmail = "" : insidedelivery = "" : commercialloc = "" : wantinsurance = "" : saturdaydelivery = "" : signaturerelease = ""
		ordComLoc = 0
		gotaddress = FALSE
		token = request.querystring("token")
		if abs(addshippinginsurance)=1 then ordComLoc = ordComLoc + 2
		set xmlDoc = Server.CreateObject("MSXML2.DOMDocument")
		xmlDoc.validateOnParse = False
		xmlDoc.loadXML (res)
		Set nodeList = xmlDoc.getElementsByTagName("SOAP-ENV:Body")
		Set n = nodeList.Item(0)
		for j = 0 to n.childNodes.length - 1
			Set e = n.childNodes.Item(i)
			if e.nodeName = "GetExpressCheckoutDetailsResponse" then
				for k = 0 To e.childNodes.length - 1
					Set t = e.childNodes.Item(k)
					if t.nodeName = "Ack" then
						if t.firstChild.nodeValue = "Success" OR t.firstChild.nodeValue = "SuccessWithWarning" then success=TRUE
					elseif t.nodeName = "GetExpressCheckoutDetailsResponseDetails" then
						set ff = t.childNodes
						for kk = 0 to ff.length - 1
							set gg = ff.item(kk)
							if gg.nodeName = "PayerInfo" then
								set hh = gg.childNodes
								for ll = 0 to hh.length - 1
									set ii = hh.item(ll)
									if ii.nodeName = "Payer" then
										if ii.hasChildNodes then ordEmail = ii.firstChild.nodeValue
									elseif ii.nodeName = "PayerID" then
										if ii.hasChildNodes then payerid = ii.firstChild.nodeValue
									elseif ii.nodeName = "PayerStatus" then
										if ii.hasChildNodes then
											ordCVV = "U"
											payer_status = lcase(ii.firstChild.nodeValue)
											if payer_status="verified" then ordCVV = "Y"
											if payer_status="unverified" then ordCVV = "N"
										end if
									elseif ii.nodeName = "PayerName" then
										set jj = ii.childNodes
										for mm = 0 to jj.length - 1
											set jjj = jj.item(mm)
											if jjj.nodeName = "FirstName" then
												if jjj.hasChildNodes then ordName = jjj.firstChild.nodeValue & IIfVr(ordName<>"", " " & ordName, ordName)
											elseif jjj.nodeName = "LastName" then
												if jjj.hasChildNodes then ordName = IIfVr(ordName<>"", ordName&" ",ordName) & jjj.firstChild.nodeValue
											end if
										next
									elseif ii.nodeName = "Address" then
										set jj = ii.childNodes
										for mm = 0 to jj.length - 1
											set jjj = jj.item(mm)
											if jjj.nodeName = "Street1" then
												if jjj.hasChildNodes then ordAddress = jjj.firstChild.nodeValue
											elseif jjj.nodeName = "Street2" then
												if jjj.hasChildNodes then ordAddress2 = jjj.firstChild.nodeValue
											elseif jjj.nodeName = "CityName" then
												if jjj.hasChildNodes then ordCity = jjj.firstChild.nodeValue
											elseif jjj.nodeName = "StateOrProvince" then
												if jjj.hasChildNodes then ordState = jjj.firstChild.nodeValue
											elseif jjj.nodeName = "Country" then
												if jjj.hasChildNodes then
													sSQL = "SELECT countryName,countryID FROM countries WHERE countryCode='" & replace(jjj.firstChild.nodeValue, "'", "''") & "'"
													rs.Open sSQL,cnn,0,1
														ordCountry = rs("countryName")
														countryid = rs("countryID")
													rs.Close
												end if
											elseif jjj.nodeName = "PostalCode" then
												if jjj.hasChildNodes then ordZip = jjj.firstChild.nodeValue
											elseif jjj.nodeName = "AddressStatus" then
												if jjj.hasChildNodes then
													ordAVS = "U"
													address_status = lcase(jjj.firstChild.nodeValue)
													gotaddress = (address_status<>"none")
													if address_status="confirmed" then ordAVS = "Y"
													if address_status="unconfirmed" then ordAVS = "N"
												end if
											end if
										next
									end if
								next
							elseif gg.nodeName = "Custom" then
								customarr = split(gg.firstChild.nodeValue, ":")
								thesessionid = customarr(0)
								ordAffiliate = customarr(1)
							elseif gg.nodeName = "ContactPhone" then
								if gg.hasChildNodes then ordPhone = gg.firstChild.nodeValue
							end if
						next
					elseif t.nodeName = "Errors" then
						set ff = t.childNodes
						for kk = 0 to ff.length - 1
							set gg = ff.item(kk)
							if gg.nodeName = "ShortMessage" then
								errormsg = gg.firstChild.nodeValue & "<br>" & errormsg
							elseif gg.nodeName = "LongMessage" then
								errormsg= errormsg & gg.firstChild.nodeValue
							elseif gg.nodeName = "ErrorCode" then
								errcode = gg.firstChild.nodeValue
							end if
						next
					end if
				next
			end if
		next
		if NOT gotaddress then
			ppexpresscancel=TRUE
		elseif success then
			paypalexpress=TRUE
			if (countryid=1 OR countryid=2) AND usestateabbrev<>TRUE then
				sSQL = "SELECT stateName FROM states WHERE stateAbbrev='" & replace(ordState,"'","''") & "'"
				rs.Open sSQL,cnn,0,1
				if NOT rs.EOF then ordState = rs("stateName")
				rs.Close
			end if
		else
			response.write "PayPal Payment Pro error: " & errormsg
		end if
	else
		response.write "PayPal Payment Pro error: " & errormsg
	end if
elseif checkoutmode="paypalexpress1" then
	success = FALSE
	call getpayprovdetails(18,username,data2pwd,data2hash,demomode,ppmethod)
	if demomode then sandbox = ".sandbox" else sandbox = ""
	if pathtossl<>"" then
		if Right(pathtossl,1) <> "/" then storeurl = pathtossl & "/" else storeurl = pathtossl
	end if
	sXML = ppsoapheader(username, data2pwd, data2hash) & _
		"<soap:Body><SetExpressCheckoutReq xmlns=""urn:ebay:api:PayPalAPI""><SetExpressCheckoutRequest><Version xmlns=""urn:ebay:apis:eBLBaseComponents"">1.00</Version>" & _
		"  <SetExpressCheckoutRequestDetails xmlns=""urn:ebay:apis:eBLBaseComponents"">" & _
		"    <OrderTotal currencyID=""" & countryCurrency & """>" & request.form("estimate") & "</OrderTotal>" & _
		"    <ReturnURL>" & storeurl & "cart.asp</ReturnURL><CancelURL>" & storeurl & "cart.asp</CancelURL>" & _
		"    <Custom>" & thesessionid & ":" & request.form("PARTNER") & "</Custom>" & _
		"    <PaymentAction>" & IIfVr(ppmethod=1, "Authorization", "Sale") & "</PaymentAction>" & _
		"  </SetExpressCheckoutRequestDetails>" & _
		"</SetExpressCheckoutRequest></SetExpressCheckoutReq></soap:Body></soap:Envelope>"
	if callxmlfunction("https://api-aa" & IIfVr(sandbox="" AND data2hash<>"", "-3t", "") & sandbox & ".paypal.com/2.0/", sXML, res, IIfVr(data2hash<>"","",username), "WinHTTP.WinHTTPRequest.5.1", errormsg, FALSE) then
		set xmlDoc = Server.CreateObject("MSXML2.DOMDocument")
		xmlDoc.validateOnParse = False
		xmlDoc.loadXML (res)
		Set nodeList = xmlDoc.getElementsByTagName("SOAP-ENV:Body")
		Set n = nodeList.Item(0)
		for j = 0 to n.childNodes.length - 1
			Set e = n.childNodes.Item(i)
			if e.nodeName = "SetExpressCheckoutResponse" then
				for k = 0 To e.childNodes.length - 1
					Set t = e.childNodes.Item(k)
					if t.nodeName = "Ack" then
						if t.firstChild.nodeValue = "Success" OR t.firstChild.nodeValue = "SuccessWithWarning" then success=TRUE
					elseif t.nodeName = "Token" then
						token = t.firstChild.nodeValue
					elseif t.nodeName = "Errors" then
						set ff = t.childNodes
						for kk = 0 to ff.length - 1
							set gg = ff.item(kk)
							if gg.nodeName = "ShortMessage" then
								errormsg = gg.firstChild.nodeValue & "<br>" & errormsg
							elseif gg.nodeName = "LongMessage" then
								errormsg= errormsg & gg.firstChild.nodeValue
							elseif gg.nodeName = "ErrorCode" then
								errcode = gg.firstChild.nodeValue
							end if
						next
					end if
				next
			end if
		next
		if success then
			response.redirect "https://www" & sandbox & ".paypal.com/webscr?cmd=_express-checkout&token=" & token
			response.write "<p align=""center"">" & xxAutFo & "</p>"
			response.write "<p align=""center"">" & xxForAut & " <a href=""https://www" & sandbox & ".paypal.com/webscr?cmd=_express-checkout&token=" & token & """>" & xxClkHere & "</a></p>"
		else
			response.write "PayPal Payment Pro error: " & errormsg
		end if
	else
		response.write "PayPal Payment Pro error: " & errormsg
	end if
elseif checkoutmode="update" then
	if estimateshipping=TRUE then session("xsshipping") = ""
	if NOT IsEmpty(session("discounts")) then session("discounts")=""
	if NOT IsEmpty(session("xscountrytax")) then session("xscountrytax")=""
	cnn.Execute("UPDATE orders SET ordTotal=0,ordShipping=0,ordStateTax=0,ordCountryTax=0,ordHSTTax=0,ordHandling=0,ordDiscount=0,ordDiscountText='' WHERE ordSessionID="&Session.SessionID&" AND ordAuthNumber=''")
	for each objItem In Request.Form
		thequant = Trim(Request.form(objItem))
		if NOT IsNumeric(thequant) then thequant=0 else thequant=abs(int(thequant))
		if Left(objItem,5)="quant" AND thequant<>"" then
			thecartid = int(Right(objItem, Len(objItem)-5))
			if thequant=0 then
				sSQL="DELETE FROM cartoptions WHERE coCartID="&thecartid
				cnn.Execute(sSQL)
				sSQL="DELETE FROM cart WHERE cartID="&thecartid
				cnn.Execute(sSQL)
			else
				totQuant = 0
				pPrice = 0
				pID = ""
				sSQL="SELECT cartQuantity,pInStock,pID,pStockByOpts,"&WSP&"pPrice FROM cart INNER JOIN products ON cart.cartProdId=products.pID WHERE cartID="&thecartid
				rs.Open sSQL,cnn,0,1
				if NOT rs.EOF then
					pID = rs("pID")
					pInStock = int(rs("pInStock"))
					pStockByOpts = cint(rs("pStockByOpts"))
					pPrice = rs("pPrice")
					cartQuantity = int(rs("cartQuantity"))
					rs.Close
					sSQL = "SELECT SUM(cartQuantity) AS cartQuant FROM cart WHERE cartCompleted=0 AND cartProdID='"&Trim(pID)&"'"
					rs.Open sSQL,cnn,0,1
					if NOT rs.EOF then
						if NOT IsNull(rs("cartQuant")) then totQuant = Int(rs("cartQuant"))
					end if
				end if
				rs.Close
				if pID<>"" then
					if stockManage<>0 then
						quantavailable = thequant
						if pStockByOpts <> 0 then
							hasalloptions=true
							sSQL = "SELECT coID,optStock,cartQuantity,coOptID FROM cart INNER JOIN (cartoptions INNER JOIN (options INNER JOIN optiongroup ON options.optGroup=optiongroup.optGrpID) ON cartoptions.coOptID=options.optID) ON cart.cartID=cartoptions.coCartID WHERE optType IN (-2,-1,1,2) AND cartID="&thecartid
							rs.Open sSQL,cnn,0,1
							if NOT rs.EOF then
								do while NOT rs.EOF
									pInStock = Int(rs("optStock"))
									totQuant = 0
									cartQuantity = Int(rs("cartQuantity"))
									sSQL = "SELECT SUM(cartQuantity) AS cartQuant FROM cart INNER JOIN cartoptions ON cart.cartID=cartoptions.coCartID WHERE cartCompleted=0 AND coOptID="&rs("coOptID")
									rs2.Open sSQL,cnn,0,1
									if NOT rs2.EOF then
										if NOT IsNull(rs2("cartQuant")) then totQuant = Int(rs2("cartQuant"))
									end if
									rs2.Close
									if Int(pInStock - totQuant + cartQuantity) < quantavailable then quantavailable = (pInStock - totQuant + cartQuantity)
									if (pInStock - totQuant + cartQuantity - thequant) < 0 then hasalloptions=false
									rs.MoveNext
								loop
								cnn.Execute("UPDATE cart SET cartQuantity="&quantavailable&" WHERE cartCompleted=0 AND cartID="&thecartid)
								if NOT hasalloptions then isInStock = false
							end if
							rs.Close
						else
							if (pInStock - totQuant + cartQuantity - thequant) < 0 then
								quantavailable = (pInStock - totQuant + cartQuantity)
								if quantavailable < 0 then quantavailable=0
								isInStock = false
							end if
							cnn.Execute("UPDATE cart SET cartQuantity="&quantavailable&" WHERE cartCompleted=0 AND cartID="&thecartid)
						end if
					else
						cnn.Execute("UPDATE cart SET cartQuantity="&thequant&" WHERE cartCompleted=0 AND cartID="&thecartid)
					end if
					call checkpricebreaks(pID,pPrice)
				end if
			end if
		elseif Left(objItem,5)="delet" then
			rs.Open "SELECT cartID FROM cart WHERE cartCompleted=0 AND cartID="&int(Right(objItem, Len(objItem)-5)),cnn,0,1
			if NOT rs.EOF then
				cnn.Execute("DELETE FROM cart WHERE cartID="&int(Right(objItem, Len(objItem)-5)))
				cnn.Execute("DELETE FROM cartoptions WHERE coCartID="&int(Right(objItem, Len(objItem)-5)))
			end if
			rs.Close
		end if
	next
end if
if checkoutmode="add" then
	if estimateshipping=TRUE then session("xsshipping") = ""
	if NOT IsEmpty(session("discounts")) then session("discounts")=""
	if NOT IsEmpty(session("xscountrytax")) then session("xscountrytax")=""
	cnn.Execute("UPDATE orders SET ordTotal=0,ordShipping=0,ordStateTax=0,ordCountryTax=0,ordHSTTax=0,ordHandling=0,ordDiscount=0,ordDiscountText='' WHERE ordSessionID="&Session.SessionID&" AND ordAuthNumber=''")
	Session.LCID = 1033
	if Trim(Request.Form("frompage"))<>"" then Session("frompage")=Request.Form("frompage") else Session("frompage")=""
	if Request.Form("quant")="" OR NOT IsNumeric(Request.Form("quant")) then
		quantity=1
	else
		quantity=abs(int(trim(Request.Form("quant"))))
	end if
	origquantity = quantity
	for jj = 1 to Request.Form.Count
		for each objElem in Request.Form
			if Request.Form(objElem) is Request.Form(jj) then objForm = objElem
		next
		if Left(objForm,4)="optn" AND trim(Request.Form(objForm))<>"" AND IsNumeric(trim(Request.Form(objForm))) then
			sSQL="SELECT optRegExp FROM options WHERE optID="&replace(Request.Form(objForm),"'","")
			rs2.Open sSQL,cnn,0,1
			if rs2.EOF then theexp="" else theexp = trim(rs2("optRegExp")&"")
			if theexp<>"" AND Left(theexp,1)<>"!" then
				theexp = replace(theexp, "%s", theid)
				if InStr(theexp, " ") > 0 then ' Search and replace
					exparr = split(theexp, " ", 2)
					theid = replace(theid, exparr(0), exparr(1), 1, 1)
				else
					theid = theexp
				end if
			end if
			rs2.Close
		end if
	next
	bExists=False
	sSQL = "SELECT cartID FROM cart WHERE cartCompleted=0 AND cartSessionID="&Session.SessionID&" AND cartProdID='"&theid&"'"
	rs.Open sSQL,cnn,0,1
	do while (NOT rs.EOF) AND (NOT bExists)
		bExists=True
		cartID=rs("cartID")
		for each objForm in Request.Form ' We have the product. Check we have all the same options
			if Left(objForm,4)="optn" then
				if trim(Request.Form("v"&objForm))<>"" then
					sSQL="SELECT coID FROM cartoptions WHERE coCartID="&cartID&" AND coOptID="&replace(Request.Form(objForm),"'","")&" AND coCartOption='"&replace(trim(Request.Form("v"&objForm)),"'","''")&"'"
					rs2.Open sSQL,cnn,0,1
					if rs2.EOF then bExists=false
					rs2.Close
				elseif trim(Request.Form(objForm))<>"" then
					sSQL="SELECT coID FROM cartoptions WHERE coCartID="&cartID&" AND coOptID="&replace(Request.Form(objForm),"'","")
					rs2.Open sSQL,cnn,0,1
					if rs2.EOF then bExists=false
					rs2.Close
				end if
			end if
			if NOT bExists then exit for
		next
		rs.MoveNext
	loop
	rs.Close
	sSQL = "SELECT "&getlangid("pName",1)&","&WSP&"pPrice,pInStock,pWeight,pStockByOpts FROM products WHERE pID='"&theid&"'"
	rs.Open sSQL,cnn,0,1
	if NOT rs.EOF then
		alldata=rs.getrows
	else
		redim alldata(1,1)
		alldata(0,0)=theid
		stockManage=0
		isInStock=false
		outofstockreason=2
	end if
	rs.Close
	if stockManage<>0 then
		bestDate = DateAdd("m",-2,now())
		outofstockreason=1
		if int(alldata(4,0)) <> 0 then
			for each objForm in Request.Form
				totQuant = 0
				if Left(objForm,4)="optn" AND trim(Request.Form(objForm))<>"" then
					sSQL="SELECT optStock FROM options INNER JOIN optiongroup ON options.optGroup=optiongroup.optGrpID WHERE optType IN (-2,-1,1,2) AND optID="&replace(Request.Form(objForm),"'","")
					rs.Open sSQL,cnn,0,1
					if NOT rs.EOF then stockQuant = rs("optStock") else stockQuant = origquantity
					rs.Close
					sSQL = "SELECT cartQuantity,cartDateAdded,cartOrderID FROM cart INNER JOIN (cartoptions INNER JOIN (options INNER JOIN optiongroup ON options.optGroup=optiongroup.optGrpID) ON cartoptions.coOptID=options.optID) ON cart.cartID=cartoptions.coCartID WHERE optType IN (-2,-1,1,2) AND cartCompleted=0 AND coOptID="&replace(Request.Form(objForm),"'","")&" ORDER BY cartDateAdded DESC"
					rs.Open sSQL,cnn,0,1
					do while NOT rs.EOF
						totQuant=totQuant+rs("cartQuantity")
						if Int(rs("cartOrderID"))=0 AND rs("cartDateAdded") > bestDate AND totQuant+stockQuant >= origquantity then bestDate = rs("cartDateAdded")
						rs.MoveNext
					loop
					rs.Close
					if stockQuant-totQuant < quantity then quantity = stockQuant-totQuant
					if (stockQuant+totQuant) < origquantity then outofstockreason=0
				end if
			next
		else
			totQuant = 0
			stockQuant = alldata(2,0)
			sSQL = "SELECT cartQuantity,cartDateAdded,cartOrderID FROM cart WHERE cartCompleted=0 AND cartProdID='"&theid&"' ORDER BY cartDateAdded DESC"
			rs.Open sSQL,cnn,0,1
			do while NOT rs.EOF
				totQuant=totQuant+rs("cartQuantity")
				if Int(rs("cartOrderID"))=0 AND rs("cartDateAdded") > bestDate AND totQuant+stockQuant >= origquantity then bestDate = rs("cartDateAdded")
				rs.MoveNext
			loop
			rs.Close
			if stockQuant-totQuant < quantity then quantity = stockQuant-totQuant
			if (stockQuant+totQuant) < origquantity then outofstockreason=0
		end if
		if quantity > 0 then isInStock = TRUE else isInStock = FALSE
	end if
	if isInStock then
		if bExists then
			sSQL = "UPDATE cart SET cartQuantity=cartQuantity+"&quantity&" WHERE cartCompleted=0 AND cartID="&cartID
			cnn.Execute(sSQL)
		else
			rs.Open "cart",cnn,1,3,&H0002
			rs.AddNew
			rs.Fields("cartSessionID")		= Session.SessionID
			rs.Fields("cartProdID")			= theid
			rs.Fields("cartQuantity")		= quantity
			rs.Fields("cartCompleted")		= 0
			rs.Fields("cartProdName")		= alldata(0,0)
			rs.Fields("cartProdPrice")		= alldata(1,0)
			rs.Fields("cartDateAdded")		= DateAdd("h",dateadjust,Now())
			rs.Update
			if mysqlserver=true then
				rs.Close
				rs.Open "SELECT LAST_INSERT_ID() AS lstIns",cnn,0,1
				cartID = rs("lstIns")
			else
				cartID = rs.Fields("cartID")
			end if
			rs.Close
			for jj = 1 to Request.Form.Count
				for each objElem in Request.Form
					if Request.Form(objElem) is Request.Form(jj) then objForm = objElem
				next
				if Left(objForm,4)="optn" then
					if Trim(Request.Form("v"&objForm))="" then
						if trim(Request.Form(objForm))<>"" then
							sSQL="SELECT optID,"&getlangid("optGrpName",16)&","&getlangid("optName",32)&","&OWSP&"optPriceDiff,optWeightDiff,optType,optFlags FROM options INNER JOIN optiongroup ON options.optGroup=optiongroup.optGrpID WHERE optID="&Replace(Request.Form(objForm),"'","")
							rs.Open sSQL,cnn,0,1
							if abs(rs("optType"))<> 3 then
								sSQL = "INSERT INTO cartoptions (coCartID,coOptID,coOptGroup,coCartOption,coPriceDiff,coWeightDiff) VALUES ("&cartID&","&rs("optID")&",'"&Replace(rs(getlangid("optGrpName",16))&"","'","''")&"','"&Replace(rs(getlangid("optName",32))&"","'","''")&"',"
								if (rs("optFlags") AND 1) = 0 then sSQL = sSQL & rs("optPriceDiff") & "," else sSQL = sSQL & vsround((rs("optPriceDiff")*alldata(1,0))/100.0, 2) & ","
								if (rs("optFlags") AND 2) = 0 then sSQL = sSQL & rs("optWeightDiff") & ")" else sSQL = sSQL & multShipWeight(alldata(3,0),rs("optWeightDiff")) & ")"
							else
								sSQL = "INSERT INTO cartoptions (coCartID,coOptID,coOptGroup,coCartOption,coPriceDiff,coWeightDiff) VALUES ("&cartID&","&rs("optID")&",'"&Replace(rs(getlangid("optGrpName",16))&"","'","''")&"','',0,0)"
							end if
							rs.Close
							cnn.Execute(sSQL)
						end if
					else
						sSQL="SELECT optID,"&getlangid("optGrpName",16)&","&getlangid("optName",32)&" FROM options INNER JOIN optiongroup ON options.optGroup=optiongroup.optGrpID WHERE optID="&replace(Request.Form(objForm),"'","")
						rs.Open sSQL,cnn,0,1
						sSQL = "INSERT INTO cartoptions (coCartID,coOptID,coOptGroup,coCartOption,coPriceDiff,coWeightDiff) VALUES ("&cartID&","&rs("optID")&",'"&Replace(rs(getlangid("optGrpName",16))&"","'","''")&"','"&replace(trim(Request.Form("v"&objForm)),"'","''")&"',0,0)"
						cnn.Execute(sSQL)
						rs.Close
					end if
				end if
			next
		end if
		call checkpricebreaks(theid, alldata(1,0))
%>
      <table border="0" cellspacing="<%=maintablespacing%>" cellpadding="<%=maintablepadding%>" width="<%=maintablewidth%>" bgcolor="<%=maintablebg%>" align="center">
        <tr> 
          <td width="100%" align="center">
            <table width="<%=innertablewidth%>" border="0" cellspacing="<%=innertablespacing%>" cellpadding="<%=innertablepadding%>" bgcolor="<%=innertablebg%>">
			  <tr>
			    <td align="center"><p>&nbsp;</p>
<%		if quantity < origquantity then
			response.write "<p><strong><font color=""#FF0000"">"&xxInsuff&"</font></strong></p><p>"&replace(xxOnlyAd,"%s",quantity)&"</p><p>"&xxWanRem&"</p>"
			response.write "<form method=""post"" action=""cart.asp""><input type=""hidden"" name=""delet" & cartID & """ value=""1""><input type=""hidden"" name=""mode"" value=""update""><input type=""submit"" value="""&xxDelete&"""> <input type=""button"" value="""&xxCntShp&""" onclick=""javascript:document.location=window.location='cart.asp'""></form>"
		else
			if cartrefreshseconds="" then cartrefreshseconds=3
			if Trim(Request.Form("frompage"))<>"" AND actionaftercart=3 then
				if cartrefreshseconds=0 then
					response.redirect trim(Request.Form("frompage"))
				else
					response.write "<meta http-equiv=""Refresh"" content="""&cartrefreshseconds&"; URL="&trim(Request.Form("frompage"))&""">"
				end if
			elseif actionaftercart=4 OR cartrefreshseconds=0 then
				response.redirect "cart.asp"&IIfVr(request.form("PARTNER")<>"","?PARTNER="&request.form("partner"),"")
			else
				response.write "<meta http-equiv=""Refresh"" content="""&cartrefreshseconds&"; URL=cart.asp"&IIfVr(request.form("PARTNER")<>"","?PARTNER="&request.form("partner"),"")&""">"
			end if
			response.write "<p>" & quantity & " <strong>" & alldata(0,0) & "</strong> "&xxAddOrd & "</p>"
			response.write "<p>" & xxPlsWait & " <a href="""
			if Trim(Request.Form("frompage"))<>"" AND actionaftercart=3 then response.write Trim(Request.Form("frompage")) else response.write "cart.asp"
			response.write """><strong>" & xxClkHere & "</strong></a>.</p>"
		end if %>
				<p>&nbsp;</p><p>&nbsp;</p>
				</td>
			  </tr>
			</table>
		  </td>
        </tr>
      </table>
<%
	else
%>
      <table border="0" cellspacing="<%=maintablespacing%>" cellpadding="<%=maintablepadding%>" width="<%=maintablewidth%>" bgcolor="<%=maintablebg%>" align="center">
        <tr> 
          <td width="100%" align="center">
            <table width="<%=innertablewidth%>" border="0" cellspacing="<%=innertablespacing%>" cellpadding="<%=innertablepadding%>" bgcolor="<%=innertablebg%>">
			  <tr>
			    <td align="center"><p>&nbsp;</p>
				<% response.write "<p>" & xxSrryItm & " <strong>" & alldata(0,0)&"</strong> " & xxIsCntly
				if outofstockreason=1 then response.write " " & xxTemprly
				if outofstockreason=2 then response.write " not available in our product database." else response.write " " & xxOutStck & "</p>"
				if outofstockreason=1 then
					response.write "<p>" & xxNotChOu & " "
					totMins = DateDiff("n",DateAdd("h",dateadjust,Now()),DateAdd("h",stockManage,bestDate))+1
					if totMins > 300 then
						response.write xxShrtWhl
					else
						if totMins >= 60 then response.write Int(totMins / 60) & " hour"
						if totMins >= 120 then response.write "s"
						totMins = totMins - (Int(totMins / 60) * 60)
						if totMins > 0 then response.write " " & totMins & " minute"
						if totMins > 1 then response.write "s"
					end if
					response.write xxChkBack & "</p>"
				end if %>
				<p><%=xxPlease%> <a href="javascript:history.go(-1)"><strong><%=xxClkHere%></strong></a> <%=xxToRetrn%></p>
				<p>&nbsp;</p>
				<p>&nbsp;</p>
				</td>
			  </tr>
			</table>
		  </td>
        </tr>
      </table>
<%
	end if
elseif checkoutmode="checkout" OR ppexpresscancel then
	Dim ordName,ordAddress,ordAddress2,ordCity,ordState,ordZip,ordCountry,ordEmail,ordPhone,ordShipName,ordShipAddress,ordShipAddress2,ordShipCity,ordShipState,ordShipZip,ordShipCountry,ordAddInfo
	Dim remember,allstates,havestate,allcountries
	allstates=""
	allcountries=""
	remember=False
	if request.form("checktmplogin")="1" then
		sSQL = "SELECT tmploginname FROM tmplogin WHERE tmploginid=" & replace(trim(request.form("sessionid")),"'","")
		rs.Open sSQL,cnn,0,1
		if NOT rs.EOF then
			Session("clientUser")=rs("tmploginname")
			rs.Close
			cnn.Execute("DELETE FROM tmplogin WHERE tmploginid=" & replace(trim(request.form("sessionid")),"'",""))
			sSQL = "SELECT clientActions,clientLoginLevel FROM clientlogin WHERE clientUser='"&replace(trim(session("clientUser")),"'","")&"'"
			rs.Open sSQL,cnn,0,1
			if NOT rs.EOF then
				Session("clientActions")=rs("clientActions")
				Session("clientLoginLevel")=rs("clientLoginLevel")
			end if
		end if
		rs.Close
	end if
	if request.cookies("id1")<>"" AND request.cookies("id2")<>"" AND IsNumeric(request.cookies("id1")) AND IsNumeric(request.cookies("id2")) then
		sSQL = "SELECT ordName,ordAddress,ordAddress2,ordCity,ordState,ordZip,ordCountry,ordEmail,ordPhone,ordShipName,ordShipAddress,ordShipAddress2,ordShipCity,ordShipState,ordShipZip,ordShipCountry,ordPayProvider,ordComLoc,ordExtra1,ordExtra2,ordExtra3,ordAddInfo FROM orders WHERE ordID="&request.cookies("id1")&" AND ordSessionID="&request.cookies("id2")
		rs.Open sSQL,cnn,0,1
		if NOT rs.EOF then
			ordName = rs("ordName")
			ordAddress = rs("ordAddress")
			ordAddress2 = rs("ordAddress2")
			ordCity = rs("ordCity")
			ordState = rs("ordState")
			ordZip = rs("ordZip")
			ordCountry = rs("ordCountry")
			ordEmail = rs("ordEmail")
			ordPhone = rs("ordPhone")
			ordShipName = rs("ordShipName")
			ordShipAddress = rs("ordShipAddress")
			ordShipAddress2 = rs("ordShipAddress2")
			ordShipCity = rs("ordShipCity")
			ordShipState = rs("ordShipState")
			ordShipZip = rs("ordShipZip")
			ordShipCountry = rs("ordShipCountry")
			ordPayProvider = rs("ordPayProvider")
			ordComLoc = rs("ordComLoc")
			ordExtra1 = rs("ordExtra1")
			ordExtra2 = rs("ordExtra2")
			ordExtra3 = rs("ordExtra3")
			ordAddInfo = rs("ordAddInfo")
			remember=True
		end if
		rs.Close
	end if
	sSQL = "SELECT stateName,stateAbbrev FROM states WHERE stateEnabled=1 ORDER BY stateName"
	rs.Open sSQL,cnn,0,1
	if NOT rs.EOF then allstates=rs.getrows
	rs.Close
	numhomecountries = 0
	nonhomecountries = 0
	sSQL = "SELECT countryName,countryOrder,"&getlangid("countryName",8)&" AS cnameshow FROM countries WHERE countryEnabled=1 ORDER BY countryOrder DESC,"&getlangid("countryName",8)
	rs.Open sSQL,cnn,0,1
	if NOT rs.EOF then allcountries=rs.getrows
	rs.Close
	if IsArray(allcountries) then
		for rowcounter=0 to UBOUND(allcountries,2)
			if allcountries(1,rowcounter)=2 then numhomecountries = numhomecountries + 1 else nonhomecountries = nonhomecountries + 1
		next
	end if
%>
<script language="javascript" type="text/javascript">
<!--
var checkedfullname=false;
var numhomecountries=0,nonhomecountries=0;
function checkform(frm)
{
<% if Trim(extraorderfield1)<>"" AND extraorderfield1required=true then %>
if(frm.ordextra1.value==""){
	alert("<%=xxPlsEntr%> \"<%=extraorderfield1%>\".");
	frm.ordextra1.focus();
	return (false);
}
<% end if %>
if(frm.name.value==""){
	alert("<%=xxPlsEntr%> \"<%=xxName%>\".");
	frm.name.focus();
	return (false);
}
gotspace=false;
var checkStr = frm.name.value;
for (i = 0; i < checkStr.length; i++){
	if(checkStr.charAt(i)==" ")
		gotspace=true;
}
if(!checkedfullname && !gotspace){
	alert("<%=xxFulNam%> \"<%=xxName%>\".");
	frm.name.focus();
	checkedfullname=true;
	return (false);
}
if(frm.email.value==""){
	alert("<%=xxPlsEntr%> \"<%=xxEmail%>\".");
	frm.email.focus();
	return (false);
}
validemail=0;
var checkStr = frm.email.value;
for (i = 0; i < checkStr.length; i++){
	if(checkStr.charAt(i)=="@")
		validemail |= 1;
	if(checkStr.charAt(i)==".")
		validemail |= 2;
}
if(validemail != 3){
	alert("<%=xxValEm%>");
	frm.email.focus();
	return (false);
}
if(frm.address.value==""){
	alert("<%=xxPlsEntr%> \"<%=xxAddress%>\".");
	frm.address.focus();
	return (false);
}
if(frm.city.value==""){
	alert("<%=xxPlsEntr%> \"<%=xxCity%>\".");
	frm.city.focus();
	return (false);
}
if(frm.country.selectedIndex < numhomecountries){
<%	if IsArray(allstates) AND xxOutState<>"" then %>
	if(frm.state.selectedIndex==0){
		alert("<%=xxPlsSlct & " " & xxState%>.");
		frm.state.focus();
		return (false);
	}
<%	end if %>
}else{
<%	if nonhomecountries>0 then %>
	if(frm.state2.value==""){
		alert("<%=xxPlsEntr%> \"<%=Replace(xxNonState,"<br />"," ")%>\".");
		frm.state2.focus();
		return (false);
	}
<%	end if %>}
if(frm.zip.value==""<% if zipoptional=TRUE then response.write " && FALSE"%>){
	alert("<%=xxPlsEntr%> \"<%=xxZip%>\".");
	frm.zip.focus();
	return (false);
}
if(frm.phone.value==""){
	alert("<%=xxPlsEntr%> \"<%=xxPhone%>\".");
	frm.phone.focus();
	return (false);
}
<% if Trim(extraorderfield2)<>"" AND extraorderfield2required=true then %>
if(frm.ordextra2.value==""){
	alert("<%=xxPlsEntr%> \"<%=extraorderfield2%>\".");
	frm.ordextra2.focus();
	return (false);
}
<% end if
   if Trim(extraorderfield3)<>"" AND extraorderfield3required=true then %>
if(frm.ordextra3.value==""){
	alert("<%=xxPlsEntr%> \"<%=extraorderfield3%>\".");
	frm.ordextra3.focus();
	return (false);
}
<% end if
   if noshipaddress<>true then %>
if(frm.saddress.value!=""){
	if(frm.sname.value==""){
		alert("<%=xxShpDtls%>\n\n<%=xxPlsEntr%> \"<%=xxName%>\".");
		frm.sname.focus();
		return (false);
	}
	if(frm.scity.value==""){
		alert("<%=xxShpDtls%>\n\n<%=xxPlsEntr%> \"<%=xxCity%>\".");
		frm.scity.focus();
		return (false);
	}
	if(frm.scountry.selectedIndex < numhomecountries){
<%	if IsArray(allstates) then %>
		if(frm.sstate.selectedIndex==0){
			alert("<%=xxShpDtls%>\n\n<%=xxPlsSlct & " " & xxState%>.");
			frm.sstate.focus();
			return (false);
		}
<%	end if %>
	}else{
<%	if nonhomecountries>0 then %>
		if(frm.sstate2.value==""){
			alert("<%=xxShpDtls%>\n\n<%=xxPlsEntr%> \"<%=Replace(xxNonState,"<br />"," ")%>\".");
			frm.sstate2.focus();
			return (false);
		}
<%	end if %>
	}
	if(frm.szip.value==""<% if zipoptional=TRUE then response.write " && FALSE"%>){
		alert("<%=xxShpDtls%>\n\n<%=xxPlsEntr%> \"<%=xxZip%>\".");
		frm.szip.focus();
		return (false);
	}
}
<% end if %>
if(frm.remember.checked==false){
	if(confirm("<%=xxWntRem%>")){
		frm.remember.checked=true
	}
}
<% if termsandconditions=TRUE then %>
if(frm.license.checked==false){
	alert("<%=xxPlsProc%>");
	frm.license.focus();
	return (false);
}
<% end if %>
return (true);
}
<% if termsandconditions=TRUE then %>
function showtermsandconds(){
newwin=window.open("termsandconditions.asp","Terms","menubar=no, scrollbars=yes, width=420, height=380, directories=no,location=no,resizable=yes,status=no,toolbar=no");
}
<% end if %>
var savestate=0;
var ssavestate=0;
function dosavestate(shp){
	thestate = eval('document.forms.mainform.'+shp+'state');
	eval(shp+'savestate = thestate.selectedIndex');
}
function checkoutspan(shp){
if(shp=='s' && document.getElementById('saddress').value=="")visib='hidden';else visib='visible';<%
if nonhomecountries>0 then response.write "thestyle = document.getElementById(shp+'outspan').style;"&vbCrLf
if IsArray(allstates) then
	response.write "theddstyle = document.getElementById(shp+'outspandd').style;"&vbCrLf
	response.write "thestate = eval('document.forms.mainform.'+shp+'state');"&vbCrLf
end if %>
thecntry = eval('document.forms.mainform.'+shp+'country');
if(thecntry.selectedIndex < numhomecountries){<%
if nonhomecountries>0 then response.write "thestyle.visibility='hidden';"&vbCrLf
if IsArray(allstates) then
	response.write "theddstyle.visibility=visib;"&vbCrLf
	response.write "thestate.disabled=false;"&vbCrLf
	response.write "eval('thestate.selectedIndex='+shp+'savestate');"&vbCrLf
end if %>
}else{<%
if nonhomecountries>0 then response.write "thestyle.visibility=visib;"&vbCrLf
if IsArray(allstates) then %>
theddstyle.visibility="hidden";
if(thestate.disabled==false){
thestate.disabled=true;
eval(shp+'savestate = thestate.selectedIndex');
thestate.selectedIndex=0;}
<% end if %>
}}
//-->
</script>
	  <table border="0" cellspacing="<%=maintablespacing%>" cellpadding="<%=maintablepadding%>" width="<%=maintablewidth%>" bgcolor="<%=maintablebg%>" align="center">
        <tr> 
          <td width="100%">
		    <form method="post" name="mainform" action="cart.asp" onsubmit="return checkform(this)">
			  <input type="hidden" name="mode" value="go" />
			  <input type="hidden" name="sessionid" value="<%=thesessionid%>" />
			  <input type="hidden" name="PARTNER" value="<%=Request.Form("PARTNER")%>" />
			  <table width="<%=innertablewidth%>" border="0" cellspacing="<%=innertablespacing%>" cellpadding="<%=innertablepadding%>" bgcolor="<%=innertablebg%>">
				<tr>
				  <td align="center" colspan="4"><strong><%=xxCstDtl%></strong></td>
				</tr>
<%	if Trim(extraorderfield1)<>"" then %>
				<tr>
				  <td align="right"><strong><% if extraorderfield1required=true then response.write "<font color='#FF0000'>*</font>"
									response.write extraorderfield1 %>:</strong></td>
				  <td colspan="3"><% if extraorderfield1html<>"" then response.write extraorderfield1html else response.write "<input type=""text"" name=""ordextra1"" size=""20"" value="""&ordExtra1&""" />"%></td>
				</tr>
<%	end if %>
				<tr>
				  <td align="right"><strong><font color='#FF0000'>*</font><%=xxName%>:</strong></td>
				  <td align="left"><input type="text" name="name" size="20" value="<%=ordName%>" /></td>
				  <td align="right"><strong><font color='#FF0000'>*</font><%=xxEmail%>:</strong></td>
				  <td align="left"><input type="text" name="email" size="20" value="<%=ordEmail%>" /></td>
				</tr>
				<tr>
				  <td align="right"><strong><font color='#FF0000'>*</font><%=xxAddress%>:</strong></td>
				  <td align="left"<% if useaddressline2=TRUE then response.write " colspan=""3"""%>><input type="text" name="address" id="address" size="25" value="<%=ordAddress%>" /></td>
<%	if useaddressline2=TRUE then %>
				</tr>
				<tr>
				  <td align="right"><strong><%=xxAddress2%>:</strong></td>
				  <td align="left"><input type="text" name="address2" size="25" value="<%=ordAddress2%>" /></td>
<%	end if %>
				  <td align="right"><strong><font color='#FF0000'>*</font><%=xxCity%>:</strong></td>
				  <td align="left"><input type="text" name="city" size="20" value="<%=ordCity%>" /></td>
				</tr>
<%	if IsArray(allstates) OR nonhomecountries<>0 then %>
				<tr>
<%		if IsArray(allstates) then %>
				  <td align="right"><strong><font color='#FF0000'><span id="outspandd" style="visibility:hidden">*</span></font><%=xxState%>:</strong></td>
				  <td align="left"><select name="state" size="1" onchange="dosavestate('')"><% havestate = show_states(ordState) %></select></td>
<%		end if
		if nonhomecountries=0 then
			response.write "<td colspan=""2"">&nbsp;</td>"
		else %>
				  <td align="right"><strong><font color='#FF0000'><span id="outspan" style="visibility:hidden">*</span></font><%=xxNonState%>:</strong></td>
				  <td align="left"><input type="text" name="state2" size="20" value="<% if not havestate then response.write ordState%>" /></td>
<%		end if
		if NOT IsArray(allstates) then response.write "<td colspan=""2"">&nbsp;</td>" %>
				</tr>
<%	end if %>
				<tr>
				  <td align="right"><strong><font color='#FF0000'>*</font><%=xxCountry%>:</strong></td>
				  <td align="left"><select name="country" size="1" onchange="checkoutspan('')">
<%	show_countries(ordCountry) %>
					</select>
				  </td>
				  <td align="right"><strong><font color='#FF0000'><% if zipoptional<>TRUE then response.write "*"%></font><%=xxZip%>:</strong></td>
				  <td align="left"><input type="text" name="zip" size="10" value="<%=ordZip%>" /></td>
				</tr>
				<tr>
				  <td align="right"><strong><font color='#FF0000'>*</font><%=xxPhone%>:</strong></td>
				  <td align="left"<%	if Trim(extraorderfield2)="" then response.write " colspan=""3""" %>><input type="text" name="phone" size="20" value="<%=ordPhone%>" /></td>
			<%	if Trim(extraorderfield2)<>"" then %>
				  <td align="right"><strong><% if extraorderfield2required=true then response.write "<font color='#FF0000'>*</font>"
									response.write extraorderfield2 %>:</strong></td>
				  <td align="left"><% if extraorderfield2html<>"" then response.write extraorderfield2html else response.write "<input type=""text"" name=""ordextra2"" size=""20"" value="""&ordExtra2&""" />"%></td>
			<%	end if %>
				</tr>
			<%	if commercialloc=TRUE then %>
				<tr>
				  <td align="right"><input type="checkbox" name="commercialloc" value="Y" <%if (ordComLoc AND 1)=1 then response.write "checked"%> /></td>
				  <td align="left" colspan="3"><font size="1"><%=xxComLoc%></font></td>
				</tr>
			<%	end if
				if saturdaydelivery=TRUE then %>
				<tr>
				  <td align="right"><input type="checkbox" name="saturdaydelivery" value="Y" <%if (ordComLoc AND 4)=4 then response.write "checked"%> /></td>
				  <td align="left" colspan="3"><font size="1"><%=xxSatDel%></font></td>
				</tr>
			<%	end if
				if abs(addshippinginsurance)=2 then %>
				<tr>
				  <td align="right"><input type="checkbox" name="wantinsurance" value="Y" <%if (ordComLoc AND 2)=2 then response.write "checked"%> /></td>
				  <td align="left" colspan="3"><font size="1"><%=xxWantIns%></font></td>
				</tr>
			<%	end if
				if allowsignaturerelease=TRUE AND signatureoption<>"" then %>
				<tr>
				  <td align="right"><input type="checkbox" name="signaturerelease" value="Y" <%if (ordComLoc AND 8)=8 then response.write "checked"%> /></td>
				  <td align="left" colspan="3"><font size="1"><%=xxSigRel%></font></td>
				</tr>
			<%	end if
				if insidedelivery=TRUE then %>
				<tr>
				  <td align="right"><input type="checkbox" name="insidedelivery" value="Y" <%if (ordComLoc AND 16)=16 then response.write "checked"%> /></td>
				  <td align="left" colspan="3"><font size="1"><%=xxInsDel%></font></td>
				</tr>
			<%	end if
				if noshipaddress<>true then %>
				<tr>
				  <td align="center" colspan="4"><strong><%=xxShpDiff%></strong></td>
				</tr>
<%					if Trim(extraorderfield3)<>"" then %>
				<tr>
				  <td align="right"><strong><% if extraorderfield3required=true then response.write "<font color='#FF0000'>*</font>"
									response.write extraorderfield3 %>:</strong></td>
				  <td colspan="3"><% if extraorderfield3html<>"" then response.write extraorderfield3html else response.write "<input type=""text"" name=""ordextra3"" size=""20"" value="""&ordExtra3&""" />"%></td>
				</tr>
<%					end if %>
				<tr>
				  <td align="right"><strong><%=xxName%>:</strong></td>
				  <td align="left" colspan="3"><input type="text" name="sname" size="20" value="<%=ordShipName%>" /></td>
				</tr>
				<tr>
				  <td align="right"><strong><%=xxAddress%>:</strong></td>
				  <td align="left"<% if useaddressline2=TRUE then response.write " colspan=""3""" %>><input type="text" name="saddress" id="saddress" size="25" value="<%=Trim(ordShipAddress)%>" /></td>
<%	if useaddressline2=TRUE then %>
				</tr>
				<tr>
				  <td align="right"><strong><%=xxAddress2%>:</strong></td>
				  <td align="left"><input type="text" name="saddress2" size="25" value="<%=ordShipAddress2%>" /></td>
<%	end if %>
				  <td align="right"><strong><%=xxCity%>:</strong></td>
				  <td align="left"><input type="text" name="scity" size="20" value="<%=ordShipCity%>" /></td>
				</tr>
<%	if IsArray(allstates) OR nonhomecountries<>0 then %>
				<tr>
<%		if IsArray(allstates) then %>
				  <td align="right"><strong><font color='#FF0000'><span id="soutspandd" style="visibility:hidden">*</span></font><%=xxState%>:</strong></td>
				  <td align="left"><select name="sstate" size="1" onchange="dosavestate('s')"><% havestate = show_states(ordShipState) %></select></td>
<%		end if
		if nonhomecountries=0 then
			response.write "<td colspan=""2"">&nbsp;</td>"
		else %>
				  <td align="right"><strong><font color='#FF0000'><span id="soutspan" style="visibility:hidden">*</span></font><%=xxNonState%>:</strong></td>
				  <td align="left"><input type="text" name="sstate2" size="20" value="<% if not havestate then response.write ordShipState%>" /></td>
<%		end if
		if NOT IsArray(allstates) then response.write "<td colspan=""2"">&nbsp;</td>" %>
				</tr>
<%	end if %>
				<tr>
				  <td align="right"><strong><%=xxCountry%>:</strong></td>
				  <td align="left"><select name="scountry" size="1" onchange="checkoutspan('s')">
<%		show_countries(ordShipCountry) %>
					</select>
				  </td>
				  <td align="right"><strong><%=xxZip%>:</strong></td>
				  <td align="left"><input type="text" name="szip" size="10" value="<%=ordShipZip%>" /></td>
				</tr>
			<%	end if ' noshipaddress %>
				<tr>
				  <td width="100%" align="center" colspan="4">
					<strong><%=xxAddInf%>.</strong><br />
					<textarea name="ordAddInfo" rows="3" wrap=virtual cols="44"><%=ordAddInfo%></textarea> 
				  </td>
				</tr>
<% if termsandconditions=TRUE then %>
				<tr>
				  <td width="100%" align="center" colspan="4"><input type="checkbox" name="license" value="1" />
					<%=xxTermsCo%>
				  </td>
				</tr>
<% end if %>
				<tr>
				  <td width="100%" align="center" colspan="4"><input type="checkbox" name="remember" value="1" <% if remember then response.write "checked"%> />
					<strong><%=xxRemMe%></strong><br />
					<font size="1"><%=xxOpCook%></font>
				  </td>
				</tr>
<%					if nogiftcertificate<>true then %>
				<tr>
				  <td align="right" colspan="2"><strong><%=xxGifNum%>:</strong></td><td colspan="2"><input type="text" name="cpncode" size="20" /></td>
				</tr>
				<tr>
				  <td align="center" colspan="4"><font size="1"><%=xxGifEnt%></font></td>
				</tr>
<%					end if
					if Session("clientLoginLevel")<>"" then minloglevel=Session("clientLoginLevel") else minloglevel=0
					sSQL = "SELECT payProvID,"&getlangid("PayProvShow",128)&" FROM payprovider WHERE payProvEnabled=1 AND payProvLevel<="&minloglevel&" AND payProvID NOT IN (19,20) ORDER BY payProvOrder"
					rs.Open sSQL,cnn,0,1
					alldata=""
					if not rs.EOF then alldata=rs.getrows
					rs.Close
					if NOT IsArray(alldata) then %>
				<tr>
				  <td colspan="4" align="center"><strong><%=xxNoPay%></strong></td>
				</tr>
<%					elseif UBOUND(alldata,2)=0 then %>
				<tr>
				  <td colspan="4" align="center"><input type="hidden" name="payprovider" value="<%=alldata(0,0)%>" /><strong><%=xxClkCmp%></strong></td>
				</tr>
<%					else %>			    <tr>
				  <td colspan="4" align="center"><p><strong><%=xxPlsChz%></strong></p>
				    <p><select name="payprovider" size="1">
<%						for rowcounter=0 to UBOUND(alldata,2)
							response.write "<option value='"&alldata(0,rowcounter)&"'"
							if ordPayProvider=alldata(0,rowcounter) then response.write " selected"
							response.write ">"&alldata(1,rowcounter)&"</option>"&vbCrLf
						next %>
				    </select></p>
				  </td>
			    </tr>
<%					end if %>
				<tr>
				  <td width="50%" align="center" colspan="4"><input type="image" src="images/checkout.gif" border="0" alt="<%=xxCOTxt%>" /></td>
				</tr>
			  </table>
			</form>
		  </td>
        </tr>
      </table>
<script language="javascript" type="text/javascript">
<%	if IsArray(allstates) then response.write "savestate = document.forms.mainform.state.selectedIndex;" & vbCrLf
	response.write "numhomecountries="&numhomecountries&";"&vbCrLf
	response.write "checkoutspan('');" & vbCrLf
	if noshipaddress<>true then
		if IsArray(allstates) then response.write "ssavestate = document.forms.mainform.sstate.selectedIndex;"&vbCrLf
		response.write "checkoutspan('s')"&vbCrLf
	end if %></script>
<%
elseif checkoutmode="go" OR paypalexpress then
%>
<!--#include file="uspsshipping.asp"-->
<%
	if NOT paypalexpress then
		thesessionid = replace(trim(request.form("sessionid")),"'","")
		ordName = trim(request.form("name"))
		ordAddress = trim(request.form("address"))
		ordAddress2 = trim(request.form("address2"))
		ordCity = trim(request.form("city"))
		ordState = trim(request.form("state2"))
		if trim(request.form("state")) <> "" then ordState = trim(request.form("state"))
		ordZip = trim(request.form("zip"))
		ordCountry = trim(request.form("country"))
		ordEmail = trim(request.form("email"))
		ordPhone = trim(request.form("phone"))
		ordShipName = trim(request.form("sname"))
		ordShipAddress = trim(request.form("saddress"))
		ordShipAddress2 = trim(request.form("saddress2"))
		ordShipCity = trim(request.form("scity"))
		ordShipState = trim(request.form("sstate2"))
		if trim(request.form("sstate")) <> "" then ordShipState = trim(request.form("sstate"))
		ordShipZip = trim(request.form("szip"))
		ordShipCountry = trim(request.form("scountry"))
		commercialloc = trim(commerciallocpost)
		wantinsurance = trim(request.form("wantinsurance"))
		saturdaydelivery = trim(request.form("saturdaydelivery"))
		signaturerelease = trim(request.form("signaturerelease"))
		insidedelivery = trim(request.form("insidedelivery"))
		if commercialloc="Y" then ordComLoc = 1 else ordComLoc = 0
		if wantinsurance="Y" OR abs(addshippinginsurance)=1 then ordComLoc = ordComLoc + 2
		if saturdaydelivery="Y" then ordComLoc = ordComLoc + 4
		if signaturerelease="Y" then ordComLoc = ordComLoc + 8
		if insidedelivery="Y" then ordComLoc = ordComLoc + 16
		ordAffiliate = trim(request.form("PARTNER"))
		ordExtra1 = trim(request.form("ordextra1"))
		ordExtra2 = trim(request.form("ordextra2"))
		ordExtra3 = trim(request.form("ordextra3"))
		ordAVS = trim(request.form("ppexp1"))
		ordCVV = trim(request.form("ppexp2"))
		ordAddInfo = trim(request.form("ordAddInfo"))
	end if
	if ordShipAddress<>"" then
		shipcountry = ordShipCountry
		shipstate = ordShipState
		destZip = ordShipZip
	else
		shipcountry = ordCountry
		shipstate = ordState
		destZip = ordZip
	end if
	sSQL = "SELECT countryID,countryCode,countryOrder FROM countries WHERE countryName='"&replace(ordCountry,"'","''")&"'"
	rs.Open sSQL,cnn,0,1
		countryID = rs("countryID")
		countryCode = rs("countryCode")
		homecountry = (rs("countryOrder")=2)
	rs.Close
	if NOT homecountry then perproducttaxrate=FALSE
	sSQL = "SELECT countryID,countryTax,countryCode,countryFreeShip,countryOrder FROM countries WHERE countryName='"&replace(shipcountry,"'","''")&"'"
	rs.Open sSQL,cnn,0,1
		countryTaxRate = rs("countryTax")
		shipCountryID = rs("countryID")
		shipCountryCode = rs("countryCode")
		freeshipapplies = (rs("countryFreeShip")=1)
		shiphomecountry = (rs("countryOrder")=2)
	rs.Close
	if homecountry then
		sSQL = "SELECT stateAbbrev FROM states WHERE "&IIfVr(usestateabbrev=TRUE,"stateAbbrev","stateName")&"='"&replace(ordState,"'","''")&"'"
		rs.Open sSQL,cnn,0,1
		if NOT rs.EOF then stateAbbrev=rs("stateAbbrev")
		rs.Close
	end if
	if shiphomecountry then
		sSQL = "SELECT stateTax,stateAbbrev,stateFreeShip FROM states WHERE "&IIfVr(usestateabbrev=TRUE,"stateAbbrev","stateName")&"='"&replace(shipstate,"'","''")&"'"
		rs.Open sSQL,cnn,0,1
		if NOT rs.EOF then
			stateTaxRate=rs("stateTax")
			shipStateAbbrev=rs("stateAbbrev")
			freeshipapplies=(freeshipapplies AND (rs("stateFreeShip")=1))
		end if
		rs.Close
	end if
	if trim(Session("clientUser")) <> "" then
		if (Session("clientActions") AND 1)=1 then stateTaxRate=0
		if (Session("clientActions") AND 2)=2 then countryTaxRate=0
	end if
	initshippingmethods()
	if mysqlserver=true then
		sSQL = "SELECT cartID,cartProdID,cartProdName,cartProdPrice,cartQuantity,pWeight,pShipping,pShipping2,pExemptions,pSection,topSection,pDims,pTax FROM cart LEFT JOIN products ON cart.cartProdID=products.pId LEFT OUTER JOIN sections ON products.pSection=sections.sectionID WHERE cartCompleted=0 AND cartSessionID="&thesessionid
	else
		sSQL = "SELECT cartID,cartProdID,cartProdName,cartProdPrice,cartQuantity,pWeight,pShipping,pShipping2,pExemptions,pSection,topSection,pDims,pTax FROM cart INNER JOIN (products LEFT OUTER JOIN sections ON products.pSection=sections.sectionID) ON cart.cartProdID=products.pID WHERE cartCompleted=0 AND cartSessionID="&thesessionid
	end if
	rs.Open sSQL,cnn,0,1
	if NOT (rs.EOF OR rs.BOF) then alldata=rs.getrows
	rs.Close
	if success AND IsArray(alldata) then
		rowcounter = 0
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
		call calculatediscounts(vsround(totalgoods,2), true, cpncode)
		if shippingpost<>"" then
			shipArr = split(shippingpost,"|")
			shipping = cDbl(shipArr(0))
			isstandardship = Int(shipArr(1))=1
			shipMethod = shipArr(2)
		else
			calculateshipping()
		end if
		if shippingpost="" AND alternaterates AND somethingToShip then checkIntOptions = True
		insuranceandtaxaddedtoshipping()
		if NOT checkIntOptions then
			call calculateshippingdiscounts(true)
			if Session("clientUser")<>"" AND Session("clientActions")<>0 then cpnmessage = cpnmessage & xxLIDis & Session("clientUser") & "<br />"
			cpnmessage = Right(cpnmessage,Len(cpnmessage)-6)
			if totaldiscounts > totalgoods then totaldiscounts = totalgoods
			calculatetaxandhandling()
			totalgoods = vsround(totalgoods,2)
			shipping = vsround(shipping,2)
			stateTax = vsround(stateTax,2)
			countryTax = vsround(countryTax,2)
			handling = vsround(handling,2)
			if addshippingtodiscounts=TRUE then totaldiscounts = totaldiscounts + freeshipamnt : freeshipamnt = 0
			freeshipamnt = vsround(freeshipamnt, 2)
			totaldiscounts = vsround(totaldiscounts, 2)
			grandtotal = vsround((totalgoods + shipping + stateTax + countryTax + handling) - (totaldiscounts + freeshipamnt), 2)
			if grandtotal < 0 then grandtotal = 0
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
			rs.Fields("ordPayProvider") = ordPayProvider
			rs.Fields("ordAuthNumber")	= "" ' Not yet authorized
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
			rs.Fields("ordShipCarrier")	= shipType
			rs.Fields("ordTotal")		= totalgoods
			rs.Fields("ordDate")		= DateAdd("h",dateadjust,Now())
			rs.Fields("ordStatus")		= 2
			rs.Fields("ordStatusDate")	= DateAdd("h",dateadjust,Now())
			rs.Fields("ordIP")			= left(request.servervariables("REMOTE_ADDR"), 48)
			rs.Fields("ordComLoc")		= ordComLoc
			rs.Fields("ordAffiliate")	= ordAffiliate
			rs.Fields("ordAddInfo")		= ordAddInfo
			rs.Fields("ordDiscount")	= totaldiscounts
			rs.Fields("ordDiscountText")= Left(cpnmessage,255)
			rs.Fields("ordExtra1")		= ordExtra1
			rs.Fields("ordExtra2")		= ordExtra2
			rs.Fields("ordExtra3")		= ordExtra3
			rs.Fields("ordAVS")			= ordAVS
			rs.Fields("ordCVV")			= ordCVV
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
			sSQL="UPDATE cart SET cartOrderID="&orderid&" WHERE cartCompleted=0 AND cartSessionID="&replace(thesessionid,"'","")
			cnn.Execute(sSQL)
			descstr=""
			addcomma = ""
			sSQL="SELECT cartQuantity,cartProdName FROM cart WHERE cartOrderID="&orderid&" AND cartCompleted=0"
			rs.Open sSQL,cnn,0,1
			do while NOT rs.EOF
				descstr=descstr&addcomma&rs("cartQuantity")&" "&rs("cartProdName")
				addcomma = ", "
				rs.MoveNext
			loop
			rs.Close
			descstr=Replace(descstr,"""","")
			if request.form("remember")="1" then
				response.write "<script src='vsadmin/savecookie.asp?id1="&orderid&"&id2="&replace(thesessionid,"'","")&"'></script>"
			end if
		end if
	else
		success=False
	end if
	if checkIntOptions AND success OR (alternaterates AND NOT success) then
		hassuccess = success
		success = False ' So not to print the order totals.
%>
	<br />
	<form method="post" name="shipform" action="cart.asp">
<%
call writehiddenvar("mode", "go")
call writehiddenvar("vrshippingoptions", "1")
call writehiddenvar("sessionid", thesessionid)
call writehiddenvar("PARTNER", ordAffiliate)
call writehiddenvar("name", ordName)
call writehiddenvar("email", ordEmail)
call writehiddenvar("address", ordAddress)
call writehiddenvar("address2", ordAddress2)
call writehiddenvar("city", ordCity)
call writehiddenvar("state", ordState)
call writehiddenvar("country", ordCountry)
call writehiddenvar("zip", ordZip)
call writehiddenvar("phone", ordPhone)
call writehiddenvar("sname", ordShipName)
call writehiddenvar("saddress", ordShipAddress)
call writehiddenvar("saddress2", ordShipAddress2)
call writehiddenvar("scity", ordShipCity)
call writehiddenvar("sstate", ordShipState)
call writehiddenvar("scountry", ordShipCountry)
call writehiddenvar("szip", ordShipZip)
call writehiddenvar("ordAddInfo", ordAddInfo)
call writehiddenvar("ordextra1", ordExtra1)
call writehiddenvar("ordextra2", ordExtra2)
call writehiddenvar("ordextra3", ordExtra3)
call writehiddenvar("ppexp1", ordAVS)
call writehiddenvar("ppexp2", ordCVV)
call writehiddenvar("cpncode", cpncode)
call writehiddenvar("payprovider", ordPayProvider)
call writehiddenvar("token", token)
call writehiddenvar("payerid", payerid)
call writehiddenvar("wantinsurance", wantinsurance)
call writehiddenvar("commercialloc", commercialloc)
call writehiddenvar("saturdaydelivery", saturdaydelivery)
call writehiddenvar("signaturerelease", signaturerelease)
call writehiddenvar("insidedelivery", insidedelivery)
call writehiddenvar("remember", request.form("remember"))
%>
            <table class="cobtbl" width="<%=maintablewidth%>" border="0" bordercolor="#B1B1B1" cellspacing="1" cellpadding="3" bgcolor="#B1B1B1">
			  <tr>
			    <td height="34" align="center" class="cobhl" bgcolor="#EBEBEB"><strong><%=xxShpOpt%></strong></td>
			  </tr>
			  <tr>
				<td height="34" align="center" class="cobll" bgcolor="#FFFFFF">
<%					if hassuccess then %>
				  <table width="100%" cellspacing="0" cellpadding="0" border="0" bgcolor="#FFFFFF">
					<tr>
					  <td height="74" align="right" width="50%" class="cobll" bgcolor="#FFFFFF"><%
						if shipType=4 then
							response.write "<img src=""images/LOGO_S.gif"" alt=""UPS"" />&nbsp;&nbsp;"
						elseif shipType=7 then
							response.write "<img src=""images/fedexsmall.gif"" alt=""FedEx"" />&nbsp;&nbsp;"
						else
							response.write "&nbsp;"
						end if %></td>
					  <td height="74" align="center" class="cobll" bgcolor="#FFFFFF"><%
						call calculateshippingdiscounts(false)
						response.write "<select name='shipping' size='1'>"
						if shipType=2 OR shipType=5 then
							if IsArray(allzones) then
								for index3=0 to numshipoptions
									response.write "<option value='"&intShipping(2,index3)&"|"&IIfVr((pzFSA AND (2 ^ index3))<>0,"1","0")&"|" & intShipping(0,index3) & "'>"
									response.write IIfVr(freeshippingapplied AND ((pzFSA AND (2 ^ index3))<>0),xxFree & " " & intShipping(0,index3),intShipping(0,index3) & " " & FormatEuroCurrency(intShipping(2,index3))) & "</option>"
								next
							end if
						else
							for index=0 to UBOUND(intShipping,2)
								if shipType=3 then
									if iTotItems=intShipping(3,index) then
										for index2=0 to UBOUND(uspsmethods,2)
											if replace(lcase(intShipping(0,index)),"-"," ") = replace(lcase(uspsmethods(0,index2)),"-"," ") then
												response.write "<option value='"&intShipping(2,index)&"|"&uspsmethods(1,index2)&"|"&uspsmethods(2,index2)&"'" & IIfVr(freeshippingapplied AND uspsmethods(1,index2)=1, " selected>",">")
												response.write uspsmethods(2,index2)
												response.write " ("&intShipping(1,index)&") " & IIfVr(freeshippingapplied AND uspsmethods(1,index2)=1, xxFree, FormatEuroCurrency(intShipping(2,index)))
												response.write "</option>"
												exit for
											end if
										next
									end if
								elseif shipType=4 OR shipType=6 OR shipType=7 then
									if intShipping(3,index) then
										response.write "<option value='"&intShipping(2,index)&"|"&intShipping(4,index)&"|"&intShipping(0,index)&"'" & IIfVr(freeshippingapplied AND intShipping(4,index)=1, " selected>",">") & intShipping(0,index)&" "
										if Trim(intShipping(1,index))<>"" then response.write "("&xxGuar&" "&intShipping(1,index)&") "
										response.write IIfVr(freeshippingapplied AND intShipping(4,index)=1, xxFree, FormatEuroCurrency(intShipping(2,index)))
										response.write "</option>"
									end if
								end if
							next
						end if
						if willpickuptext<>"" then
							if willpickupcost="" then willpickupcost=0
							response.write "<option value="""&willpickupcost&"|1|"& replace(willpickuptext,"""","&quot;") &""">"
							response.write willpickuptext & " " & FormatEuroCurrency(willpickupcost) & "</option>"
						end if
						response.write "</select>"
					%></td>
					  <td height="74" align="left" width="50%" class="cobll" bgcolor="#FFFFFF">&nbsp;</td>
					</tr>
				  </table>
<%					else
						response.write "<input type=""hidden"" name=""shipping"" value="""">" & errormsg
					end if %>
				</td>
			  </tr>
<%			if alternaterates then %>
			  <tr>
			    <td height="34" align="center" class="cobhl" bgcolor="#EBEBEB"><strong>Or select an alternate shipping carrier to compare rates.</strong></td>
			  </tr>
			  <tr>
				<td height="34" align="center" class="cobll" bgcolor="#FFFFFF">
					<select name="altrates" size="1" onchange="document.forms.shipform.shipping.value='';document.forms.shipform.shipping.disabled=true;document.forms.shipform.submit();"><%
				if alternateratesups<>"" OR origShipType=4 then response.write "<option value=""4"""&IIfVr(shipType=4," selected","")&">"&alternateratesups&"</option>"
				if alternateratesusps<>"" OR origShipType=3 then response.write "<option value=""3"""&IIfVr(shipType=3," selected","")&">"&alternateratesusps&"</option>"
				if alternateratesweightbased<>"" OR origShipType=2 then response.write "<option value=""2"""&IIfVr(shipType=2," selected","")&">"&alternateratesweightbased&"</option>"
				if alternateratescanadapost<>"" OR origShipType=6 then response.write "<option value=""6"""&IIfVr(shipType=6," selected","")&">"&alternateratescanadapost&"</option>"
				if alternateratesfedex<>"" OR origShipType=7 then response.write "<option value=""7"""&IIfVr(shipType=7," selected","")&">"&alternateratesfedex&"</option>"
						%></select>
				</td>
			  </tr>
<%			end if %>
			  <tr>
			    <td height="34" align="center" class="cobll" bgcolor="#FFFFFF"><table width="100%" cellspacing="0" cellpadding="0" border="0">
				    <tr>
					  <td class="cobll" bgcolor="#FFFFFF" width="16" height="26" align="right" valign="bottom">&nbsp;</td>
					  <td class="cobll" bgcolor="#FFFFFF" width="100%" align="center"><input type="image" value="Checkout" border="0" src="images/checkout.gif" alt="<%=xxCOTxt%>" /></td>
					  <td class="cobll" bgcolor="#FFFFFF" width="16" height="26" align="right" valign="bottom"><img src="images/tablebr.gif" alt="" /></td>
					</tr>
				  </table>
				</td>
			  </tr>
			</table>
		<% if shipType=4 then %>
			<p align="center">&nbsp;<br /><font size="1">UPS&reg;, UPS & Shield Design&reg; and UNITED PARCEL SERVICE&reg; 
			  are<br />registered trademarks of United Parcel Service of America, Inc.</font></p>
		<% elseif shipType=7 then %>
			<p align="center">&nbsp;<br /><font size="1">FedEx&reg; is a registered service mark of Federal Express Corporation. FedEx logos used by permission. All rights reserved.
</font></p>
		<% end if %>
	</form>
<%	elseif NOT success then %>
      <table border="0" cellspacing="<%=maintablespacing%>" cellpadding="<%=maintablepadding%>" width="<%=maintablewidth%>" bgcolor="<%=maintablebg%>" align="center">
        <tr> 
          <td width="100%">
            <table width="<%=innertablewidth%>" border="0" cellspacing="<%=innertablespacing%>" cellpadding="<%=innertablepadding%>" bgcolor="<%=innertablebg%>">
			  <tr>
			    <td align="center"><p>&nbsp;</p><p><strong><%=xxSryErr%></strong></p><p><strong><%="<br />"&errormsg%></strong></p><p>&nbsp;</p></td>
			  </tr>
			</table>
		  </td>
        </tr>
      </table>
<%	elseif ordPayProvider<>"" then
		blockuser=checkuserblock(ordPayProvider)
		if blockuser then
			orderid = 0
			thesessionid = 0
			xxMstClk = multipurchaseblockmessage
		else
			call getpayprovdetails(ordPayProvider,data1,data2,data3,demomode,ppmethod)
		end if
		if pathtossl<>"" then
			if Right(pathtossl,1) <> "/" then pathtossl = pathtossl & "/"
			storeurl = pathtossl
		end if
		if grandtotal > 0 AND ordPayProvider="1" then ' PayPal
%>
	<form method="post" action="https://www.<% if demomode then response.write "sandbox." %>paypal.com/cgi-bin/webscr">
	<input type="hidden" name="cmd" value="_ext-enter" />
	<input type="hidden" name="redirect_cmd" value="_xclick" />
	<input type="hidden" name="rm" value="2" />
	<input type="hidden" name="business" value="<%=data1%>" />
	<input type="hidden" name="return" value="<%=storeurl%>thanks.asp" />
	<input type="hidden" name="notify_url" value="<%=storeurl%>vsadmin/ppconfirm.asp" />
	<input type="hidden" name="item_name" value="<%=Left(descstr,127)%>" />
	<input type="hidden" name="custom" value="<%=orderid%>" />
<%			if paypallc<>"" then call writehiddenvar("lc", paypallc)
			Session.LCID = 1033
			if splitpaypalshipping then
				call writehiddenvar("shipping", FormatNumber(vsround((shipping + handling) - freeshipamnt, 2),2,-1,0,0))
				call writehiddenvar("amount", FormatNumber(vsround((totalgoods + stateTax + countryTax) - totaldiscounts, 2),2,-1,0,0))
			else
				call writehiddenvar("amount", FormatNumber(grandtotal,2,-1,0,0))
			end if
			Session.LCID = saveLCID %>
	<input type="hidden" name="currency_code" value="<%=countryCurrency%>" />
	<input type="hidden" name="bn" value="ecommercetemplates.asp.ecommplus" />
<%			thename = Trim(Request.form("name"))
			if thename<>"" then
				if InStr(thename," ") > 0 then
					namearr = Split(thename," ",2)
					response.write "<input type=""hidden"" name='first_name' value='"&namearr(0)&"' />"&vbCrLf
					response.write "<input type=""hidden"" name='last_name' value='"&namearr(1)&"' />"&vbCrLf
				else
					response.write "<input type=""hidden"" name='last_name' value='"&thename&"' />"&vbCrLf
				end if
			end if %>
	<input type="hidden" name="address1" value="<%=Request.form("address")%>" />
	<input type="hidden" name="address2" value="<%=Request.form("address2")%>" />
	<input type="hidden" name="city" value="<%=Request.form("city")%>" />
<%			if countryID=1 AND stateAbbrev<>"" then %>
	<input type="hidden" name="state" value="<%=stateAbbrev%>" />
<%			else %>
	<input type="hidden" name="state" value="<%if Trim(Request.form("state"))<>"" then response.write Trim(Request.form("state")) else response.write Trim(Request.form("state2"))%>" />
<%			end if %>
	<input type="hidden" name="country" value="<%=countryCode%>" />
	<input type="hidden" name="email" value="<%=Request.form("email")%>" />
	<input type="hidden" name="zip" value="<%=Request.form("zip")%>" />
	<input type="hidden" name="cancel_return" value="<%=storeurl%>sorry.asp" />
<%		elseif grandtotal > 0 AND ordPayProvider="2" then ' 2Checkout
			courl="https://www.2checkout.com/cgi-bin/sbuyers/cartpurchase.2c"
			if IsNumeric(data1) then
				if data1>200000 OR use2checkoutv2=TRUE then courl="https://www2.2checkout.com/2co/buyer/purchase"
			end if %>
	<form method="post" action="<%=courl%>">
	<input type="hidden" name="cart_order_id" value="<%=orderid%>" />
	<input type="hidden" name="merchant_order_id" value="<%=orderid%>" />
	<input type="hidden" name="sid" value="<%=data1%>" />
	<%			Session.LCID = 1033 %>
	<input type="hidden" name="total" value="<%=FormatNumber(grandtotal,2,-1,0,0)%>" />
	<%			Session.LCID = saveLCID %>
	<input type="hidden" name="card_holder_name" value="<%=Request.form("name")%>" />
	<input type="hidden" name="street_address" value="<%=Request.form("address") & IIfVr(trim(Request.form("address2"))<>"",", " & Request.form("address2"), "")%>" />
	<%			if countryID=1 OR countryID=2 then %>
	<input type="hidden" name="city" value="<%=Request.form("city")%>" />
	<input type="hidden" name="state" value="<%if Trim(Request.form("state"))<>"" then response.write Trim(Request.form("state")) else response.write Trim(Request.form("state2"))%>" />
	<%			else
					if Trim(Request.form("state"))<>"" then thestate = Trim(Request.form("state")) else thestate = Trim(Request.form("state2")) %>
	<input type="hidden" name="city" value="<%=Request.form("city") & IIfVr(thestate<>"", ", " & thestate, "")%>" />
	<input type="hidden" name="state" value="Outside US and Canada" />
	<%			end if %>
	<input type="hidden" name="zip" value="<%=Request.form("zip")%>" />
	<input type="hidden" name="country" value="<%=countryCode%>" />
	<input type="hidden" name="email" value="<%=Request.form("email")%>" />
	<input type="hidden" name="phone" value="<%=Request.form("phone")%>" />
	<input type="hidden" name="id_type" value="1" />
<%			sSQL = "SELECT cartID,cartProdID,pName,pPrice,cartQuantity,"&IIfVr(digidownloads=TRUE,"pDownload,","")&"pDescription FROM cart INNER JOIN products on cart.cartProdID=products.pID WHERE cartCompleted=0 AND cartSessionID=" &  thesessionid
			rs.Open sSQL,cnn,0,1
			index=1
			do while NOT rs.EOF
				thedesc = left(trim(replace(strip_tags2(rs("pDescription")&""),vbNewLine,"\n")), 255)
				if thedesc = "" then thedesc = left(trim(replace(strip_tags2(rs("pName")&""),vbNewLine,"\n")), 255)
				response.write "<input type=""hidden"" name=""c_prod_" & index & """ value=""" & replace(replace(rs("cartProdID"),"""","&quot;"), ",", "&#44;") & "," & rs("cartQuantity") & """ />" & vbCrLf
				response.write "<input type=""hidden"" name=""c_name_" & index & """ value=""" & replace(strip_tags2(rs("pName")),"""","&quot;") & """ />" & vbCrLf
				response.write "<input type=""hidden"" name=""c_description_" & index & """ value=""" & replace(thedesc,"""","&quot;") & """ />" & vbCrLf
				response.write "<input type=""hidden"" name=""c_price_" & index & """ value=""" & FormatNumber(rs("pPrice"),2,-1,0,0) & """ />" & vbCrLf
				if digidownloads=TRUE then
					if trim(rs("pDownload")&"")<>"" then response.write "<input type=""hidden"" name=""c_tangible_" & index & """ value=""N"" />" & vbCrLf
				end if
				index = index+1
				rs.MoveNext
			loop
			rs.Close
			if trim(Request.form("sname")) <> "" OR trim(Request.form("saddress")) <> "" then %>
	<input type="hidden" name="ship_name" value="<%=Request.form("sname")%>" />
	<input type="hidden" name="ship_street_address" value="<%=Request.form("saddress")& IIfVr(trim(Request.form("saddress2"))<>"",", " & Request.form("saddress2"), "")%>" />
	<input type="hidden" name="ship_city" value="<%=Request.form("scity")%>" />
	<input type="hidden" name="ship_state" value="<%if Trim(Request.form("sstate"))<>"" then response.write Trim(Request.form("sstate")) else response.write Trim(Request.form("sstate2"))%>" />
	<input type="hidden" name="ship_zip" value="<%=Request.form("szip")%>" />
	<input type="hidden" name="ship_country" value="<%=Request.form("scountry")%>" />
<%			end if
			if demomode then call writehiddenvar("demo", "Y")
			call writehiddenvar("pay_method", "CC")
			call writehiddenvar("fixed", "Y")
		elseif grandtotal > 0 AND ordPayProvider="3" then ' Authorize.net SIM
			if secretword<>"" then
				data1 = upsdecode(data1, secretword)
				data2 = upsdecode(data2, secretword)
			end if
%>
	<form method="post" action="https://secure.authorize.net/gateway/transact.dll">
	<input type="hidden" name="x_Version" value="3.0" />
	<input type="hidden" name="x_Login" value="<%=data1%>" />
	<input type="hidden" name="x_Show_Form" value="PAYMENT_FORM" />
<%			if ppmethod=1 then %>
	<input type="hidden" name="x_type" value="AUTH_ONLY" />
<%			end if
			thename = Trim(Request.form("name"))
			if thename<>"" then
				if InStr(thename," ") > 0 then
					namearr = Split(thename," ",2)
					response.write "<input type=""hidden"" name=""x_First_Name"" value="""&replace(namearr(0),"""","&quot;")&""" />"&vbCrLf
					response.write "<input type=""hidden"" name=""x_Last_Name"" value="""&replace(namearr(1),"""","&quot;")&""" />"&vbCrLf
				else
					response.write "<input type=""hidden"" name=""x_Last_Name"" value="""&replace(thename,"""","&quot;")&""" />"&vbCrLf
				end if
			end if
			Randomize
			sequence = Int(1000 * Rnd)
			if authnetadjust<>"" then
				tstamp = GetSecondsSince1970() + authnetadjust
			else
				tstamp = GetSecondsSince1970()
			end if
			fingerprint = HMAC(data2, data1 & "^" & sequence & "^" & tstamp & "^" & FormatNumber(grandtotal,2,-1,0,0) & "^")
%>
	<input type="hidden" name="x_fp_sequence" value="<%=sequence%>" />
	<input type="hidden" name="x_fp_timestamp" value="<%=tstamp%>" />
	<input type="hidden" name="x_fp_hash" value="<%=fingerprint%>" />
	<input type="hidden" name="x_address" value="<%=Request.form("address")& IIfVr(trim(Request.form("address2"))<>"",", " & Request.form("address2"), "")%>" />
	<input type="hidden" name="x_city" value="<%=Request.form("city")%>" />
	<input type="hidden" name="x_country" value="<%=Request.form("country")%>" />
	<input type="hidden" name="x_phone" value="<%=Request.form("phone")%>" />
	<input type="hidden" name="x_state" value="<%if Trim(Request.form("state"))<>"" then response.write Trim(Request.form("state")) else response.write Trim(Request.form("state2"))%>" />
	<input type="hidden" name="x_zip" value="<%=Request.form("zip")%>" />
	<input type="hidden" name="x_cust_id" value="<%=orderid%>" />
	<input type="hidden" name="x_Invoice_Num" value="<%=orderid%>" />
	<input type="hidden" name="x_ect_ordid" value="<%=orderid%>" />
	<input type="hidden" name="x_Description" value="<%=Left(descstr,255)%>" />
	<input type="hidden" name="x_email" value="<%=Request.form("email")%>" />
<%			if trim(Request.form("sname")) <> "" OR trim(Request.form("saddress")) <> "" then
				thename = Trim(Request.form("sname"))
				if thename<>"" then
					if InStr(thename," ") > 0 then
						namearr = Split(thename," ",2)
						response.write "<input type=""hidden"" name='x_Ship_To_First_Name' value='"&namearr(0)&"'>"&vbCrLf
						response.write "<input type=""hidden"" name='x_Ship_To_Last_Name' value='"&namearr(1)&"'>"&vbCrLf
					else
						response.write "<input type=""hidden"" name='x_Ship_To_Last_Name' value='"&thename&"'>"&vbCrLf
					end if
				end if %>
	<input type="hidden" name="x_ship_to_address" value="<%=Request.form("saddress")& IIfVr(trim(Request.form("saddress2"))<>"",", " & Request.form("saddress2"), "")%>" />
	<input type="hidden" name="x_ship_to_city" value="<%=Request.form("scity")%>" />
	<input type="hidden" name="x_ship_to_country" value="<%=Request.form("scountry")%>" />
	<input type="hidden" name="x_ship_to_state" value="<%if Trim(Request.form("sstate"))<>"" then response.write Trim(Request.form("sstate")) else response.write Trim(Request.form("sstate2"))%>" />
	<input type="hidden" name="x_ship_to_zip" value="<%=Request.form("szip")%>" />
<%			end if %>
	<input type="hidden" name="x_Amount" value="<%=FormatNumber(grandtotal,2,-1,0,0)%>" />
	<input type="hidden" name="x_Relay_Response" value="True" />
	<input type="hidden" name="x_Relay_URL" value="<%=storeurl%>vsadmin/wpconfirm.asp" />
<%			if demomode then %>
	<input type="hidden" name="x_Test_Request" value="TRUE" />
<%			end if
		elseif grandtotal = 0 OR ordPayProvider="4" then ' Email %>
	<form method="post" action="<%=storeurl%>thanks.asp">
	<input type="hidden" name="emailorder" value="<%=orderid%>" />
	<input type="hidden" name="thesessionid" value="<%=thesessionid%>" />
<%		elseif grandtotal > 0 AND ordPayProvider="17" then ' Email 2 %>
	<form method="post" action="<%=storeurl%>thanks.asp">
	<input type="hidden" name="secondemailorder" value="<%=orderid%>" />
	<input type="hidden" name="thesessionid" value="<%=thesessionid%>" />
<%		elseif grandtotal > 0 AND ordPayProvider="5" then ' WorldPay %>
	<form method="post" action="https://select.worldpay.com/wcc/purchase">
	<input type="hidden" name="instId" value="<%=data1%>" />
	<input type="hidden" name="cartId" value="<%=orderid%>" />
<%			Session.LCID = 1033 %>
	<input type="hidden" name="amount" value="<%=FormatNumber(grandtotal,2,-1,0,0)%>" />
<%			Session.LCID = saveLCID %>
	<input type="hidden" name="currency" value="<%=countryCurrency%>" />
	<input type="hidden" name="desc" value="<%=Left(descstr,255)%>" />
	<input type="hidden" name="name" value="<%=Request.form("name")%>" />
	<input type="hidden" name="address" value="<%=Request.form("address")& IIfVr(trim(Request.form("address2"))<>"",", " & Request.form("address2"), "")%>&#10;<%=Request.form("city")%>&#10;<%
			if trim(request.form("state"))<>"" then
				response.write Request.form("state")
			else
				response.write Request.form("state2")
			end if%>" />
	<input type="hidden" name="postcode" value="<%=Request.form("zip")%>" />
	<input type="hidden" name="country" value="<%=countryCode%>" />
	<input type="hidden" name="tel" value="<%=Request.form("phone")%>" />
	<input type="hidden" name="email" value="<%=Request.form("email")%>" />
	<input type="hidden" name="authMode" value="<% if ppmethod=1 then response.write "E" else response.write "A" %>" />
<%			if demomode then %>
	<input type="hidden" name="testMode" value="100" />
<%			end if
			data2arr = split(data2,"&",2)
			if UBOUND(data2arr) >= 0 then data2 = data2arr(0)
			if data2<>"" then
				response.write "<input type=""hidden"" name=""signatureFields"" value=""amount:currency:cartId"" />"
				Session.LCID = 1033
				response.write "<input type=""hidden"" name=""signature"" value="""&calcmd5(data2&":"&FormatNumber(grandtotal,2,-1,0,0)&":"&countryCurrency&":"&orderid)&""" />"
				Session.LCID = saveLCID
			end if
		elseif grandtotal > 0 AND ordPayProvider="6" then ' NOCHEX %>
	<form method="post" action="https://www.nochex.com/nochex.dll/checkout">
	<input type="hidden" name="email" value="<%=data1%>" />
	<input type="hidden" name="returnurl" value="<%=storeurl & IIfVr(TRUE, "thanks.asp?ncretval="&orderid&"&ncsessid="&thesessionid, "")%>" />
	<input type="hidden" name="responderurl" value="<%=storeurl%>vsadmin/ncconfirm.asp" />
	<input type="hidden" name="description" value="<%=Left(descstr,255)%>" />
	<input type="hidden" name="ordernumber" value="<%=orderid%>" />
	<input type="hidden" name="amount" value="<%=FormatNumber(grandtotal,2,-1,0,0)%>" />
	<input type="hidden" name="firstline" value="<%=Request.form("address")& IIfVr(trim(Request.form("address2"))<>"",", " & Request.form("address2"), "")%>" />
	<input type="hidden" name="town" value="<%=Request.form("city")%>" />
	<input type="hidden" name="county" value="<%if Trim(Request.form("state"))<>"" then response.write Trim(Request.form("state")) else response.write Trim(Request.form("state2"))%>" />
	<input type="hidden" name="postcode" value="<%=Request.form("zip")%>" />
	<input type="hidden" name="email_address_sender" value="<%=Request.form("email")%>" />
<%			thename = Trim(Request.form("name"))
			if thename<>"" then
				if InStr(thename," ") > 0 then
					namearr = Split(thename," ",2)
					response.write "<input type=""hidden"" name=""firstname"" value="""&replace(namearr(0),"""","&quot;")&""" />"&vbCrLf
					response.write "<input type=""hidden"" name=""lastname"" value="""&replace(namearr(1),"""","&quot;")&""" />"&vbCrLf
				else
					response.write "<input type=""hidden"" name=""lastname"" value="""&replace(thename,"""","&quot;")&""" />"&vbCrLf
				end if
			end if
			if demomode then response.write "<input type=""hidden"" name=""status"" value=""test"" />"
		elseif grandtotal > 0 AND ordPayProvider="7" then ' VeriSign Payflow Pro %>
	<form method="post" action="cart.asp" onsubmit="return isvalidcard(this)">
	<input type="hidden" name="mode" value="authorize" />
	<input type="hidden" name="method" value="7" />
	<input type="hidden" name="ordernumber" value="<%=orderid%>" />
<%		elseif grandtotal > 0 AND ordPayProvider="8" then ' Payflow Link
			paymentlink = "https://payments.verisign.com/payflowlink"
			if data2="VSA" then paymentlink="https://payments.verisign.com.au/payflowlink" %>
	<form method="post" action="<%=paymentlink%>">
	<input type="hidden" name="LOGIN" value="<%=data1%>" />
	<input type="hidden" name="PARTNER" value="<%=data2%>" />
	<input type="hidden" name="CUSTID" value="<%=orderid%>" />
	<input type="hidden" name="AMOUNT" value="<%=FormatNumber(grandtotal,2,-1,0,0)%>" />
	<input type="hidden" name="TYPE" value="S" />
	<input type="hidden" name="DESCRIPTION" value="<%=Left(descstr,255)%>" />
	<input type="hidden" name="NAME" value="<%=Request.form("name")%>" />
	<input type="hidden" name="ADDRESS" value="<%=Request.form("address")& IIfVr(trim(Request.form("address2"))<>"",", " & Request.form("address2"), "")%>" />
	<input type="hidden" name="CITY" value="<%=Request.form("city")%>" />
	<input type="hidden" name="STATE" value="<%if Trim(Request.form("state"))<>"" then response.write Trim(Request.form("state")) else response.write Trim(Request.form("state2"))%>" />
	<input type="hidden" name="ZIP" value="<%=Request.form("zip")%>" />
	<input type="hidden" name="COUNTRY" value="<%=Request.form("country")%>" />
	<input type="hidden" name="EMAIL" value="<%=Request.form("email")%>" />
	<input type="hidden" name="PHONE" value="<%=Request.form("phone")%>" />
	<input type="hidden" name="METHOD" value="CC" />
	<input type="hidden" name="ORDERFORM" value="TRUE" />
	<input type="hidden" name="SHOWCONFIRM" value="FALSE" />
<%			if trim(Request.form("sname")) <> "" OR trim(Request.form("saddress")) <> "" then %>
	<input type="hidden" name="NAMETOSHIP" value="<%=Request.form("sname")%>" />
	<input type="hidden" name="ADDRESSTOSHIP" value="<%=Request.form("saddress")& IIfVr(trim(Request.form("saddress2"))<>"",", " & Request.form("saddress2"), "")%>" />
	<input type="hidden" name="CITYTOSHIP" value="<%=Request.form("scity")%>" />
	<input type="hidden" name="STATETOSHIP" value="<%if Trim(Request.form("sstate"))<>"" then response.write Trim(Request.form("sstate")) else response.write Trim(Request.form("sstate2"))%>" />
	<input type="hidden" name="ZIPTOSHIP" value="<%=Request.form("szip")%>" />
	<input type="hidden" name="COUNTRYTOSHIP" value="<%=Request.form("scountry")%>" />
<%			end if
		elseif grandtotal > 0 AND ordPayProvider="9" then ' Secpay %>
	<form method="post" action="https://www.secpay.com/java-bin/ValCard">
	<input type="hidden" name="merchant" value="<%=data1%>" />
	<input type="hidden" name="trans_id" value="<%=orderid%>" />
	<input type="hidden" name="amount" value="<%=FormatNumber(grandtotal,2,-1,0,0)%>" />
	<input type="hidden" name="callback" value="<%=storeurl%>vsadmin/wpconfirm.asp" />
	<input type="hidden" name="currency" value="<%=countryCurrency%>" />
	<input type="hidden" name="cb_post" value="true" />
	<input type="hidden" name="bill_name" value="<%=Request.form("name")%>" />
	<input type="hidden" name="bill_addr_1" value="<%=Request.form("address")%>" />
	<input type="hidden" name="bill_addr_2" value="<%=Request.form("address2")%>" />
	<input type="hidden" name="bill_city" value="<%=Request.form("city")%>" />
	<input type="hidden" name="bill_state" value="<%if Trim(Request.form("state"))<>"" then response.write Trim(Request.form("state")) else response.write Trim(Request.form("state2"))%>" />
	<input type="hidden" name="bill_post_code" value="<%=Request.form("zip")%>" />
	<input type="hidden" name="bill_country" value="<%=Request.form("country")%>" />
	<input type="hidden" name="bill_email" value="<%=Request.form("email")%>" />
	<input type="hidden" name="bill_tel" value="<%=Request.form("phone")%>" />
<%			if trim(Request.form("sname")) <> "" OR trim(Request.form("saddress")) <> "" then %>
	<input type="hidden" name="ship_name" value="<%=Request.form("sname")%>" />
	<input type="hidden" name="ship_addr_1" value="<%=Request.form("saddress")%>" />
	<input type="hidden" name="ship_addr_2" value="<%=Request.form("saddress2")%>" />
	<input type="hidden" name="ship_city" value="<%=Request.form("scity")%>" />
	<input type="hidden" name="ship_state" value="<%if Trim(Request.form("sstate"))<>"" then response.write Trim(Request.form("sstate")) else response.write Trim(Request.form("sstate2"))%>" />
	<input type="hidden" name="ship_post_code" value="<%=Request.form("szip")%>" />
	<input type="hidden" name="ship_country" value="<%=Request.form("scountry")%>" />
<%			end if
			data2arr = split(data2,"&",2)
			if UBOUND(data2arr) >= 0 then data2md5 = data2arr(0)
			if UBOUND(data2arr) > 0 then data2tpl = data2arr(1)
			if trim(data2md5) <> "" then
				Session.LCID = 1033
%>	<input type="hidden" name="digest" value="<%=calcmd5(orderid & FormatNumber(grandtotal,2,-1,0,0) & data2md5)%>" />
	<input type="hidden" name="md_flds" value="trans_id:amount:callback" />
<%				Session.LCID = saveLCID
			end if
			if trim(data2tpl) <> "" then response.write "<input type=""hidden"" name=""template"" value=""" & urldecode(data2tpl) & """ />"
			if ppmethod=1 then response.write "<input type=""hidden"" name=""deferred"" value=""reuse:5:5"" />"
			if requirecvv=TRUE then response.write "<input type=""hidden"" name=""req_cv2"" value=""true"" />"
			if demomode then response.write "<input type=""hidden"" name=""options"" value=""test_status=true,dups=false"" />"
		elseif grandtotal > 0 AND ordPayProvider="10" then ' Capture Card %>
	<form method="post" action="thanks.asp" onsubmit="return isvalidcard(this)">
	<input type="hidden" name="docapture" value="vsprods" />
	<input type="hidden" name="ordernumber" value="<%=orderid%>" />
<%		elseif grandtotal > 0 AND (ordPayProvider="11" OR ordPayProvider="12") then ' PSiGate %>
	<form method="post" action="https://order.psigate.com/psigate.asp" <% if ordPayProvider="12" then response.write "onsubmit=""return isvalidcard(this)""" %>>
	<input type="hidden" name="MerchantID" value="<%=data1%>" />
	<input type="hidden" name="Oid" value="<%=orderid%>" />
<%			Session.LCID = 1033 %>
	<input type="hidden" name="FullTotal" value="<%=FormatNumber(grandtotal,2,-1,0,0)%>" />
<%			Session.LCID = saveLCID %>
	<input type="hidden" name="ThanksURL" value="<%=storeurl%>thanks.asp" />
	<input type="hidden" name="NoThanksURL" value="<%=storeurl%>thanks.asp" />
	<input type="hidden" name="Chargetype" value="<% if ppmethod=1 then response.write "1" else response.write "0" %>" />
	<% if ordPayProvider="11" then %><input type="hidden" name="Bname" value="<%=Request.form("name")%>" /><% end if %>
	<input type="hidden" name="Baddr1" value="<%=Request.form("address")%>" />
	<input type="hidden" name="Baddr2" value="<%=Request.form("address2")%>" />
	<input type="hidden" name="Bcity" value="<%=Request.form("city")%>" />
	<input type="hidden" name="IP" value="<%=left(request.servervariables("REMOTE_ADDR"), 48)%>" />
<%			if countryID=1 AND stateAbbrev<>"" then %>
	<input type="hidden" name="Bstate" value="<%=stateAbbrev%>" />
<%			else %>
	<input type="hidden" name="Bstate" value="<%if Trim(Request.form("state"))<>"" then response.write Trim(Request.form("state")) else response.write Trim(Request.form("state2"))%>" />
<%			end if %>
	<input type="hidden" name="Bzip" value="<%=Request.form("zip")%>" />
	<input type="hidden" name="Bcountry" value="<%=countryCode%>" />
	<input type="hidden" name="Email" value="<%=Request.form("email")%>" />
	<input type="hidden" name="Phone" value="<%=Request.form("phone")%>" />
<%			if trim(Request.form("sname")) <> "" OR trim(Request.form("saddress")) <> "" then %>
	<input type="hidden" name="Sname" value="<%=Request.form("sname")%>" />
	<input type="hidden" name="Saddr1" value="<%=Request.form("saddress")%>" />
	<input type="hidden" name="Saddr2" value="<%=Request.form("saddress2")%>" />
	<input type="hidden" name="Scity" value="<%=Request.form("scity")%>" />
	<input type="hidden" name="Sstate" value="<%if Trim(Request.form("sstate"))<>"" then response.write Trim(Request.form("sstate")) else response.write Trim(Request.form("sstate2"))%>" />
	<input type="hidden" name="Szip" value="<%=Request.form("szip")%>" />
	<input type="hidden" name="Scountry" value="<%=Request.form("scountry")%>" />
<%			end if
			if demomode then %>
	<input type="hidden" name="Result" value="1" />
<%			end if
		elseif grandtotal > 0 AND ordPayProvider="13" then ' Authorize.net AIM %>
	<form method="post" action="cart.asp" onsubmit="return isvalidcard(this)">
	<input type="hidden" name="mode" value="authorize" />
	<input type="hidden" name="method" value="13" />
	<input type="hidden" name="ordernumber" value="<%=orderid%>" />
	<input type="hidden" name="description" value="<%=Left(descstr,254)%>" />
<%		elseif grandtotal > 0 AND ordPayProvider="14" then ' Custom Pay Provider %>
<!--#include file="customppsend.asp"-->
<%		elseif grandtotal > 0 AND ordPayProvider="15" then ' Netbanx %>
	<form method="post" action="https://www.netbanx.com/cgi-bin/payment/<%=data1%>">
	<input type="hidden" name="order_id" value="<%=orderid%>" />
	<input type="hidden" name="payment_amount" value="<%=FormatNumber(grandtotal,2,-1,0,0)%>" />
	<input type="hidden" name="currency_code" value="<%=countryCurrency%>" />
	<input type="hidden" name="cardholder_name" value="<%=Request.form("name")%>" />
	<input type="hidden" name="email" value="<%=Request.form("email")%>" />
	<input type="hidden" name="postcode" value="<%=Request.form("zip")%>" />
<%		elseif grandtotal > 0 AND ordPayProvider="16" then ' Linkpoint
			if demomode then theurl="https://staging.linkpt.net/lpc/servlet/lppay" else theurl="https://www.linkpointcentral.com/lpc/servlet/lppay"
			lpsubtotal = vsround(totalgoods - totaldiscounts, 2)
			lpshipping = vsround((shipping + handling) - freeshipamnt, 2)
			lptax = vsround(stateTax + countryTax, 2)
			randomize
			sequence = Int(1000000 * Rnd) + 1000000
  %><form action="<%=theurl%>" method="post"<%if data2="1" then response.write " onsubmit=""return isvalidcard(this)"""%>>
	<input type="hidden" name="storename" value="<%=data1%>" />
	<input type="hidden" name="mode" value="payonly" />
	<input type="hidden" name="ponumber" value="<%=orderid%>" />
	<input type="hidden" name="oid" value="<%=orderid&"."&sequence%>" />
	<input type="hidden" name="responseURL" value="<%=storeurl%>thanks.asp" />
	<input type="hidden" name="subtotal" value="<%=FormatNumber(lpsubtotal,2,-1,0,0)%>" />
	<input type="hidden" name="chargetotal" value="<%=FormatNumber(lpsubtotal+lpshipping+lptax,2,-1,0,0)%>" />
	<input type="hidden" name="shipping" value="<%=FormatNumber(lpshipping,2,-1,0,0)%>" />
	<input type="hidden" name="tax" value="<%=FormatNumber(lptax,2,-1,0,0)%>" />
	<%if data2<>"1" then %><input type="hidden" name="bname" value="<%=Request.form("name")%>" /><% end if %>
	<input type="hidden" name="baddr1" value="<%=Request.form("address")%>" />
	<input type="hidden" name="baddr2" value="<%=Request.form("address2")%>" />
	<input type="hidden" name="bcity" value="<%=Request.form("city")%>" />
<%			if countryID=1 AND stateAbbrev<>"" then %>
	<input type="hidden" name="bstate" value="<%=stateAbbrev%>" />
<%			else %>
	<input type="hidden" name="bstate2" value="<%if Trim(Request.form("state"))<>"" then response.write Trim(Request.form("state")) else response.write Trim(Request.form("state2"))%>" />
<%			end if %>
	<input type="hidden" name="bzip" value="<%=Request.form("zip")%>" />
	<input type="hidden" name="bcountry" value="<%=countryCode%>" />
	<input type="hidden" name="email" value="<%=Request.form("email")%>" />
	<input type="hidden" name="phone" value="<%=Request.form("phone")%>" />
	<input type="hidden" name="txntype" value="<% if ppmethod=1 then response.write "preauth" else response.write "sale" %>" />
<%			if trim(Request.form("sname")) <> "" OR trim(Request.form("saddress")) <> "" then %>
	<input type="hidden" name="sname" value="<%=Request.form("sname")%>" />
	<input type="hidden" name="saddr1" value="<%=Request.form("saddress")%>" />
	<input type="hidden" name="saddr2" value="<%=Request.form("saddress2")%>" />
	<input type="hidden" name="scity" value="<%=Request.form("scity")%>" />
	<input type="hidden" name="sstate" value="<%if Trim(Request.form("sstate"))<>"" then response.write Trim(Request.form("sstate")) else response.write Trim(Request.form("sstate2"))%>" />
	<input type="hidden" name="szip" value="<%=Request.form("szip")%>" />
	<input type="hidden" name="scountry" value="<%=shipCountryCode%>" />
<%			end if
			if demomode then %>
	<input type="hidden" name="txnmode" value="test" />
<%			end if
		elseif grandtotal > 0 AND ordPayProvider="18" then ' PayPal Payment Pro %>
	<form method="post" action="cart.asp" onsubmit="return isvalidcard(this)">
	<input type="hidden" name="mode" value="authorize" />
	<input type="hidden" name="method" value="18" />
	<input type="hidden" name="ordernumber" value="<%=orderid%>" />
	<input type="hidden" name="description" value="<%=replace(left(descstr,254),"""","&quot;")%>" />
<%		elseif grandtotal > 0 AND ordPayProvider="19" then ' PayPal Express Payment %>
	<form method="post" action="thanks.asp">
	<input type="hidden" name="token" value="<%=token%>" />
	<input type="hidden" name="method" value="paypalexpress" />
	<input type="hidden" name="ordernumber" value="<%=orderid%>" />
	<input type="hidden" name="payerid" value="<%=payerid%>" />
	<input type="hidden" name="email" value="<%=ordEmail%>" />
<%		end if
	end if
	if success then
%><br />
            <table class="cobtbl" width="<%=maintablewidth%>" border="0" bordercolor="#B1B1B1" cellspacing="1" cellpadding="3" bgcolor="#B1B1B1">
			  <tr>
			    <td class="cobhl" bgcolor="#EBEBEB" height="30" colspan="2" align="center"><strong><%=xxChkCmp%></strong></td>
			  </tr>
<%	if cpncode<>"" AND NOT gotcpncode then %>
			  <tr>
			    <td class="cobhl" bgcolor="#EBEBEB" height="30" align="right" width="50%"><strong><%=xxGifCer%>:</strong></td>
				<td class="cobll" bgcolor="#FFFFFF" height="30" align="left" width="50%"><font size="1"><%
							if shippingpost="" then jumpback=1 else jumpback=2
							response.write Replace(Replace(xxNoGfCr,"%s",cpncode,1,1),"%s",jumpback,1,1)%></font></td>
			  </tr>
<%	end if
	if cpnmessage<>"" then %>
			  <tr>
			    <td class="cobhl" bgcolor="#EBEBEB" height="30" align="right" width="50%"><strong><%=xxAppDs%>:</strong></td>
				<td class="cobll" bgcolor="#FFFFFF" height="30" align="left" width="50%"><%=cpnmessage%></td>
			  </tr>
<%	end if %>
			  <tr>
			    <td class="cobhl" bgcolor="#EBEBEB" height="30" align="right" width="50%"><strong><%=xxTotGds%>:</strong></td>
				<td class="cobll" bgcolor="#FFFFFF" height="30" align="left" width="50%"><%=FormatEuroCurrency(totalgoods)%></td>
			  </tr>
<%	if combineshippinghandling=TRUE then %>
			  <tr>
			    <td class="cobhl" bgcolor="#EBEBEB" height="30" align="right" width="50%"><strong><%=xxShipHa%>:</strong></td>
				<td class="cobll" bgcolor="#FFFFFF" height="30" align="left" width="50%"><%=FormatEuroCurrency((shipping+handling)-freeshipamnt)%></td>
			  </tr>
<%	else
		if shipType<>0 then %>
			  <tr>
			    <td class="cobhl" bgcolor="#EBEBEB" height="30" align="right" width="50%"><strong><%=xxShippg%>:</strong></td>
				<td class="cobll" bgcolor="#FFFFFF" height="30" align="left" width="50%"><%=FormatEuroCurrency(shipping-freeshipamnt)%></td>
			  </tr>
<%		end if
		if handling<>0 then %>
			  <tr>
			    <td class="cobhl" bgcolor="#EBEBEB" height="30" align="right" width="50%"><strong><%=xxHndlg%>:</strong></td>
				<td class="cobll" bgcolor="#FFFFFF" height="30" align="left" width="50%"><%=FormatEuroCurrency(handling)%></td>
			  </tr>
<%		end if
	end if
	if totaldiscounts<>0 then %>
			  <tr>
			    <td class="cobhl" bgcolor="#EBEBEB" height="30" align="right" width="50%"><strong><%=xxTotDs%>:</strong></td>
				<td class="cobll" bgcolor="#FFFFFF" height="30" align="left" width="50%"><font color="#FF0000"><%=FormatEuroCurrency(totaldiscounts)%></font></td>
			  </tr>
			  <tr>
			    <td class="cobhl" bgcolor="#EBEBEB" height="30" align="right" width="50%"><strong><%=xxSubTot%>:</strong></td>
				<td class="cobll" bgcolor="#FFFFFF" height="30" align="left" width="50%"><%=FormatEuroCurrency((totalgoods+shipping+handling)-(totaldiscounts+freeshipamnt))%></td>
			  </tr>
<%	end if
	if usehst then %>
			  <tr>
			    <td class="cobhl" bgcolor="#EBEBEB" height="30" align="right" width="50%"><strong><%=xxHST%>:</strong></td>
				<td class="cobll" bgcolor="#FFFFFF" height="30" align="left" width="50%"><%=FormatEuroCurrency(stateTax+countryTax)%></td>
			  </tr>
<%	else
		if stateTax<>0.0 then %>
			  <tr>
			    <td class="cobhl" bgcolor="#EBEBEB" height="30" align="right" width="50%"><strong><%=xxStaTax%>:</strong></td>
				<td class="cobll" bgcolor="#FFFFFF" height="30" align="left" width="50%"><%=FormatEuroCurrency(stateTax)%></td>
			  </tr>
<%		end if
		if countryTax<>0.0 then %>
			  <tr>
			    <td class="cobhl" bgcolor="#EBEBEB" height="30" align="right" width="50%"><strong><%=xxCntTax%>:</strong></td>
				<td class="cobll" bgcolor="#FFFFFF" height="30" align="left" width="50%"><%=FormatEuroCurrency(countryTax)%></td>
			  </tr>
<%		end if
	end if %>
			  <tr>
			    <td class="cobhl" bgcolor="#EBEBEB" height="30" align="right" width="50%"><strong><%=xxGndTot%>:</strong></td>
				<td class="cobll" bgcolor="#FFFFFF" height="30" align="left" width="50%"><%=FormatEuroCurrency(grandtotal)%></td>
			  </tr>
<% if grandtotal > 0 AND (ordPayProvider="7" OR ordPayProvider="10" OR ordPayProvider="12" OR ordPayProvider="13" OR (ordPayProvider="16" AND (data2&"")="1") OR ordPayProvider="18") then ' VeriSign Payflow Pro OR Capture Card OR PSiGate SSL OR Auth.NET AIM OR PayPal Pro
		if ordPayProvider="7" OR ordPayProvider="12" OR ordPayProvider="13" OR ordPayProvider="16" OR ordPayProvider="18" then data1 = "XXXXXXX0XXXXXXXXXXXXXXXXX"
		isPSiGate = (ordPayProvider="12")
		isLinkpoint = (ordPayProvider="16")
		if isPSiGate then
			sscardname="bname"
			sscardnum = "CardNumber"
			ssexmon = "ExpMonth"
			ssexyear = "ExpYear"
		elseif isLinkpoint then
			sscardname="bname"
			sscardnum = "cardnumber"
			ssexmon = "expmonth"
			ssexyear = "expyear"
			sscvv2 = "cvm"
		else
			sscardname="cardname"
			sscardnum = "ACCT"
			ssexmon = "EXMON"
			ssexyear = "EXYEAR"
			sscvv2 = "CVV2"
		end if
		acceptecheck = (acceptecheck=true) AND (ordPayProvider="13")
%>
<input type="hidden" name="vrshippingoptions" value="<%=request.form("vrshippingoptions")%>" />
<input type="hidden" name="sessionid" value="<%=thesessionid%>" />
<script language="javascript" type="text/javascript">
<!--
var isswitchcard=false;
function isCreditCard(st){
  // Encoding only works on cards with less than 19 digits
  if (st.length > 19)
    return (false);
  sum = 0; mul = 1; l = st.length;
  for (i = 0; i < l; i++) {
    digit = st.substring(l-i-1,l-i);
    tproduct = parseInt(digit ,10)*mul;
    if (tproduct >= 10)
      sum += (tproduct % 10) + 1;
    else
      sum += tproduct;
    if (mul == 1)
      mul++;
    else
      mul = mul - 1;
  }
  if ((sum % 10) == 0)
    return (true);
  else
    return (false);
}
function isVisa(cc){ // 4111 1111 1111 1111
  if (((cc.length == 16) || (cc.length == 13)) && (cc.substring(0,1) == 4))
    return isCreditCard(cc);
  return false;
}
function isMasterCard(cc){ // 5500 0000 0000 0004
  firstdig = cc.substring(0,1);
  seconddig = cc.substring(1,2);
  if ((cc.length == 16) && (firstdig == 5) && ((seconddig >= 1) && (seconddig <= 5)))
    return isCreditCard(cc);
  return false;
}
function isAmericanExpress(cc){ // 340000000000009
  firstdig = cc.substring(0,1);
  seconddig = cc.substring(1,2);
  if ((cc.length == 15) && (firstdig == 3) && ((seconddig == 4) || (seconddig == 7)))
    return isCreditCard(cc);
  return false;
}
function isDinersClub(cc){ // 30000000000004
  firstdig = cc.substring(0,1);
  seconddig = cc.substring(1,2);
  if ((cc.length == 14) && (firstdig == 3) &&
      ((seconddig == 0) || (seconddig == 6) || (seconddig == 8)))
    return isCreditCard(cc);
  return false;
}
function isDiscover(cc){ // 6011000000000004
  first4digs = cc.substring(0,4);
  if ((cc.length == 16) && (first4digs == "6011"))
    return isCreditCard(cc);
  return false;
}
function isAusBankcard(cc){ // 5610591000000009
  first4digs = cc.substring(0,4);
  if ((cc.length == 16) && (first4digs == "5610"))
    return isCreditCard(cc);
  return false;
}
function isEnRoute(cc){ // 201400000000009
  first4digs = cc.substring(0,4);
  if ((cc.length == 15) && ((first4digs == "2014") || (first4digs == "2149")))
    return isCreditCard(cc);
  return false;
}
function isJCB(cc){
  first4digs = cc.substring(0,4);
  if ((cc.length == 16) && ((first4digs == "3088") || (first4digs == "3096") || (first4digs == "3112") || (first4digs == "3158") || (first4digs == "3337") || (first4digs == "3528")))
    return isCreditCard(cc);
  return false;
}
function isSwitch(cc){ // 675911111111111128
  first4digs = cc.substring(0,4);
  if ((cc.length == 16 || cc.length == 17 || cc.length == 18 || cc.length == 19) && ((first4digs == "4903") || (first4digs == "4911") || (first4digs == "4936") || (first4digs == "5641") || (first4digs == "6333") || (first4digs == "6759") || (first4digs == "6334") || (first4digs == "6767"))){
    isswitchcard=isCreditCard(cc);
    return(isswitchcard);
  }
  return false;
}
function isvalidcard(theForm){
  cc = theForm.<%=sscardnum%>.value;
  newcode = "";
  isswitchcard=false;
  l = cc.length;
  for(i=0;i<l;i++){
	digit = cc.substring(i,i+1);
	digit = parseInt(digit ,10);
	if(!isNaN(digit)) newcode += digit;
  }
  cc=newcode;
  if (theForm.<%=sscardname%>.value==""){
	alert("<%=xxPlsEntr & " \""" & xxCCName & "\""" %>");
	theForm.<%=sscardname%>.focus();
    return false;
  }
<% if acceptecheck=true then %>
if(cc!="" && theForm.accountnum.value!=""){
alert("Please enter either Credit Card OR ECheck details");
return(false);
}else if(theForm.accountnum.value!=""){
  if(theForm.accountname.value==""){
    alert("Please enter a value in the field \"Account Name\".");
	theForm.accountname.focus();
    return false;
  }
  if(theForm.bankname.value==""){
    alert("Please enter a value in the field \"Bank Name\".");
	theForm.bankname.focus();
    return false;
  }
  if(theForm.routenumber.value==""){
    alert("Please enter a value in the field \"Routing Number\".");
	theForm.routenumber.focus();
    return false;
  }
  if(theForm.accounttype.selectedIndex==0){
    alert("Please select your account type: (Checking / Savings).");
	theForm.accounttype.focus();
    return false;
  }
<%		if wellsfargo=true then %>
  if(theForm.orgtype.selectedIndex==0){
    alert("Please select your account type: (Personal / Business).");
	theForm.orgtype.focus();
    return false;
  }
  if(theForm.taxid.value=="" && theForm.licensenumber.value==""){
    alert("Please enter either a Tax ID number or Drivers License Details.");
	theForm.taxid.focus();
    return false;
  }
  if(theForm.taxid.value==""){
	  if(theForm.licensestate.selectedIndex==0){
		alert("Please select your Drivers License State.");
		theForm.licensestate.focus();
		return false;
	  }
	  if(theForm.dldobmon.selectedIndex==0){
		alert("Please select your Drivers License D.O.B. Month.");
		theForm.dldobmon.focus();
		return false;
	  }
	  if(theForm.dldobday.selectedIndex==0){
		alert("Please select your Drivers License D.O.B. Day.");
		theForm.dldobday.focus();
		return false;
	  }
	  if(theForm.dldobyear.selectedIndex==0){
		alert("Please select your Drivers License D.O.B. year.");
		theForm.dldobyear.focus();
		return false;
	  }
  }
<%		end if %>
}else{
<% end if %>
  if (true <% 
		if Mid(data1,1,1)="X" then response.write "&& !isVisa(cc) "
		if Mid(data1,2,1)="X" then response.write "&& !isMasterCard(cc) "
		if Mid(data1,3,1)="X" then response.write "&& !isAmericanExpress(cc) "
		if Mid(data1,4,1)="X" then response.write "&& !isDinersClub(cc) "
		if Mid(data1,5,1)="X" then response.write "&& !isDiscover(cc) "
		if Mid(data1,6,1)="X" then response.write "&& !isEnRoute(cc) "
		if Mid(data1,7,1)="X" then response.write "&& !isJCB(cc) "
		if Mid(data1,8,1)="X" then response.write "&& !isSwitch(cc) "
		if Mid(data1,9,1)="X" then response.write "&& !isAusBankcard(cc)" %>){
	<% if acceptecheck=true then xxValCC="Please enter a valid credit card number or bank account details if paying by ECheck." %>
	alert("<%=xxValCC%>");
	theForm.<%=sscardnum%>.focus();
    return false;
  }
  if(theForm.<%=ssexmon%>.selectedIndex==0){
    alert("<%=xxCCMon%>");
	theForm.<%=ssexmon%>.focus();
    return false;
  }
  if(theForm.<%=ssexyear%>.selectedIndex==0){
    alert("<%=xxCCYear%>");
	theForm.<%=ssexyear%>.focus();
    return false;
  }
<% if Mid(data1,8,1)="X" then %>
  if(theForm.IssNum.value=="" && isswitchcard){
    alert("Please enter an issue number / start date for Switch/Solo cards.");
	theForm.IssNum.focus();
    return false;
  }
<% end if
   if requirecvv=true then %>
  if(theForm.<%=sscvv2%>.value==""){
    alert("<%=xxPlsEntr & " \""" & replace(xx34code,"""","\""") & "\"""%>");
	theForm.<%=sscvv2%>.focus();
    return false;
  }
<% end if
   if acceptecheck=true then response.write "}" %>
  return true;
}
//-->
</script>
<%	if request.servervariables("HTTPS")<>"on" AND (Request.ServerVariables("SERVER_PORT_SECURE") <> "1") AND nochecksslserver<>true then %>
			  <tr height="30">
			    <td class="cobhl" bgcolor="#EBEBEB" align="center" colspan="2"><strong><font color="#FF0000">This site may not be secure. Do not enter real Credit Card numbers.</font></strong></td>
			  </tr>
<%	end if %>
			  <tr>
			    <td class="cobhl" bgcolor="#EBEBEB" height="30" colspan="2" align="center"><strong><%=xxCCDets%></strong></td>
			  </tr>
			  <tr height="30">
			    <td class="cobhl" bgcolor="#EBEBEB" align="right" width="50%"><strong><%=xxCCName%>:</strong></td>
				<td class="cobll" bgcolor="#FFFFFF" align="left" width="50%"><input type="text" name="<%=sscardname%>" size="21" value="<%=request.form("name")%>" AUTOCOMPLETE="off" /></td>
			  </tr>
			  <tr height="30">
			    <td class="cobhl" bgcolor="#EBEBEB" align="right" width="50%"><strong><%=xxCrdNum%>:</strong></td>
				<td class="cobll" bgcolor="#FFFFFF" align="left" width="50%"><input type="text" name="<%=sscardnum%>" size="21" AUTOCOMPLETE="off" /></td>
			  </tr>
			  <tr height="30">
			    <td class="cobhl" bgcolor="#EBEBEB" align="right" width="50%"><strong><%=xxExpEnd%>:</strong></td>
				<td class="cobll" bgcolor="#FFFFFF" align="left" width="50%">
				  <select name="<%=ssexmon%>" size="1">
					<option value=""><%=xxMonth%></option>
					<%	for index=1 to 12
							if index < 10 then themonth = "0" & index else themonth = index
							response.write "<option value='"&themonth&"'>"&themonth&"</option>"&vbCrLf
						next %>
				  </select> / <select name="<%=ssexyear%>" size="1">
					<option value=""><%=xxYear%></option>
					<%	thisyear=DatePart("yyyy", Date())
						for index=thisyear to thisyear+10
							if isPSiGate then
								tmpYear = right(index,2)
								if Len(tmpYear)<2 then tmpYear = "0" & tmpYear
								response.write "<option value='"&tmpYear&"'>"&index&"</option>"&vbCrLf
							else
								response.write "<option value='"&index&"'>"&index&"</option>"&vbCrLf
							end if
						next %>
				  </select>
				</td>
			  </tr>
			<%	if NOT isPSiGate then %>
			  <tr height="30">
			    <td class="cobhl" bgcolor="#EBEBEB" align="right" width="50%"><strong><%=xx34code%>:</strong></td>
				<td class="cobll" bgcolor="#FFFFFF" align="left" width="50%"><input type="text" name="<%=sscvv2%>" size="4" AUTOCOMPLETE="off" /> <strong><%if requirecvv<>true then response.write xxIfPres%></strong></td>
			  </tr>
			<%	end if
				if Mid(data1,8,1)="X" then %>
			  <tr height="30">
			    <td class="cobhl" bgcolor="#EBEBEB" align="right" width="50%"><strong>Issue Number / Start Date:</strong></td>
				<td class="cobll" bgcolor="#FFFFFF" align="left" width="50%"><input type="text" name="IssNum" size="4" AUTOCOMPLETE="off" /> <strong>(Switch/Solo Only)</strong></td>
			  </tr>
<%				end if
				if acceptecheck=true then ' Auth.net %>
			  <tr>
			    <td class="cobhl" bgcolor="#EBEBEB" height="30" colspan="2" align="center"><strong>ECheck Details</strong><br /><font size="1">Please enter either Credit Card OR ECheck details</font></td>
			  </tr>
			  <tr height="30">
			    <td class="cobhl" bgcolor="#EBEBEB" align="right" width="50%"><strong>Account Name:</strong></td>
				<td class="cobll" bgcolor="#FFFFFF" align="left" width="50%"><input type="text" name="accountname" size="21" AUTOCOMPLETE="off" value="<%=request.form("name")%>" /></td>
			  </tr>
			  <tr height="30">
			    <td class="cobhl" bgcolor="#EBEBEB" align="right" width="50%"><strong>Account Number:</strong></td>
				<td class="cobll" bgcolor="#FFFFFF" align="left" width="50%"><input type="text" name="accountnum" size="21" AUTOCOMPLETE="off" /></td>
			  </tr>
			  <tr height="30">
			    <td class="cobhl" bgcolor="#EBEBEB" align="right" width="50%"><strong>Bank Name:</strong></td>
				<td class="cobll" bgcolor="#FFFFFF" align="left" width="50%"><input type="text" name="bankname" size="21" AUTOCOMPLETE="off" /></td>
			  </tr>
			  <tr height="30">
			    <td class="cobhl" bgcolor="#EBEBEB" align="right" width="50%"><strong>Routing Number:</strong></td>
				<td class="cobll" bgcolor="#FFFFFF" align="left" width="50%"><input type="text" name="routenumber" size="10" AUTOCOMPLETE="off" /></td>
			  </tr>
			  <tr height="30">
			    <td class="cobhl" bgcolor="#EBEBEB" align="right" width="50%"><strong>Account Type:</strong></td>
				<td class="cobll" bgcolor="#FFFFFF" align="left" width="50%"><select name="accounttype" size="1"><option value=""><%=xxPlsSel%></option><option value="CHECKING">Checking</option><option value="SAVINGS">Savings</option><option value="BUSINESSCHECKING">Business Checking</option></select></td>
			  </tr>
<%					if wellsfargo=true then %>
			  <tr height="30">
			    <td class="cobhl" bgcolor="#EBEBEB" align="right" width="50%"><strong>Personal or Business Acct.:</strong></td>
				<td class="cobll" bgcolor="#FFFFFF" align="left" width="50%"><select name="orgtype" size="1"><option value=""><%=xxPlsSel%></option><option value="I">Personal</option><option value="B">Business</option></select></td>
			  </tr>
			  <tr height="30">
			    <td class="cobhl" bgcolor="#EBEBEB" align="right" width="50%"><strong>Tax ID:</strong></td>
				<td class="cobll" bgcolor="#FFFFFF" align="left" width="50%"><input type="text" name="taxid" size="21" AUTOCOMPLETE="off" /></td>
			  </tr>
			  <tr>
			    <td class="cobhl" bgcolor="#EBEBEB" height="30" colspan="2" align="center"><font size="1">If you have provided a Tax ID then the following information is not necessary</font></td>
			  </tr>
			  <tr height="30">
			    <td class="cobhl" bgcolor="#EBEBEB" align="right" width="50%"><strong>Drivers License Number:</strong></td>
				<td class="cobll" bgcolor="#FFFFFF" align="left" width="50%"><input type="text" name="licensenumber" size="21" AUTOCOMPLETE="off" /></td>
			  </tr>
			  <tr height="30">
			    <td class="cobhl" bgcolor="#EBEBEB" align="right" width="50%"><strong>Drivers License State:</strong></td>
				<td class="cobll" bgcolor="#FFFFFF" align="left" width="50%"><select size="1" name="licensestate"><option value=""><%=xxPlsSel%></option><%
					sSQL = "SELECT stateName,stateAbbrev FROM states ORDER BY stateName"
					rs.Open sSQL,cnn,0,1
					do while not rs.EOF
						response.write "<option value="""&Replace(rs("stateAbbrev"),"""","&quot;")&""""
						response.write ">"&rs("stateName")&"</option>"&vbCrLf
						rs.MoveNext
					loop
					rs.Close %></select></td>
			  </tr>
			  <tr height="30">
			    <td class="cobhl" bgcolor="#EBEBEB" align="right" width="50%"><strong>Date Of Birth On License:</strong></td>
				<td class="cobll" bgcolor="#FFFFFF" align="left" width="50%">
				  <select name="dldobmon" size="1">
					<option value=""><%=xxMonth%></option>
					<% for index=1 to 12 %>
					<option value="<%=index%>"><%=MonthName(index)%></option>
					<% next %>
				  </select>
				  <select name="dldobday" size="1">
					<option value="">Day</option>
					<% for index=1 to 31 %>
					<option value="<%=index%>"><%=index%></option>
					<% next %>
				  </select>
				  <select name="dldobyear" size="1">
					<option value=""><%=xxYear%></option>
					<% for index=Year(date())-100 to Year(date()) %>
					<option value="<%=index%>"><%=index%></option>
					<% next %>
				  </select>
				</td>
			  </tr>
<%					end if
				end if
	end if %>
			  <tr>
			    <td class="cobhl" bgcolor="#EBEBEB" height="30" colspan="2" align="center"><strong><%=xxMstClk%></strong></td>
			  </tr>
			  <tr>
				<td class="cobll" bgcolor="#FFFFFF" colspan="2" align="center"><table width="100%" cellspacing="0" cellpadding="0" border="0">
				    <tr>
					  <td class="cobll" bgcolor="#FFFFFF" width="16" height="26" align="right" valign="bottom">&nbsp;</td>
					  <td class="cobll" bgcolor="#FFFFFF" width="100%" align="center"><% if orderid<>0 then %><input type="image" src="images/checkout.gif" border="0" alt="<%=xxCOTxt%>" /><% end if %></td>
					  <td class="cobll" bgcolor="#FFFFFF" width="16" height="26" align="right" valign="bottom"><img src="images/tablebr.gif" alt="" /></td>
					</tr>
				  </table></td>
			  </tr>
			</table>
	</form>
<%
	end if ' success
elseif checkoutmode="authorize" then
	blockuser=checkuserblock("")
	ordID = replace(Request.Form("ordernumber"), "'", "")
	gobackplaces=1
	call getpayprovdetails(Request.Form("method"),data1,data2,data3,demomode,ppmethod)
	if Request.Form("method")="7" then ' PayFlow Pro
		vsdetails = Split(data1, "&")
		if UBOUND(vsdetails) > 0 then
			vs1=vsdetails(0)
			vs2=vsdetails(1)
			vs3=vsdetails(2)
			vs4=vsdetails(3)
		end if
		sSQL = "SELECT ordZip,ordShipping,ordStateTax,ordCountryTax,ordHandling,ordTotal,ordDiscount,ordAddress,ordAddress2,ordAuthNumber FROM orders WHERE ordID="&ordID
		rs.Open sSQL,cnn,0,1
		vsAUTHCODE = (rs("ordAuthNumber")&"")
		theaddress = rs("ordAddress") & IIfVr(trim(rs("ordAddress2")&"")<>"", ", " & trim(rs("ordAddress2")), "")
		parmList = "TRXTYPE=" & IIfVr(ppmethod=1,"A","S") & "&TENDER=C&ZIP["&Len(rs("ordZip"))&"]="&rs("ordZip") & "&STREET["&len(theaddress)&"]="&theaddress
		parmList = parmList & "&NAME["&Len(Request.Form("cardname"))&"]="&Request.Form("cardname")
		parmList = parmList & "&COMMENT1="&ordID & "&ACCT=" & replace(request.form("ACCT")," ", "") & "&CUSTIP=" & request.servervariables("REMOTE_ADDR")
		parmList = parmList & "&PWD=" & vs4 & "&USER=" & vs1 & "&VENDOR=" & vs2 & "&PARTNER=" & vs3 & "&CVV2=" & Trim(request.form("CVV2"))
		parmList = parmList & "&EXPDATE=" & request.form("EXMON") & Right(request.form("EXYEAR"),2)
		parmList = parmList & "&AMT=" & FormatNumber((rs("ordShipping")+rs("ordStateTax")+rs("ordCountryTax")+rs("ordTotal")+rs("ordHandling"))-rs("ordDiscount"),2,-1,0,0)
		rs.Close
		if vsAUTHCODE="" then
			success=true
			if blockuser then
				success=FALSE
				vsRESPMSG = multipurchaseblockmessage
			else
				Set client = Server.CreateObject("PFProCOMControl.PFProCOMControl.1")
				if vs3="VSA" then
					theurl = "payflow.verisign.com.au"
					if demomode then theurl = "payflow-test.verisign.com.au"
				else
					theurl = "payflow.verisign.com"
					if demomode then theurl = "test-payflow.verisign.com"
				end if
				Ctx1 = client.CreateContext(theurl, 443, 30, "", 0, "", "")
				curString = client.SubmitTransaction(Ctx1, parmList, Len(parmList))
				client.DestroyContext (Ctx1)
				Do while Len(curString) <> 0
					'get the next name value pair
					if InStr(curString,"&") Then
						varString = Left(curString, InStr(curString , "&" ) -1)
					else
						varString = curString
					end if
					'get the name part of the name/value pair
					name = Left(varString, InStr(varString, "=" ) -1)
					value = Right(varString, Len(varString) - (Len(name)+1))
					if name="RESULT" then
						vsRESULT=value
					elseif name="PNREF" then
						vsPNREF=value
					elseif name="RESPMSG" then
						vsRESPMSG=value
					elseif name="AUTHCODE" then
						vsAUTHCODE=value
					elseif name="AVSADDR" then
						vsAVSADDR=value
					elseif name="AVSZIP" then
						vsAVSZIP=value
					elseif name="IAVS" then
						vsIAVS=value
					elseif name="CVV2MATCH" then
						vsCVV2=value
					end if
					'skip over the &
					if Len(curString) <> Len(varString) then curString = Right(curString, Len(curString) - (Len(varString)+1)) else curString = ""
				Loop
			end if
			if success then
				if vsRESULT="0" OR vsRESULT="126" then
					if vsRESULT="126" then underreview="Fraud Review:<br />" : vsRESPMSG="Approved" else underreview=""
					do_stock_management(ordID)
					cnn.Execute("UPDATE cart SET cartCompleted=1 WHERE cartOrderID="&ordID)
					cnn.Execute("UPDATE orders SET ordStatus=3,ordAVS='"&replace(vsAVSADDR&vsAVSZIP, "'", "")&"',ordCVV='"&replace(vsCVV2, "'", "")&"',ordAuthNumber='"&replace(underreview&vsAUTHCODE, "'", "")&"' WHERE ordID="&ordID)
					vsRESULT="0"
				end if
			end if
			set client = nothing
		else
			vsRESULT="0"
			vsRESPMSG="Approved"
		end if
	elseif Request.Form("method")="13" then ' Auth.net AIM
		if secretword<>"" then
			data1 = upsdecode(data1, secretword)
			data2 = upsdecode(data2, secretword)
		end if
		sSQL = "SELECT ordID,ordName,ordCity,ordState,ordCountry,ordPhone,ordHandling,ordZip,ordEmail,ordShipping,ordStateTax,ordCountryTax,ordTotal,ordDiscount,ordAddress,ordAddress2,ordIP,ordAuthNumber,ordShipName,ordShipAddress,ordShipAddress2,ordShipCity,ordShipState,ordShipCountry,ordShipZip FROM orders WHERE ordID="&ordID
		rs.Open sSQL,cnn,0,1
		vsAUTHCODE = trim(rs("ordAuthNumber")&"")
		parmList = "x_version=3.1&x_delim_data=True&x_relay_response=False&x_delim_char=|&x_duplicate_window=15"
		parmList = parmList & "&x_login="&data1&"&x_tran_key="&data2&"&x_cust_id="&rs("ordID")&"&x_Invoice_Num="&rs("ordID")
		parmList = parmList & "&x_amount="&FormatNumber((rs("ordShipping")+rs("ordStateTax")+rs("ordCountryTax")+rs("ordTotal")+rs("ordHandling"))-rs("ordDiscount"),2,-1,0,0)
		parmList = parmList & "&x_currency_code="&countryCurrency&"&x_Description=" & left(server.urlencode(request.form("description")),254)
		if trim(request.form("accountnum"))<>"" then
			parmList = parmList & "&x_method=ECHECK&x_echeck_type=WEB&x_recurring_billing=NO"
			parmList = parmList & "&x_bank_acct_name=" & server.urlencode(trim(request.form("accountname"))) & "&x_bank_acct_num=" & server.urlencode(trim(request.form("accountnum")))
			parmList = parmList & "&x_bank_name=" & server.urlencode(trim(request.form("bankname"))) & "&x_bank_aba_code=" & server.urlencode(trim(request.form("routenumber")))
			parmList = parmList & "&x_bank_acct_type=" & server.urlencode(trim(request.form("accounttype"))) & "&x_type=AUTH_CAPTURE"
			if wellsfargo=true then
				parmList = parmList & "&x_customer_organization_type=" & trim(request.form("orgtype"))
				if trim(request.form("taxid"))<>"" then
					parmList = parmList & "&x_customer_tax_id=" & server.urlencode(trim(request.form("taxid")))
				else
					parmList = parmList & "&x_drivers_license_num=" & server.urlencode(trim(request.form("licensenumber"))) & "&x_drivers_license_state=" & server.urlencode(trim(request.form("licensestate"))) & "&x_drivers_license_dob=" & server.urlencode(trim(request.form("dldobyear")) & "/" & trim(request.form("dldobmon")) & "/" & trim(request.form("dldobday")))
				end if
			end if
		else
			parmList = parmList & "&x_method=CC&x_card_num=" & server.urlencode(trim(request.form("ACCT"))) & "&x_exp_date=" & request.form("EXMON") & Right(request.form("EXYEAR"),2)
			if Trim(request.form("CVV2"))<>"" then parmList = parmList & "&x_card_code=" & server.urlencode(Trim(request.form("CVV2")))
			if ppmethod=1 then parmList = parmList & "&x_type=AUTH_ONLY" else parmList = parmList & "&x_type=AUTH_CAPTURE"
		end if
		thename = Trim(trim(request.form("cardname")))
		if thename<>"" then
			if InStr(thename," ") > 0 then
				namearr = Split(thename," ",2)
				parmList = parmList & "&x_first_name=" & server.urlencode(namearr(0)) & "&x_last_name=" & server.urlencode(namearr(1))
			else
				parmList = parmList & "&x_last_name=" & server.urlencode(thename)
			end if
		end if
		parmList = parmList & "&x_address="&server.urlencode(rs("ordAddress"))
		if trim(rs("ordAddress2")&"")<>"" then parmList = parmList & server.urlencode(", "&rs("ordAddress2"))
		parmList = parmList & "&x_city="&server.urlencode(rs("ordCity")) & "&x_state="&server.urlencode(rs("ordState")) & "&x_zip="&server.urlencode(rs("ordZip")) & "&x_country="&server.urlencode(rs("ordCountry")) & "&x_phone="&server.urlencode(rs("ordPhone")) & "&x_email="&server.urlencode(rs("ordEmail"))
		thename = trim(rs("ordShipName"))
		if thename<>"" OR rs("ordShipAddress")<>"" then
			if thename<>"" then
				if InStr(thename," ") > 0 then
					namearr = Split(thename," ",2)
					parmList = parmList & "&x_ship_to_first_name=" & server.urlencode(namearr(0)) & "&x_ship_to_last_name=" & server.urlencode(namearr(1))
				else
					parmList = parmList & "&x_ship_to_last_name=" & server.urlencode(thename)
				end if
			end if
			parmList = parmList & "&x_ship_to_address="&server.urlencode(rs("ordShipAddress"))
			if trim(rs("ordShipAddress2")&"")<>"" then parmList = parmList & server.urlencode(", "&rs("ordShipAddress2"))
			parmList = parmList & "&x_ship_to_city="&server.urlencode(rs("ordShipCity")) & "&x_ship_to_state="&server.urlencode(rs("ordShipState")) & "&x_ship_to_zip="&server.urlencode(rs("ordShipZip")) & "&x_ship_to_country="&server.urlencode(rs("ordShipCountry"))
		end if
		if Trim(rs("ordIP"))<>"" then parmList = parmList & "&x_customer_ip="&server.urlencode(Trim(rs("ordIP")))
		if demomode then parmList = parmList & "&x_test_request=TRUE"
		rs.Close
		if vsAUTHCODE="" then
			success=true
			if blockuser then
				success=FALSE
				vsRESPMSG = multipurchaseblockmessage
			else
				set objHttp = Server.CreateObject("Msxml2.ServerXMLHTTP")
				objHttp.open "POST", "https://secure.authorize.net/gateway/transact.dll", false
				objHttp.Send parmList
				If err.number <> 0 OR objHttp.status <> 200 Then
					errormsg = "Error, couldn't connect to Authorize.net server"
				Else
					varString = Split(objHttp.responseText, "|")
					vsRESULT=varString(0)
					vsERRCODE=varString(2)
					vsRESPMSG=varString(3)
					if vsERRCODE <> "1" AND demomode then vsRESPMSG = vsERRCODE & " - " & vsRESPMSG
					vsAUTHCODE=varString(4)
					vsAVSADDR=varString(5)
					vsTRANSID=varString(6)
					vsCVV2=varString(38)
					if Int(vsRESULT)=1 then
						vsRESULT="0" ' Keep in sync with Payflow Pro
						do_stock_management(ordID)
						cnn.Execute("UPDATE cart SET cartCompleted=1 WHERE cartOrderID="&ordID)
						cnn.Execute("UPDATE orders SET ordStatus=3,ordAVS='"&vsAVSADDR&"',ordCVV='"&vsCVV2&"',ordAuthNumber='"&vsAUTHCODE&"',ordTransID='"&vsTRANSID&"' WHERE ordID="&ordID)
					elseif Int(vsRESULT)=27 then
						gobackplaces=IIfVr(request.form("vrshippingoptions")="1", 3, 2)
					end if
				End If
				set objHttp = nothing
			end if
		else
			vsRESULT="0"
			vsRESPMSG="This transaction has been approved."
			if InStr(vsAUTHCODE,"-") > 0 then vsAUTHCODE = Right(vsAUTHCODE,Len(vsAUTHCODE)-InStr(vsAUTHCODE,"-"))
		end if
	elseif Request.Form("method")="18" then ' PayPal Pro
		on error resume next
		Server.ScriptTimeout = 120
		on error goto 0
		sSQL = "SELECT ordID,ordName,ordCity,ordState,ordCountry,ordPhone,ordHandling,ordZip,ordEmail,ordShipping,ordStateTax,ordCountryTax,ordTotal,ordDiscount,ordAddress,ordAddress2,ordIP,ordAuthNumber,ordShipName,ordShipAddress,ordShipAddress2,ordShipCity,ordShipState,ordShipCountry,ordShipZip FROM orders WHERE ordID=" & ordID
		rs.Open sSQL,cnn,0,1
		ordState = rs("ordState")
		ordShipState = rs("ordShipState")
		sSQL = "SELECT countryCode FROM countries WHERE countryName='" & replace(rs("ordCountry"),"'","''") & "'"
		rs2.Open sSQL,cnn,0,1
			countryCode = rs2("countryCode")
		rs2.Close
		sSQL = "SELECT countryCode FROM countries WHERE countryName='" & replace(rs("ordShipCountry"),"'","''") & "'"
		rs2.Open sSQL,cnn,0,1
			if NOT rs2.EOF then shipCountryCode = rs2("countryCode")
		rs2.Close
		if countryCode = "US" OR countryCode = "CA" then
			sSQL = "SELECT stateAbbrev FROM states WHERE stateName='" & replace(ordState,"'","''") & "'"
			rs2.Open sSQL,cnn,0,1
				if NOT rs2.EOF then ordState=rs2("stateAbbrev")
			rs2.Close
		end if
		if shipCountryCode="US" OR shipCountryCode="CA" then
			sSQL = "SELECT stateAbbrev FROM states WHERE stateName='" & replace(ordShipState,"'","''") & "'"
			rs2.Open sSQL,cnn,0,1
				if NOT rs2.EOF then ordShipState=rs2("stateAbbrev")
			rs2.Close
		end if
		vsAUTHCODE = trim(rs("ordAuthNumber")&"")
		thename = trim(request.form("cardname"))
		if thename<>"" then
			if InStr(thename," ") > 0 then
				namearr = Split(thename," ",2)
				firstname = namearr(0)
				lastname = namearr(1)
			else
				firstname = ""
				lastname = thename
			end if
		end if
		cardnum = replace(trim(request.form("ACCT")), " ", "")
		cartype = "Visa"
		if left(cardnum, 1)="5" then
			cartype="MasterCard"
		elseif left(cardnum, 1)="6" then
			cartype="Discover"
		elseif left(cardnum, 1)="3" then
			cartype="Amex"
		end if
		data2hash = data3
		sXML = ppsoapheader(data1, data2, data2hash) & _
			"  <soap:Body><DoDirectPaymentReq xmlns=""urn:ebay:api:PayPalAPI"">" & _
			"    <DoDirectPaymentRequest><Version xmlns=""urn:ebay:apis:eBLBaseComponents"">1.00</Version>" & _
			"      <DoDirectPaymentRequestDetails xmlns=""urn:ebay:apis:eBLBaseComponents"">" & _
			"        <PaymentAction>" & IIfVr(ppmethod=1, "Authorization", "Sale") & "</PaymentAction>" & _
			"        <PaymentDetails>" & _
			"          <OrderTotal currencyID=""" & countryCurrency & """>" & FormatNumber((rs("ordShipping")+rs("ordStateTax")+rs("ordCountryTax")+rs("ordTotal")+rs("ordHandling"))-rs("ordDiscount"),2,-1,0,0) & "</OrderTotal>" & _
			"          <ButtonSource>ecommercetemplates.asp.ecommplus</ButtonSource>"
		if trim(rs("ordShipAddress"))<>"" then
			sXML = sXML & "<ShipToAddress><Name>" & vrxmlencode(rs("ordShipName")) & "</Name><Street1>" & vrxmlencode(rs("ordShipAddress")) & "</Street1><Street2>" & vrxmlencode(rs("ordShipAddress2")) & "</Street2><CityName>" & rs("ordShipCity") & "</CityName><StateOrProvince>" & ordShipState & "</StateOrProvince><Country>" & shipCountryCode & "</Country><PostalCode>" & rs("ordShipZip") & "</PostalCode></ShipToAddress>"
		else
			sXML = sXML & "<ShipToAddress><Name>" & vrxmlencode(rs("ordName")) & "</Name><Street1>" & vrxmlencode(rs("ordAddress")) & "</Street1><Street2>" & vrxmlencode(rs("ordAddress2")) & "</Street2><CityName>" & rs("ordCity") & "</CityName><StateOrProvince>" & ordState & "</StateOrProvince><Country>" & countryCode & "</Country><PostalCode>" & rs("ordZip") & "</PostalCode></ShipToAddress>>"
		end if
		sXML = sXML & "</PaymentDetails>" & _
			"        <CreditCard>" & _
			"          <CreditCardType>" & cartype & "</CreditCardType><CreditCardNumber>" & vrxmlencode(cardnum) & "</CreditCardNumber>" & _
			"          <ExpMonth>" & request.form("EXMON") & "</ExpMonth><ExpYear>" & request.form("EXYEAR") & "</ExpYear>" & _
			"          <CardOwner>" & _
			"            <Payer>" & vrxmlencode(rs("ordEmail")) & "</Payer>" & _
			"            <PayerName><FirstName>" & firstname & "</FirstName><LastName>" & lastname & "</LastName></PayerName>" & _
			"            <PayerCountry>" & countryCode & "</PayerCountry>" & _
			"            <Address><Street1>" & vrxmlencode(rs("ordAddress")) & "</Street1><Street2>" & vrxmlencode(rs("ordAddress2")) & "</Street2><CityName>" & rs("ordCity") & "</CityName><StateOrProvince>" & ordState & "</StateOrProvince><Country>" & countryCode & "</Country><PostalCode>" & rs("ordZip") & "</PostalCode></Address>" & _
			"          </CardOwner>" & _
			"          <CVV2>" & trim(request.form("CVV2")) & "</CVV2>" & _
			"        </CreditCard>" & _
			"        <IPAddress>" & trim(rs("ordIP")) & "</IPAddress><MerchantSessionId>" & rs("ordID") & "</MerchantSessionId>" & _
			"      </DoDirectPaymentRequestDetails>" & _
			"    </DoDirectPaymentRequest></DoDirectPaymentReq></soap:Body></soap:Envelope>"
		rs.Close
		if demomode then sandbox = ".sandbox" else sandbox = ""
		vsRESULT="-1"
		if vsAUTHCODE="" then
			if blockuser then
				success=FALSE
				vsRESPMSG = multipurchaseblockmessage
			else
				success = callxmlfunction("https://api-aa" & IIfVr(sandbox="" AND data2hash<>"", "-3t", "") & sandbox & ".paypal.com/2.0/", sXML, res, IIfVr(data2hash<>"","",data1), "WinHTTP.WinHTTPRequest.5.1", vsRESPMSG, TRUE)
			end if
			if success then
				vsAUTHCODE="":vsERRCODE="":vsRESPMSG="":vsAVSADDR="":vsTRANSID="":vsCVV2=""
				set xmlDoc = Server.CreateObject("MSXML2.DOMDocument")
				xmlDoc.validateOnParse = False
				xmlDoc.loadXML (res)
				Set nodeList = xmlDoc.getElementsByTagName("SOAP-ENV:Body")
				Set n = nodeList.Item(0)
				for j = 0 to n.childNodes.length - 1
					Set e = n.childNodes.Item(i)
					if e.nodeName = "DoDirectPaymentResponse" then
						for k9 = 0 To e.childNodes.length - 1
							Set t = e.childNodes.Item(k9)
							if t.nodeName = "Ack" then
								if t.firstChild.nodeValue = "Success" OR t.firstChild.nodeValue = "SuccessWithWarning" then
									vsRESULT = 1
									vsRESPMSG = "Success"
								end if
							elseif t.nodeName = "TransactionID" then
								vsAUTHCODE = t.firstChild.nodeValue
							elseif t.nodeName = "AVSCode" then
								if t.hasChildNodes then vsAVSADDR = t.firstChild.nodeValue
							elseif t.nodeName = "CVV2Code" then
								if t.hasChildNodes then vsCVV2 = t.firstChild.nodeValue
							elseif t.nodeName = "Errors" then
								themsg=""
								thecode=""
								iswarning=FALSE
								set ff = t.childNodes
								for kk = 0 to ff.length - 1
									set gg = ff.item(kk)
									if gg.nodeName = "ShortMessage" then
										' vsRESPMSG = gg.firstChild.nodeValue & "<br>" & errormsg
									elseif gg.nodeName = "LongMessage" then
										themsg = gg.firstChild.nodeValue
									elseif gg.nodeName = "ErrorCode" then
										thecode = gg.firstChild.nodeValue
									elseif gg.nodeName = "SeverityCode" then
										if gg.hasChildNodes then iswarning = (gg.firstChild.nodeValue="Warning")
									end if
								next
								if NOT iswarning then
									vsRESPMSG = themsg & "<br />" & vsRESPMSG
									vsERRCODE = thecode
								end if
							end if
						next
					end if
				next
				if int(vsRESULT)=1 then
					vsRESULT="0" ' Keep in sync with Payflow Pro
					do_stock_management(ordID)
					cnn.Execute("UPDATE cart SET cartCompleted=1 WHERE cartOrderID="&ordID)
					cnn.Execute("UPDATE orders SET ordStatus=3,ordAVS='"&vsAVSADDR&"',ordCVV='"&vsCVV2&"',ordAuthNumber='"&vsAUTHCODE&"',ordTransID='"&vsTRANSID&"' WHERE ordID="&ordID)
				elseif vsERRCODE<>"" then
					vsERRCODE = int(vsERRCODE)
					if vsERRCODE=10505 OR (vsERRCODE>=10701 AND vsERRCODE<=10751) then
						gobackplaces=IIfVr(request.form("vrshippingoptions")="1", 3, 2)
					end if
				end if
			end if
		else
			vsRESULT="0"
			vsRESPMSG="This transaction has been approved."
			if InStr(vsAUTHCODE,"-") > 0 then vsAUTHCODE = Right(vsAUTHCODE,Len(vsAUTHCODE)-InStr(vsAUTHCODE,"-"))
		end if
	end if
%>	<br />
	<form method="post" action="thanks.asp" name="checkoutform">
	<input type="hidden" name="xxpreauth" value="<%=ordID%>" />
	<input type="hidden" name="xxpreauthmethod" value="<%=request.form("method")%>" />
	<input type="hidden" name="thesessionid" value="<%=thesessionid%>" />
            <table class="cobtbl" width="<%=maintablewidth%>" border="0" bordercolor="#B1B1B1" cellspacing="1" cellpadding="3" bgcolor="#B1B1B1">
<%		if vsRESULT="0" then %>
			  <tr height="30">
			    <td class="cobhl" bgcolor="#EBEBEB" align="center" colspan="2"><strong><%=xxTnxOrd%></strong></td>
			  </tr>
			  <tr height="30">
			    <td class="cobhl" bgcolor="#EBEBEB" align="right" width="50%"><strong><%=xxTrnRes%>:</strong></td>
				<td class="cobll" bgcolor="#FFFFFF" width="50%"><strong><%=vsRESPMSG%></strong></td>
			  </tr>
			  <tr height="30">
			    <td class="cobhl" bgcolor="#EBEBEB" align="right" width="50%"><strong><%=xxOrdNum%>:</strong></td>
				<td class="cobll" bgcolor="#FFFFFF" width="50%"><strong><%=ordID%></strong></td>
			  </tr>
			  <tr height="30">
			    <td class="cobhl" bgcolor="#EBEBEB" align="right" width="50%"><strong><%=xxAutCod%>:</strong></td>
				<td class="cobll" bgcolor="#FFFFFF" width="50%"><strong><%=vsAUTHCODE%></strong></td>
			  </tr>
			  <tr height="30">
				<td class="cobll" bgcolor="#FFFFFF" colspan="2">
				  <table width="100%" cellspacing="0" cellpadding="0" border="0">
				    <tr>
					  <td width="16" height="26" align="right" valign="bottom">&nbsp;</td>
					  <td class="cobll" bgcolor="#FFFFFF" width="100%" align="center">&nbsp;<br />
					  <input type="submit" value="Click to Confirm Order and View Receipt" /><br />&nbsp;
					  </td>
					  <td width="16" height="26" align="right" valign="bottom"><img src="images/tablebr.gif" alt="" /></td>
					</tr>
				  </table>
				</td>
			  </tr>
<%			if forcesubmit=true then
				if forcesubmittimeout="" then forcesubmittimeout=5000
				response.write "<script language=""javascript"" type=""text/javascript"">setTimeout('document.checkoutform.submit()',"&forcesubmittimeout&");</script>" & vbCrLf
			end if
		else %>
			  <tr height="30">
			    <td class="cobhl" bgcolor="#EBEBEB" align="center" colspan="2"><strong><%=xxSorTrn%></strong></td>
			  </tr>
			  <tr height="30">
			    <td class="cobhl" bgcolor="#EBEBEB" align="right" width="50%"><strong><%=xxTrnRes%>:</strong></td>
				<td class="cobll" bgcolor="#FFFFFF" width="50%"><strong><%=IIfVr(vsERRCODE<>"", "(" & vsERRCODE & ") ", "") & vsRESPMSG%></strong></td>
			  </tr>
			  <tr height="30">
				<td class="cobll" bgcolor="#FFFFFF" colspan="2">
				  <table width="100%" cellspacing="0" cellpadding="0" border="0">
				    <tr>
					  <td width="16" height="26" align="right" valign="bottom">&nbsp;</td>
					  <td class="cobll" bgcolor="#FFFFFF" width="100%" align="center">&nbsp;<br />
					  <input type="button" value="<%=xxGoBack%>" onclick="javascript:history.go(-<%=gobackplaces%>)" /><br />&nbsp;
					  </td>
					  <td width="16" height="26" align="right" valign="bottom"><img src="images/tablebr.gif" alt="" /></td>
					</tr>
				  </table>
				</td>
			  </tr>
<%		end if %>
			</table>
	</form>
<%
elseif request.querystring("token") = "" AND checkoutmode <> "paypalexpress1" AND cartisincluded<>TRUE then
	addextrarows=0
	wantstateselector=FALSE
	wantcountryselector=FALSE
	wantzipselector=FALSE
	if estimateshipping=TRUE then
		addextrarows=1
		if shipType=2 OR shipType=5 then ' weight / price based
			wantcountryselector=TRUE
			if splitUSZones then
				addextrarows=3
				wantstateselector=TRUE
			else
				addextrarows=2
			end if
		elseif shipType=3 OR shipType=4 OR shipType=6 OR shipType=7 then
			addextrarows=3
			wantzipselector=TRUE
			wantcountryselector=TRUE
		end if
		shiphomecountry=TRUE
		if cartisincluded<>TRUE then
			if request.form("state")<>"" then
				shipstate = request.form("state")
				session("state") = request.form("state")
			elseif session("state")<>"" then
				shipstate = session("state")
			else
				shipstate = defaultshipstate
			end if
			if request.form("zip")<>"" then
				destZip = trim(request.form("zip"))
				session("zip") = trim(request.form("zip"))
			elseif session("zip")<>"" then
				destZip = session("zip")
			else
				if NOT (nodefaultzip=TRUE) then destZip = origZip
			end if
			if request.form("country")<>"" then
				shipcountry = request.form("country")
				session("country") = request.form("country")
				if trim(request.form("state"))="" then shipstate=""
			elseif session("country")<>"" then
				shipcountry = session("country")
			else
				shipCountryCode = origCountryCode
				shipcountry = origCountry
			end if
		end if
		sSQL = "SELECT countryID,countryTax,countryCode,countryFreeShip,countryOrder FROM countries WHERE countryName='"&replace(shipcountry,"'","''")&"'"
		rs.Open sSQL,cnn,0,1
		if NOT rs.EOF then
			if trim(Session("clientUser")) <> "" AND (Session("clientActions") AND 2)=2 then countryTaxRate=0 else countryTaxRate = rs("countryTax")
			shipCountryID = rs("countryID")
			shipCountryCode = rs("countryCode")
			freeshipapplies = (rs("countryFreeShip")=1)
			shiphomecountry = (rs("countryOrder")=2)
		end if
		rs.Close
		if session("xsshipping")="" then initshippingmethods()
	end if
	if showtaxinclusive then addextrarows=addextrarows+1
	alldata=""
	if mysqlserver=true then
		sSQL = "SELECT cartID,cartProdID,cartProdName,cartProdPrice,cartQuantity,pWeight,pShipping,pShipping2,pExemptions,pSection,topSection,pDims,pTax,"&getlangid("pDescription",2)&" FROM cart INNER JOIN products ON cart.cartProdID=products.pID LEFT OUTER JOIN sections ON products.pSection=sections.sectionID WHERE cartCompleted=0 AND cartSessionID="&Session.SessionID
	else
		sSQL = "SELECT cartID,cartProdID,cartProdName,cartProdPrice,cartQuantity,pWeight,pShipping,pShipping2,pExemptions,pSection,topSection,pDims,pTax,"&getlangid("pDescription",2)&" FROM cart INNER JOIN (products LEFT OUTER JOIN sections ON products.pSection=sections.sectionID) ON cart.cartProdID=products.pID WHERE cartCompleted=0 AND cartSessionID="&Session.SessionID
	end if
	rs.Open sSQL,cnn,0,1
	if NOT (rs.EOF OR rs.BOF) then alldata=rs.getrows
	rs.Close
%>	<br />
	<form method="post" action="cart.asp" name="checkoutform">
	<input type="hidden" name="mode" value="update" />
            <table class="cobtbl" width="<%=maintablewidth%>" border="0" bordercolor="#B1B1B1" cellspacing="1" cellpadding="3" bgcolor="#B1B1B1">
<%	if IsArray(alldata) then
		if NOT isInStock then %>
			  <tr height="30">
			    <td class="cobll" bgcolor="#FFFFFF" colspan="6" align="center"><font color="#FF0000"><strong><%=xxNoStok%></strong></font></td>
			  </tr>
<%		end if %>
			  <tr height="30">
			    <td class="cobhl" bgcolor="#EBEBEB" width="15%"><strong><%=xxCODets%></strong></td>
			    <td class="cobhl" bgcolor="#EBEBEB" width="33%"><strong><%=xxCOName%></strong></td>
				<td class="cobhl" bgcolor="#EBEBEB" width="14%" align="center"><strong><%=xxCOUPri%></strong></td>
				<td class="cobhl" bgcolor="#EBEBEB" width="14%" align="center"><strong><%=xxQuant%></strong></td>
				<td class="cobhl" bgcolor="#EBEBEB" width="14%" align="center"><strong><%=xxTotal%></strong></td>
				<td class="cobhl" bgcolor="#EBEBEB" width="10%" align="center"><strong><%=xxCOSel%></strong></td>
			  </tr>
<%		totaldiscounts = 0
		changechecker = ""
		googlelineitems = ""
		For index=0 to UBOUND(alldata,2)
			changechecker = changechecker & "if(document.checkoutform.quant" & alldata(0,index) & ".value!=" & alldata(4,index) & ") dowarning=true;" & vbCrLf
			theoptions = ""
			theoptionspricediff = 0
			sSQL = "SELECT coOptGroup,coCartOption,coPriceDiff,coWeightDiff FROM cartoptions WHERE coCartID="&alldata(0,index) & " ORDER BY coID"
			rs.Open sSQL,cnn,0,1
			do while NOT rs.EOF
				theoptionspricediff = theoptionspricediff + rs("coPriceDiff")
				alldata(5,index)=cDbl(alldata(5,index))+cDbl(rs("coWeightDiff"))
				theoptions = theoptions & "<tr height=""25"">"
				theoptions = theoptions & "<td class=""cobhl"" bgcolor=""#EBEBEB"" align=""right""><font style=""font-size: 10px""><strong>"&rs("coOptGroup")&":</strong></font></td>"
				theoptions = theoptions & "<td class=""cobll"" bgcolor=""#FFFFFF""><font style=""font-size: 10px"">" & "&nbsp;- " & replace(rs("coCartOption")&"", vbCrLf, "<br>") & "</font></td>"
				theoptions = theoptions & "<td class=""cobll"" bgcolor=""#FFFFFF"" align=""right""><font style=""font-size: 10px"">" & IIfVr(rs("coPriceDiff")=0 OR hideoptpricediffs=true,"- ", FormatEuroCurrency(rs("coPriceDiff"))) & "</font></td>"
				theoptions = theoptions & "<td class=""cobll"" bgcolor=""#FFFFFF"" align=""right"">&nbsp;</td>"
				theoptions = theoptions & "<td class=""cobll"" bgcolor=""#FFFFFF"" align=""right""><font style=""font-size: 10px"">" & IIfVr(rs("coPriceDiff")=0 OR hideoptpricediffs=true,"- ", FormatEuroCurrency(rs("coPriceDiff")*alldata(4,index))) & "</font></td>"
				theoptions = theoptions & "<td class=""cobll"" bgcolor=""#FFFFFF"" align=""center"">&nbsp;</td>"
				theoptions = theoptions & "</tr>" & vbCrLf
				totalgoods = totalgoods + (rs("coPriceDiff")*alldata(4,index))
				rs.MoveNext
			loop
			Session.LCID = 1033
			googlelineitems = googlelineitems & "<item><merchant-private-item-data><product-id>"&xmlencodecharref(alldata(1,index))&"</product-id></merchant-private-item-data><item-name>"&xmlencodecharref(strip_tags2(alldata(2,index)&""))&"</item-name><item-description>"&xmlencodecharref(left(strip_tags2(alldata(13,index)&""),301))&"</item-description><unit-price currency="""&countryCurrency&""">"&FormatNumber(alldata(3,index) + theoptionspricediff,2,-1,0,0)&"</unit-price><quantity>"&alldata(4,index)&"</quantity></item>"
			Session.LCID = saveLCID
			rs.Close %>
              <tr height="30">
			    <td class="cobhl" bgcolor="#EBEBEB"><strong><%=alldata(1,index)%></strong></td>
			    <td class="cobll" bgcolor="#FFFFFF"><% Response.write alldata(2,index) %></td>
				<td class="cobll" bgcolor="#FFFFFF" align="right"><%=IIfVr(hideoptpricediffs=true,FormatEuroCurrency(alldata(3,index)+theoptionspricediff),FormatEuroCurrency(alldata(3,index)))%></td>
				<td class="cobll" bgcolor="#FFFFFF" align="center"><input type="text" name="quant<%=alldata(0,index)%>" value="<%=alldata(4,index)%>" size="2" maxlength="5" /></td>
				<td class="cobll" bgcolor="#FFFFFF" align="right"><%=IIfVr(hideoptpricediffs=true,FormatEuroCurrency((alldata(3,index)+theoptionspricediff)*alldata(4,index)),FormatEuroCurrency(alldata(3,index)*alldata(4,index)))%></td>
				<td class="cobll" bgcolor="#FFFFFF" align="center"><input type="checkbox" name="delet<%=alldata(0,index)%>" /></td>
			  </tr>
<%			response.write theoptions
			runTot=(alldata(3,index)*Int(alldata(4,index)))
			totalquantity = totalquantity + Int(alldata(4,index))
			totalgoods = totalgoods + runTot
			alldata(3,index) = alldata(3,index) + theoptionspricediff
			if trim(Session("clientUser"))<>"" then alldata(8,index) = (alldata(8,index) OR Session("clientActions"))
			if (shipType=2 OR shipType=3 OR shipType=4 OR shipType=6 OR shipType=7) AND cDbl(alldata(5,index))<=0.0 then alldata(8,index) = (alldata(8,index) OR 4)
			if perproducttaxrate=TRUE then
				if isnull(alldata(12,index)) then alldata(12,index)=countryTaxRate
				if (alldata(8,index) AND 2)<>2 then countryTax = countryTax + ((alldata(12,index) * runTot) / 100.0)
			else
				if (alldata(8,index) AND 2)=2 then countrytaxfree = countrytaxfree + runTot + (theoptionspricediff*Int(alldata(4,index)))
			end if
			if (alldata(8,index) AND 4)=4 then shipfreegoods = shipfreegoods + runTot else somethingToShip=TRUE
			if estimateshipping=TRUE AND session("xsshipping")="" then call addproducttoshipping(alldata, index)
		Next
		call calculatediscounts(totalgoods, false, "")
		if totaldiscounts > totalgoods then totaldiscounts = totalgoods
		if totaldiscounts=0 then
			session("discounts")=""
		else
			session("discounts")=totaldiscounts
			addextrarows = addextrarows + 1
			glicpnmessage = Right(cpnmessage,Len(cpnmessage)-6)
			glicpnmessage = Left(glicpnmessage,Len(glicpnmessage)-6)
			googlelineitems = googlelineitems & "<item><merchant-private-item-data><discountflag>true</discountflag></merchant-private-item-data><item-name>"&xmlencodecharref(strip_tags2(xxAppDs))&"</item-name><item-description>"&xmlencodecharref(strip_tags2(replace(glicpnmessage,"<br />"," - ")))&"</item-description><unit-price currency="""&countryCurrency&""">-"&FormatNumber(totaldiscounts,2,-1,0,0)&"</unit-price><quantity>1</quantity></item>"
		end if
		if addextrarows > 0 then %>
              <tr height="30">
				<td class="cobhl" bgcolor="#EBEBEB" rowspan="<%=addextrarows+3%>">&nbsp;</td>
				<td class="cobll" bgcolor="#FFFFFF" align="right" colspan="3"><strong><%=xxSubTot%>:</strong></td>
				<td class="cobll" bgcolor="#FFFFFF" align="right"><%=FormatEuroCurrency(totalgoods)%></td>
				<td class="cobll" bgcolor="#FFFFFF" align="center"><a href="javascript:document.checkoutform.submit()"><strong><%=xxDelete%></strong></a></td>
			  </tr>
<%		end if
		if totaldiscounts>0 then %>
			  <tr height="30">
				<td class="cobll" bgcolor="#FFFFFF" align="right" colspan="3"><font color="#FF0000"><strong><%=xxDsApp%></strong></font></td>
				<td class="cobll" bgcolor="#FFFFFF" align="right"><font color="#FF0000"><%=FormatEuroCurrency(totaldiscounts)%></font></td>
				<td class="cobll" bgcolor="#FFFFFF" align="center">&nbsp;</td>
			  </tr>
<%		end if
		if estimateshipping=TRUE then
			if session("xsshipping")="" then
				if calculateshipping() then
					if IsNumeric(shipinsuranceamt) AND abs(addshippinginsurance)=1 then shipping = shipping + IIfVr(addshippinginsurance=1,((cDbl(totalgoods)*cDbl(shipinsuranceamt))/100.0),shipinsuranceamt)
					if taxShipping=1 AND showtaxinclusive then shipping = shipping + (cDbl(shipping)*(cDbl(countryTaxRate)))/100.0
					calculateshippingdiscounts(false)
					session("xsshipping")=shipping-freeshipamnt
				end if
			else
				shipping = session("xsshipping")
			end if
			if errormsg<>"" then %>
              <tr height="30">
				<td class="cobll" bgcolor="#FFFFFF" align="right" colspan="3"><strong><%=xxShpEst%>:</strong></td>
				<td class="cobll" bgcolor="#FFFFFF" colspan="2"><font style="font-size: 10px" color="#FF0000"><strong><%=errormsg%></strong></font></td>
			  </tr>
<%			else %>
              <tr height="30">
				<td class="cobll" bgcolor="#FFFFFF" align="right" colspan="3"><strong><%=xxShpEst%>:</strong></td>
				<td class="cobll" bgcolor="#FFFFFF" align="right"><% if freeshipamnt=shipping then response.write "<p align=""center""><font color=""#FF0000""><strong>" & xxFree & "</strong></font></p>" else response.write FormatEuroCurrency(shipping-freeshipamnt)%></td>
				<td class="cobll" bgcolor="#FFFFFF" align="center">&nbsp;</td>
			  </tr>
<%			end if
			if wantstateselector then
				sSQL = "SELECT stateName,stateAbbrev FROM states WHERE stateEnabled=1 ORDER BY stateName"
				rs.Open sSQL,cnn,0,1
				if NOT rs.EOF then allstates=rs.getrows
				rs.Close
				if IsArray(allstates) then %>
              <tr height="30">
				<td class="cobll" bgcolor="#FFFFFF" align="right" colspan="3"><strong><%=xxAllSta%>:</strong></td>
				<td class="cobll" bgcolor="#FFFFFF" colspan="2"><select name="state" size="1"><% show_states(shipstate) %></select></td>
			  </tr>
<%				end if
			end if
			if wantcountryselector then %>
              <tr height="30">
				<td class="cobll" bgcolor="#FFFFFF" align="right" colspan="3"><strong><%=xxCountry%>:</strong></td>
				<td class="cobll" bgcolor="#FFFFFF" colspan="2"><select name="country" size="1"><%
				sSQL = "SELECT countryName,countryCode,"&getlangid("countryName",8)&" AS cnameshow FROM countries WHERE countryEnabled=1 ORDER BY countryOrder DESC,"&getlangid("countryName",8)
				rs.Open sSQL,cnn,0,1
				do while NOT rs.EOF
					response.write "<option value="""&rs("countryName")&""""
					if shipcountry=rs("countryName") then response.write " selected"
					response.write ">"&rs("cnameshow")&"</option>"&vbCrLf
					rs.MoveNext
				loop
				rs.Close %></select></td>
			  </tr>
<%			end if
			if wantzipselector then %>
              <tr height="30">
				<td class="cobll" bgcolor="#FFFFFF" align="right" colspan="3"><strong><%=xxZip%>:</strong></td>
				<td class="cobll" bgcolor="#FFFFFF" colspan="2"><input type="text" name="zip" size="8" value="<%=replace(destZip,"""","&quot;")%>"></td>
			  </tr>
<%			end if
		end if
		if showtaxinclusive then
			countryTax = vsround((((totalgoods-countrytaxfree)+IIfVr(taxShipping=2,shipping-freeshipamnt,0))-totaldiscounts)*countryTaxRate/100.0, 2)
			session("xscountrytax")=countryTax
%>			  <tr height="30">
				<td class="cobll" bgcolor="#FFFFFF" align="right" colspan="3"><strong><%=xxCntTax%>:</strong></td>
				<td class="cobll" bgcolor="#FFFFFF" align="right"><%=FormatEuroCurrency(countryTax)%></td>
				<td class="cobll" bgcolor="#FFFFFF" align="center">&nbsp;</td>
			  </tr>
<%		else
			countryTax=0
		end if %>
              <tr height="30">
				<% if addextrarows=0 then %>
				<td class="cobhl" bgcolor="#EBEBEB" rowspan="2">&nbsp;</td>
				<% end if %>
				<td class="cobll" bgcolor="#FFFFFF" align="right" colspan="3"><strong><%=xxGndTot%>:</strong></td>
				<td class="cobll" bgcolor="#FFFFFF" align="right"><%=FormatEuroCurrency((totalgoods+shipping+countryTax)-(totaldiscounts+freeshipamnt))%></td>
				<td class="cobll" bgcolor="#FFFFFF" align="center"><% if addextrarows=0 then response.write "<a href=""javascript:document.checkoutform.submit()""><strong>" & xxDelete & "</strong></a>" else response.write "&nbsp;" end if %></td>
			  </tr>
			  <tr height="30">
				<td class="cobll" bgcolor="#FFFFFF" colspan="5">
				  <table width="100%" cellspacing="0" cellpadding="0" border="0">
				    <tr>
					  <td class="cobll" bgcolor="#FFFFFF" width="50%" align="center"><a href="<% if Trim(Session("frompage"))<>"" AND (actionaftercart=2 OR actionaftercart=3) then response.write Session("frompage") else response.write xxHomeURL%>"><strong><%=xxCntShp%></strong></a></td>
					  <td class="cobll" bgcolor="#FFFFFF" width="50%" align="center"><a href="javascript:document.checkoutform.submit()"><strong><%=xxUpdTot%></strong></a></td>
					  <td class="cobll" bgcolor="#FFFFFF" width="16" height="26" align="right" valign="bottom"><img src="images/tablebr.gif" alt="" /></td>
					</tr>
				  </table>
				</td>
			  </tr>
<script language="javascript" type="text/javascript">
<!--
function changechecker(){
dowarning=false;
<%=changechecker%>
if(dowarning){
	if(confirm('<%=replace(xxWrnChQ,"'","\'")%>')){
		document.checkoutform.submit();
		return false;
	}else
		return(true);
}
return true;
}
//--></script>
<%	else
		cartEmpty=True %>
              <tr>
			    <td class="cobll" bgcolor="#FFFFFF" colspan="6" align="center">
				  <p>&nbsp;</p>
				  <p><%=xxSryEmp%></p>
				  <p>&nbsp;</p>
<script language="javascript" type="text/javascript">
<!--
if(document.cookie=="") document.write("<%=Replace(xxNoCk & " " & xxSecWar, """", "\""")%>");
//--></script>
<noscript><%=xxNoJS & " " & xxSecWar%></noscript>
				  <p><a href="<% if Trim(Session("frompage"))<>"" AND (actionaftercart=2 OR actionaftercart=3) then response.write Session("frompage") else response.write xxHomeURL%>"><strong><%=xxCntShp%></strong></a></p>
				  <p>&nbsp;</p>
				</td>
			  </tr>
<%	end if
%>			</table>
	</form>
<%
end if
if request.querystring("token") = "" AND checkoutmode <> "paypalexpress1" AND checkoutmode<>"go" AND checkoutmode<>"checkout" AND checkoutmode<>"add" AND checkoutmode<>"authorize" AND NOT cartEmpty AND cartisincluded<>TRUE then
	sub generatemerchantcalcshiptypes(theshiptype)
		if googledefaultshipping="" then googledefaultshipping="999.99"
		if NOT somethingToShip then
		elseif theshiptype=1 then
			sXML = sXML & "<merchant-calculated-shipping name="""&xmlencodecharref(xxShipHa)&"""><price currency="""&countryCurrency&""">"&googledefaultshipping&"</price></merchant-calculated-shipping>"
		elseif theshiptype=2 OR theshiptype=5 then
			Dim gshipmethods()
			redim gshipmethods(10)
			numshipmethods=0
			for index3=1 to 5
				sSQL = "SELECT DISTINCT pzMethodName"&index3&" FROM postalzones WHERE pzName<>'' AND pzMethodName"&index3&"<>''"
				rs.Open sSQL,cnn,0,1
				do while NOT rs.EOF
					gotshipmethod=false
					for index4=0 to numshipmethods
						if gshipmethods(index4)=rs("pzMethodName"&index3) then gotshipmethod=true
					next
					if NOT gotshipmethod then gshipmethods(numshipmethods)=rs("pzMethodName"&index3) : numshipmethods = numshipmethods + 1
					rs.MoveNext
				loop
				rs.Close
			next
			for index3=0 to numshipmethods-1
				sXML = sXML & "<merchant-calculated-shipping name="""&xmlencodecharref(gshipmethods(index3)&"")&"""><price currency="""&countryCurrency&""">"&googledefaultshipping&"</price></merchant-calculated-shipping>"
			next
		elseif theshiptype=3 OR theshiptype=4 OR theshiptype=6 OR theshiptype=7 then
			if theshiptype=3 then startid=0
			if theshiptype=4 then startid=1
			if theshiptype=6 then startid=2
			if theshiptype=7 then startid=3
			sSQL = "SELECT DISTINCT uspsShowAs,uspsFSA FROM uspsmethods WHERE (uspsID>"&(startid*100)&" AND uspsID<"&((startid+1)*100)&") AND uspsUseMethod=1 ORDER BY uspsFSA DESC,uspsShowAs"
			rs.Open sSQL,cnn,0,1
			do while NOT rs.EOF
				sXML = sXML & "<merchant-calculated-shipping name="""&xmlencodecharref(rs("uspsShowAs")&"")&"""><price currency="""&countryCurrency&""">"&googledefaultshipping&"</price></merchant-calculated-shipping>"
				rs.MoveNext
			loop
			rs.Close
		end if
	end sub
	function writegoogleparams(data1, data2, demomode)
		sSQL = "SELECT cpnID FROM coupons WHERE cpnIsCoupon=1 AND cpnNumAvail>0 AND cpnEndDate>="&datedelim&VSUSDate(Date())&datedelim
		rs.Open sSQL,cnn,0,1
		if rs.EOF then acoupondefined="false" else acoupondefined="true"
		rs.Close
		b64pad="="
		' response.write "data1:"&data1&", data2:"&data2&"<br>"
		sXML = "<?xml version=""1.0"" encoding=""UTF-8""?><checkout-shopping-cart xmlns=""http://checkout.google.com/schema/2""><shopping-cart>"
		sXML = sXML & "<items>" & googlelineitems & "</items>"
		sXML = sXML & "<merchant-private-data><privateitems><sessionid>"&thesessionid&"</sessionid><partner>"&IIfVr(trim(request.querystring("PARTNER"))<>"",trim(request.querystring("PARTNER")),Trim(request.cookies("PARTNER")))&"</partner><clientuser>"&Session("clientUser")&"</clientuser></privateitems></merchant-private-data>"
		sXML = sXML & "</shopping-cart>"
		sXML = sXML & "<checkout-flow-support><merchant-checkout-flow-support><platform-id>236638029623651</platform-id>"
		sXML = sXML & "<edit-cart-url>"&storeurl&"cart.asp</edit-cart-url>"
		sXML = sXML & "<continue-shopping-url>"&storeurl&"categories.asp</continue-shopping-url>"
		sXML = sXML & "<shipping-methods>"
		generatemerchantcalcshiptypes(shipType)
		if adminIntShipping<>0 AND adminIntShipping<>shipType then generatemerchantcalcshiptypes(adminIntShipping)
		if willpickuptext<>"" then
			if willpickupcost="" then willpickupcost=0
			' sXML = sXML & "<pickup name="""&willpickuptext&"""><price currency="""&countryCurrency&""">"&willpickupcost&"</price></pickup>"
			sXML = sXML & "<merchant-calculated-shipping name=""" & xmlencodecharref(willpickuptext) & """><price currency=""" & countryCurrency & """>" & willpickupcost & "</price></merchant-calculated-shipping>"
		end if
		sXML = sXML & "</shipping-methods>"
		sXML = sXML & "<request-buyer-phone-number>true</request-buyer-phone-number>"
		sXML = sXML & "<tax-tables merchant-calculated=""true""><default-tax-table><tax-rules></tax-rules></default-tax-table></tax-tables>"
		sXML = sXML & "<merchant-calculations><merchant-calculations-url>"&gcallbackpath&"</merchant-calculations-url><accept-merchant-coupons>"&acoupondefined&"</accept-merchant-coupons><accept-gift-certificates>false</accept-gift-certificates></merchant-calculations></merchant-checkout-flow-support></checkout-flow-support>"
		sXML = sXML & "</checkout-shopping-cart>"
		' response.write Replace(Replace(sxml,"</","&lt;/"),"<","<br />&lt;")&"<HR>"
		thecart = vrbase64_encrypt(sXML)
		thesignature = b64_hmac_sha1(data2,sXML)
		theurl = "https://"&IIfVr(demomode, "sandbox", "checkout")&".google.com/cws/v2/Merchant/"&data1&"/checkout"
			' theurl = theurl & IIfVr(demomode, "/diagnose", "")
		call writehiddenvar("cart", thecart)
		call writehiddenvar("signature", thesignature)
		' response.write "signature:" & thesignature & "<br>"
		writegoogleparams = theurl
	end function
	requiressl = false
	if pathtossl="" then
		sSQL = "SELECT payProvID FROM payprovider WHERE payProvEnabled=1 AND (payProvID IN (7,10,12,13,18) OR (payProvID=16 AND payProvData2='1'))" ' All the ones that require SSL
		rs.Open sSQL,cnn,0,1
		if NOT rs.EOF then requiressl = true
		rs.Close
	end if
	if requiressl OR pathtossl<>"" then
		if pathtossl<>"" then
			if Right(pathtossl,1) <> "/" then pathtossl = pathtossl & "/"
			cartpath = pathtossl & "cart.asp"
			gcallbackpath = pathtossl & "vsadmin/gcallback.asp"
		else
			cartpath = Replace(storeurl,"http:","https:") & "cart.asp"
			gcallbackpath = Replace(storeurl,"http:","https:") & "vsadmin/gcallback.asp"
		end if
	else
		cartpath="cart.asp"
		gcallbackpath = storeurl & "vsadmin/gcallback.asp"
	end if
%>
	  <table border="0" cellspacing="<%=maintablespacing%>" cellpadding="<%=maintablepadding%>" width="<%=maintablewidth%>" bgcolor="<%=maintablebg%>" align="center">
        <tr> 
          <td width="100%">
		    <form method="post" name="mainform" action="<%=cartpath%>" onsubmit="return changechecker(this)">
			  <input type="hidden" name="mode" value="checkout" />
			  <input type="hidden" name="sessionid" value="<%=Session.SessionID%>" />
			  <input type="hidden" name="PARTNER" value="<%=IIfVr(trim(request.querystring("PARTNER"))<>"",trim(request.querystring("PARTNER")),Trim(request.cookies("PARTNER")))%>" />
			  <input type="hidden" name="estimate" value="<%=FormatNumber((totalgoods+shipping)-(totaldiscounts+freeshipamnt),2,-1,0,0) %>" />
<%				if Trim(Session("clientUser"))<>"" then
					cnn.Execute("DELETE FROM tmplogin WHERE tmplogindate < " & datedelim & VSUSDate(Date()-3) & datedelim & " OR tmploginid=" & Session.SessionID)
					cnn.Execute("INSERT INTO tmplogin (tmploginid, tmploginname, tmplogindate) VALUES (" & Session.SessionID & ",'" & Trim(Session("clientUser")) & "'," & datedelim & VSUSDate(Date()) & datedelim & ")")
					response.write "<input type=""hidden"" name=""checktmplogin"" value=""1"" />"
					if (Session("clientActions") AND 8) = 8 OR (Session("clientActions") AND 16) = 16 then
						if minwholesaleamount<>"" then minpurchaseamount=minwholesaleamount
						if minwholesalemessage<>"" then minpurchasemessage=minwholesalemessage
					end if
				end if
%>
			  <table width="<%=innertablewidth%>" border="0" cellspacing="<%=innertablespacing%>" cellpadding="<%=innertablepadding%>" bgcolor="<%=innertablebg%>">
<%		if totalgoods<minpurchaseamount then %>
				<tr><td width="100%" align="center" colspan="2"><%=minpurchasemessage%></td></tr>
<%		else
			sSQL = "SELECT payProvID,payProvData1,payProvData2,payProvDemo FROM payprovider WHERE payProvEnabled=1 ORDER BY payProvOrder"
			rs.Open sSQL,cnn,0,1
			if NOT rs.EOF then checkoutmethods = rs.GetRows else checkoutmethods=""
			rs.Close
			if isarray(checkoutmethods) then
				regularcheckoutshown=false
				for index=0 to UBOUND(checkoutmethods, 2)
					if checkoutmethods(0, index)=19 then %>
				<tr><td align="center" colspan="2"><%=xxPPPBlu%></td></tr>
				<tr><td colspan="2" align="center"><input type="image" src="images/ppexpress.gif" border="0" onclick="javascript:document.forms.mainform.mode.value='paypalexpress1';" alt="PayPal Express" /></td></tr>
<%					elseif checkoutmethods(0, index)=20 then
						theurl = writegoogleparams(checkoutmethods(1, index), checkoutmethods(2, index), checkoutmethods(3, index))
						if xxGooCo<>"" then %><tr><td align="center" colspan="2"><strong><%=xxGooCo%></strong></td></tr><% end if %>
				<tr><td colspan="2" align="center"><input type="image" name="GBuy" alt="Google Checkout" src="http://checkout.google.com/buttons/checkout.gif?merchant_id=<%=checkoutmethods(1, index)&IIfVr(googlebuttonparams<>"", googlebuttonparams, "&w=160&h=43&style=white&variant=text&loc=en_US")%>" onclick="document.forms.mainform.onsubmit='';document.forms.mainform.action='<%=theurl%>';"></td></tr>
<%					elseif NOT regularcheckoutshown then
						regularcheckoutshown=TRUE %>
				<tr><td width="100%" align="center" colspan="2"><strong><%=xxPrsChk%></strong></td></tr>
				<tr><td align="center" colspan="2"><input type="image" src="images/checkout.gif" border="0" onclick="javascript:document.forms.mainform.mode.value='checkout';" alt="<%=xxCOTxt%>" /></td></tr>
<%					end if
				next
			end if
		end if %>
			  </table>
			</form>
		  </td>
        </tr>
      </table>
<%
end if
if cartisincluded<>TRUE then
	cnn.Close
	set rs = nothing
	set rs2 = nothing
	set cnn = nothing
end if
%>