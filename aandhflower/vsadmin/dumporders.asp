<%
Response.Buffer = True
%>
<!--#include file="db_conn_open.asp"-->
<!--#include file="includes.asp"-->
<!--#include file="inc/incfunctions.asp"-->
<%
Dim sd, ed, rs, cnn, sSQL, sSQL2, hasdetails, sslok
function twodp(theval)
	twodp=FormatNumber(theval,2,-1,0,0)
end function
function xmlstrip(name2)
	name2=replace(name2&"","&","chr(11)")
	name2=replace(name2,chr(146),"chr(146)")
	name2=replace(name2,chr(150),"chr(150)")
	name2=replace(name2,"-","chr(45)")
	name2=replace(name2,"'","chr(39)chr(39)")
	name2=replace(name2,"€","chr(128)")
	name2=replace(name2,chr(163),"chr(163)")
	name2=replace(name2,chr(130),"chr(130)")
	name2=replace(name2,chr(138),"chr(138)")
	name2=replace(name2,chr(153),"")
	name2=replace(name2,chr(250),"u")
	name2=replace(name2,chr(225),"a")
	name2=replace(name2,chr(241),"n")
	name2=replace(name2,chr(252),"chr(129)")
	name2=replace(name2,chr(246),"chr(148)")	
	name2=replace(name2,chr(174),"")
	name2=replace(name2,"""","")
	name2=replace(name2,chr(147),"")
	name2=replace(name2,chr(148),"")
	name2=replace(name2,chr(169),"")
	name2=replace(name2,"å","a")
	tmp_str=""
	for i=1 to len(name2)
		ch_code=Asc(Mid(name2,i,1))
		if ch_code>130 then tmp_str=tmp_str & "chr("&ch_code&")" else tmp_str=tmp_str & Mid(name2,i,1)
	next
	xmlstrip=tmp_str
end function
function getsearchparams()
	tmpsql = ""
	if request.form("powersearch")="1" then
		fromdate = Trim(request.form("fromdate"))
		todate = Trim(request.form("todate"))
		ordid = Trim(Replace(Replace(request.form("ordid"),"'",""),"""",""))
		origsearchtext = Trim(Replace(request.form("searchtext"),"""","&quot;"))
		searchtext = Trim(Replace(request.form("searchtext"),"'","''"))
		ordstatus = Trim(request.form("ordstatus"))
		tmpsql = tmpsql & " WHERE 1=1"
		if ordid<>"" then
			if IsNumeric(ordid) then
				tmpsql = tmpsql & " AND ordID=" & ordid
			else
				success=false
				errmsg="The order id you specified seems to be invalid - " & ordid
				tmpsql = tmpsql & " AND ordID=0"
			end if
		else
			if fromdate<>"" then
				if IsNumeric(fromdate) then
					thefromdate = (Date()-fromdate)
				else
					err.number=0
					on error resume next
					thefromdate = DateValue(fromdate)
					if err.number <> 0 then
						thefromdate = Date()
						success=false
						errmsg="One of your date values was invalid - " & fromdate
					end if
					on error goto 0
				end if
				if todate="" then
					thetodate = thefromdate
				elseif IsNumeric(todate) then
					thetodate = (Date()-todate)
				else
					err.number=0
					on error resume next
					thetodate = DateValue(todate)
					if err.number <> 0 then
						thetodate = Date()
						success=false
						errmsg="One of your date values was invalid - " & todate
					end if
					on error goto 0
				end if
				if thefromdate > thetodate then
					tmpdate = thetodate
					thetodate = thefromdate
					thefromdate = tmpdate
				end if
				sd = thefromdate
				ed = thetodate
				tmpsql = tmpsql & " AND ordDate BETWEEN " & datedelim & VSUSDate(thefromdate) & datedelim & " AND " & datedelim & VSUSDate(thetodate+1) & datedelim
			end if
			if ordstatus<>"" AND NOT InStr(ordstatus,"9999")>0 then tmpsql = tmpsql & " AND ordStatus IN (" & ordstatus & ")"
			if searchtext<>"" then tmpsql = tmpsql & " AND (ordAuthNumber LIKE '%"&searchtext&"%' OR ordName LIKE '%"&searchtext&"%' OR ordEmail LIKE '%"&searchtext&"%' OR ordAddress LIKE '%"&searchtext&"%' OR ordCity LIKE '%"&searchtext&"%' OR ordState LIKE '%"&searchtext&"%' OR ordZip LIKE '%"&searchtext&"%' OR ordPhone LIKE '%"&searchtext&"%')"
		end if
		tmpsql = tmpsql & " ORDER BY ordID"
	else
		tmpsql = tmpsql & " WHERE ordDate BETWEEN "&datedelim & VSUSDate(sd) & datedelim & " AND " & datedelim & VSUSDate(DateValue(ed)+1) & datedelim & " ORDER BY ordID"
	end if
	getsearchparams = tmpsql
end function
if storesessionvalue="" then storesessionvalue="virtualstore"
if Session("loggedon") <> storesessionvalue OR disallowlogin=TRUE then response.redirect "login.asp"
hasdetails = request.form("act")="dumpdetails"
Response.ContentType = "unknown/exe"
if request.form("act")="stockinventory" then
Response.AddHeader "Content-Disposition","attachment;filename=stockinventory.csv"
elseif request.form("act")="dump2COinventory" then
Response.AddHeader "Content-Disposition","attachment;filename=inventory2co.csv"
elseif request.form("act")="fullinventory" then
Response.AddHeader "Content-Disposition","attachment;filename=inventory.csv"
elseif request.form("act")="dumpaffiliate" then
Response.AddHeader "Content-Disposition","attachment;filename=affilreport.csv"
elseif request.form("act")="quickbooks" then
elseif request.form("act")="ouresolutionsxmldump" then
Response.AddHeader "Content-Disposition","attachment;filename=oes_ordersdata.xml"
elseif hasdetails then
Response.AddHeader "Content-Disposition","attachment;filename=orderdetails.csv"
else
Response.AddHeader "Content-Disposition","attachment;filename=dumporders.csv"
end if
sslok=true
if request.servervariables("HTTPS")<>"on" AND (Request.ServerVariables("SERVER_PORT_SECURE") <> "1") AND nochecksslserver<>true then sslok=false
Set rs = Server.CreateObject("ADODB.RecordSet")
Set rs2 = Server.CreateObject("ADODB.RecordSet")
Set cnn=Server.CreateObject("ADODB.Connection")
cnn.open sDSN
alreadygotadmin = getadminsettings()
if Request.Form("sd") = "" then
	sd=Date()
else
	sd=Request.Form("sd")
end if
if Request.Form("ed") = "" then
	ed=Date()
else
	ed=Request.Form("ed")
end if
if request.form("act")="dumpaffiliate" then
	tdt = DateValue(sd)
	tdt2 = DateValue(ed)+1
	Response.write "Affiliate report for " & sd & " to " & ed & vbCrLf
	Response.write """ID"",""Name"",""Address"",""City"",""State"",""Zip"",""Country"",""Email"",""Total""" & vbCrLf
	if mysqlserver=true then
		sSQL = "SELECT affilID,affilName,affilAddress,affilCity,affilState,affilZip,affilCountry,affilEmail,SUM(ordTotal-ordDiscount) AS sumTot FROM affiliates LEFT JOIN orders ON affiliates.affilID=orders.ordAffiliate WHERE ordStatus>=3 AND ordDate BETWEEN " & datedelim & VSUSDate(tdt) & datedelim & " AND " & datedelim & VSUSDate(tdt2) & datedelim & " OR orders.ordAffiliate IS NULL GROUP BY affilID ORDER BY affilID"
	else
		sSQL = "SELECT affilID,affilName,affilAddress,affilCity,affilState,affilZip,affilCountry,affilEmail,(SELECT Sum(ordTotal-ordDiscount) FROM orders WHERE ordStatus>=3 AND ordAffiliate=affilID AND ordDate BETWEEN " & datedelim & VSUSDate(tdt) & datedelim & " AND " & datedelim & VSUSDate(tdt2) & datedelim & ") FROM affiliates ORDER BY affilID"
	end if
	rs.Open sSQL,cnn,0,1
	do while NOT rs.EOF
		response.write """"&replace(rs("affilID")&"","""","""""")&""","
		response.write """"&replace(rs("affilName")&"","""","""""")&""","
		response.write """"&replace(rs("affilAddress")&"","""","""""")&""","
		response.write """"&replace(rs("affilCity")&"","""","""""")&""","
		response.write """"&replace(rs("affilState")&"","""","""""")&""","
		response.write """"&replace(rs("affilZip")&"","""","""""")&""","
		response.write """"&replace(rs("affilCountry")&"","""","""""")&""","
		response.write """"&replace(rs("affilEmail")&"","""","""""")&""","
		response.write """"&rs(8)&""""&vbCrLf
		rs.MoveNext
	loop
	rs.Close
elseif request.form("act")="stockinventory" then
	sSQL2 = "SELECT pID,pName,pPrice,pInStock,pStockByOpts FROM products"
	rs.Open sSQL2,cnn,0,1
	response.write "pID,pName,pPrice,pInStock,optID,OptionGroup,Option" & vbCrLf
	do while NOT rs.EOF
		if rs("pStockByOpts") <> 0 then
			rs2.Open "SELECT optID,optGrpName,optName,optStock FROM optiongroup INNER JOIN (options INNER JOIN prodoptions ON options.optGroup=prodoptions.poOptionGroup) ON optiongroup.optGrpID=options.optGroup WHERE prodoptions.poProdID='"&replace(rs("pID"),"'","''")&"'",cnn,0,1
			do while NOT rs2.EOF
				response.write """"&replace(rs("pID")&"","""","""""")&""","
				response.write """"&replace(rs("pName")&"","""","""""")&""","
				response.write """"&rs("pPrice")&""","
				response.write rs2("optStock")&","
				response.write trim(rs2("optID"))&","
				response.write """"&replace(rs2("optGrpName")&"","""","""""")&""","
				response.write """"&replace(rs2("optName")&"","""","""""")&""""&vbCrLf
				rs2.MoveNext
			loop
			rs2.Close
		else
			response.write """"&replace(rs("pID")&"","""","""""")&""","
			response.write """"&replace(rs("pName")&"","""","""""")&""","
			response.write """"&rs("pPrice")&""","
			response.write rs("pInStock")&",,,"&vbCrLf
		end if
		rs.MoveNext
	loop
	rs.Close
elseif request.form("act")="fullinventory" then
	fieldlist = "pID,pName"
	for index=2 to adminlanguages+1
		if (adminlangsettings AND 1)=1 then fieldlist = fieldlist & ",pName"&index
	next
	fieldlist = fieldlist & ",pSection,pImage,pLargeimage,pPrice,pWholesalePrice,pListPrice,pShipping,pShipping2,pWeight,pDisplay,pSell,pExemptions,pInStock,pDims,pTax,pDropship"
	if digidownloads=TRUE then fieldlist = fieldlist & ",pDownload"
	fieldlist = fieldlist & ",pStaticPage,pStockByOpts,pDescription"
	for index=2 to adminlanguages+1
		if (adminlangsettings AND 2)=2 then sSQL2 = sSQL2 & ",pDescription"&index
	next
	fieldlist = fieldlist & ",pLongDescription"
	for index=2 to adminlanguages+1
		if (adminlangsettings AND 4)=4 then sSQL2 = sSQL2 & ",pLongDescription"&index
	next
	rs.Open "SELECT " & fieldlist & " FROM products",cnn,0,1
	fieldlistarr = split(fieldlist, ",")
	fieldlistcnt = UBOUND(fieldlistarr)
	for index=0 to fieldlistcnt
		response.write """"&fieldlistarr(index)&""""
		if index < fieldlistcnt then response.write ","
	next
	response.write vbCrLf
	do while NOT rs.EOF
		for index=0 to fieldlistcnt
			fieldtype = rs.fields(fieldlistarr(index)).type
			if fieldtype=11 then
				response.write IIfVr(rs(fieldlistarr(index)),"1","0")
			elseif (fieldtype >= 2 AND fieldtype<=5) OR (fieldtype >= 14 AND fieldtype <= 21) then
				response.write rs(fieldlistarr(index))
			else
				response.write """"&replace(rs(fieldlistarr(index))&"","""","""""")&""""
			end if
			if index < fieldlistcnt then response.write ","
		next
		response.write vbCrLf
		rs.MoveNext
	loop
	rs.Close
elseif request.form("act")="dump2COinventory" then
	sSQL2 = "SELECT payProvData1 FROM payprovider WHERE payProvID=2"
	rs.Open sSQL2,cnn,0,1
	response.write rs("payProvData1") & vbCrLf
	rs.Close
	sSQL2 = "SELECT pID,pName,pPrice,"&IIfVr(digidownloads=TRUE,"pDownload,","")&"pDescription FROM products"
	rs.Open sSQL2,cnn,0,1
	do while NOT rs.EOF
		response.write replace(rs("pID"),",","&#44;")&","
		response.write replace(replace(strip_tags2(rs("pName")),",","&#44;"),vbNewline," ")&","
		response.write ","
		response.write rs("pPrice")&","
		response.write ",,"
		if digidownloads=TRUE then
			response.write IIfVr(trim(rs("pDownload")&"")<>"", "N", "Y")&","
		else
			response.write "Y,"
		end if
		response.write replace(replace(strip_tags2(rs("pDescription")&""),",","&#44;"),vbNewline,"\n")&vbCrLf
		rs.MoveNext
	loop
	rs.Close
elseif request.form("act")="quickbooks" then
	sSQL2 = "SELECT ordID,ordName,ordAddress,ordAddress2,ordCity,ordState,ordZip,ordCountry,ordEmail,ordPhone,ordExtra1,ordExtra2,ordExtra3,ordShipName,ordShipAddress,ordShipAddress2,ordShipCity,ordShipState,ordShipZip,ordShipCountry,payProvName,ordAuthNumber,ordTotal,ordDate,ordStateTax,ordCountryTax,ordHSTTax,ordShipping,ordHandling,ordShipType,ordDiscount,ordAddInfo FROM orders INNER JOIN payprovider ON payprovider.payProvID=orders.ordPayProvider"
	sSQL2 = sSQL2 & getsearchparams()
	response.write "!TRNS	DATE	ACCNT	NAME	CLASS	AMOUNT	MEMO" & vbCrLf
	response.write "!SPL	DATE	ACCNT	NAME	AMOUNT	MEMO" & vbCrLf
	response.write "!ENDTRNS" & vbCrLf
	rs.Open sSQL2,cnn,0,1
	do while NOT rs.EOF
		response.write "TRNS" & vbTab & """" & vsusdate(rs("ordDate")) & """"
		rs.MoveNext
	loop
	rs.Close
elseif request.form("act")="ouresolutionsxmldump" then
	response.write "<?xml version=""1.0""?>" & vbCrLf
	response.write "<DATABASE NAME=""DataBaseCopy.mdb"" >" & vbCrLf
	sSQL = "SELECT ordID,cartProdId,cartProdName,cartProdPrice,cartQuantity,cartID FROM cart INNER JOIN orders ON cart.cartOrderId=orders.ordID"
	sSQL = sSQL & getsearchparams()
	rs.Open sSQL,cnn,0,1
	do while NOT rs.EOF
		theoptionspricediff=0
		sSQL = "SELECT coPriceDiff,coOptGroup,coCartOption FROM cartoptions WHERE coCartID=" & rs("cartID")
		rs2.Open sSQL,cnn,0,1
		do while NOT rs2.EOF
			theoptionspricediff = theoptionspricediff + rs2("coPriceDiff")
			rs2.MoveNext
		loop
		rs2.Close
		theunitprice = rs("cartProdPrice")+theoptionspricediff
		sSQL = "SELECT pName,pDescription,pDropShip FROM products WHERE pID='"&rs("cartProdID")&"'"
		rs2.Open sSQL,cnn,0,1
		if NOT rs2.EOF then
			prodname = strip_tags2(rs2("pName")&"")
			proddesc = strip_tags2(rs2("pDescription")&"")
			supplier = rs2("pDropShip")
		else
			prodname = ""
			proddesc = ""
			supplier = 0
		end if
		if ouresolutionsxml=1 then
			itemname = strip_tags2(rs("cartProdID")) & "chr(60)brchr(62)" & proddesc
		elseif ouresolutionsxml=3 then
			itemname = strip_tags2(rs("cartProdID"))
		elseif ouresolutionsxml=4 then
			itemname = prodname
		else ' default to "2"
			itemname = prodname & "chr(60)brchr(62)" & proddesc
		end if
		rs2.Close
		response.write "<DATA TABLE='oitems' ORDERITEMID='"&rs("cartID")&"' ORDERID='"&rs("ordID")&"' CATALOGID='"&rs("cartID")&"' NUMITEMS='"&rs("cartQuantity")&"' ITEMNAME='"&xmlstrip(itemname)&"' UNITPRICE='"&twodp(theunitprice)&"' DUALPRICE='0' SUPPLIERID='"&supplier&"' ADDRESS='' />" & vbCrLf
		rs.MoveNext
	loop
	rs.Close
	sSQL = "SELECT ordID,ordName,ordAddress,ordAddress2,ordCity,ordState,ordZip,ordCountry,ordEmail,ordPhone,ordExtra1,ordExtra2,ordExtra3,ordShipName,ordShipAddress,ordShipAddress2,ordShipCity,ordShipState,ordShipZip,ordShipCountry,ordPayProvider,ordAuthNumber,ordTotal,ordDate,ordStateTax,ordCountryTax,ordHSTTax,ordShipping,ordHandling,ordShipType,ordDiscount,ordAffiliate,ordDiscountText,ordStatus,statPrivate,ordAddInfo FROM orders INNER JOIN orderstatus ON orders.ordStatus=orderstatus.statID"
	sSQL = sSQL & getsearchparams()
	rs.Open sSQL,cnn,0,1
	do while NOT rs.EOF
		ordGrandTotal = (rs("ordTotal")+rs("ordStateTax")+rs("ordCountryTax")+rs("ordHSTTax")+rs("ordShipping")+rs("ordHandling"))-rs("ordDiscount")
		thename = xmlstrip(trim(rs("ordName")&""))
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
		response.write "<DATA TABLE='orders' ORDERID='"&rs("ordID")&"' OCUSTOMERID='"&rs("ordID")&"' ODATE='"&DateValue(rs("ordDate"))&"' ORDERAMOUNT='"&twodp(ordGrandTotal)&"' OFIRSTNAME='"&firstname&"' OLASTNAME='"&lastname&"' OEMAIL='"&xmlstrip(rs("ordEmail"))&"' OADDRESS='"&xmlstrip(rs("ordAddress")&IIfVr(trim(rs("ordAddress2")&"")<>"",", " & rs("ordAddress2"), ""))&"' OCITY='"&xmlstrip(rs("ordCity"))&"' OPOSTCODE='"&xmlstrip(rs("ordZip"))&"' OSTATE='"&xmlstrip(rs("ordState"))&"' OCOUNTRY='"&xmlstrip(rs("ordCountry"))&"' OPHONE='"&right(xmlstrip(replace(replace(replace(rs("ordPhone")&""," ", ""),".",""),"-","")), 10)&"' OFAX='' OCOMPANY='"&IIfVr(extra1iscompany=TRUE,xmlstrip(rs("ordExtra1")), "")&"' OCARDTYPE='' "
		if dumpccnumber then
			if sslok=false then
				response.write "OCARDNO='No SSL' OCARDNAME='No SSL' OCARDEXPIRES='No SSL' OCARDADDRESS='No SSL' "
			else
				rs2.Open "SELECT ordCNum FROM orders WHERE ordID=" & rs("ordID"),cnn,0,1
				ordCNum = rs2("ordCNum")
				encryptmethod = LCase(encryptmethod&"")
				if encryptmethod="aspencrypt" OR encryptmethod="" then
					response.write "OCARDNO='Encrypted' OCARDNAME='Encrypted' OCARDEXPIRES='Encrypted' OCARDADDRESS='Encrypted' "
				elseif Trim(ordCNum)="" OR IsNull(ordCNum) then
					response.write "OCARDNO='' OCARDNAME='' OCARDEXPIRES='' OCARDADDRESS='' "
				elseif encryptmethod="none" then
					cnumarr = Split(ordCNum, "&")
					if IsArray(cnumarr) then
						response.write "OCARDNO='"&cnumarr(0)&"' OCARDNAME='"&cnumarr(3)&"' OCARDEXPIRES='"&cnumarr(1)&"' OCARDADDRESS='"&rs("ordAddress")&IIfVr(trim(rs("ordAddress2")&"")<>"",", " & rs("ordAddress2"), "")&"' "
					else
						response.write "OCARDNO='' OCARDNAME='' OCARDEXPIRES='' OCARDADDRESS='' "
					end if
				end if
				rs2.Close
			end if
		else
			response.write "OCARDNO='' OCARDNAME='' OCARDEXPIRES='' OCARDADDRESS='' "
		end if
		response.write "OPROCESSED='' OCOMMENT='"&xmlstrip(rs("ordAddInfo"))&"' OTAX='"&twodp(rs("ordStateTax")+rs("ordCountryTax")+rs("ordHSTTax"))&"' OPROMISEDSHIPDATE='' OSHIPPEDDATE='' OSHIPMETHOD='0' OSHIPCOST='"&twodp(rs("ordShipping"))&"' "
		response.write "OSHIPNAME='"&xmlstrip(rs("ordShipName"))&"' OSHIPCOMPANY='' OSHIPEMAIL='' OSHIPMETHODTYPE='"&xmlstrip(rs("ordShipType"))&"' OSHIPADDRESS='"&xmlstrip(rs("ordShipAddress")&IIfVr(trim(rs("ordShipAddress2")&"")<>"",", " & rs("ordShipAddress2"), ""))&"' OSHIPTOWN='"&xmlstrip(rs("ordShipCity"))&"' OSHIPZIP='"&xmlstrip(rs("ordShipZip"))&"' OSHIPCOUNTRY='"&xmlstrip(rs("ordShipCountry"))&"' OSHIPSTATE='"&xmlstrip(rs("ordShipState"))&"' "
		response.write "OPAYMETHOD='"&rs("ordPayProvider")&"' OTHER1='"&IIfVr(extra1iscompany=TRUE,"",xmlstrip(rs("ordExtra1")))&"' OTHER2='"&xmlstrip(rs("ordExtra2"))&"' OTIME='' OAUTHORIZATION='"&xmlstrip(rs("ordAuthNumber"))&"' OERRORS='' ODISCOUNT='"&twodp(rs("ordDiscount"))&"' OSTATUS='"&xmlstrip(rs("statPrivate"))&"' OAFFID='' ODUALTOTAL='0' ODUALTAXES='0' ODUALSHIPPING='0' ODUALDISCOUNT='0' OHANDLING='"&twodp(rs("ordHandling"))&"' COUPON='"&xmlstrip(strip_tags2(rs("ordDiscountText")&""))&"' COUPONDISCOUNT='0' COUPONDISCOUNTDUAL='0' GIFTCERTIFICATE='' GIFTAMOUNTUSED='0' GIFTAMOUNTUSEDDUAL='0' CANCELED='"&IIfVr(rs("ordStatus")<2,"True","False")&"' />" & vbCrLf
		rs.MoveNext
	loop
	rs.Close
	response.write "</DATABASE>" & vbCrLf
else
	if hasdetails then
		sSQL2 = "SELECT ordID,ordName,ordAddress,ordAddress2,ordCity,ordState,ordZip,ordCountry,ordEmail,ordPhone,ordExtra1,ordExtra2,ordExtra3,ordShipName,ordShipAddress,ordShipAddress2,ordShipCity,ordShipState,ordShipZip,ordShipCountry,payProvName,ordAuthNumber,ordTotal,ordDate,ordStateTax,ordCountryTax,ordHSTTax,ordShipping,ordHandling,ordShipType,statPrivate,cartProdId,cartProdName,cartProdPrice,cartQuantity,cartID,ordDiscount,ordAddInfo FROM cart INNER JOIN ((orderstatus RIGHT OUTER JOIN orders ON orders.ordStatus=orderstatus.statID) INNER JOIN payprovider ON payprovider.payProvID=orders.ordPayProvider) ON cart.cartOrderId=orders.ordID"
	else
		sSQL2 = "SELECT ordID,ordName,ordAddress,ordAddress2,ordCity,ordState,ordZip,ordCountry,ordEmail,ordPhone,ordExtra1,ordExtra2,ordExtra3,ordShipName,ordShipAddress,ordShipAddress2,ordShipCity,ordShipState,ordShipZip,ordShipCountry,payProvName,ordAuthNumber,ordTotal,ordDate,ordStateTax,ordCountryTax,ordHSTTax,ordShipping,ordHandling,ordShipType,statPrivate,ordDiscount,ordAddInfo FROM orderstatus RIGHT OUTER JOIN (orders INNER JOIN payprovider ON payprovider.payProvID=orders.ordPayProvider) ON orders.ordStatus=orderstatus.statID"
	end if
	sSQL2 = sSQL2 & getsearchparams()
	rs.Open sSQL2,cnn,0,1
	response.write """OrderID"","
	if extraorderfield1<>"" then response.write """" & replace(extraorderfield1,"""","""""") & ""","
	response.write """CustomerName"",""Address"","
	if useaddressline2=TRUE then response.write """Address2"","
	response.write """City"",""State"",""Zip"",""Country"",""Email"",""Phone"","
	if extraorderfield2<>"" then response.write """" & replace(extraorderfield2,"""","""""") & ""","
	if extraorderfield3<>"" then response.write """" & replace(extraorderfield3,"""","""""") & ""","
	response.write """ShipName"",""ShipAddress"","
	if useaddressline2=TRUE then response.write """ShipAddress2"","
	response.write """ShipCity"",""ShipState"",""ShipZip"",""ShipCountry"",""PaymentMethod"",""AuthNumber"",""Total"",""Date"",""StateTax"",""CountryTax"","
	if canadataxsystem=true then response.write """HST"","
	response.write """Shipping"",""Handling"",""Discounts"",""AddInfo"",""ShippingMethod"",""Status"""
	if dumpccnumber then response.write ",""Card Number"",""Expiry Date"",""CVV Code"",""Issue Number"""
	if hasdetails then response.write ",""ProductID"",""ProductName"",""ProductPrice"",""Quantity"",""Options"""
	response.write vbCrLf
	do while NOT rs.EOF
			response.write rs("ordID")&","
			if extraorderfield1<>"" then response.write """"&replace(rs("ordExtra1")&"","""","""""")&""","
			response.write """"&replace(rs("ordName")&"","""","""""")&""","
			response.write """"&replace(rs("ordAddress")&"","""","""""")&""","
			if useaddressline2=TRUE then response.write """"&replace(rs("ordAddress2")&"","""","""""")&""","
			response.write """"&replace(rs("ordCity")&"","""","""""")&""","
			response.write """"&replace(rs("ordState")&"","""","""""")&""","
			response.write """"&replace(rs("ordZip")&"","""","""""")&""","
			response.write """"&replace(rs("ordCountry")&"","""","""""")&""","
			response.write """"&replace(rs("ordEmail")&"","""","""""")&""","
			response.write """"&replace(rs("ordPhone")&"","""","""""")&""","
			if extraorderfield2<>"" then response.write """"&replace(rs("ordExtra2")&"","""","""""")&""","
			if extraorderfield3<>"" then response.write """"&replace(rs("ordExtra3")&"","""","""""")&""","
			response.write """"&replace(rs("ordShipName")&"","""","""""")&""","
			response.write """"&replace(rs("ordShipAddress")&"","""","""""")&""","
			if useaddressline2=TRUE then response.write """"&replace(rs("ordShipAddress2")&"","""","""""")&""","
			response.write """"&replace(rs("ordShipCity")&"","""","""""")&""","
			response.write """"&replace(rs("ordShipState")&"","""","""""")&""","
			response.write """"&replace(rs("ordShipZip")&"","""","""""")&""","
			response.write """"&replace(rs("ordShipCountry")&"","""","""""")&""","
			response.write """"&replace(rs("payProvName")&"","""","""""")&""","
			response.write """"&replace(rs("ordAuthNumber")&"","""","""""")&""","
			response.write """"&rs("ordTotal")&""","
			response.write """"&rs("ordDate")&""","
			response.write """"&rs("ordStateTax")&""","
			response.write """"&rs("ordCountryTax")&""","
			if canadataxsystem=true then response.write """"&rs("ordHSTTax")&""","
			response.write """"&rs("ordShipping")&""","
			response.write """"&rs("ordHandling")&""","
			response.write """"&rs("ordDiscount")&""","
			response.write """"&replace(rs("ordAddInfo")&"","""","""""")&""","
			response.write """"&replace(rs("ordShipType")&"","""","""""")&""","
			response.write """"&replace(rs("statPrivate")&"","""","""""")&""""
			if dumpccnumber then
				if sslok=false then
					response.write ",No SSL,No SSL,No SSL,No SSL"
				else
					rs2.Open "SELECT ordCNum FROM orders WHERE ordID=" & rs("ordID"),cnn,0,1
					ordCNum = rs2("ordCNum")
					encryptmethod = LCase(encryptmethod&"")
					if encryptmethod="aspencrypt" OR encryptmethod="" then
						response.write """Encrypted"",""Encrypted"",""Encrypted"",""Encrypted"""
					elseif Trim(ordCNum)="" OR IsNull(ordCNum) then
						response.write ",""(no data)"","""","""","""""
					elseif encryptmethod="none" then
						cnumarr = Split(ordCNum, "&")
						if IsArray(cnumarr) then
							response.write ","""""""&cnumarr(0)&""""""""
							if UBOUND(cnumarr)>=1 then response.write ","""""""&cnumarr(1)&"""""""" else response.write ","""""
							if UBOUND(cnumarr)>=2 then response.write ","""&cnumarr(2)&"""" else response.write ","""""
							if UBOUND(cnumarr)>=3 then response.write ","""&cnumarr(3)&"""" else response.write ","""""
						else
							response.write ",""(no data)"","""","""","""""
						end if
					end if
					rs2.Close
				end if
			end if
			if hasdetails then
				theOptions = ""
				thePriceDiff = 0
				rs2.Open "SELECT coPriceDiff,coOptGroup,coCartOption FROM cartoptions WHERE coCartID=" & rs("cartID"),cnn,0,1
				do while NOT rs2.EOF
					theOptions = theOptions & "," & """" & replace(rs2("coOptGroup")&"","""","""""") & " - " & replace(rs2("coCartOption"),"""","""""") & """"
					thePriceDiff = thePriceDiff + rs2("coPriceDiff")
					rs2.MoveNext
				loop
				response.write ","""&replace(rs("cartProdId")&"","""","""""")&""""
				response.write ","""&replace(rs("cartProdName")&"","""","""""")&""""
				response.write ","&rs("cartProdPrice")+thePriceDiff
				response.write ","&rs("cartQuantity")
				response.write theOptions
				rs2.Close
			end if
			response.write vbCrLf
		rs.MoveNext
	loop
	rs.Close
end if
cnn.Close
set rs = nothing
set rs2 = nothing
set cnn = nothing
%>
