<%
Dim gasaReferer,gasaThisSite,datedelim
Dim splitUSZones,countryCurrency,useEuro,storeurl,stockManage,delAfter,handling,adminCanPostUser,packtogether,origZip,shipType,adminIntShipping,saveLCID,delccafter,adminTweaks,currRate1,currSymbol1,currRate2,currSymbol2,currRate3,currSymbol3,upsUser,upsPw
Dim origCountry,origCountryCode,uspsUser,uspsPw,upsAccess,fedexaccount,fedexmeter,adminUnits,adminlanguages,adminlangsettings,useStockManagement,adminProdsPerPage,countryTax,countryTaxRate,currLastUpdate,currConvUser,currConvPw,emailAddr,sendEmail,emailObject,themailhost,theuser,thepass
incfunctionsdefined=true
function ip2long(ip2lip)
ipret = -1
iparr = split(ip2lip, ".")
if isarray(iparr) then
if UBOUND(iparr)=3 then
if isnumeric(iparr(0)) AND isnumeric(iparr(1)) AND isnumeric(iparr(2)) AND isnumeric(iparr(3)) then
ipret = (iparr(0) * 16777216) + (iparr(1) * 65536) + (iparr(2) * 256) + (iparr(3))
end if
end if
end if
ip2long = ipret
end function
if Trim(request.querystring("PARTNER"))<>"" OR Trim(request.querystring("REFERER"))<>"" then
	if expireaffiliate = "" then expireaffiliate=30
	if Trim(request.querystring("PARTNER"))<>"" then thereferer=Trim(request.querystring("PARTNER")) else thereferer=Trim(request.querystring("REFERER"))
	response.write "<script src='vsadmin/savecookie.asp?PARTNER="&thereferer&"&EXPIRES="&expireaffiliate&"'></script>"
end if
if mysqlserver=true then sqlserver=true
if sqlserver=true then datedelim = "'" else datedelim = "#"
codestr="2952710692840328509902143349209039553396765"
if emailencoding="" then emailencoding="iso-8859-1"
if adminencoding="" then adminencoding="iso-8859-1"
if Session("languageid") <> "" then languageid=Session("languageid")
function getadminsettings()
	if NOT alreadygotadmin then
		if saveadmininapplication AND Application("getadminsettings")<>"" then
			splitUSZones = Application("splitUSZones")
			if orlocale<>"" then saveLCID = orlocale else saveLCID = Application("saveLCID")
			Session.LCID = saveLCID
			countryCurrency = Application("countryCurrency")
			useEuro = Application("useEuro")
			storeurl = Application("storeurl")
			stockManage = Application("adminStockManage")
			useStockManagement = Application("useStockManagement")
			adminProdsPerPage = Application("adminProdsPerPage")
			countryTax = Application("countryTax")
			countryTaxRate = Application("countryTax")
			delAfter = Application("delAfter")
			delccafter = Application("delccafter")
			handling = Application("handling")
			adminCanPostUser = Application("adminCanPostUser")
			packtogether = Application("packtogether")
			origZip = Application("origZip")
			shipType = Application("shipType")
			adminIntShipping = Application("adminIntShipping")
			origCountry = Application("origCountry")
			origCountryCode = Application("origCountryCode")
			uspsUser = Application("uspsUser")
			uspsPw = Application("uspsPw")
			upsUser = Application("upsUser")
			upsPw = Application("upsPw")
			upsAccess = Application("upsAccess")
			fedexaccount = Application("fedexaccount")
			fedexmeter = Application("fedexmeter")
			adminUnits = Application("adminUnits")
			emailObject = Application("emailObject")
			themailhost = Application("themailhost")
			theuser = Application("theuser")
			thepass = Application("thepass")
			emailAddr = Application("emailAddr")
			sendEmail = Application("sendEmail")
			adminTweaks = Application("adminTweaks")
			adminlanguages = Application("adminlanguages")
			adminlangsettings = Application("adminlangsettings")
			currRate1 = Application("currRate1")
			currSymbol1 = Application("currSymbol1")
			currRate2 = Application("currRate2")
			currSymbol2 = Application("currSymbol2")
			currRate3 = Application("currRate3")
			currSymbol3 = Application("currSymbol3")
			currConvUser = Application("currConvUser")
			currConvPw = Application("currConvPw")
			currLastUpdate = Application("currLastUpdate")
		else
			sSQL = "SELECT adminEmail,emailObject,smtpserver,emailUser,emailPass,adminEmailConfirm,adminTweaks,adminProdsPerPage,adminStoreURL,adminHandling,adminPacking,adminDelUncompleted,adminDelCC,adminUSZones,adminStockManage,adminShipping,adminIntShipping,adminCanPostUser,adminZipCode,adminUnits,adminUSPSUser,adminUSPSpw,adminUPSUser,adminUPSpw,adminUPSAccess,FedexAccountNo,FedexMeter,adminlanguages,adminlangsettings,currRate1,currSymbol1,currRate2,currSymbol2,currRate3,currSymbol3,currConvUser,currConvPw,currLastUpdate,countryLCID,countryCurrency,countryName,countryCode,countryTax FROM admin INNER JOIN countries ON admin.adminCountry=countries.countryID WHERE adminID=1"
			rs.Open sSQL,cnn,0,1
			splitUSZones = (Int(rs("adminUSZones"))=1)
			if orlocale<>"" then
				Session.LCID = orlocale
			elseif rs("countryLCID")<>0 then
				Session.LCID = rs("countryLCID")
			end if
			saveLCID = Session.LCID
			countryCurrency = rs("countryCurrency")
			if orcurrencyisosymbol<>"" then countryCurrency=orcurrencyisosymbol
			useEuro = (countryCurrency="EUR")
			storeurl = rs("adminStoreURL")
			stockManage = rs("adminStockManage")
			useStockManagement = (rs("adminStockManage")<>0)
			adminProdsPerPage = rs("adminProdsPerPage")
			countryTax=cDbl(rs("countryTax"))
			countryTaxRate=cDbl(rs("countryTax"))
			delAfter = Int(rs("adminDelUncompleted"))
			delccafter = Int(rs("adminDelCC"))
			handling = cDbl(rs("adminHandling"))
			adminCanPostUser = trim(rs("adminCanPostUser"))
			packtogether = Int(rs("adminPacking"))=1
			origZip = rs("adminZipCode")
			shipType = Int(rs("adminShipping"))
			adminIntShipping = Int(rs("adminIntShipping"))
			origCountry = rs("countryName")
			origCountryCode = rs("countryCode")
			uspsUser = rs("adminUSPSUser")
			uspsPw = rs("adminUSPSpw")
			upsUser = upsdecode(rs("adminUPSUser"), "")
			upsPw = upsdecode(rs("adminUPSpw"), "")
			upsAccess = rs("adminUPSAccess")
			fedexaccount = rs("FedexAccountNo")
			fedexmeter = rs("FedexMeter")
			adminUnits=Int(rs("adminUnits"))
			emailObject = rs("emailObject")
			themailhost = Trim(rs("smtpserver")&"")
			theuser = Trim(rs("emailUser")&"")
			thepass = Trim(rs("emailPass")&"")
			emailAddr = rs("adminEmail")
			sendEmail = Int(rs("adminEmailConfirm"))=1
			adminTweaks = Int(rs("adminTweaks"))
			adminlanguages = Int(rs("adminlanguages"))
			adminlangsettings = Int(rs("adminlangsettings"))
			currRate1=cDbl(rs("currRate1"))
			currSymbol1=trim(rs("currSymbol1")&"")
			currRate2=cDbl(rs("currRate2"))
			currSymbol2=trim(rs("currSymbol2")&"")
			currRate3=cDbl(rs("currRate3"))
			currSymbol3=trim(rs("currSymbol3")&"")
			currConvUser=rs("currConvUser")
			currConvPw=rs("currConvPw")
			currLastUpdate=rs("currLastUpdate")
			rs.Close
			if saveadmininapplication=TRUE then
				Application.Lock()
				Application("splitUSZones") = splitUSZones
				Application("saveLCID") = saveLCID
				Application("countryCurrency") = countryCurrency
				Application("useEuro") = useEuro
				Application("storeurl") = storeurl
				Application("adminStockManage") = stockManage
				Application("useStockManagement") = useStockManagement
				Application("adminProdsPerPage") = adminProdsPerPage
				Application("countryTax") = countryTax
				Application("delAfter") = delAfter
				Application("delccafter") = delccafter
				Application("handling") = handling
				Application("adminCanPostUser") = adminCanPostUser
				Application("packtogether") = packtogether
				Application("origZip") = origZip
				Application("shipType") = shipType
				Application("adminIntShipping") = adminIntShipping
				Application("origCountry") = origCountry
				Application("origCountryCode") = origCountryCode
				Application("uspsUser") = uspsUser
				Application("uspsPw") = uspsPw
				Application("upsUser") = upsUser
				Application("upsPw") = upsPw
				Application("upsAccess") = upsAccess
				Application("fedexaccount") = fedexaccount
				Application("fedexmeter") = fedexmeter
				Application("adminUnits") = adminUnits
				Application("emailObject") = emailObject
				Application("themailhost") = themailhost
				Application("theuser") = theuser
				Application("thepass") = thepass
				Application("emailAddr") = emailAddr
				Application("sendEmail") = sendEmail
				Application("adminTweaks") = adminTweaks
				Application("adminlanguages") = adminlanguages
				Application("adminlangsettings") = adminlangsettings
				Application("currRate1") = currRate1
				Application("currSymbol1") = currSymbol1
				Application("currRate2") = currRate2
				Application("currSymbol2") = currSymbol2
				Application("currRate3") = currRate3
				Application("currSymbol3") = currSymbol3
				Application("currConvUser") = currConvUser
				Application("currConvPw") = currConvPw
				Application("currLastUpdate") = currLastUpdate
				Application("getadminsettings")=TRUE
				Application.UnLock()
			end if
		end if
	end if
	' Overrides
	if orstoreurl<>"" then storeurl=orstoreurl
	if (left(LCase(storeurl),7) <> "http://") AND (left(LCase(storeurl),8) <> "https://") then storeurl = "http://" & storeurl
	if Right(storeurl,1) <> "/" then storeurl = storeurl & "/"
	if oremailaddr<>"" then emailAddr=oremailaddr
	if adminIntShipping="" then adminIntShipping=0 ' failsafe
	getadminsettings = TRUE
end function
function strip_tags2(mistr)
Set toregexp = new RegExp
toregexp.pattern = "<[^>]+>"
toregexp.ignorecase = TRUE
toregexp.global = TRUE
mistr = toregexp.replace(mistr, "")
Set toregexp = Nothing
strip_tags2 = replace(mistr, """", "&quot;")
end function
function cleanforurl(surl)
if isempty(urlfillerchar) then urlfillerchar="_"
Set toregexp = new RegExp
toregexp.pattern = "<[^>]+>"
toregexp.ignorecase = TRUE
toregexp.global = TRUE
surl = replace(lcase(toregexp.replace(surl, ""))," ",urlfillerchar)
toregexp.pattern = "[^a-z\"&urlfillerchar&"0-9]"
cleanforurl = toregexp.replace(surl, "")
end function
function vrxmlencode(xmlstr)
	xmlstr = replace(xmlstr, "&", "&amp;")
	xmlstr = replace(xmlstr, "<", "&lt;")
	xmlstr = replace(xmlstr, ">", "&gt;")
	xmlstr = replace(xmlstr, "'", "&apos;")
	vrxmlencode = replace(xmlstr, """", "&quot;")
end function
function xmlencodecharref(xmlstr)
	xmlstr = replace(xmlstr, "&reg;", "")
	xmlstr = replace(xmlstr, "&", "&#x26;")
	xmlstr = replace(xmlstr, "<", "&#x3c;")
	xmlstr = replace(xmlstr, "®", "")
	xmlstr = replace(xmlstr, ">", "&#x3e;")
	tmp_str=""
	for i=1 to len(xmlstr)
		ch_code=Asc(Mid(xmlstr,i,1))
		if ch_code<=130 then tmp_str=tmp_str & Mid(xmlstr,i,1)
	next
	xmlencodecharref = tmp_str
end function
function getlangid(col, bfield)
	if languageid="" or languageid=1 then
		getlangid = col
	else
		if (adminlangsettings AND bfield)<>bfield then getlangid = col else getlangid = col & languageid
	end if
end function
function upsencode(thestr, propcodestr)
	if propcodestr="" then localcodestr=codestr else localcodestr=propcodestr
	newstr=""
	for index=1 to Len(localcodestr)
		thechar = Mid(localcodestr,index,1)
		if NOT IsNumeric(thechar) then
			thechar = asc(thechar) MOD 10
		end if
		newstr = newstr & thechar
	next
	localcodestr = newstr
	do while Len(localcodestr) < 40
		localcodestr = localcodestr & localcodestr
	loop
	newstr=""
	for index=1 to Len(thestr)
		thechar = Mid(thestr,index,1)
		newstr=newstr & Chr(asc(thechar)+Int(Mid(localcodestr,index,1)))
	next
	upsencode=newstr
end function
function upsdecode(thestr, propcodestr)
	if propcodestr="" then localcodestr=codestr else localcodestr=propcodestr
	newstr=""
	for index=1 to Len(localcodestr)
		thechar = Mid(localcodestr,index,1)
		if NOT IsNumeric(thechar) then
			thechar = asc(thechar) MOD 10
		end if
		newstr = newstr & thechar
	next
	localcodestr = newstr
	do while Len(localcodestr) < 40
		localcodestr = localcodestr & localcodestr
	loop
	if IsNull(thestr) then
		upsdecode=""
	else
		newstr=""
		for index=1 to Len(thestr)
			thechar = Mid(thestr,index,1)
			newstr=newstr & Chr(asc(thechar)-Int(Mid(localcodestr,index,1)))
		next
		upsdecode=newstr
	end if
end function
function VSUSDate(thedate)
	if mysqlserver=true then
		VSUSDate = DatePart("yyyy",thedate) & "-" & DatePart("m",thedate) & "-" & DatePart("d",thedate)
	elseif sqlserver=true then
		VSUSDate = right(DatePart("yyyy",thedate),2) & IIfVr(DatePart("m",thedate)<10,"0","") & DatePart("m",thedate) & IIfVr(DatePart("d",thedate)<10,"0","") & DatePart("d",thedate)
	else
		VSUSDate = DatePart("m",thedate) & "/" & DatePart("d",thedate) & "/" & DatePart("yyyy",thedate)
	end if
end function
function VSUSDateTime(thedate)
	if mysqlserver=true then
		VSUSDateTime = DatePart("yyyy",thedate) & "-" & DatePart("m",thedate) & "-" & DatePart("d",thedate) & " " & DatePart("h",thedate) & ":" & DatePart("n",thedate) & ":" & DatePart("s",thedate)
	elseif sqlserver=true then
		VSUSDateTime = right(DatePart("yyyy",thedate),2) & IIfVr(DatePart("m",thedate)<10,"0","") & DatePart("m",thedate) & IIfVr(DatePart("d",thedate)<10,"0","") & DatePart("d",thedate) & " " & DatePart("h",thedate) & ":" & DatePart("n",thedate) & ":" & DatePart("s",thedate)
	else
		VSUSDateTime = DatePart("m",thedate) & "/" & DatePart("d",thedate) & "/" & DatePart("yyyy",thedate) & " " & DatePart("h",thedate) & ":" & DatePart("n",thedate) & ":" & DatePart("s",thedate)
	end if
end function
function FormatEuroCurrency(amount)
	if overridecurrency=true then
		if orcpreamount=true then FormatEuroCurrency = orcsymbol & FormatNumber(amount,orcdecplaces) else FormatEuroCurrency = FormatNumber(amount,orcdecplaces) & orcsymbol
	else
		if useEuro then FormatEuroCurrency = FormatNumber(amount,2) & " &euro;" else FormatEuroCurrency = FormatCurrency(amount,-1,-2,0,-2)
	end if
end function
function FormatEmailEuroCurrency(amount)
	if overridecurrency=true then
		if orcpreamount=true then FormatEmailEuroCurrency = orcemailsymbol & FormatNumber(amount,orcdecplaces) else FormatEmailEuroCurrency = FormatNumber(amount,orcdecplaces) & orcemailsymbol
	else
		if useEuro then FormatEmailEuroCurrency = FormatNumber(amount,2) & " Euro" else FormatEmailEuroCurrency = FormatCurrency(amount,-1,-2,0,-2)
	end if
end function
Sub do_stock_management(smOrdId)
	smOrdId = Trim(smOrdId)
	If NOT IsNumeric(smOrdId) OR smOrdId="" then smOrdId=0
	Set rsl = Server.CreateObject("ADODB.RecordSet")
	if stockManage <> 0 then
		sSQL="SELECT cartID,cartProdID,cartQuantity,pStockByOpts FROM cart INNER JOIN products ON cart.cartProdID=products.pID WHERE (cartCompleted=0 OR cartCompleted=2) AND cartOrderID=" & smOrdId
		rsl.Open sSQL,cnn,0,1
		do while NOT rsl.EOF
			if cint(rsl("pStockByOpts")) <> 0 then
				sSQL = "SELECT coOptID FROM cartoptions INNER JOIN (options INNER JOIN optiongroup ON options.optGroup=optiongroup.optGrpID) ON cartoptions.coOptID=options.optID WHERE optType IN (-2,-1,1,2) AND coCartID=" & rsl("cartID")
				rs.Open sSQL,cnn,0,1
				do while NOT rs.EOF
					sSQL = "UPDATE options SET optStock=optStock-"&rsl("cartQuantity")&" WHERE optID="&rs("coOptID")
					cnn.Execute(sSQL)
					rs.MoveNext
				loop
				rs.Close
			else
				sSQL = "UPDATE products SET pInStock=pInStock-"&rsl("cartQuantity")&" WHERE pID='"&rsl("cartProdID")&"'"
				cnn.Execute(sSQL)
			end if
			rsl.MoveNext
		loop
		rsl.Close
	end if
	set rsl = nothing
End Sub
Sub productdisplayscript(doaddprodoptions)
if currSymbol1<>"" AND currFormat1="" then currFormat1="%s <strong>" & currSymbol1 & "</strong>"
if currSymbol2<>"" AND currFormat2="" then currFormat2="%s <strong>" & currSymbol2 & "</strong>"
if currSymbol3<>"" AND currFormat3="" then currFormat3="%s <strong>" & currSymbol3 & "</strong>"
%>
<script language="javascript" type="text/javascript">
<!--
<%	if NOT (pricecheckerisincluded=TRUE) then %>
var aPC = new Array();<%
		if useStockManagement then %>
var aPS = new Array();
function checkStock(x,i){
if(i!='' && aPS[i] > 0)return(true);
alert('<%=replace(xxOptOOS,"'","\'")%>');
x.focus();return(false);
}<%		end if %>
var isW3 = (document.getElementById&&true);
var tax=<%=replace(countryTaxRate,",",".") %>;
function dummyfunc(){};
function pricechecker(i){
if(i!='')return(aPC[i]);return(0);}
function enterValue(x){
alert('<%=replace(xxPrdEnt,"'","\'")%>');
x.focus();return(false);}
function chooseOption(x){
alert('<%=replace(xxPrdChs,"'","\'")%>');
x.focus();return(false);}
function dataLimit(x){
alert('<%=replace(xxPrd255,"'","\'")%>');
x.focus();return(false);}
function formatprice(i, currcode, currformat){
<%	tempStr = FormatEuroCurrency(0)
	tempStr2 = FormatNumber(0,2)
	response.write "var pTemplate='" & tempStr & "';" & vbCrLf
	response.write "if(currcode!='') pTemplate=' " & tempStr2 & "' + (currcode!=' '?'<strong>'+currcode+'<\/strong>':'');"
	if InStr(tempStr,",")<>0 OR InStr(tempStr,".")<>0 then %>
if(currcode==' JPY')i = Math.round(i).toString();
else if(i==Math.round(i))i=i.toString()+".00";
else if(i*10.0==Math.round(i*10.0))i=i.toString()+"0";
else if(i*100.0==Math.round(i*100.0))i=i.toString();
<%	end if
	response.write "if(currcode!='')pTemplate = currformat.toString().replace(/%s/,i.toString());"
	response.write "else pTemplate = pTemplate.toString().replace(/\d[,.]*\d*/,i.toString());"
	if InStr(tempStr,",")<>0 then
		response.write "return(pTemplate.replace(/\./,','));"
	else
		response.write "return(pTemplate);"
	end if
%>}
function openEFWindow(id){
window.open('emailfriend.asp?id='+id,'email_friend','menubar=no, scrollbars=no, width=400, height=400, directories=no,location=no,resizable=yes,status=no,toolbar=no')
}<%		pricecheckerisincluded=TRUE
	end if
if doaddprodoptions AND prodlist<>"" then
	Session.LCID = 1033
	sSQL = "SELECT DISTINCT optID,"&OWSP&"optPriceDiff,optStock FROM options INNER JOIN prodoptions ON options.optGroup=prodoptions.poOptionGroup WHERE prodoptions.poProdID IN (" & prodlist & ")"
	rs2.Open sSQL,cnn,0,1
	do while NOT rs2.EOF
		if useStockManagement then response.write "aPS["&rs2("optID")&"]="&rs2("optStock")&";"
		response.write "aPC["&rs2("optID")&"]="&rs2("optPriceDiff")&";"
		rs2.MoveNext
	loop
	rs2.Close
	Session.LCID = saveLCID
end if
%>
//-->
</script><%
End Sub
Sub updatepricescript(doaddprodoptions) %>
<script language="javascript" type="text/javascript">
<!--
function formvalidator<%=Count%>(theForm){
<%
prodoptions=""
hasonepriceoption=false
if doaddprodoptions then
	sSQL = "SELECT poOptionGroup,optType,optFlags FROM prodoptions INNER JOIN optiongroup ON optiongroup.optGrpID=prodoptions.poOptionGroup WHERE poProdID='"&rs("pID")&"' ORDER BY poID"
	rs2.Open sSQL,cnn,0,1
	if NOT rs2.EOF then prodoptions=rs2.getrows
	rs2.Close
	if IsArray(prodoptions) then
		for rowcounter=0 to UBOUND(prodoptions,2)
			if Int(prodoptions(1,rowcounter))=3 then
				response.write "if(theForm.voptn"&rowcounter&".value=='')return(enterValue(theForm.voptn"&rowcounter&"));"&vbCrLf
				response.write "if(theForm.voptn"&rowcounter&".value.length>255)return(dataLimit(theForm.voptn"&rowcounter&"));"&vbCrLf
			elseif abs(prodoptions(1,rowcounter))=2 then
				hasonepriceoption=true
				if Int(prodoptions(1,rowcounter))=2 then response.write "if(theForm.optn"&rowcounter&".selectedIndex==0)return(chooseOption(theForm.optn"&rowcounter&"));"&vbCrLf
				if useStockManagement AND cint(rs("pStockByOpts")) <> 0 then response.write "if(!checkStock(theForm.optn"&rowcounter&",theForm.optn"&rowcounter&".options[theForm.optn"&rowcounter&".selectedIndex].value))return(false);"&vbCrLf
			elseif abs(prodoptions(1,rowcounter))=1 then
				hasonepriceoption=true
				response.write "havefound='';"
				if Int(prodoptions(1,rowcounter))=1 then response.write "for(var i=0; i<theForm.optn"&rowcounter&".length; i++) if(theForm.optn"&rowcounter&"[i].checked)havefound=theForm.optn"&rowcounter&"[i].value;if(havefound=='')return(chooseOption(theForm.optn"&rowcounter&"[0]));"&vbCrLf
				if useStockManagement AND cint(rs("pStockByOpts")) <> 0 then response.write "if(havefound!=''){if(!checkStock(theForm.optn"&rowcounter&"[0],havefound))return(false);}"&vbCrLf
			end if
		next
	end if
end if
if customvalidator<>"" then response.write customvalidator
%>return (true);
}
<%
if noprice<>true AND NOT (rs("pPrice")=0 AND pricezeromessage<>"") AND hasonepriceoption then
	saveLCID = Session.LCID
	Session.LCID = 1033
	response.write "function updateprice"&Count&"(){"&vbCrLf
	response.write "var totAdd=" & rs("pPrice") & ";" & vbCrLf
	response.write "if(!isW3) return;" & vbCrLf
	for rowcounter=0 to UBOUND(prodoptions,2)
		if abs(int(prodoptions(1,rowcounter)))=2 then
			if (prodoptions(2,rowcounter) AND 1) = 1 then
				response.write "totAdd=totAdd+(("&rs("pPrice")&"*pricechecker(document.forms.tForm"&Count&".optn"&rowcounter&".options[document.forms.tForm"&Count&".optn"&rowcounter&".selectedIndex].value))/100.0);"&vbCrLf
			else
				response.write "totAdd=totAdd+pricechecker(document.forms.tForm"&Count&".optn"&rowcounter&".options[document.forms.tForm"&Count&".optn"&rowcounter&".selectedIndex].value);"&vbCrLf
			end if
		elseif abs(int(prodoptions(1,rowcounter)))=1 then
			if (prodoptions(2,rowcounter) AND 1) = 1 then
				response.write "for(var i=0; i<document.forms.tForm"&Count&".optn"&rowcounter&".length; i++) if (document.forms.tForm"&Count&".optn"&rowcounter&"[i].checked) totAdd=totAdd+(("&rs("pPrice")&"*pricechecker(document.forms.tForm"&Count&".optn"&rowcounter&"[i].value))/100.0);"&vbCrLf
			else
				response.write "for(var i=0; i<document.forms.tForm"&Count&".optn"&rowcounter&".length; i++) if (document.forms.tForm"&Count&".optn"&rowcounter&"[i].checked) totAdd=totAdd+pricechecker(document.forms.tForm"&Count&".optn"&rowcounter&"[i].value);"&vbCrLf
			end if
		end if
	next
	if noupdateprice<>TRUE then response.write "document.getElementById('pricediv" & Count & "').innerHTML=formatprice(Math.round(totAdd*100.0)/100.0, '', '');" & vbCrLf
	if showtaxinclusive=true AND (rs("pExemptions") AND 2)<>2 then response.write "document.getElementById('pricedivti" & Count & "').innerHTML=formatprice(Math.round((totAdd+(totAdd*tax/100.0))*100.0)/100.0, '', '');" & vbCrLf
	extracurr = ""
	if currRate1<>0 AND currSymbol1<>"" then extracurr = "+formatprice(Math.round((totAdd*"&currRate1&")*100.0)/100.0, ' " & currSymbol1 & "','" & replace(currFormat1,"'","\'") & "')+'"&replace(currencyseparator,"'","\'")&"'"
	if currRate2<>0 AND currSymbol2<>"" then extracurr = extracurr & "+formatprice(Math.round((totAdd*"&currRate2&")*100.0)/100.0, ' " & currSymbol2 & "','" & replace(currFormat2,"'","\'") & "')+'"&replace(currencyseparator,"'","\'")&"'"
	if currRate3<>0 AND currSymbol3<>"" then extracurr = extracurr & "+formatprice(Math.round((totAdd*"&currRate3&")*100.0)/100.0, ' " & currSymbol3 & "','" & replace(currFormat3,"'","\'") & "');" & vbCrLf
	if extracurr<>"" then response.write "document.getElementById('pricedivec" & Count & "').innerHTML=''" & extracurr & vbCrLf
	Session.LCID = saveLCID
	response.write "}" & vbCrLf
end if
%>//-->
</script><%
End Sub
function checkDPs(currcode)
	if currcode="JPY" then checkDPs=0 else checkDPs=2
end function
Sub checkCurrencyRates(currConvUser,currConvPw,currLastUpdate,byRef currRate1,currSymbol1,byRef currRate2,currSymbol2,byRef currRate3,currSymbol3)
	ccsuccess = true
	if currConvUser<>"" AND currConvPw<>"" AND currLastUpdate < Now()-1 then
		sstr = ""
		if currSymbol1<>"" then sstr = sstr & "&curr=" & currSymbol1
		if currSymbol2<>"" then sstr = sstr & "&curr=" & currSymbol2
		if currSymbol3<>"" then sstr = sstr & "&curr=" & currSymbol3
		if sstr="" then
			cnn.Execute("UPDATE admin SET currLastUpdate="&datedelim&VSUSDate(Now())&datedelim)
			Application.Lock()
			Application("getadminsettings")=""
			Application.UnLock()
			exit sub
		end if
		sstr = "?source=" & countryCurrency & "&user=" & currConvUser & "&pw=" & currConvPw & sstr
		set objHttp = Server.CreateObject("Msxml2.ServerXMLHTTP")
		objHttp.open "POST", "default.asp" & sstr, false
		objHttp.Send "X"
		if (objHttp.status <> 200 ) then
			' HTTP error handling
		else
			Set xmlDoc = objHttp.responseXML
			Set t2 = xmlDoc.getElementsByTagName("currencyRates").Item(0)
			for j = 0 to t2.childNodes.length - 1
				Set n = t2.childNodes.Item(j)
				if n.nodename="currError" then
					response.write n.firstChild.nodeValue
					ccsuccess = false
				elseif n.nodename="selectedCurrency" then
					currRate = 0
					for i = 0 To n.childNodes.length - 1
						Set e = n.childNodes.Item(i)
						if e.nodeName="currSymbol" then
							currSymbol = e.firstChild.nodeValue
						elseif e.nodeName="currRate" then
							currRate = e.firstChild.nodeValue
						end if
					next
					saveLCID = Session.LCID
					Session.LCID = 1033
					if currSymbol1 = currSymbol then
						currRate1 = cDbl(currRate)
						cnn.Execute("UPDATE admin SET currRate1="&currRate&" WHERE adminID=1")
					end if
					if currSymbol2 = currSymbol then
						currRate2 = cDbl(currRate)
						cnn.Execute("UPDATE admin SET currRate2="&currRate&" WHERE adminID=1")
					end if
					if currSymbol3 = currSymbol then
						currRate3 = cDbl(currRate)
						cnn.Execute("UPDATE admin SET currRate3="&currRate&" WHERE adminID=1")
					end if
					Session.LCID = saveLCID
				end if
			next
			if ccsuccess then cnn.Execute("UPDATE admin SET currLastUpdate="&datedelim&VSUSDate(Now())&datedelim)
			Application.Lock()
			Application("getadminsettings")=""
			Application.UnLock()
		end if
		set objHttp = nothing
	end if
End Sub
function IIfVr(theExp,theTrue,theFalse)
if theExp then IIfVr=theTrue else IIfVr=theFalse
end function
function getsectionids(thesecid, delsections)
	secid = thesecid
	iterations = 0
	iteratemore = true
	if Session("clientLoginLevel")<>"" then minloglevel=Session("clientLoginLevel") else minloglevel=0
	if delsections then nodel = "" else nodel = "sectionDisabled<="&minloglevel&" AND "
	do while iteratemore AND iterations<10
		sSQL2 = "SELECT DISTINCT sectionID,rootSection FROM sections WHERE " & nodel & "(topSection IN ("&secid&") OR (sectionID IN ("&secid&") AND rootSection=1))"
		secid = ""
		iteratemore = false
		rs2.Open sSQL2,cnn,0,1
		addcomma = ""
		do while NOT rs2.EOF
			if rs2("rootSection")=0 then iteratemore = true
			secid = secid & addcomma & rs2("sectionID")
			addcomma = ","
			rs2.MoveNext
		loop
		rs2.Close
		iterations = iterations + 1
	loop
	if secid="" then getsectionids = "0" else getsectionids = secid
end function
if Trim(Session("clientUser"))="" then
	clientUser = Trim(Replace(Request.Cookies("WRITECLL"),"'",""))
	if clientUser<>"" then
		Set clientRS = Server.CreateObject("ADODB.RecordSet")
		Set clientCnn=Server.CreateObject("ADODB.Connection")
		clientCnn.open sDSN
		sSQL = "SELECT clientUser,clientActions,clientLoginLevel,clientPercentDiscount FROM clientlogin WHERE clientUser='"&clientUser&"' AND clientPW='"&Trim(Replace(Request.Cookies("WRITECLP"),"'",""))&"'"
		clientRS.Open sSQL,clientCnn,0,1
		if NOT clientRS.EOF then
			Session("clientUser")=clientRS("clientUser")
			Session("clientActions")=clientRS("clientActions")
			Session("clientLoginLevel")=clientRS("clientLoginLevel")
			Session("clientPercentDiscount")=(100.0-cDbl(clientRS("clientPercentDiscount")))/100.0
		end if
		clientRS.Close
		clientCnn.Close
		set clientRS = nothing
		set clientCnn = nothing
	end if
end if
function callxmlfunction(cfurl, cfxml, byref res, cfcert, cxfobj, byref cferr, settimeouts)
	set objHttp = Server.CreateObject(cxfobj)
	if settimeouts then objHttp.setTimeouts 30000, 30000, 0, 0
	objHttp.open "POST", cfurl, false
	objHttp.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
	' if cfcert<>"" then objHttp.setOption 3, "LOCAL_MACHINE\My\" & cfcert
	if cfcert<>"" then objHttp.SetClientCertificate("LOCAL_MACHINE\My\" & cfcert)
	' response.write Replace(Replace(cfxml,"</","&lt;/"),"<","<br />&lt;")&"<HR>"
	err.number=0
	objHttp.Send cfxml
	If err.number <> 0 OR objHttp.status <> 200 Then
		cferr = "Error, couldn't connect to server " & err.number
		callxmlfunction = FALSE
	Else
		res = objHttp.responseText
		callxmlfunction = TRUE
		' response.write Replace(Replace(objHttp.responseText,"</","&lt;/"),"<","<br />&lt;")&"<HR>"
	End If
	set objHttp = nothing
end function
function getpayprovdetails(ppid,ppdata1,ppdata2,ppdata3,ppdemo,ppmethod)
	sSQL = "SELECT payProvData1,payProvData2,payProvData3,payProvDemo,payProvMethod FROM payprovider WHERE payProvEnabled=1 AND payProvID=" & replace(ppid, "'", "")
	rs.Open sSQL,cnn,0,1
	if NOT rs.EOF then
		ppdata1 = trim(rs("payProvData1")&"")
		ppdata2 = trim(rs("payProvData2")&"")
		ppdata3 = trim(rs("payProvData3")&"")
		ppdemo=(cint(rs("payProvDemo"))=1)
		ppmethod=Int(rs("payProvMethod"))
		getpayprovdetails = TRUE
	else
		getpayprovdetails = FALSE
	end if
	rs.Close
end function
sub writehiddenvar(hvname,hvval)
response.write "<input type=""hidden"" name=""" & hvname & """ value=""" & replace(hvval,"""","&quot;") & """ />" & vbCrLf
end sub
function ppsoapheader(username, password, threetokenhash)
ppsoapheader = "<?xml version=""1.0"" encoding=""utf-8""?><soap:Envelope xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema""><soap:Header><RequesterCredentials xmlns=""urn:ebay:api:PayPalAPI""><Credentials xmlns=""urn:ebay:apis:eBLBaseComponents""><Username>" & username & "</Username><Password>" & password & "</Password>" & IIfVr(threetokenhash<>"","<Signature>"&threetokenhash&"</Signature>","") & "</Credentials></RequesterCredentials></soap:Header>"
end function
function displayproductoptions(grpnmstyle,grpnmstyleend, byRef optpricediff)
	optionshtml = ""
	optpricediff = 0
	pricediff = 0
	for rowcounter=0 to UBOUND(prodoptions,2)
		opthasstock = false
		sSQL="SELECT optID,"&getlangid("optName",32)&","&getlangid("optGrpName",16)&","&OWSP&"optPriceDiff,optType,optGrpSelect,optFlags,optStock,optPriceDiff AS optDims,optDefault FROM options INNER JOIN optiongroup ON options.optGroup=optiongroup.optGrpID WHERE optGroup="&prodoptions(0,rowcounter)&" ORDER BY optID"
		rs2.Open sSQL,cnn,0,1
		if NOT rs2.EOF then
			if abs(int(rs2("optType")))=3 then
				opthasstock=true
				fieldHeight = cInt((cDbl(rs2("optDims"))-Int(rs2("optDims")))*100.0)
				optionshtml = optionshtml & "<tr><td align='right' width='30%'>"&grpnmstyle&rs2(getlangid("optGrpName",16))&":"&grpnmstyleend&"</td><td align=""left""> <input type='hidden' name='optn"&rowcounter&"' value='"&rs2("optID")&"' />"
				if fieldHeight<>1 then
					optionshtml = optionshtml & "<textarea wrap='virtual' name='voptn"&rowcounter&"' cols='"&int(rs2("optDims"))&"' rows='"&fieldHeight&"'>"
					optionshtml = optionshtml & rs2(getlangid("optName",32))&"</textarea>"
				else
					optionshtml = optionshtml & "<input maxlength='255' type='text' name='voptn"&rowcounter&"' size='"&int(rs2("optDims"))&"' value="""&Replace(rs2(getlangid("optName",32)),"""","&quot;")&""" />"
				end if
				optionshtml = optionshtml & "</td></tr>"
			elseif abs(int(rs2("optType")))=1 then
				optionshtml = optionshtml & "<tr><td align='right' valign=""baseline"" width='30%'>"&grpnmstyle&rs2(getlangid("optGrpName",16))&":"&grpnmstyleend&"</td><td align=""left""> "
				do while not rs2.EOF
					optionshtml = optionshtml & "<input type=""radio"" style=""vertical-align:middle"" onclick="""&IIfVr((rs("pPrice")=0 AND pricezeromessage<>"") OR noprice=true,"dummyfunc","updateprice"&Count)&"();"" name=""optn"&rowcounter&""" "
					if cint(rs2("optDefault"))<>0 then optionshtml = optionshtml & "checked "
					optionshtml = optionshtml & "value='"&rs2("optID")&"' /><span "
					if useStockManagement AND cint(rs("pStockByOpts"))<>0 AND rs2("optStock") <= 0 then optionshtml = optionshtml & " class=""oostock"" " else opthasstock=true
					optionshtml = optionshtml & ">"&rs2(getlangid("optName",32))
					if hideoptpricediffs<>true AND cDbl(rs2("optPriceDiff"))<>0 then
						optionshtml = optionshtml & " ("
						if cDbl(rs2("optPriceDiff")) > 0 then optionshtml = optionshtml & "+"
						if (rs2("optFlags") AND 1) = 1 then pricediff = (rs("pPrice")*rs2("optPriceDiff"))/100.0 else pricediff = rs2("optPriceDiff")
						optionshtml = optionshtml & FormatEuroCurrency(pricediff)&")"
						if rs2("optDefault")<>0 then optpricediff = optpricediff + pricediff
					end if
					if useStockManagement AND showinstock=TRUE AND noshowoptionsinstock<>TRUE AND cint(rs("pStockByOpts"))<>0 then optionshtml = optionshtml & replace(xxOpSkTx, "%s", rs2("optStock"))
					optionshtml = optionshtml & "</span>"
					if (rs2("optFlags") AND 4) <> 4 then optionshtml = optionshtml & "<br />"&vbCrLf
					rs2.MoveNext
				loop
				optionshtml = optionshtml & "</td></tr>"
			else
				optionshtml = optionshtml & "<tr><td align='right' width='30%'>"&grpnmstyle&rs2(getlangid("optGrpName",16))&":"&grpnmstyleend&"</td><td align=""left""> <select class=""prodoption"" onchange="""&IIfVr((rs("pPrice")=0 AND pricezeromessage<>"") OR noprice=true,"dummyfunc","updateprice"&Count)&"();"" name=""optn"&rowcounter&""" size=""1"">"
				gotdefaultdiff = FALSE
				firstpricediff = 0
				if cint(rs2("optGrpSelect"))<>0 then
					optionshtml = optionshtml & "<option value=''>"&xxPlsSel&"</option>"
				else
					if (rs2("optFlags") AND 1) = 1 then firstpricediff = (rs("pPrice")*rs2("optPriceDiff"))/100.0 else firstpricediff = rs2("optPriceDiff")
				end if
				do while not rs2.EOF
					optionshtml = optionshtml & "<option "
					if useStockManagement AND cint(rs("pStockByOpts"))<>0 AND rs2("optStock") <= 0 then optionshtml = optionshtml & "class=""oostock"" " else opthasstock=true
					if cint(rs2("optDefault"))<>0 then optionshtml = optionshtml & "selected "
					optionshtml = optionshtml & "value='"&rs2("optID")&"'>"&rs2(getlangid("optName",32))
					if hideoptpricediffs<>true AND cDbl(rs2("optPriceDiff"))<>0 then
						optionshtml = optionshtml & " ("
						if cDbl(rs2("optPriceDiff")) > 0 then optionshtml = optionshtml & "+"
						if (rs2("optFlags") AND 1) = 1 then pricediff = (rs("pPrice")*rs2("optPriceDiff"))/100.0 else pricediff = rs2("optPriceDiff")
						optionshtml = optionshtml & FormatEuroCurrency(pricediff)&")"
						if rs2("optDefault")<>0 then optpricediff = optpricediff + pricediff : gotdefaultdiff=TRUE
					end if
					if useStockManagement AND showinstock=TRUE AND noshowoptionsinstock<>TRUE AND cint(rs("pStockByOpts"))<>0 then optionshtml = optionshtml & replace(xxOpSkTx, "%s", rs2("optStock"))
					optionshtml = optionshtml & "</option>"&vbCrLf
					rs2.MoveNext
				loop
				if NOT gotdefaultdiff then optpricediff = optpricediff + firstpricediff
				optionshtml = optionshtml & "</select></td></tr>"
			end if
		end if
		rs2.Close
		optionshavestock = (optionshavestock AND opthasstock)
	next
	displayproductoptions = optionshtml
end function
if enableclientlogin=true then
	if Session("clientUser")<>"" then
	elseif request.form("checktmplogin")="1" AND request.form("sessionid")<>"" then
		Set clientRS = Server.CreateObject("ADODB.RecordSet")
		Set clientCnn=Server.CreateObject("ADODB.Connection")
		clientCnn.open sDSN
		sSQL = "SELECT tmploginname FROM tmplogin WHERE tmploginid=" & request.form("sessionid")
		clientRS.Open sSQL,clientCnn,0,1
		if NOT clientRS.EOF then
			Session("clientUser")=clientRS("tmploginname")
			clientRS.Close
			clientCnn.Execute("DELETE FROM tmplogin WHERE tmploginid=" & request.form("sessionid"))
			sSQL = "SELECT clientActions,clientLoginLevel,clientPercentDiscount FROM clientlogin WHERE clientUser='"&replace(session("clientUser"),"'","")&"'"
			clientRS.Open sSQL,clientCnn,0,1
			if NOT clientRS.EOF then
				Session("clientActions")=clientRS("clientActions")
				Session("clientLoginLevel")=clientRS("clientLoginLevel")
				Session("clientPercentDiscount")=(100.0-cDbl(clientRS("clientPercentDiscount")))/100.0
			end if
		end if
		clientRS.Close
		clientCnn.Close
		set clientRS = nothing
		set clientCnn = nothing
	elseif Request.Cookies("WRITECLL")<>"" then
		Set clientRS = Server.CreateObject("ADODB.RecordSet")
		Set clientCnn=Server.CreateObject("ADODB.Connection")
		clientCnn.open sDSN
		sSQL = "SELECT clientUser,clientActions,clientLoginLevel FROM clientlogin WHERE clientUser='"&replace(Request.Cookies("WRITECLL"),"'","")&"' AND clientPW='"&replace(Request.Cookies("WRITECLP"),"'","")&"'"
		clientRS.Open sSQL,clientCnn,0,1
		if NOT clientRS.EOF then
			Session("clientUser")=clientRS("clientUser")
			Session("clientActions")=clientRS("clientActions")
			Session("clientLoginLevel")=clientRS("clientLoginLevel")
		end if
		clientRS.Close
		clientCnn.Close
		set clientRS = nothing
		set clientCnn = nothing
	end if
	if requiredloginlevel<>"" then
		if Session("clientLoginLevel")<requiredloginlevel then
			if Int(requiredloginlevel)>Session("clientLoginLevel") then Response.redirect "clientlogin.asp?refurl=" & server.urlencode(request.servervariables("URL") & IIfVr(request.servervariables("QUERY_STRING")<>"" ,"?"&request.servervariables("QUERY_STRING"), ""))
		end if
	end if
end if
function urldecode(encodedstring)
	strIn  = encodedstring : strOut = "" : intPos = Instr(strIn, "+")
	do While intPos
		strLeft = "" : strRight = ""
		if intPos > 1 then strLeft = Left(strIn, intPos - 1)
		if intPos < len(strIn) then strRight = Mid(strIn, intPos + 1)
		strIn = strLeft & " " & strRight
		intPos = InStr(strIn, "+")
		intLoop = intLoop + 1
	Loop
	intPos = InStr(strIn, "%")
	do while intPos AND Len(strIn)-intPos > 2
		if intPos > 1 then strOut = strOut & Left(strIn, intPos - 1)
		strOut = strOut & Chr(CInt("&H" & mid(strIn, intPos + 1, 2)))
		if intPos > (len(strIn) - 3) then strIn = "" else strIn = Mid(strIn, intPos + 3)
		intPos = InStr(strIn, "%")
	Loop
	urldecode = strOut & strIn
end function
function vrmax(a,b)
	if a > b then vrmax=a else vrmax=b
end function
function vrmin(a,b)
	if a < b then vrmin=a else vrmin=b
end function
%>
<SCRIPT LANGUAGE=JScript RUNAT=SERVER>
function vrbase64_encrypt(origstr){
	var tcharset = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/";
	var str = "";
	for(var i = 0; i < origstr.length; i += 3){
		triplet = (origstr.charCodeAt(i) << 16) | (origstr.charCodeAt(i+1) << 8) | (origstr.charCodeAt(i+2) << 0)
		for(var j = 0; j < 4; j++){
			if(i + j > origstr.length) str += "="; else str += tcharset.charAt((triplet >> 6*(3-j)) & 0x3F);
		}
	}
	return str;
}
function vsround(amnt, decpl){
	return(Math.round(amnt * Math.pow(10,decpl),decpl) / Math.pow(10,decpl));
}
function long2ip(ip2lip){
	retval = "here";
	//retval = int(ip2lip >> 24) + "." // + (int(ip2lip / 65536) AND 255) & "." & (int(ip2lip / 256) AND 255) & "." & (ip2lip AND 255)
	retval = ((ip2lip >> 24) & 255) + "." + ((ip2lip >> 16) & 255) + "." + ((ip2lip >> 8) & 255) + "." + (ip2lip & 255);
	return(retval);
}
</SCRIPT>