<%
prodoptions=""
' id,name,discounts,listprice,price,priceinctax,options,quantity,currency,instock,buy
if cpdcolumns="" then cpdcolumns="id,name,discounts,listprice,price,priceinctax,instock,quantity,buy"
cpdarray=split(lcase(cpdcolumns),",")
noproductoptions=TRUE
showtaxinclusive=FALSE
hascurrency=FALSE
noupdateprice=TRUE
for cpdindex=0 to UBOUND(cpdarray)
	select case cpdarray(cpdindex)
	case "options"
		noproductoptions=FALSE
	case "price"
		noupdateprice=FALSE
	case "priceinctax"
		showtaxinclusive=TRUE
	case "currency"
		hascurrency=TRUE
	end select
next
saveLCID = Session.LCID
productdisplayscript(noproductoptions<>true) %>
			<table width="<%=innertablewidth%>" border="0" cellspacing="<%=innertablespacing%>" cellpadding="<%=innertablepadding%>" bgcolor="<%=innertablebg%>">
<%	if IsEmpty(showcategories) OR showcategories=true then %>
			  <tr>
				<td class="prodnavigation" colspan="2" align="left"><strong><p class="prodnavigation"><% response.write tslist %></p></strong></td>
				<td align="right"><% if nobuyorcheckout<>true then %><a href="cart.asp"><img src="images/checkout.gif" border="0" alt="<%=xxCOTxt%>" /></a><% else response.write "&nbsp;" end if %></td>
			  </tr>
<%	end if
if nowholesalediscounts=true AND Session("clientUser")<>"" then
	if ((Session("clientActions") AND 8) = 8) OR ((Session("clientActions") AND 16) = 16) then noshowdiscounts=true
end if
if noshowdiscounts<>true then
	Session.LCID = 1033
	sSQL = "SELECT DISTINCT "&getlangid("cpnName",1024)&" FROM coupons LEFT OUTER JOIN cpnassign ON coupons.cpnID=cpnassign.cpaCpnID WHERE ("
	addor = ""
	if catid<>"0" then
		sSQL = sSQL & addor & "((cpnSitewide=0 OR cpnSitewide=3) AND cpaType=1 AND cpaAssignment IN ('"&Replace(topsectionids,",","','")&"'))"
		addor = " OR "
	end if
	tdt = Date()
	sSQL = sSQL & addor & "(cpnSitewide=1 OR cpnSitewide=2)) AND cpnNumAvail>0 AND cpnEndDate>="&datedelim&VSUSDate(tdt)&datedelim&" AND cpnIsCoupon=0"
	Session.LCID = saveLCID
	rs2.Open sSQL,cnn,0,1
	if NOT rs2.EOF then %>
			  <tr>
				<td align="left" colspan="3">
				  <p><strong><%=xxDsProd%></strong><br /><font color="#FF0000" size="1">
				  <%	do while NOT rs2.EOF
							response.write rs2(getlangid("cpnName",1024)) & "<br />"
							rs2.MoveNext
						loop %></font></p>
				</td>
			  </tr>
<%	end if
	rs2.Close
end if
%>
			  <tr>
				<td colspan="3" align="center" class="pagenums"><p class="pagenums"><%
					If iNumOfPages > 1 AND pagebarattop=1 Then Response.Write writepagebar(CurPage, iNumOfPages) & "<br />" %>
				  <img src="images/clearpixel.gif" width="300" height="8" alt="" /></p></td>
			  </tr>
<%	if rs.EOF then
		response.write "<tr><td colspan=""3"" align=""center""><p>" & xxNoPrds & "</p></td></tr>"
	else
	response.write "<tr><td colspan=""3""><table class=""cpd"" width="""&maintablewidth&""" border=""0"" bordercolor=""#B1B1B1"" cellspacing=""1"" cellpadding=""3"" bgcolor=""#B1B1B1"">"
	if cpdheaders<>"" then
		cpdheadarray=split(cpdheaders,",")
		response.write "<tr>"
		for cpdindex=0 to UBOUND(cpdheadarray)
			if cpdindex<=UBOUND(cpdarray) then classid=cpdarray(cpdindex) else classid=""
			response.write "<td class=""cpdhl"" bgcolor=""#EBEBEB""><div class=""cpdhl"&classid&""">"&cpdheadarray(cpdindex)&"</div></td>"
		next
		response.write "</tr>"
	end if
	if NOT hascurrency then currSymbol1="" : currSymbol2="" : currSymbol3=""
	Do While Not rs.EOF And Count < rs.PageSize
		if forcedetailslink=TRUE OR Trim(rs(getlangid("pLongDescription",4)))<>"" OR NOT (Trim(rs("pLargeImage")&"")="" OR Trim(rs("pLargeImage"))="prodimages/") then
			if cint(rs("pStaticPage"))<>0 then
				startlink="<a href='"&cleanforurl(rs("pName"))&".asp"&IIfVr(catid<>"" AND catid<>"0" AND int(catid)<>rs("pSection") AND nocatid<>TRUE,"?cat="&catid,"")&"'>"
				endlink="</a>"
			elseif detailslink<>"" then
				startlink=replace(replace(detailslink,"%largeimage%", rs("pLargeImage")),"%pid%", rs("pId"))
				endlink=detailsendlink
			else
				startlink="<a href='proddetail.asp?prod="&Server.URLEncode(rs("pId"))&IIfVr(catid<>"" AND catid<>"0" AND int(catid)<>rs("pSection") AND nocatid<>TRUE,"&amp;cat="&catid,"")&"'>"
				endlink="</a>"
			end if
		else
			startlink=""
			endlink=""
		end if
		Session.LCID = 1033
		if NOT isrootsection then
			if IsNull(rs("pSection")) then thetopts = 0 else thetopts = rs("pSection")
			gotdiscsection = false
			for cpnindex=0 to adminProdsPerPage-1
				if aDiscSection(0,cpnindex)=thetopts then
					gotdiscsection = true
					exit for
				elseif aDiscSection(0,cpnindex)="" then
					exit for
				end if
			next
			aDiscSection(0,cpnindex) = thetopts
			if NOT gotdiscsection then
				topcpnids = thetopts
				for index=0 to 10
					if thetopts=0 then
						exit for
					else
						sSQL = "SELECT topSection FROM sections WHERE sectionID=" & thetopts
						rs2.Open sSQL,cnn,0,1
						if NOT rs2.EOF then
							thetopts = rs2("topSection")
							topcpnids = topcpnids & "," & thetopts
						else
							rs2.Close
							exit for
						end if
						rs2.Close
					end if
				next
				aDiscSection(1,cpnindex) = topcpnids
			else
				topcpnids = aDiscSection(1,cpnindex)
			end if
		end if
		alldiscounts = ""
		if noshowdiscounts<>true then
			tdt = Date()
			sSQL = "SELECT DISTINCT "&getlangid("cpnName",1024)&" FROM coupons LEFT OUTER JOIN cpnassign ON coupons.cpnID=cpnassign.cpaCpnID WHERE (cpnSitewide=0 OR cpnSitewide=3) AND cpnNumAvail>0 AND cpnEndDate>="&datedelim&VSUSDate(tdt)&datedelim&" AND cpnIsCoupon=0 AND ((cpaType=2 AND cpaAssignment='"&rs("pID")&"')"
			if NOT isrootsection then sSQL = sSQL & " OR (cpaType=1 AND cpaAssignment IN ('"&Replace(topcpnids,",","','")&"') AND NOT cpaAssignment IN ('"&Replace(topsectionids,",","','")&"'))"
			sSQL = sSQL & ")"
			rs2.Open sSQL,cnn,0,1
			do while NOT rs2.EOF
				alldiscounts = alldiscounts & rs2(getlangid("cpnName",1024)) & "<br />"
				rs2.MoveNext
			loop
			rs2.Close
		end if
		Session.LCID = saveLCID
		optionshavestock=true
		if currencyseparator="" then currencyseparator=" "
		response.write "<form method=""post"" name=""tForm"&Count&""" action=""cart.asp"" onsubmit=""return formvalidator"&Count&"(this)""><tr class=""cpdtr"" bgcolor=""#EBEBEB"">"
		updatepricescript(noproductoptions<>true)
		totprice = rs("pPrice")
		if noproductoptions=FALSE then
			if IsArray(prodoptions) then
				optionshtml = displayproductoptions("<strong><span class=""prodoption"">","</span></strong>",optdiff)
				totprice = totprice + optdiff
			else
				optionshtml = ""
			end if
		end if
		for cpdindex=0 to UBOUND(cpdarray)
			select case cpdarray(cpdindex)
			case "id" %>
			<td class="cpdll" bgcolor="#FFFFFF"><div class="prod3id"><%=startlink & rs("pID") & endlink %></div></td>
<%			case "name" %>
			<td class="cpdll" bgcolor="#FFFFFF"><div class="prod3name"><%=rs(getlangid("pName",1)) %></div></td>
<%			case "description" %>
			<td class="cpdll" bgcolor="#FFFFFF"><div class="prod3description"><%
				shortdesc = rs(getlangid("pDescription",2))
				if shortdescriptionlimit="" then response.write shortdesc else response.write left(shortdesc, shortdescriptionlimit) & IIfVr(len(shortdesc)>shortdescriptionlimit, "...", "") %></div></td>
<%			case "image" %>
			<td class="cpdll" bgcolor="#FFFFFF"><% if Trim(rs("pImage"))="" or IsNull(rs("pImage")) or Trim(rs("pImage"))="prodimages/" then response.write "&nbsp;" else response.write startlink & "<img class=""prod3image"" src="""&rs("pImage")&""" border=""0"" alt="""&strip_tags2(rs(getlangid("pName",1))&"")&""" />"&endlink %></td>
<%			case "discounts" %>
			<td class="cpdll" bgcolor="#FFFFFF"><div class="prod3discounts"><% if alldiscounts<>"" then response.write alldiscounts else response.write "&nbsp;" %></div></td>
<%			case "details" %>
			<td class="cpdll" bgcolor="#FFFFFF"><div class="prod3details"><% if startlink <> "" then response.write startlink & "<strong>"&xxPrDets&"</strong></a>&nbsp;" else response.write "&nbsp;" %></div></td>
<%			case "options" %>
			<td class="cpdll" bgcolor="#FFFFFF">
<%
if IsArray(prodoptions) then
	response.write "<div class=""prod3options""><table border='0' cellspacing='1' cellpadding='1' width='100%'>"
	response.write optionshtml & "</table></div>"
else
	response.write "&nbsp;"
end if
%>
                </td>
<%			case "listprice" %>
			<td class="cpdll" bgcolor="#FFFFFF"><div class="prod3listprice"><% if cDbl(rs("pListPrice"))<>0.0 then response.write FormatEuroCurrency(rs("pListPrice")) else response.write "&nbsp;" %></div></td>
<%			case "price" %>
			<td class="cpdll" bgcolor="#FFFFFF"><div class="prod3price"><%	if cDbl(totprice)=0 AND pricezeromessage<>"" then
							response.write pricezeromessage
						else
							response.write "<span class=""price"" id=""pricediv" & Count & """>" & FormatEuroCurrency(totprice) & "</span>"
						end if %></div></td>
<%			case "priceinctax" %>
			<td class="cpdll" bgcolor="#FFFFFF"><div class="prod3pricetaxinc"><%	if cDbl(totprice)=0 AND pricezeromessage<>"" then
							response.write pricezeromessage
						else
							response.write "<span class=""price"" id=""pricedivti" & Count & """>"
							if (rs("pExemptions") AND 2)=2 then response.write FormatEuroCurrency(totprice) else response.write FormatEuroCurrency(totprice+(totprice*countryTaxRate/100.0))
							response.write "</span>"
						end if %></div></td>
<%			case "currency" %>
			<td class="cpdll" bgcolor="#FFFFFF"><%	if cDbl(totprice)=0 AND pricezeromessage<>"" then
							response.write "&nbsp;"
						else
							extracurr = ""
							if currRate1<>0 AND currSymbol1<>"" then extracurr = replace(currFormat1, "%s", FormatNumber(totprice*currRate1, checkDPs(currSymbol1))) & currencyseparator
							if currRate2<>0 AND currSymbol2<>"" then extracurr = extracurr & replace(currFormat2, "%s", FormatNumber(totprice*currRate2, checkDPs(currSymbol2))) & currencyseparator
							if currRate3<>0 AND currSymbol3<>"" then extracurr = extracurr & replace(currFormat3, "%s", FormatNumber(totprice*currRate3, checkDPs(currSymbol3)))
							if extracurr<>"" then response.write "<div class=""prod3currency""><span class=""extracurr"" id=""pricedivec" & Count & """>" & extracurr & "</span></div>"
						end if %></td>
<%			case "quantity" %>
			<td class="cpdll" bgcolor="#FFFFFF"><div class="prod3quant"><input type="text" name="quant" size="2" maxlength="6" value="1" /></div></td>
<%			case "instock" %>
			<td class="cpdll" bgcolor="#FFFFFF"><div class="prod3instock"><% if cint(rs("pStockByOpts"))<>0 then response.write "-" else response.write rs("pInStock") %></div></td>
<%			case "buy" %>
			<td class="cpdll" bgcolor="#FFFFFF"><div class="prod3buy"><%
	if useStockManagement then
		if cint(rs("pStockByOpts"))<>0 then isInStock = optionshavestock else isInStock = Int(rs("pInStock")) > 0
	else
		isInStock = cint(rs("pSell")) <> 0
	end if
	if isInStock then
%><input type="hidden" name="id" value="<%=rs("pID")%>" />
<input type="hidden" name="mode" value="add" />
<input type="hidden" name="frompage" value="<%=Request.ServerVariables("URL")&IIfVr(Trim(Request.ServerVariables("QUERY_STRING"))<>"","?","")&Request.ServerVariables("QUERY_STRING")%>" />
<%	if custombuybutton<>"" then response.write custombuybutton else response.write "<input align=""middle"" type=""image"" src=""images/buy.gif"" alt="""&xxAddToC&""" />"
	else
		response.write "<strong>"&xxOutStok&"</strong>"
	end if %></div></td>
<%			end select
		next
		response.write "</tr></form>"
		Count = Count + 1
		rs.MoveNext
	loop
	response.write "</table></td></tr>"
	end if
%>			  <tr>
				<td colspan="3" align="center" class="pagenums"><p class="pagenums"><%
					If iNumOfPages > 1 AND nobottompagebar<>true Then Response.Write writepagebar(CurPage, iNumOfPages) %><br />
				<img src="images/clearpixel.gif" width="300" height="1" alt="" /></p></td>
			  </tr>
			</table>