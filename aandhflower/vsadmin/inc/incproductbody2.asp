<%
	prodoptions=""
	saveLCID = Session.LCID
	productdisplayscript(noproductoptions<>true) %>
			<table class="<%=cs%>products" width="<%=innertablewidth%>" border="0" cellspacing="<%=innertablespacing%>" cellpadding="<%=innertablepadding%>" bgcolor="<%=innertablebg%>">
<%	if productcolumns="" then productcolumns=1
	if IsEmpty(showcategories) OR showcategories=true then %>
			  <tr>
				<td colspan="<%=productcolumns%>">
				  <table width="100%" border="0" cellspacing="0" cellpadding="0">
					<tr>
					  <td class="prodnavigation" align="left"><strong><p class="prodnavigation"><% response.write tslist %></p></strong></td>
					  <td align="right"><% if nobuyorcheckout<>true then %><a href="cart.asp"><img src="images/checkout.gif" border="0" alt="<%=xxCOTxt%>" /></a><% else response.write "&nbsp;" end if %></td>
					</tr>
				  </table>
				</td>
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
	if NOT rs2.EOF then
%>
			  <tr>
				<td align="left" colspan="<%=productcolumns%>">
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
		If iNumOfPages > 1 AND pagebarattop=1 Then 
%>
			  <tr>
				<td colspan="<%=productcolumns%>" align="center" class="pagenums"><p class="pagenums"><%
					Response.Write writepagebar(CurPage, iNumOfPages) & "<br />" %><img src="images/clearpixel.gif" width="300" height="5" alt="" /></p></td>
			  </tr>
<%
		end if
	if rs.EOF then
		response.write "<tr><td colspan=""3"" align=""center""><p>" & xxNoPrds & "</p></td></tr>"
	else
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
		if Count MOD productcolumns = 0 then response.write "<tr>" %>
				<td width="<%=Int(100 / productcolumns)%>%" align="center" valign="top" class="<%=cs%>product"><div class="<%=cs%>product">
<%	if currencyseparator="" then currencyseparator=" "
	updatepricescript(noproductoptions<>true)
	thedesc = trim(rs(getlangid("pDescription",2)))
	if shortdescriptionlimit<>"" then thedesc = left(thedesc, shortdescriptionlimit) & IIfVr(len(thedesc)>shortdescriptionlimit, "...", "")
%>				<form method="post" name="tForm<%=Count%>" action="cart.asp" style="margin: 0;padding: 0;" onsubmit="return formvalidator<%=Count%>(this)">
				  <table width="100%" border="0" cellspacing="4" cellpadding="4">
			  <% if showproductid=true then response.write "<tr><td><div class="""&cs&"prodid""><strong>" & xxPrId & ":</strong> " & rs("pID") & "</div></td></tr>" %>
				    <tr>
					  <td width="100%" align="center" class="<%=cs%>prodimage">
						<% if Trim(rs("pImage"))="" or IsNull(rs("pImage")) or Trim(rs("pImage"))="prodimages/" then %>
						  &nbsp;
						<% else %>
						  <%=startlink%><img class="<%=cs%>prodimage" src="<%=rs("pImage")%>" border="0" alt="<%=strip_tags2(rs(getlangid("pName",1))&"")%>" /><%=endlink%>
						<% end if %>
					  </td>
					</tr>
					<tr>
					  <td width="100%">
					    <strong><div class="<%=cs%>prodname"><%=startlink & rs(getlangid("pName",1)) & endlink & xxDot%></div></strong><%
						if alldiscounts<>"" then response.write "<font color=""#FF0000""><strong><span class=""discountsapply"">"&xxDsApp&"</span></strong><br /><font size=""1""><div class="""&cs&"proddiscounts"">" & alldiscounts & "</div></font></font>"
						if showinstock=TRUE then if cint(rs("pStockByOpts"))=0 then response.write "<div class="""&cs&"prodinstock""><strong>" & xxInStoc & ":</strong> " & rs("pInStock") & "</div>" %>
<%	if thedesc<>"" then response.write "<div class="""&cs&"proddescription"">" & thedesc & "</div>" else response.write "<br />"
optionshavestock=true
totprice = rs("pPrice")
if IsArray(prodoptions) AND noproductoptions<>true then
	response.write "<div class="""&cs&"prodoptions""><table border='0' cellspacing='1' cellpadding='1' width='100%'>"
	response.write displayproductoptions("<strong><span class=""prodoption"">","</span></strong>", optdiff)
	totprice = totprice + optdiff
	response.write "</table></div>"
end if
					if noprice<>true then
						if cDbl(rs("pListPrice"))<>0.0 then response.write "<div class="""&cs&"listprice"">" & Replace(xxListPrice, "%s", FormatEuroCurrency(rs("pListPrice"))) & "</div>"
						if totprice=0 AND pricezeromessage<>"" then
							response.write "<div class="""&cs&"prodprice"">" & pricezeromessage & "</div>"
						else
							response.write "<div class="""&cs&"prodprice""><strong>" & xxPrice & ":</strong> <span class=""price"" id=""pricediv" & Count & """>" & FormatEuroCurrency(totprice) & "</span> "
							if showtaxinclusive=true AND (rs("pExemptions") AND 2)<>2 then response.write Replace(ssIncTax,"%s", "<span id=""pricedivti" & Count & """>" & FormatEuroCurrency(totprice+(totprice*countryTaxRate/100.0)) & "</span> ")
							response.write "</div>"
							extracurr = ""
							if currRate1<>0 AND currSymbol1<>"" then extracurr = replace(currFormat1, "%s", FormatNumber(totprice*currRate1, checkDPs(currSymbol1))) & currencyseparator
							if currRate2<>0 AND currSymbol2<>"" then extracurr = extracurr & replace(currFormat2, "%s", FormatNumber(totprice*currRate2, checkDPs(currSymbol2))) & currencyseparator
							if currRate3<>0 AND currSymbol3<>"" then extracurr = extracurr & replace(currFormat3, "%s", FormatNumber(totprice*currRate3, checkDPs(currSymbol3)))
							if extracurr<>"" then response.write "<div class="""&cs&"prodcurrency""><span class=""extracurr"" id=""pricedivec" & Count & """>" & extracurr & "</span></div>"
						end if
					end if %>
					  </td>
					</tr><%
if nobuyorcheckout<>true then
	response.write "<tr><td align=""center"">"
	if useStockManagement then
		if cint(rs("pStockByOpts"))<>0 then isInStock = optionshavestock else isInStock = Int(rs("pInStock")) > 0
	else
		isInStock = cint(rs("pSell")) <> 0
	end if
	if isInStock then
%>
<input type="hidden" name="id" value="<%=rs("pID")%>" />
<input type="hidden" name="mode" value="add" />
<input type="hidden" name="frompage" value="<%=Request.ServerVariables("URL")&IIfVr(Trim(Request.ServerVariables("QUERY_STRING"))<>"","?","")&Request.ServerVariables("QUERY_STRING")%>" />
<%		if showquantonproduct=true then response.write "<input type=""text"" name=""quant"" size=""2"" maxlength=""5"" value=""1"" />&nbsp;"
		if custombuybutton<>"" then response.write custombuybutton else response.write "<input align=""middle"" type=""image"" src=""images/buy.gif"" alt="""&xxAddToC&""" />"
	else
		response.write "<strong>"&xxOutStok&"</strong>"
	end if
	response.write "</td></tr>"
end if%>		  </table>
				  </form></div>
				</td><%
		Count = Count + 1
		rs.MoveNext
		if Count MOD productcolumns = 0 then
			response.write "</tr>"
			if noproductseparator<>TRUE then
				if Not rs.EOF And Count < rs.PageSize then
					response.write "<tr>"
					for index=1 to productcolumns
						response.write "<td class=""prodseparator"">" & IIfVr(prodseparator<>"", prodseparator, "<hr class=""prodseparator"" width=""70%"" align=""center"" />") & "</td>"
					next
					response.write "</tr>"
				end if
			end if
		end if
	loop
	if Count MOD productcolumns <> 0 then
		do while Count MOD productcolumns <> 0
			response.write "<td class="""&cs&"noproduct"" width="""&Int(100 / productcolumns)&"%"" align=""center"">&nbsp;</td>"
			Count = Count + 1
		loop
		response.write "</tr>"
	end if
	end if
%>			  <tr>
				<td colspan="<%=productcolumns%>" align="center" class="pagenums"><p class="pagenums"><%
					If iNumOfPages > 1 AND nobottompagebar<>true Then Response.Write writepagebar(CurPage, iNumOfPages) %><br />
				<img src="images/clearpixel.gif" width="300" height="1" alt="" /></p></td>
			  </tr>
			</table>