<%
Dim sSQL,rs,alldata,cnn,rowcounter,iNumOfPages,CurPage,Count,weburl,longdesc,currFormat1,currFormat2,currFormat3
if Trim(explicitid)<>"" then prodid=Trim(explicitid) else prodid=Trim(request.querystring("prod"))
prodlist="'" & Replace(prodid,"'","''") & "'"
WSP = ""
OWSP = ""
Count=0
if pricecheckerisincluded<>TRUE then pricecheckerisincluded=FALSE
sub writepreviousnextlinks()
	if previousid<>"" then
		if previousidstatic then
			response.write "<a href='"&cleanforurl(previousidname)&".asp"&IIfVr(request.querystring("cat")<>"","?cat="&request.querystring("cat"),"")&"'>"
		else
			response.write "<a href=""proddetail.asp?prod=" & previousid & IIfVr(request.querystring("cat")<>"","&cat="&request.querystring("cat"),"") & """>"
		end if
	end if
	response.write "<strong>&laquo; "&xxPrev&"</strong>"
	if previousid<>"" then response.write "</a>"
	response.write " | "
	if nextid<>"" then
		if nextidstatic then
			response.write "<a href='"&cleanforurl(nextidname)&".asp"&IIfVr(request.querystring("cat")<>"","?cat="&request.querystring("cat"),"")&"'>"
		else
			response.write "<a href=""proddetail.asp?prod=" & nextid & IIfVr(request.querystring("cat")<>"","&cat="&request.querystring("cat"),"") & """>"
		end if
	end if
	response.write "<strong>"&xxNext&" &raquo;</strong>"
	if nextid<>"" then response.write "</a>"
end sub
Set rs = Server.CreateObject("ADODB.RecordSet")
Set rs2 = Server.CreateObject("ADODB.RecordSet")
Set cnn=Server.CreateObject("ADODB.Connection")
cnn.open sDSN
alreadygotadmin = getadminsettings()
Call checkCurrencyRates(currConvUser,currConvPw,currLastUpdate,currRate1,currSymbol1,currRate2,currSymbol2,currRate3,currSymbol3)
if Session("clientUser")<>"" then
	if (Session("clientActions") AND 8) = 8 then
		WSP = "pWholesalePrice AS "
		if wholesaleoptionpricediff=TRUE then OWSP = "optWholesalePriceDiff AS "
	end if
	if (Session("clientActions") AND 16) = 16 then
		Session.LCID = 1033
		WSP = Session("clientPercentDiscount") & "*pPrice AS "
		if wholesaleoptionpricediff=TRUE then OWSP = Session("clientPercentDiscount") & "*optPriceDiff AS "
		Session.LCID = saveLCID
	end if
end if
Session("frompage")=Request.ServerVariables("URL")&IIfVr(Trim(Request.ServerVariables("QUERY_STRING"))<>"","?","")&Request.ServerVariables("QUERY_STRING")
' Previous and next
sSQL = "SELECT pId,"&getlangid("pName",1)&",pImage,"&WSP&"pPrice,pSection,pListPrice,pSell,pStockByOpts,pStaticPage,pInStock,pExemptions,"&IIfVr(detailslink<>"","'' AS ","")&"pLargeImage,"&getlangid("pDescription",2)&","&getlangid("pLongDescription",4)&" FROM products WHERE pDisplay<>0 AND pId='"&Replace(prodid,"'","''")&"'"
rs.Open sSQL,cnn,0,1,&H0001
if rs.EOF then
	response.write "<p align=""center"">&nbsp;<br />Sorry, this product is not currently available.<br />&nbsp;</p>"
else
tslist = ""
if IsNull(rs("pSection")) then catid = 0 else catid = rs("pSection")
if Trim(request.querystring("cat"))<>"" AND IsNumeric(request.querystring("cat")) AND Trim(request.querystring("cat"))<>"0" then catid = request.querystring("cat")
thetopts = catid
topsectionids = catid
isrootsection=false
for index=0 to 10
	if thetopts=0 then
		tslist = "<a href=""categories.asp"">"&xxHome&"</a> " & tslist
		exit for
	elseif index=10 then
		tslist = "<strong>Loop</strong>" & tslist
	else
		sSQL = "SELECT sectionID,topSection,"&getlangid("sectionName",256)&",rootSection,sectionurl FROM sections WHERE sectionID=" & thetopts
		rs2.Open sSQL,cnn,0,1
		if NOT rs2.EOF then
			if trim(rs2("sectionurl")&"")<>"" then
				tslist = " &raquo; <a href="""& rs2("sectionurl") & """>" & rs2(getlangid("sectionName",256)) & "</a>" & tslist
			elseif rs2("rootSection")=1 then
				tslist = " &raquo; <a href=""products.asp?cat="& rs2("sectionID") & """>" & rs2(getlangid("sectionName",256)) & "</a>" & tslist
			else
				tslist = " &raquo; <a href=""categories.asp?cat="& rs2("sectionID") & """>" & rs2(getlangid("sectionName",256)) & "</a>" & tslist
			end if
			thetopts = rs2("topSection")
			topsectionids = topsectionids & "," & thetopts
		else
			tslist = "Top Section Deleted" & tslist
			rs2.Close
			exit for
		end if
		rs2.Close
	end if
next
nextid=""
previousid=""
sectionids = getsectionids(catid, false)
sSQL = "SELECT "&IIfVr(mysqlserver<>TRUE,"TOP 1 ", "")&"products.pId,pName,pStaticPage FROM products LEFT JOIN multisections ON products.pId=multisections.pId WHERE (products.pSection IN (" & sectionids & ") OR multisections.pSection IN (" & sectionids & "))"&IIfVr(useStockManagement AND noshowoutofstock=TRUE, " AND (pInStock>0 OR pStockByOpts<>0)", "")&" AND pDisplay<>0 AND products.pId > '"&Replace(prodid,"'","''")&"' ORDER BY products.pId ASC"&IIfVr(mysqlserver=TRUE," LIMIT 0,1", "")
rs2.Open sSQL,cnn,0,1
if NOT rs2.EOF then
	nextid = Server.URLEncode(rs2("pId"))
	nextidname = rs2("pName")
	nextidstatic = (cint(rs2("pStaticPage"))<>0)
end if
rs2.Close
sSQL = "SELECT "&IIfVr(mysqlserver<>TRUE,"TOP 1 ", "")&"products.pId,pName,pStaticPage FROM products LEFT JOIN multisections ON products.pId=multisections.pId WHERE (products.pSection IN (" & sectionids & ") OR multisections.pSection IN (" & sectionids & "))"&IIfVr(useStockManagement AND noshowoutofstock=TRUE, " AND (pInStock>0 OR pStockByOpts<>0)", "")&" AND pDisplay<>0 AND products.pId < '"&Replace(prodid,"'","''")&"' ORDER BY products.pId DESC"&IIfVr(mysqlserver=TRUE," LIMIT 0,1", "")
rs2.Open sSQL,cnn,0,1
if NOT rs2.EOF then
	previousid = Server.URLEncode(rs2("pId"))
	previousidname = rs2("pName")
	previousidstatic = (cint(rs2("pStaticPage"))<>0)
end if
rs2.Close
saveLCID = Session.LCID
prodoptions=""
productdisplayscript(true)
if currencyseparator="" then currencyseparator=" "
updatepricescript(true) %>
      <table border="0" cellspacing="<%=maintablespacing%>" cellpadding="<%=maintablepadding%>" width="<%=maintablewidth%>" bgcolor="<%=maintablebg%>" align="center">
        <tr> 
          <td width="100%">
		    <form method="post" name="tForm<%=Count%>" action="cart.asp" onsubmit="return formvalidator<%=Count%>(this)">
<%	if IsEmpty(showcategories) OR showcategories=true then %>
            <table width="<%=innertablewidth%>" border="0" cellspacing="<%=innertablespacing%>" cellpadding="<%=innertablepadding%>" bgcolor="<%=innertablebg%>">
              <tr> 
                <td class="prodnavigation" colspan="3" align="left" valign="top"><strong><p class="prodnavigation"><%=tslist%><br />
                  <img src="images/clearpixel.gif" width="300" height="8" alt="" /></p></strong></td>
                <td align="right" valign="top"><% if nobuyorcheckout<>true then %><a href="cart.asp"><img src="images/checkout.gif" border="0" alt="<%=xxCOTxt%>" /></a><% else response.write "&nbsp;" end if %></td>
              </tr>
			</table>
<%	end if
	alldiscounts = ""
	if nowholesalediscounts=true AND Session("clientUser")<>"" then
		if ((Session("clientActions") AND 8) = 8) OR ((Session("clientActions") AND 16) = 16) then noshowdiscounts=true
	end if
	if noshowdiscounts<>true then
		Session.LCID = 1033
		tdt = Date()
		sSQL = "SELECT DISTINCT "&getlangid("cpnName",1024)&" FROM coupons LEFT OUTER JOIN cpnassign ON coupons.cpnID=cpnassign.cpaCpnID WHERE cpnNumAvail>0 AND cpnEndDate>="&datedelim&VSUSDate(tdt)&datedelim&" AND cpnIsCoupon=0 AND "
		sSQL = sSQL & "((cpnSitewide=1 OR cpnSitewide=2) "
		sSQL = sSQL & "OR (cpnSitewide=0 AND cpaType=2 AND cpaAssignment='"&rs("pID")&"') "
		sSQL = sSQL & "OR ((cpnSitewide=0 OR cpnSitewide=3) AND cpaType=1 AND cpaAssignment IN ('"&Replace(topsectionids,",","','")&"')))"
		Session.LCID = saveLCID
		rs2.Open sSQL,cnn,0,1
		do while NOT rs2.EOF
			alldiscounts = alldiscounts & rs2(getlangid("cpnName",1024)) & "<br />"
			rs2.MoveNext
		loop
		rs2.Close
	end if
	if usedetailbodyformat=1 OR usedetailbodyformat="" then %>
            <table width="<%=innertablewidth%>" border="0" cellspacing="<%=innertablespacing%>" cellpadding="<%=innertablepadding%>" bgcolor="<%=innertablebg%>">
              <tr> 
                <td width="100%" colspan="4"> 
<%						if showproductid=true then response.write "<div class=""detailid""><strong>" & xxPrId & ":</strong> " & rs("pID") & "</div>" %><strong><div class="detailname"><% response.write rs(getlangid("pName",1))&xxDot
						if alldiscounts<>"" then response.write " <font color=""#FF0000""><span class=""discountsapply"">"&xxDsApp&"</span></div></strong><font size=""1""><div class=""detaildiscounts"">" & alldiscounts & "</div></font></font>" else response.write "</div></strong>"
						if showinstock=TRUE then if cint(rs("pStockByOpts"))=0 then response.write "<div class=""detailinstock""><strong>" & xxInStoc & ":</strong> " & rs("pInStock") & "</div>" %>
                </td>
              </tr>
              <tr> 
                <td width="100%" colspan="4" align="center" class="detailimage"> <% if NOT (Trim(rs("pLargeImage"))="" OR IsNull(rs("pLargeImage")) OR Trim(rs("pLargeImage"))="prodimages/") then %> 
                  <img class="prodimage" src="<%=rs("pLargeImage")%>" border="0" alt="<%=strip_tags2(rs(getlangid("pName",1))&"")%>" /> <% elseif NOT (Trim(rs("pImage"))="" OR IsNull(rs("pImage")) OR Trim(rs("pImage"))="prodimages/") then %> 
                  <img class="prodimage" src="<%=rs("pImage")%>" border="0" alt="<%=strip_tags2(rs(getlangid("pName",1))&"")%>" /> <% else %> &nbsp; <% end if %> 
                </td>
              </tr>
              <tr> 
                <td width="100%" colspan="4"> 
                  <p align="left"><% shortdesc = Trim(rs(getlangid("pDescription",2)))
				longdesc = Trim(rs(getlangid("pLongDescription",4)))
				if longdesc<>"" then
					response.write "<div class=""detaildescription"">"&longdesc&"</div>"
				elseif shortdesc<>"" then
					response.write "<div class=""detaildescription"">"&shortdesc&"</div>"
				else
					response.write "&nbsp;"
				end if %></p>
<%
optionshavestock=true
totprice = rs("pPrice")
if IsArray(prodoptions) then
	response.write "<div class=""detailoptions"" align=""center""><table border='0' cellspacing='1' cellpadding='1'>"
	response.write displayproductoptions("<strong><span class=""detailoption"">","</strong>",optdiff)
	totprice = totprice + optdiff
	response.write "</table></div>"
end if
%>              </td>
              </tr>
              <tr>
			    <td width="20%"><% if useemailfriend then %>
<a href="javascript:openEFWindow('<%=Server.URLEncode(prodid)%>')"><strong><%=xxEmFrnd%></strong></a>
<% else %>
&nbsp;
<% end if %></td>
                <td width="60%" align="center" colspan="2">
				<%	if noprice=true then
						response.write "&nbsp;"
					else
						if cDbl(rs("pListPrice"))<>0.0 then response.write Replace(xxListPrice, "%s", FormatEuroCurrency(rs("pListPrice"))) & "<br />"
						if cDbl(totprice)=0 AND pricezeromessage<>"" then
							response.write "<div class=""detailprice"">" & pricezeromessage & "</div>"
						else
							response.write "<div class=""detailprice""><strong>" & xxPrice & ":</strong> <span class=""price"" id=""pricediv" & Count & """>" & FormatEuroCurrency(totprice) & "</span> "
							if showtaxinclusive=true AND (rs("pExemptions") AND 2)<>2 then response.write Replace(ssIncTax,"%s", "<span id=""pricedivti" & Count & """>" & FormatEuroCurrency(totprice+(totprice*countryTax/100.0)) & "</span> ")
							response.write "</div>"
							extracurr = ""
							if currRate1<>0 AND currSymbol1<>"" then extracurr = replace(currFormat1, "%s", FormatNumber(totprice*currRate1, checkDPs(currSymbol1))) & currencyseparator
							if currRate2<>0 AND currSymbol2<>"" then extracurr = extracurr & replace(currFormat2, "%s", FormatNumber(totprice*currRate2, checkDPs(currSymbol2))) & currencyseparator
							if currRate3<>0 AND currSymbol3<>"" then extracurr = extracurr & replace(currFormat3, "%s", FormatNumber(totprice*currRate3, checkDPs(currSymbol3)))
							if extracurr<>"" then response.write "<div class=""detailcurrency""><span class=""extracurr"" id=""pricedivec" & Count & """>" & extracurr & "</span></div>"
						end if
					end if %>
				</td> 
                <td width="20%" align="right">
<%
if nobuyorcheckout=true then
	response.write "&nbsp;"
else
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
<%		if showquantondetail=true then response.write "<input type=""text"" name=""quant"" size=""2"" maxlength=""5"" value=""1"" />&nbsp;"
		if custombuybutton<>"" then response.write custombuybutton else response.write "<input type=""image"" align=""middle"" src=""images/buy.gif"" alt="""&xxAddToC&""" />"
	else
		response.write "<strong>"&xxOutStok&"</strong>"
	end if
end if			%></td>
            </tr>
<%
if previousid<>"" OR nextid<>"" then
	response.write "<tr><td align=""center"" colspan=""4"" class=""pagenums""><p class=""pagenums"">&nbsp;<br />"
	call writepreviousnextlinks()
	response.write "</p></td></tr>"
end if
rs.Close
cnn.Close
%> 
            </table>
<% else ' if usedetailbodyformat=2 %>
			<table width="<%=innertablewidth%>" border="0" cellspacing="<%=innertablespacing%>" cellpadding="<%=innertablepadding%>" bgcolor="<%=innertablebg%>">
              <tr> 
                <td width="30%" align="center" class="detailimage"> <%
				if NOT (Trim(rs("pLargeImage"))="" OR IsNull(rs("pLargeImage")) OR Trim(rs("pLargeImage"))="prodimages/") then %> 
                  <img class="prodimage" src="<%=rs("pLargeImage")%>" border="0" alt="<%=strip_tags2(rs(getlangid("pName",1))&"")%>" /> <%
				elseif NOT (Trim(rs("pImage"))="" OR IsNull(rs("pImage")) OR Trim(rs("pImage"))="prodimages/") then %> 
                  <img class="prodimage" src="<%=rs("pImage")%>" border="0" alt="<%=strip_tags2(rs(getlangid("pName",1))&"")%>" /> <%
				else %> &nbsp; <%
				end if %> 
                </td>
				<td>&nbsp;</td>
				<td width="70%" valign="top"> 
<%				totprice = rs("pPrice")
				optionshavestock=true
				if IsArray(prodoptions) then
					optionshavestock=true
					optionshtml = displayproductoptions("<span class=""detailoption"">","</span>", optdiff)
					totprice = totprice + optdiff
				end if
				if showproductid=true then response.write "<div class=""detailid""><strong>" & xxPrId & ":</strong> " & rs("pID") & "</div>" %><strong><div class="detailname"><% response.write rs(getlangid("pName",1))&xxDot
				if alldiscounts<>"" then response.write " <font color=""#FF0000""><span class=""discountsapply"">"&xxDsApp&"</span></font></div></strong><font size=""1"" color=""#FF0000""><div class=""detaildiscounts"">" & alldiscounts & "</div></font>" else response.write "</div></strong>"
				if showinstock=TRUE then if cint(rs("pStockByOpts"))=0 then response.write "<div class=""detailinstock""><strong>" & xxInStoc & ":</strong> " & rs("pInStock") & "</div>"
				response.write "<br />"
				shortdesc = Trim(rs(getlangid("pDescription",2)))
				longdesc = Trim(rs(getlangid("pLongDescription",4)))
				if longdesc<>"" then
					response.write "<div class=""detaildescription"">"&longdesc&"</div>"
				elseif shortdesc<>"" then
					response.write "<div class=""detaildescription"">"&shortdesc&"</div>"
				end if
				if noprice=true then
					response.write "&nbsp;"
				else
					if cDbl(rs("pListPrice"))<>0.0 then response.write Replace(xxListPrice, "%s", FormatEuroCurrency(rs("pListPrice"))) & "<br />"
					if cDbl(totprice)=0 AND pricezeromessage<>"" then
						response.write "<div class=""detailprice"">" & pricezeromessage & "</div>"
					else
						response.write "<div class=""detailprice""><strong>" & xxPrice & ":</strong> <span class=""price"" id=""pricediv" & Count & """>" & FormatEuroCurrency(totprice) & "</span> "
						if showtaxinclusive=true AND (rs("pExemptions") AND 2)<>2 then response.write Replace(ssIncTax,"%s", "<span id=""pricedivti" & Count & """>" & FormatEuroCurrency(totprice+(totprice*countryTax/100.0)) & "</span> ")
						response.write "</div>"
						extracurr = ""
						if currRate1<>0 AND currSymbol1<>"" then extracurr = replace(currFormat1, "%s", FormatNumber(totprice*currRate1, checkDPs(currSymbol1))) & currencyseparator
						if currRate2<>0 AND currSymbol2<>"" then extracurr = extracurr & replace(currFormat2, "%s", FormatNumber(totprice*currRate2, checkDPs(currSymbol2))) & currencyseparator
						if currRate3<>0 AND currSymbol3<>"" then extracurr = extracurr & replace(currFormat3, "%s", FormatNumber(totprice*currRate3, checkDPs(currSymbol3)))
						if extracurr<>"" then response.write "<div class=""detailcurrency""><span class=""extracurr"" id=""pricedivec" & Count & """>" & extracurr & "</span></div>"
					end if
					response.write "<hr width=""80%"" />"
				end if
if IsArray(prodoptions) then
	response.write "<div class=""detailoptions"" align=""center""><table border='0' cellspacing='1' cellpadding='1' width='100%'>"
	response.write optionshtml
	if nobuyorcheckout<>true AND (showquantondetail=TRUE OR IsEmpty(showquantondetail)) then
%>
	<tr><td align="right"><%=xxQuant%>:</td><td align="left"><input type="text" name="quant" maxlength="5" size="4" value="1" /></td></tr>
<%
	end if
	response.write "</table></div>"
else
	if nobuyorcheckout<>true AND (showquantondetail=TRUE OR IsEmpty(showquantondetail)) then
%>
	<table border='0' cellspacing='1' cellpadding='1' width='100%'>
	<tr><td align="right"><%=xxQuant%>:</td><td><input type="text" name="quant" maxlength="5" size="4" value="1" /></td></tr>
	</table>
<%
	end if
end if
%>
<p align="center">
<%
if nobuyorcheckout=true then
	response.write "&nbsp;"
else
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
<%		if custombuybutton<>"" then response.write custombuybutton else response.write "<input type=""image"" src=""images/buy.gif"" alt="""&xxAddToC&""" /><br />"
	else
		response.write "<strong>"&xxOutStok&"</strong><br />"
	end if
end if
if previousid<>"" OR nextid<>"" then
	response.write "</p><p class=""pagenums"" align=""center"">"
	call writepreviousnextlinks()
	response.write "<br />"
end if %>
<hr width="80%" /></p>
<% if useemailfriend then %>
<p align="center"><a href="javascript:openEFWindow('<%=Server.URLEncode(prodid)%>')"><strong><%=xxEmFrnd%></strong></a></p>
<% end if %>
</td>
            </tr>
<%
rs.Close
cnn.Close
%> 
            </table>
<% end if ' usedetailbodyformat
%>
			</form>
          </td>
        </tr>
      </table>
<%
end if ' rs.EOF
set rs = nothing
set rs2 = nothing
set cnn = nothing
%>