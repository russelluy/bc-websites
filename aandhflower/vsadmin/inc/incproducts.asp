<%
Dim rs,alldata,cnn,rowcounter,iNumOfPages,CurPage,Count,weburl,startlink,endlink,catid,currFormat1,currFormat2,currFormat3,aDiscSection()
catid = Trim(Replace(Request.QueryString("id"),"'",""))
if Trim(Replace(Request.QueryString("cat"),"'",""))<>"" then catid=Trim(Replace(Request.QueryString("cat"),"'",""))
if NOT IsNumeric(catid) then catid="0"
if explicitid<>"" AND IsNumeric(explicitid) then catid=explicitid
WSP = ""
OWSP = ""
TWSP = "pPrice"
sectionurl="products.asp"
iNumOfPages = 0
Function writepagebar(CurPage, iNumPages)
	Dim sLink, i, sStr, startPage, endPage
	sLink = "<a href="""&sectionurl&"?"
	for each objQS in request.querystring
		if objQS<>"cat" AND objQS<>"id" AND objQS<>"pg" then sLink = sLink & objQS & "=" & request.querystring(objQS) & "&"
	next
	if catid<>"0" AND explicitid="" then sLink = sLink & "cat="&catid&"&pg=" else sLink = sLink & "pg="
	startPage = vrmax(1,Int(CDbl(CurPage)/10.0)*10)
	endPage = vrmin(iNumPages,(Int(CDbl(CurPage)/10.0)*10)+10)
	if CurPage > 1 then
		sStr = sLink & "1" & """><strong><font face=""Verdana"">&laquo;</font></strong></a> " & sLink & CurPage-1 & """>"&xxPrev&"</a> | "
	else
		sStr = "<strong><font face=""Verdana"">&laquo;</font></strong> "&xxPrev&" | "
	end if
	for i=startPage to endPage
		if i=CurPage then
			sStr = sStr & "<span class=""currpage"">" & i & "</span> | "
		else
			sStr = sStr & sLink & i & """>"
			if i=startPage AND i > 1 then sStr=sStr&"..."
			sStr = sStr & i
			if i=endPage AND i < iNumPages then sStr=sStr&"..."
			sStr = sStr & "</a> | "
		end if
	next
	if CurPage < iNumPages then
		writepagebar = sStr & sLink & CurPage+1 & """>"&xxNext&"</a> " & sLink & iNumPages & """><strong><font face=""Verdana"">&raquo;</font></strong></a>"
	else
		writepagebar = sStr & " "&xxNext&" <strong><font face=""Verdana"">&raquo;</font></strong>"
	end if
	writepagebar = replace(replace(writepagebar,"&pg=1""",""""),"?pg=1""","""")
End function
Set rs = Server.CreateObject("ADODB.RecordSet")
Set rs2 = Server.CreateObject("ADODB.RecordSet")
Set cnn=Server.CreateObject("ADODB.Connection")
cnn.open sDSN
alreadygotadmin = getadminsettings()
if orprodsperpage<>"" then adminProdsPerPage=orprodsperpage
Redim aDiscSection(2,adminProdsPerPage)
Call checkCurrencyRates(currConvUser,currConvPw,currLastUpdate,currRate1,currSymbol1,currRate2,currSymbol2,currRate3,currSymbol3)
if Session("clientUser")<>"" then
	if (Session("clientActions") AND 8) = 8 then
		WSP = "pWholesalePrice AS "
		TWSP = "pWholesalePrice"
		if wholesaleoptionpricediff=TRUE then OWSP = "optWholesalePriceDiff AS "
	end if
	if (Session("clientActions") AND 16) = 16 then
		Session.LCID = 1033
		WSP = Session("clientPercentDiscount") & "*pPrice AS "
		TWSP = Session("clientPercentDiscount") & "*pPrice"
		if wholesaleoptionpricediff=TRUE then OWSP = Session("clientPercentDiscount") & "*optPriceDiff AS "
		Session.LCID = saveLCID
	end if
end if
tslist = ""
thetopts = catid
topsectionids = catid
isrootsection=false
sectiondisabled=false
if Session("clientLoginLevel")<>"" then minloglevel=Session("clientLoginLevel") else minloglevel=0
for index=0 to 10
	if thetopts=0 then
		tslist = "<a href=""categories.asp"">"&xxHome&"</a> " & tslist
		exit for
	elseif index=10 then
		tslist = "<strong>Loop</strong>" & tslist
	else
		sSQL = "SELECT sectionID,topSection,"&getlangid("sectionName",256)&",rootSection,sectionDisabled,sectionurl FROM sections WHERE sectionID=" & thetopts
		rs.Open sSQL,cnn,0,1
		if NOT rs.EOF then
			if rs("sectionID")=Int(catid) then isrootsection = (rs("rootSection")=1)
			if rs("sectionDisabled")>minloglevel then catid=-1
			if rs("sectionID")=Int(catid) AND isrootsection then
				tslist = " &raquo; " & rs(getlangid("sectionName",256)) & tslist
				if explicitid<>"" AND trim(rs("sectionurl")&"")<>"" then sectionurl=rs("sectionurl")
			elseif trim(rs("sectionurl")&"")<>"" then
				tslist = " &raquo; <a href="""& rs("sectionurl") & """>" & rs(getlangid("sectionName",256)) & "</a>" & tslist
				if explicitid<>"" AND rs("sectionID")=Int(catid) then sectionurl=rs("sectionurl")
			elseif rs("rootSection")=1 then
				tslist = " &raquo; <a href=""products.asp?cat="& rs("sectionID") & """>" & rs(getlangid("sectionName",256)) & "</a>" & tslist
			else
				tslist = " &raquo; <a href=""categories.asp?cat="& rs("sectionID") & """>" & rs(getlangid("sectionName",256)) & "</a>" & tslist
			end if
			thetopts = rs("topSection")
			topsectionids = topsectionids & "," & thetopts
		else
			tslist = "Top Section Deleted" & tslist
			rs.Close
			exit for
		end if
		rs.Close
	end if
next
if NOT isrootsection AND xxAlProd<>"" then tslist = tslist & " &raquo; "&xxAlProd
if catid="0" then
	disabledsections = ""
	addcomma=""
	rs.Open "SELECT sectionID FROM sections WHERE sectionDisabled>"&minloglevel,cnn,0,1
	do while NOT rs.EOF
		disabledsections = disabledsections & addcomma & rs("sectionID")
		addcomma=","
		rs.MoveNext
	loop
	rs.Close
	sSQL = "SELECT pId FROM products WHERE pDisplay<>0"
	if disabledsections<>"" then sSQL = sSQL & " AND NOT (products.pSection IN (" & getsectionids(disabledsections, true) & "))"
else
	sectionids = getsectionids(catid, false)
	sSQL = "SELECT DISTINCT products.pId,"&getlangid("pName",1)&","&WSP&"pPrice,pOrder FROM products LEFT JOIN multisections ON products.pId=multisections.pId WHERE pDisplay<>0 AND (products.pSection IN (" & sectionids & ") OR multisections.pSection IN (" & sectionids & "))"
end if
if useStockManagement AND noshowoutofstock=TRUE then sSQL = sSQL & " AND (pInStock>0 OR pStockByOpts<>0)"
if request.form("sortby")<>"" then session("sortby")=int(request.form("sortby"))
if session("sortby")<>"" then sortBy=int(session("sortby"))
if sortBy=2 then
	sSortBy = " ORDER BY products.pID"
elseif sortBy=3 then
	sSortBy = " ORDER BY "&TWSP
elseif sortBy=4 then
	sSortBy = " ORDER BY "&TWSP&" DESC"
elseif sortBy=5 then
	sSortBy = ""
elseif sortBy=6 then
	sSortBy = " ORDER BY pOrder"
elseif sortBy=7 then
	sSortBy = " ORDER BY pOrder DESC"
else
	sSortBy = " ORDER BY "&getlangid("pName",1)
end if
rs.CursorLocation = 3 ' adUseClient
rs.CacheSize = adminProdsPerPage
rs.Open sSQL & sSortBy, cnn
if NOT rs.EOF then
	rs.MoveFirst
	rs.PageSize = adminProdsPerPage
	If Request.QueryString("pg") = "" Then
		CurPage = 1
	Else
		CurPage = Int(Request.QueryString("pg"))
	End If
	iNumOfPages = Int((rs.RecordCount + (adminProdsPerPage-1)) / adminProdsPerPage)
	rs.AbsolutePage = CurPage
end if
Count = 0
if NOT rs.EOF then
	prodlist = ""
	addcomma=""
	Do While Not rs.EOF And Count < rs.PageSize
		prodlist = prodlist & addcomma & "'" & rs("pId") & "'"
		rs.MoveNext
		Count = Count + 1
		addcomma=","
	loop
	rs.Close
	Count = 0
	sSQL = "SELECT pId,"&getlangid("pName",1)&",pImage,"&WSP&"pPrice,pListPrice,pSection,pSell,pStockByOpts,pStaticPage,pInStock,pExemptions,pLargeImage,"&getlangid("pDescription",2)&","&getlangid("pLongDescription",4)&" FROM products WHERE pId IN (" & prodlist & ")" & sSortBy
	rs.Open sSQL, cnn, 0, 1
end if
Session("frompage")=Request.ServerVariables("URL")&IIfVr(Trim(Request.ServerVariables("QUERY_STRING"))<>"","?","")&Request.ServerVariables("QUERY_STRING")
%>
      <table border="0" cellspacing="<%=maintablespacing%>" cellpadding="<%=maintablepadding%>" width="<%=maintablewidth%>" bgcolor="<%=maintablebg%>" align="center">
        <tr> 
          <td colspan="3" width="100%">
<% if useproductbodyformat=3 then %>
<!--#include file="incproductbody3.asp"-->
<% elseif useproductbodyformat=2 then %>
<!--#include file="incproductbody2.asp"-->
<% else %>
<!--#include file="incproductbody.asp"-->
<% end if %>
		  </td>
        </tr>
      </table>
<%
	rs.Close
cnn.Close
set rs = nothing
set rs2 = nothing
set cnn = nothing
%>