<%
WSP = ""
OWSP = ""
TWSP = "pPrice"
cs=csstyleprefix
if pricecheckerisincluded<>TRUE then pricecheckerisincluded=FALSE
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
Set rs = Server.CreateObject("ADODB.RecordSet")
Set rs2 = Server.CreateObject("ADODB.RecordSet")
Set cnn=Server.CreateObject("ADODB.Connection")
if crosssellcolumns="" then if productcolumns="" then crosssellcolumns=3 else crosssellcolumns=productcolumns
if crosssellrows="" then crosssellrows=1
numberofproducts=crosssellcolumns*crosssellrows
productcolumns=crosssellcolumns
if csnobuyorcheckout=TRUE then nobuyorcheckout=TRUE
if csnoshowdiscounts=TRUE then noshowdiscounts=TRUE
if csnoproductoptions=TRUE then noproductoptions=TRUE
if IsEmpty(forcedetailslink) then forcedetailslink=TRUE
iNumOfPages=1
showcategories=FALSE
isrootsection=TRUE
catid = "0"
if IsEmpty(Count) then Count=0 else Count=(Count+crosssellcolumns)-(Count MOD crosssellcolumns)
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
if NOT (prodlist<>"") then prodlist=""
if request.form("mode") <> "checkout" AND request.form("mode") <> "add" AND request.form("mode") <> "go" AND request.form("mode") <> "paypalexpress1" then
	cnn.open sDSN
	alreadygotadmin = getadminsettings()
	crosssellactionarr = split(crosssellaction, ",")
	for csindex=0 to UBOUND(crosssellactionarr)
		crosssellaction=trim(crosssellactionarr(csindex))
		addcomma="" : relatedlist=""
		if crosssellaction="alsobought" then ' Those who bought what's in your cart also bought.
			if csalsoboughttitle="" then crossselltitle="Customers who bought these products also bought." else crossselltitle=csalsoboughttitle
			if prodlist="" then
				addcomma=""
				sSQL = "SELECT cartProdID FROM cart WHERE cartCompleted=0 AND cartSessionID="&replace(Session.SessionID,"'","")
				rs.Open sSQL, cnn, 0, 1
					do while NOT rs.EOF
						prodlist = prodlist & addcomma & "'" & replace(rs("cartProdID"),"'","''") & "'"
						addcomma=","
						rs.MoveNext
					loop
				rs.Close
			end if
			addcomma="" : sessionlist="" : thecount=0
			if prodlist<>"" then
				sSQL = "SELECT "&IIfVr(mysqlserver<>true,"TOP 100","")&" cartSessionID,COUNT(cartSessionID),MAX(cartDateAdded) as maxdateadded FROM cart WHERE cartProdID IN ("&prodlist&") GROUP BY cartSessionID HAVING COUNT(cartSessionID) > 1 ORDER BY "&IIfVr(mysqlserver=TRUE,"maxdateadded","MAX(cartDateAdded)")&" DESC"&IIfVr(mysqlserver=true," LIMIT 0,100","")
				' response.write sSQL & "<br>"
				rs.Open sSQL, cnn, 0, 1
					do while NOT rs.EOF AND thecount<100
						sessionlist = sessionlist & addcomma & replace(rs("cartSessionID"),"'","''")
						addcomma=","
						thecount=thecount+1
						rs.MoveNext
					loop
				rs.Close
			end if
			if prodlist<>"" AND sessionlist<>"" then
				sSQL = "SELECT "&IIfVr(mysqlserver<>true,"TOP "&numberofproducts,"")&" cartProdID FROM cart WHERE cartSessionID IN ("&sessionlist&") AND cartProdID NOT IN ("&prodlist&") ORDER BY cartDateAdded DESC"&IIfVr(mysqlserver=true," LIMIT 0,"&numberofproducts,"")
				' response.write sSQL & "<br>"
				rs.Open sSQL, cnn, 0, 1
					addcomma="" : relatedlist="" : thecount=0
					do while NOT rs.EOF AND thecount<numberofproducts
						relatedlist = relatedlist & addcomma & "'" & replace(rs("cartProdID"),"'","''") & "'"
						addcomma=","
						thecount=thecount+1
						rs.MoveNext
					loop
				rs.Close
			end if
		elseif crosssellaction="recommended" then ' Top x recommended products (Needs v5.1)
			if csrecommendedtitle="" then crossselltitle="These products are our current recommendations for you." else crossselltitle=csrecommendedtitle
			if prodlist="" then
				addcomma=""
				sSQL = "SELECT cartProdID FROM cart WHERE cartCompleted=0 AND cartSessionID="&replace(Session.SessionID,"'","")
				rs.Open sSQL, cnn, 0, 1
					do while NOT rs.EOF
						prodlist = prodlist & addcomma & "'" & replace(rs("cartProdID"),"'","''") & "'"
						addcomma=","
						rs.MoveNext
					loop
				rs.Close
			end if
			sSQL = "SELECT pID FROM products WHERE pRecommend<>0"
			if prodlist<>"" then sSQL = sSQL & " AND pID NOT IN (" & prodlist & ")"
			rs.Open sSQL, cnn, 0, 1
				addcomma="" : relatedlist=""
				do while NOT rs.EOF
					relatedlist = relatedlist & addcomma & "'" & replace(rs("pID"),"'","''") & "'"
					addcomma=","
					rs.MoveNext
				loop
			rs.Close
		elseif crosssellaction="related" then ' Products recommended with this product (Would need v5.1)
			if csrelatedtitle="" then crossselltitle="These products are recommended with items in your cart." else crossselltitle=csrelatedtitle
			if prodlist="" then
				addcomma=""
				sSQL = "SELECT cartProdID FROM cart WHERE cartCompleted=0 AND cartSessionID="&replace(Session.SessionID,"'","")
				rs.Open sSQL, cnn, 0, 1
					do while NOT rs.EOF
						prodlist = prodlist & addcomma & "'" & replace(rs("cartProdID"),"'","''") & "'"
						addcomma=","
						rs.MoveNext
					loop
				rs.Close
			end if
			if prodlist<>"" then
				sSQL = "SELECT rpRelProdID FROM relatedprods WHERE rpProdID IN ("&prodlist&") AND rpRelProdID NOT IN ("&prodlist&")"
				rs.Open sSQL, cnn, 0, 1
					addcomma="" : relatedlist=""
					do while NOT rs.EOF
						relatedlist = relatedlist & addcomma & "'" & replace(rs("rpRelProdID"),"'","''") & "'"
						addcomma=","
						rs.MoveNext
					loop
				rs.Close
			end if
		elseif crosssellaction="bestsellers" then ' Top X best sellers
			if csbestsellerstitle="" then crossselltitle="These are our current best sellers." else crossselltitle=csbestsellerstitle
			sSQL = "SELECT "&IIfVr(mysqlserver<>true,"TOP "&numberofproducts,"")&" cartProdID,COUNT(cartProdID) AS pidcount FROM cart INNER JOIN products ON cart.cartProdID=products.pID WHERE pDisplay<>0 "&IIfVr(crosssellsection<>"", " AND pSection IN ("&crosssellsection&")", "")&IIfVr(crosssellnotsection<>"", " AND pSection NOT IN ("&crosssellnotsection&")", "")&" GROUP BY cartProdID ORDER BY "&IIfVr(mysqlserver=true,"pidcount","COUNT(cartProdID)")&" DESC"&IIfVr(mysqlserver=true," LIMIT 0,"&numberofproducts,"")
			relatedlist="" : thecount=0
			rs.Open sSQL, cnn, 0, 1
				do while NOT rs.EOF AND thecount<numberofproducts
					relatedlist = relatedlist & addcomma & "'" & replace(rs("cartProdID"),"'","''") & "'"
					addcomma=","
					thecount=thecount+1
					rs.MoveNext
				loop
			rs.Close
		else
			if crosssellaction<>"" then response.write "<p>Unrecognized crosssell action " & crosssellaction & "</p>"
		end if
		if relatedlist<>"" then
			saveprodlist=prodlist
			prodlist=relatedlist
			sSQL = "SELECT pId,"&getlangid("pName",1)&",pImage,"&WSP&"pPrice,pListPrice,pSection,pSell,pStockByOpts,pStaticPage,pInStock,pExemptions,pLargeImage,'' AS "&getlangid("pDescription",2)&","&getlangid("pLongDescription",4)&" FROM products WHERE pId IN (" & relatedlist & ")"
			if useStockManagement AND noshowoutofstock=TRUE then sSQL = sSQL & " AND (pInStock>0 OR pStockByOpts<>0)"
			sSQL = sSQL & sSortBy
			' response.write replace(sSQL,",", ", ") & "<br>"
			rs.CursorLocation = 3 ' adUseClient
			rs.CacheSize = numberofproducts
			rs.Open sSQL, cnn
			if NOT rs.EOF then
				response.write "<p class=""cstitle""><strong>"&crossselltitle&"</strong></p>"
				rs.MoveFirst
				rs.PageSize = 100
				rs.AbsolutePage = 1
%>
<!--#include file="incproductbody2.asp"-->
<%
			end if
			rs.Close
			prodlist=saveprodlist
		end if
	next
	cnn.Close
end if
set rs = nothing
set cnn = nothing
%>