<%
Dim sSQL,rs,rs2,alldata,cnn,rowcounter,success,Count,startlink,endlink,weburl,CurPage,iNumOfPages,subCats,lasttsid,sText,index,sJoin,aText,aFields(3),currFormat1,currFormat2,currFormat3,aDiscSection()
catid="0"
showcategories=FALSE
gotcriteria=FALSE
isrootsection=FALSE
topsectionids="0"
thecat = request("scat")
if Left(request("scat"),2)="ms" then thecat = Right(request("scat"),Len(request("scat"))-2)
thecat = replace(thecat,"'","")
catarr = split(thecat,",")
catzero = ""
if UBOUND(catarr)>=0 then
	if IsNumeric(catarr(0)) then catzero=Int(catarr(0))
end if
WSP = ""
OWSP = ""
TWSP = "pPrice"
minprice = ""
maxprice = ""
if Trim(request("sminprice"))<>"" AND IsNumeric(Trim(request("sminprice"))) then minprice = cDbl(replace(request("sminprice"),"$",""))
if Trim(request("sprice"))<>"" AND IsNumeric(Trim(request("sprice"))) then maxprice = cDbl(replace(request("sprice"),"$",""))
Sub writemenulevel(id,itlevel)
	Dim wmlindex
	if itlevel<10 then
		FOR wmlindex=0 TO ubound(alldata,2)
			if alldata(2,wmlindex)=id then
				response.write "<option value='"&alldata(0,wmlindex)&"'"
				if catzero=alldata(0,wmlindex) then response.write " selected>" else response.write ">"
				for index = 0 to itlevel-2
					response.write "&nbsp;&nbsp;&raquo;&nbsp;"
				next
				response.write alldata(1,wmlindex)&"</option>" & vbCrLf
				if alldata(3,wmlindex)=0 then call writemenulevel(alldata(0,wmlindex),itlevel+1)
			end if
		NEXT
	end if
end Sub
Function writepagebar(CurPage, iNumPages)
	Dim sLink, i, sStr, startPage, endPage
	sLink = "<a href="""&request.servervariables("url")&"?nobox="&request("nobox")&"&scat="&request("scat")&"&stext="&server.urlencode(request("stext"))&"&stype="&request("stype")&"&sprice="&server.urlencode(maxprice)&IIfVr(minprice<>"","&sminprice="&minprice,"")&"&pg="
	startPage = vrmax(1,CInt(Int(CDbl(CurPage)/10.0)*10))
	endPage = vrmin(iNumPages,CInt(Int(CDbl(CurPage)/10.0)*10)+10)
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
if Session("clientLoginLevel")<>"" then minloglevel=Session("clientLoginLevel") else minloglevel=0
sSQL = "SELECT sectionID,"&getlangid("sectionName",256)&",topSection,rootSection FROM sections WHERE sectionDisabled<="&minloglevel&" "
if onlysubcats=true then
	sSQL = sSQL & "AND rootSection=1 ORDER BY " & getlangid("sectionName",256)
else
	sSQL = sSQL & "ORDER BY sectionOrder"
end if
rs.Open sSQL,cnn,0,1
if rs.eof or rs.bof then
	success=false
else
	alldata=rs.getrows
	success=true
end if
rs.Close
if Request.Form("posted")="1" OR Request.QueryString("pg")<>"" then
	if thecat<>"" then
		sSQL = "SELECT DISTINCT products.pId,"&getlangid("pName",1)&","&WSP&"pPrice,pOrder FROM multisections RIGHT JOIN products ON products.pId=multisections.pId WHERE pDisplay <> 0 "
		gotcriteria=true
		sectionids = getsectionids(thecat, false)
		if sectionids<>"" then sSQL = sSQL & "AND (products.pSection IN (" & sectionids & ") OR multisections.pSection IN (" & sectionids & ")) "
	else
		sSQL = "SELECT products.pId,"&getlangid("pName",1)&","&WSP&"pPrice FROM products WHERE pDisplay <> 0 "
	end if
	session.LCID = 1033
	if Trim(request("sprice"))<>"" AND IsNumeric(Trim(request("sprice"))) then
		gotcriteria=true
		sSQL = sSQL & "AND "&TWSP&"<="&cDbl(replace(request("sprice"),"$",""))&" "
	end if
	if minprice<>"" then
		gotcriteria=true
		sSQL = sSQL & "AND "&TWSP&">="&minprice&" "
	end if
	session.LCID = saveLCID
	if Trim(request("stext"))<>"" then
		gotcriteria=true
		sText = replace(Trim(request("stext")),"'","''")
		aText = Split(sText)
		aFields(0)="products.pId"
		aFields(1)=getlangid("pName",1)
		aFields(2)=getlangid("pDescription",2)
		aFields(3)=getlangid("pLongDescription",4)
		if request("stype")="exact" then
			sSQL=sSQL & "AND (products.pId LIKE '%"&sText&"%' OR "&getlangid("pName",1)&" LIKE '%"&sText&"%' OR "&getlangid("pDescription",2)&" LIKE '%"&sText&"%' OR "&getlangid("pLongDescription",2)&" LIKE '%"&sText&"%') "
		else
			sJoin="AND "
			if request("stype")="any" then sJoin="OR "
			sSQL=sSQL & "AND ("
			for index=0 to 3
				sSQL=sSQL & "("
				for rowcounter=0 to UBOUND(aText)
					sSQL=sSQL & aFields(index) & " LIKE '%"&aText(rowcounter)&"%' "
					if rowcounter<UBOUND(aText) then sSQL=sSQL & sJoin
				next
				sSQL=sSQL & ") "
				if index < UBOUND(aFields) then sSQL=sSQL & "OR "
			next
			sSQL=sSQL & ") "
		end if
	end if
	if request.form("sortby")<>"" then session("sortby")=int(request.form("sortby"))
	if session("sortby")<>"" then sortBy=int(session("sortby"))
	if sortBy=2 then
		sSortBy = " ORDER BY products.pId"
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
	if NOT gotcriteria then sSQL = "SELECT products.pId FROM products WHERE pDisplay<>0"
	disabledsections = ""
	addcomma=""
	rs.Open "SELECT sectionID FROM sections WHERE sectionDisabled>"&minloglevel,cnn,0,1
	do while NOT rs.EOF
		disabledsections = disabledsections & addcomma & rs("sectionID")
		addcomma=","
		rs.MoveNext
	loop
	rs.Close
	if disabledsections<>"" then sSQL = sSQL & " AND NOT (products.pSection IN (" & getsectionids(disabledsections, true) & "))"
	if useStockManagement AND noshowoutofstock=TRUE then sSQL = sSQL & " AND (pInStock>0 OR pStockByOpts<>0)"
	rs.CursorLocation = 3 ' adUseClient
	rs.CacheSize = adminProdsPerPage
	rs.Open sSQL & sSortBy, cnn
	if rs.eof or rs.bof then
		success=false
		iNumOfPages=0
	else
		success=true
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
end if
Session("frompage")=Request.ServerVariables("URL")&IIfVr(Trim(Request.ServerVariables("QUERY_STRING"))<>"","?","")&Request.ServerVariables("QUERY_STRING")
if request("nobox")<>"true" then
%>
	  <br />
	  <form method="post" action="search.asp">		  
		  <input type="hidden" name="posted" value="1" />
            <table class="cobtbl" width="<%=maintablewidth%>" border="0" bordercolor="#B1B1B1" cellspacing="1" cellpadding="3" bgcolor="#B1B1B1">
			  <tr> 
                <td class="cobhl" align="center" colspan="4" bgcolor="#EBEBEB" height="30">
                  <strong><%=xxSrchPr%></strong>
                </td>
              </tr>
			  <tr> 
                <td class="cobhl" width="25%" align="right" bgcolor="#EBEBEB"><%=xxSrchFr%>:</td>
				<td class="cobll" width="25%" bgcolor="#FFFFFF"><input type="text" name="stext" size="20" value="<%=server.htmlencode(request("stext"))%>" /></td>
				<td class="cobhl" width="25%" align="right" bgcolor="#EBEBEB"><%=xxSrchMx%>:</td>
				<td class="cobll" width="25%" bgcolor="#FFFFFF"><input type="text" name="sprice" size="10" value="<%=server.htmlencode(request("sprice"))%>" /></td>
			  </tr>
			  <tr>
			    <td class="cobhl" width="25%" align="right" bgcolor="#EBEBEB"><%=xxSrchTp%>:</td>
				<td class="cobll" width="25%" bgcolor="#FFFFFF"><select name="stype" size="1">
					<option value=""><%=xxSrchAl%></option>
					<option value="any" <% if request("stype")="any" then response.write "selected"%>><%=xxSrchAn%></option>
					<option value="exact" <% if request("stype")="exact" then response.write "selected"%>><%=xxSrchEx%></option>
					</select>
				</td>
				<td class="cobhl" width="25%" align="right" bgcolor="#EBEBEB"><%=xxSrchCt%>:</td>
				<td class="cobll" width="25%" bgcolor="#FFFFFF">
				  <select name="scat" size="1">
				  <option value=""><%=xxSrchAC%></option>
					<% if IsArray(alldata) then call writemenulevel(0,1) %>
				  </select>
				</td>
              </tr>
			  <tr>
			    <td class="cobhl" bgcolor="#EBEBEB">&nbsp;</td>
			    <td class="cobll" bgcolor="#FFFFFF" colspan="3"><table width="100%" cellspacing="0" cellpadding="0" border="0">
				    <tr>
					  <td class="cobll" bgcolor="#FFFFFF" width="66%" align="center"><input type="submit" value="<%=xxSearch%>" /></td>
					  <td class="cobll" bgcolor="#FFFFFF" width="34%" height="26" align="right" valign="bottom"><img src="images/tablebr.gif" alt="" /></td>
					</tr>
				  </table></td>
			  </tr>
			</table>
		</form>
<%
end if
if request.form("posted")="1" OR Request.QueryString("pg")<>"" then
%>
		<table border="0" cellspacing="<%=maintablespacing%>" cellpadding="<%=maintablepadding%>" width="<%=maintablewidth%>" bgcolor="<%=maintablebg%>" align="center">
<%
	if rs.EOF then
%>
		<tr>
		  <td align="center">
		    <p>&nbsp;</p>
		    <p><strong><%=xxSrchNM%></strong></p>
			<p>&nbsp;</p>
		  </td>
		</tr>
<%
	else
%>
        <tr>
          <td width="100%">
<% if usesearchbodyformat=3 then %>
<!--#include file="incproductbody3.asp"-->
<% elseif usesearchbodyformat=2 then %>
<!--#include file="incproductbody2.asp"-->
<% else %>
<!--#include file="incproductbody.asp"-->
<% end if %>
          </td>
        </tr>
<%
	end if
%>
      </table>
<%
	rs.Close
end if
cnn.Close
set rs = nothing
set rs2 = nothing
set cnn = nothing
%>
