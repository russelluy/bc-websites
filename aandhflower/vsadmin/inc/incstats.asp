<%
if storesessionvalue="" then storesessionvalue="virtualstore"
if Session("loggedon") <> storesessionvalue OR disallowlogin=TRUE then response.end
success=true
Set rs = Server.CreateObject("ADODB.RecordSet")
Set cnn=Server.CreateObject("ADODB.Connection")
cnn.open sDSN
alreadygotadmin = getadminsettings()
fromdate = Trim(request.form("fromdate"))
todate = Trim(request.form("todate"))
if fromdate<>"" then
	if IsNumeric(fromdate) then
		thefromdate = (thedate-fromdate)
	else
		err.number=0
		on error resume next
		thefromdate = DateValue(fromdate)
		if err.number <> 0 then
			thefromdate = thedate
			success=false
			errmsg=yyDatInv & " - " & fromdate
		end if
		on error goto 0
	end if
	if todate="" then
		thetodate = thefromdate
	elseif IsNumeric(todate) then
		thetodate = (thedate-todate)
	else
		err.number=0
		on error resume next
		thetodate = DateValue(todate)
		if err.number <> 0 then
			thetodate = thedate
			success=false
			errmsg=yyDatInv & " - " & todate
		end if
		on error goto 0
	end if
	if thefromdate > thetodate then
		tmpdate = thetodate
		thetodate = thefromdate
		thefromdate = tmpdate
	end if
else
	thefromdate = Date()-365
	thetodate = Date()
end if
%>
      <table border="0" cellspacing="0" cellpadding="0" width="100%" bgcolor="" align="center">
        <tr>
          <td width="100%" align="center">
            <table width="550" border="0" cellspacing="0" cellpadding="3" bgcolor="">
			  <tr> 
                <td width="100%" align="center"><strong>Sales reports from <%=thefromdate%> to <%=thetodate%></strong><br />&nbsp;</td>
			  </tr>
			  <tr> 
                <td width="100%" align="center"><strong>Sales results</strong><br />&nbsp;</td>
			  </tr>
<%
sSQL = "SELECT COUNT(ordID) AS numorders,SUM(ordTotal) AS theordtot,SUM(ordHandling) AS tothandling,SUM(ordStateTax) AS totstatetax,SUM(ordCountryTax) AS totcountrytax,SUM(ordHSTTax) AS tothsttax,SUM(ordDiscount) AS totdiscount, SUM(ordShipping) AS totshipping FROM orders WHERE ordStatus>=3 AND ordDate BETWEEN " & datedelim & VSUSDate(thefromdate) & datedelim & " AND " & datedelim & VSUSDate(thetodate+1) & datedelim
rs.Open sSQL,cnn,0,1
if NOT rs.EOF then
	response.write "<tr><td align=""left""><table border=""0"" cellspacing=""0"" cellpadding=""0"" width=""100%"" bgcolor="""" align=""left"">"
	response.write "<tr><td><strong>Orders</strong></td><td><strong>Order Total</strong></td><td><strong>Shipping</strong></td><td><strong>Handling</strong></td><td><strong>Discounts</strong></td><td><strong>State Tax</strong></td><td><strong>Country Tax</strong></td><td><strong>Grand Total</strong></td></tr>"
	response.write "<tr><td>" & rs("numorders") & "</td><td>" & FormatEuroCurrency(rs("theordtot")) & "</td><td>" & FormatEuroCurrency(rs("totshipping")) & "</td><td>" & FormatEuroCurrency(rs("tothandling")) & "</td><td>" & FormatEuroCurrency(rs("totdiscount")) & "</td><td>" & FormatEuroCurrency(rs("totstatetax")) & "</td><td>" & FormatEuroCurrency(rs("totcountrytax")) & "</td><td>" & FormatEuroCurrency((rs("theordtot")+rs("totshipping")+rs("tothandling")+rs("totstatetax")+rs("totcountrytax")+rs("tothsttax"))-rs("totdiscount")) & "</td></tr>"
	response.write "</table></td></tr>"
end if
rs.Close
%>
			  <tr> 
                <td width="100%" align="center"><strong>Top 100 Sales</strong><br />&nbsp;</td>
			  </tr>
<%
sSQL = "SELECT "&IIfVr(mysqlserver<>true,"TOP 100","")&" SUM(cartQuantity) AS thecount,cartProdID,cartProdName FROM cart WHERE cartCompleted=1 AND cartDateAdded BETWEEN " & datedelim & VSUSDate(thefromdate) & datedelim & " AND " & datedelim & VSUSDate(thetodate+1) & datedelim & " GROUP BY cartProdID,cartProdName ORDER BY "&IIfVr(mysqlserver=TRUE, "cartQuantity", "SUM(cartQuantity)")&" DESC"&IIfVr(mysqlserver=true," LIMIT 0,100","")
rs.Open sSQL,cnn,0,1
if NOT rs.EOF then
	response.write "<tr><td align=""left""><table border=""0"" cellspacing=""0"" cellpadding=""0"" width=""100%"" bgcolor="""" align=""left"">"
	response.write "<tr><td><strong>Prod ID</strong></td><td><strong>Prod Name</strong></td><td align=""center""><strong>Quant sold</strong></td></tr>"
	do while NOT rs.EOF
		response.write "<tr><td>" & rs("cartProdID") & "</td><td>" & rs("cartProdName") & "</td><td align=""center"">" & rs("thecount") & "</td></tr>"
		rs.MoveNext
	loop
	response.write "</table></td></tr>"
end if
rs.Close
%>

			  <tr> 
                <td width="100%" align="center"><strong>Top Countries</strong><br />&nbsp;</td>
			  </tr>
<%
sSQL = "SELECT "&IIfVr(mysqlserver<>true,"TOP 100","")&" COUNT(ordCountry) AS thecount,ordCountry FROM orders WHERE ordStatus>=3 AND ordDate BETWEEN " & datedelim & VSUSDate(thefromdate) & datedelim & " AND " & datedelim & VSUSDate(thetodate+1) & datedelim & " GROUP BY ordCountry ORDER BY "&IIfVr(mysqlserver=TRUE, "ordCountry", "COUNT(ordCountry)")&" DESC"&IIfVr(mysqlserver=true," LIMIT 0,100","")
rs.Open sSQL,cnn,0,1
if NOT rs.EOF then
	response.write "<tr><td align=""left""><table border=""0"" cellspacing=""0"" cellpadding=""0"" width=""100%"" bgcolor="""" align=""left"">"
	response.write "<tr><td><strong>Country Name</strong></td><td align=""center""><strong>Sales</strong></td></tr>"
	do while NOT rs.EOF
		response.write "<tr><td>" & rs("ordCountry") & "</td><td align=""center"">" & rs("thecount") & "</td></tr>"
		rs.MoveNext
	loop
	response.write "</table></td></tr>"
end if
rs.Close
%>
			  <tr> 
                <td width="100%" align="center"><br /><a href="admin.asp"><strong><%=yyAdmHom%></strong></a><br />&nbsp;</td>
			  </tr>
            </table></td>
        </tr>
<%
cnn.Close
set rs = nothing
set cnn = nothing
%>
      </table>