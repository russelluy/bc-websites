<%
if Trim(Request.Form("sessionid")) <> "" then
	thesession = Trim(Request.Form("sessionid"))
else
	thesession = Session.SessionID
end if
thesession = Replace(thesession, "'", "")
function FormatMCCurrency(amount)
	if overridecurrency=true then
		if orcpreamount=true then
			FormatMCCurrency = orcsymbol & FormatNumber(amount,orcdecplaces)
		else
			FormatMCCurrency = FormatNumber(amount,orcdecplaces) & orcsymbol
		end if
	else
		if useEuro then
			FormatMCCurrency = FormatNumber(amount,2) & " &euro;"
		else
			FormatMCCurrency = FormatCurrency(amount)
		end if
	end if
end function
mcgndtot=0
mcpdtxt=""
totquant=0
shipping=0
discounts=0
Set rs = Server.CreateObject("ADODB.RecordSet")
Set rs2 = Server.CreateObject("ADODB.RecordSet")
Set cnn=Server.CreateObject("ADODB.Connection")
cnn.open sDSN
if incfunctionsdefined=TRUE then
	alreadygotadmin = getadminsettings()
else
	sSQL = "SELECT countryLCID,countryCurrency,adminStoreURL FROM admin INNER JOIN countries ON admin.adminCountry=countries.countryID WHERE adminID=1"
	rs.Open sSQL,cnn,0,1
	if orlocale<>"" then
		Session.LCID = orlocale
	elseif rs("countryLCID")<>0 then
		Session.LCID = rs("countryLCID")
	end if
	useEuro = (rs("countryCurrency")="EUR")
	storeurl = rs("adminStoreURL")
	if (left(LCase(storeurl),7) <> "http://") AND (left(LCase(storeurl),8) <> "https://") then storeurl = "http://" & storeurl
	if Right(storeurl,1) <> "/" then storeurl = storeurl & "/"
	rs.Close
end if
sSQL = "SELECT cartID,cartProdID,cartProdName,cartProdPrice,cartQuantity FROM cart WHERE cartCompleted=0 AND cartSessionID="&thesession
rs2.Open sSQL,cnn,0,1
do while NOT rs2.EOF
	optPriceDiff=0
	mcpdtxt = mcpdtxt & "<tr><td class=""mincart"" bgcolor=""#F0F0F0"">"&rs2("cartQuantity") &" " & rs2("cartProdName") & "</td></tr>"
	sSQL = "SELECT SUM(coPriceDiff) AS sumDiff FROM cartoptions WHERE coCartID="&rs2("cartID")
	rs.Open sSQL,cnn,0,1
	if NOT IsNull(rs("sumDiff")) then optPriceDiff=rs("sumDiff")
	rs.Close
	subtot = ((rs2("cartProdPrice")+optPriceDiff)*Int(rs2("cartQuantity")))
	totquant = totquant + 1
	mcgndtot=mcgndtot+subtot
	rs2.MoveNext
loop
rs2.Close
cnn.Close
set msrs = nothing
set msrs2 = nothing
set cnn = nothing
%>
      <table class="mincart" width="130" bgcolor="#FFFFFF">
        <tr> 
          <td class="mincart" bgcolor="#C6E0C3" align="center"><img src="images/littlecart1.gif" align="top" width="16" height="15" alt="<%=xxMCSC%>" /> 
            &nbsp;<strong><a href="<%=storeurl%>cart.asp"><%=xxMCSC%></a></strong></td>
        </tr>
<%		if request.form("mode")="update" then %>
		<tr> 
          <td class="mincart" bgcolor="#C6E0C3" align="center"><%=xxMainWn%></td>
        </tr>
<%		else %>
        <tr> 
          <td class="mincart" bgcolor="#C6E0C3" align="center"> 
<%			response.write totquant & " " & xxMCIIC %></td>
        </tr>
<%			response.write mcpdtxt
			if mcpdtxt<>"" AND session("discounts")<>"" then
				discounts = cDbl(session("discounts")) %>
        <tr> 
          <td class="mincart" bgcolor="#C6E0C3" align="center"><font color="#FF0000"><%=xxDscnts & " " & FormatMCCurrency(discounts)%></font></td>
        </tr>
<%			end if
			if mcpdtxt<>"" AND session("xsshipping")<>"" then
				shipping = cDbl(session("xsshipping"))
				if shipping=0 then showshipping="<font color=""#FF0000""><strong>"&xxFree&"</strong></font>" else showshipping=FormatMCCurrency(shipping) %>
        <tr> 
          <td class="mincart" bgcolor="#C6E0C3" align="center"><%=xxMCShpE & " " & showshipping%></td>
        </tr>
<%			end if %>
        <tr> 
          <td class="mincart" bgcolor="#C6E0C3" align="center"><%=xxTotal & " " & FormatMCCurrency((mcgndtot+shipping)-discounts)%></td>
        </tr>
<%		end if %>
        <tr> 
          <td class="mincart" bgcolor="#C6E0C3" align="center"><font face='Verdana'>&raquo;</font> <a href="<%=storeurl%>cart.asp"><strong><%=xxMCCO%></strong></a></td>
        </tr>
</table>