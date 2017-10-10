<%
Dim sSQL,rs,alldata,cnn,rowcounter,success,startlink,secdesc
theid=Trim(Replace(Request.QueryString("id"),"'",""))
if Trim(Replace(Request.QueryString("cat"),"'",""))<>"" then theid=Trim(Replace(Request.QueryString("cat"),"'",""))
if theid="" OR theid="ALL" OR NOT IsNumeric(theid) then theid="0"
if explicitid<>"" AND IsNumeric(explicitid) then theid=explicitid
if NOT IsNumeric(categorycolumns) OR categorycolumns="" then categorycolumns=1
cellwidth = Int(100/categorycolumns)
if usecategoryformat=3 then
	afterimage="<br />"
	beforedesc=""
elseif usecategoryformat=2 then
	afterimage=""
	beforedesc=""
else
	usecategoryformat=1
	afterimage=""
	beforedesc="</td></tr><tr><td class=""catdesc"" colspan=""2"">"
end if
border=0
if IsEmpty(catseparator) then catseparator = "<br />&nbsp;"
Set rs = Server.CreateObject("ADODB.RecordSet")
Set cnn=Server.CreateObject("ADODB.Connection")
cnn.open sDSN
alreadygotadmin = getadminsettings()
Session("frompage")=Request.ServerVariables("URL")&IIfVr(Trim(Request.ServerVariables("QUERY_STRING"))<>"","?","")&Request.ServerVariables("QUERY_STRING")
tslist = ""
thetopts = theid
topsectionids = theid
if Session("clientLoginLevel")<>"" then minloglevel=Session("clientLoginLevel") else minloglevel=0
success = true
for index=0 to 10
	if thetopts=0 then
		if theid="0" then
			tslist = xxHome & " " & tslist
		else
			tslist = "<a href=""categories.asp"">"&xxHome&"</a> " & tslist
		end if
		exit for
	elseif index=10 then
		tslist = "<strong>Loop</strong>" & tslist
	else
		sSQL = "SELECT sectionID,topSection,"&getlangid("sectionName",256)&",rootSection,sectionDisabled,sectionurl FROM sections WHERE sectionID=" & thetopts
		rs.Open sSQL,cnn,0,1
		if NOT rs.EOF then
			if rs("sectionDisabled")>minloglevel then
				success = false
			elseif rs("sectionID")=Int(theid) then
				tslist = " &raquo; " & rs(getlangid("sectionName",256)) & tslist
			elseif trim(rs("sectionurl")&"")<>"" then
				tslist = " &raquo; <a href="""& rs("sectionurl") & """>" & rs(getlangid("sectionName",256)) & "</a>" & tslist
			elseif rs("rootSection")=1 then
				tslist = " &raquo; <a href=""products.asp?cat="& rs("sectionID") & """>" & rs(getlangid("sectionName",256)) & "</a>" & tslist
			else
				tslist = " &raquo; <a href=""categories.asp?cat="& rs("sectionID") & """>" & rs(getlangid("sectionName",256)) & "</a>" & tslist
			end if
			thetopts = rs("topSection")
			topsectionids = topsectionids & "," & thetopts
		else
			tslist = "Top Section Not Available" & tslist
			rs.Close
			exit for
		end if
		rs.Close
	end if
next
if xxAlProd<>"" then tslist = tslist & " &raquo; <a href=""products.asp"&IIfVr(theid="0","","?cat="&theid)&""">"&xxAlProd&"</a>"
sSQL = "SELECT sectionID,"&getlangid("sectionName",256)&",rootSection,sectionImage,sectionOrder,"&getlangid("sectionDescription",512)&",sectionurl FROM sections WHERE topSection=" & theid & " AND sectionDisabled<="&minloglevel&" ORDER BY sectionOrder"
rs.Open sSQL,cnn,0,1
if NOT success OR rs.eof OR rs.bof then
	success=false
	mess1 = "<p>&nbsp;</p>" & xxNoCats
else
	alldata=rs.getrows
	success=true
	if xxClkPrd<>"" then mess1 = xxClkPrd & "<br />&nbsp;"
end if
rs.Close
if (usecategoryformat=1 OR usecategoryformat=2) then numcolumns=2*categorycolumns else numcolumns=categorycolumns
%>
      <table border="0" cellspacing="<%=maintablespacing%>" cellpadding="<%=maintablepadding%>" width="<%=maintablewidth%>" bgcolor="<%=maintablebg%>" align="center">
        <tr> 
          <td width="100%">
            <table width="<%=innertablewidth%>" border="<%=border%>" cellspacing="<%=innertablespacing%>" cellpadding="<%=innertablepadding%>" bgcolor="<%=innertablebg%>">
<%	if mess1<>"" then %>
			  <tr>
				<td align="center"<% if numcolumns>1 then response.write " colspan='"&numcolumns&"'"%>>
				  <p><strong><%=mess1%></strong></p>
				</td>
			  </tr>
<%
	end if
if nowholesalediscounts=true AND Session("clientUser")<>"" then
	if ((Session("clientActions") AND 8) = 8) OR ((Session("clientActions") AND 16) = 16) then noshowdiscounts=true
end if
if success then
	tdt = Date()
	if noshowdiscounts<>true then
		if theid="0" then
			sSQL = "SELECT DISTINCT "&getlangid("cpnName",1024)&" FROM coupons WHERE (cpnSitewide=1 OR cpnSitewide=2) AND cpnNumAvail>0 AND cpnEndDate>="&datedelim&VSUSDate(tdt)&datedelim&" AND cpnIsCoupon=0"
		else
			sSQL = "SELECT DISTINCT "&getlangid("cpnName",1024)&" FROM coupons LEFT OUTER JOIN cpnassign ON coupons.cpnID=cpnassign.cpaCpnID WHERE (((cpnSitewide=0 OR cpnSitewide=3) AND cpaType=1 AND cpaAssignment IN ('"&Replace(topsectionids,",","','")&"')) OR cpnSitewide=1 OR cpnSitewide=2) AND cpnNumAvail>0 AND cpnEndDate>="&datedelim&VSUSDate(tdt)&datedelim&" AND cpnIsCoupon=0"
		end if
		rs.Open sSQL,cnn,0,1
		if NOT rs.EOF then %>
			  <tr>
				<td align="left"<% if numcolumns>1 then response.write " colspan='"&numcolumns&"'"%>>
				  <p><strong><%=xxDsCat%></strong><br /><font color="#FF0000" size="1">
				  <%	do while NOT rs.EOF
							response.write rs(getlangid("cpnName",1024)) & "<br />"
							rs.MoveNext
						loop %>&nbsp;</font></p>
				</td>
			  </tr>
<%		end if
		rs.Close
	end if
	response.write "</table>"
	if IsEmpty(showcategories) OR showcategories=true then
		response.write "<table width="""&innertablewidth&""" border="""&border&""" cellspacing="""&innertablespacing&""" cellpadding="""&innertablepadding&""" bgcolor="""&innertablebg&"""><tr>"
		if allproductsimage<>"" then response.write "<td class=""catimage"" width=""5%"" align=""right""><a href='products.asp'><img class=""catimage"" src="""&allproductsimage&""" border=""0"" alt="""&xxAlProd&""" /></a>" & afterimage & "</td>"
		response.write "<td class=""catnavigation"">"
		response.write "<p class=""catnavigation""><strong>" & tslist & "</strong></p>"
		response.write "<p class=""navdesc"">" & xxAlPrCa & catseparator & "</p>"
		response.write "</td></tr>" & vbCrLf
		response.write "</table>"
	end if
	response.write "<table width="""&innertablewidth&""" border="""&border&""" cellspacing="""&IIfVr(usecategoryformat=1 AND categorycolumns>1,0,innertablespacing)&""" cellpadding="""&IIfVr(usecategoryformat=1 AND categorycolumns>1,0,innertablepadding)&""" bgcolor="""&innertablebg&""">"
	tdt = Date()
	columncount=0
	FOR rowcounter= 0 TO ubound(alldata,2)
		if trim(alldata(6,rowcounter)&"")<>"" then
			startlink="<a href='"&alldata(6,rowcounter)&"'>"
		elseif alldata(2,rowcounter)=0 then
			startlink="<a href='categories.asp?cat="&alldata(0,rowcounter)&"'>"
		else
			startlink="<a href='products.asp?cat="&alldata(0,rowcounter)&"'>"
		end if
		sSQL = "SELECT DISTINCT "&getlangid("cpnName",1024)&" FROM coupons LEFT OUTER JOIN cpnassign ON coupons.cpnID=cpnassign.cpaCpnID WHERE (cpnSitewide=0 OR cpnSitewide=3) AND cpnNumAvail>0 AND cpnEndDate>="&datedelim&VSUSDate(tdt)&datedelim&" AND cpnIsCoupon=0 AND cpaType=1 AND cpaAssignment='"&alldata(0,rowcounter)&"'"
		alldiscounts = ""
		if noshowdiscounts<>true then
			rs.Open sSQL,cnn,0,1
			do while NOT rs.EOF
				alldiscounts = alldiscounts & rs(getlangid("cpnName",1024)) & "<br />"
				rs.MoveNext
			loop
			rs.Close
		end if
		secdesc = Trim(alldata(5,rowcounter))
		noimage = (Trim(alldata(3,rowcounter))="" OR IsNull(alldata(3,rowcounter)))
		if columncount=0 then response.write "<tr>"
		if usecategoryformat=1 AND categorycolumns>1 then response.write "<td width=""" & cellwidth & "%"" valign=""top""><table width=""100%"" border="""&border&""" cellspacing="""&innertablespacing&""" cellpadding="""&innertablepadding&"""><tr>"
		if (usecategoryformat=1 OR usecategoryformat=2) AND NOT noimage then
			cellwidth = cellwidth - 5
			response.write "<td class=""catimage"" width=""5%"" align=""right"">" & startlink&"<img alt="""&replace(alldata(1,rowcounter),"""","")&""" class=""catimage"" src="""&alldata(3,rowcounter)&""" border=""0"" /></a>" & afterimage & "</td>"
		end if
		response.write "<td class=""catname"" width=""" & IIfVr(usecategoryformat=1 AND categorycolumns>1,95,cellwidth) & "%""" & IIfVr((usecategoryformat=1 OR usecategoryformat=2) AND noimage," colspan='2'","") & ">"
		if (usecategoryformat=1 OR usecategoryformat=2) AND NOT noimage then cellwidth = cellwidth + 5
		if usecategoryformat<>1 AND usecategoryformat<>2 AND NOT noimage then response.write startlink&"<img alt="""&replace(alldata(1,rowcounter),"""","")&""" class=""catimage"" src="""&alldata(3,rowcounter)&""" border=""0"" /></a>" & afterimage
		response.write "<p class=""catname""><strong>"&startlink&alldata(1,rowcounter)&"</a>"&xxDot&"</strong>"
		if alldiscounts<>"" then response.write " <font color=""#FF0000""><strong>"&xxDsApp&"</strong><br /><font size=""1"">" & alldiscounts & "</font></font>"
		if secdesc <> "" then response.write "</p>" else response.write catseparator & "</p>"
		if secdesc <> "" then response.write beforedesc & "<p class=""catdesc"">" & secdesc & catseparator & "</p>"
		response.write "</td>" & vbCrLf
		if usecategoryformat=1 AND categorycolumns>1 then response.write "</tr></table></td>"
		columncount = columncount + 1
		if columncount=categorycolumns then
			response.write "</tr>"
			columncount=0
		end if
	next
	if columncount<categorycolumns AND columncount<>0 then
		do while columncount<categorycolumns
			response.write "<td " & IIfVr(usecategoryformat=2, " colspan='2'" , "") & ">&nbsp;</td>"
			columncount = columncount + 1
		loop
		response.write "</tr>"
	end if
end if
cnn.Close
response.write "</table><table width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""0"" bgcolor="""">"
response.write "<tr><td><img src=""images/clearpixel.gif"" width=""300"" height=""1"" alt="""" /></td></tr>"
set rs = nothing
set cnn = nothing
%>
            </table>
          </td>
        </tr>
      </table>
