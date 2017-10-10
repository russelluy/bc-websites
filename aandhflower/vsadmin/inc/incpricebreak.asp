<%
if storesessionvalue="" then storesessionvalue="virtualstore"
if Session("loggedon") <> storesessionvalue OR disallowlogin=TRUE then response.end
Dim sSQL,rs,alldata,alladmin,success,cnn,rowcounter,errmsg
success=true
maxcatsperpage = 100
maxpricebreaks = 25
Set rs = Server.CreateObject("ADODB.RecordSet")
Set rs2 = Server.CreateObject("ADODB.RecordSet")
Set cnn=Server.CreateObject("ADODB.Connection")
cnn.open sDSN
sSQL = ""
dropdown = (request.form("proddrop")="OK")
if request.form("posted")="1" then
	if request.form("act")="delete" then
		sSQL = "DELETE FROM pricebreaks WHERE pbProdID='" & trim(replace(request.form("id"),"'","''")) & "'"
		cnn.Execute(sSQL)
		response.write "<meta http-equiv=""refresh"" content=""2; url=adminpricebreak.asp?pg="& request.form("pg") &""">"
	elseif request.form("act")="domodify" then
		theprod=trim(request.form("pid"))
		sSQL = "SELECT pID FROM products WHERE pID='" & replace(theprod,"'","''") & "'"
		rs.Open sSQL,cnn,0,1
		if rs.EOF then
			success=false
			errmsg = "The specified product id (" & theprod & ") does not exist."
		end if
		rs.Close
		if success then
			cnn.Execute("DELETE FROM pricebreaks WHERE pbProdID='" & theprod & "'")
			for index=1 to maxpricebreaks
				thequant=trim(request.form("quant"&index))
				if NOT IsNumeric(thequant) then thequant=0
				price=trim(request.form("price"&index))
				if NOT IsNumeric(price) then price=0
				wprice=trim(request.form("wprice"&index))
				if NOT IsNumeric(wprice) then wprice=0
				if thequant<>0 AND (price<>0 OR wprice<>0) then
					sSQL = "INSERT INTO pricebreaks (pbProdID,pbQuantity,pPrice,pWholesalePrice) VALUES ('"&replace(theprod,"'","''")&"',"
					sSQL = sSQL & thequant & ","
					sSQL = sSQL & price & ","
					sSQL = sSQL & wprice & ")"
					cnn.Execute(sSQL)
				end if
			next
			response.write "<meta http-equiv=""refresh"" content=""2; url=adminpricebreak.asp?pg="& request.form("pg") &""">"
		end if
	elseif request.form("act")="doaddnew" then
		theprod=trim(request.form("pid"))
		sSQL = "SELECT pbProdID FROM pricebreaks WHERE pbProdID='" & replace(theprod,"'","''") & "'"
		rs.Open sSQL,cnn,0,1
		if NOT rs.EOF then
			success=false
			errmsg = "Price breaks already exist for this product id. You should use the ""Modify"" option on the price breaks admin page"
		end if
		rs.Close
		sSQL = "SELECT pID FROM products WHERE pID='" & replace(theprod,"'","''") & "'"
		rs.Open sSQL,cnn,0,1
		if rs.EOF then
			success=false
			errmsg = "The specified product id (" & theprod & ") does not exist."
		end if
		rs.Close
		if success then
			for index=1 to maxpricebreaks
				thequant=trim(request.form("quant"&index))
				if NOT IsNumeric(thequant) then thequant=0
				price=trim(request.form("price"&index))
				if NOT IsNumeric(price) then price=0
				wprice=trim(request.form("wprice"&index))
				if NOT IsNumeric(wprice) then wprice=0
				if thequant<>0 AND (price<>0 OR wprice<>0) then
					sSQL = "INSERT INTO pricebreaks (pbProdID,pbQuantity,pPrice,pWholesalePrice) VALUES ('"&replace(theprod,"'","''")&"',"
					sSQL = sSQL & thequant & ","
					sSQL = sSQL & price & ","
					sSQL = sSQL & wprice & ")"
					cnn.Execute(sSQL)
				end if
			next
			response.write "<meta http-equiv=""refresh"" content=""2; url=adminpricebreak.asp?pg="& request.form("pg") &""">"
		end if
	end if
end if
%>
<script language="javascript" type="text/javascript">
<!--
function formvalidator(theForm)
{
<% if dropdown then %>
  if (theForm.pid.selectedIndex == 0){
    alert("<%=yyPlsSel%> \"<%=yyPrId%>\".");
<% else %>
  if (theForm.pid.value == ""){
    alert("<%=yyPlsEntr%> \"<%=yyPrId%>\".");
<% end if %>
    theForm.pid.focus();
    return (false);
  }
  return (true);
}
//-->
</script>
      <table border="0" cellspacing="<%=maintablespacing%>" cellpadding="<%=maintablepadding%>" width="<%=maintablewidth%>" bgcolor="<%=maintablebg%>" align="center">
<% if request.form("posted")="1" AND (request.form("act")="modify" OR request.form("act")="clone") then
		if dropdown AND request.form("act")="clone" then
			allprodids=""
			sSQL = "SELECT pID FROM products LEFT JOIN pricebreaks ON products.pID=pricebreaks.pbProdID WHERE pbProdID IS NULL ORDER BY pID"
			rs.Open sSQL,cnn,0,1
			if NOT rs.EOF then
				allprodids=rs.getrows
			end if
			rs.Close
		end if
%>
        <tr>
		<form name="mainform" method="post" action="adminpricebreak.asp" onsubmit="return formvalidator(this)">
		  <td width="100%" align="center">
			<input type="hidden" name="posted" value="1" />
			<% if request.form("act")="clone" then %>
			<input type="hidden" name="act" value="doaddnew" />
			<% else %>
			<input type="hidden" name="act" value="domodify" />
			<input type="hidden" name="pid" value="<%=Request.Form("id")%>" />
			<% end if %>
			<input type="hidden" name="pg" value="<%=Request.Form("pg")%>" />
            <table width="320" border="0" cellspacing="0" cellpadding="1" bgcolor="">
			  <tr> 
                <td colspan="3" align="center"><strong><%=yyPBKAdm%></strong><br />&nbsp;</td>
			  </tr>
			  <tr>
				<td colspan="3" align="center" valign="top"><strong><%=yyPBFID%>:</strong> <%
				if dropdown AND request.form("act")="clone" then
					response.write "<select size=""1"" name=""pid""><option value="""">"&yySelect&"</option>"
					if IsArray(allprodids) then
						for index=0 to UBOUND(allprodids, 2)
							response.write "<option value="""&allprodids(0,index)&""">"&allprodids(0,index)&"</option>"&vbCrLf
						next
					end if
					response.write "</select>"
				elseif request.form("act")="clone" then
					response.write "<input type=""text"" name=""pid"" size=""20"" />"
				else
					response.write Request.Form("id")
				end if %></td>
			  </tr>
			  <tr>
				<td align="center" valign="top"><strong><font size="1"><%=yyQuaFro%></font></strong></td>
				<td align="center" valign="top"><strong><font size="1"><%=yyPrPri%></font></strong></td>
				<td align="center" valign="top"><strong><font size="1"><%=yyWhoPri%></font></strong></td>
			  </tr>
<%			sSQL = "SELECT pbQuantity,pPrice,pWholesalePrice FROM pricebreaks WHERE pbProdID='"&trim(replace(Request.Form("id"),"'","''"))&"' ORDER BY pbQuantity"
			index=1
			rs.Open sSQL,cnn,0,1
			do while NOT rs.EOF %>
			  <tr>
				<td align="center" valign="top"><input type="text" name="quant<%=index%>" size="12" value="<%=rs("pbQuantity")%>" /></td>
				<td align="center" valign="top"><input type="text" name="price<%=index%>" size="12" value="<%=rs("pPrice")%>" /></td>
				<td align="center" valign="top"><input type="text" name="wprice<%=index%>" size="12" value="<%=rs("pWholesalePrice")%>" /></td>
			  </tr>
<%				rs.MoveNext
				index=index+1
			loop
			rs.Close
			for index2=index to maxpricebreaks %>
			  <tr>
				<td align="center" valign="top"><input type="text" name="quant<%=index2%>" size="12" value="" /></td>
				<td align="center" valign="top"><input type="text" name="price<%=index2%>" size="12" value="" /></td>
				<td align="center" valign="top"><input type="text" name="wprice<%=index2%>" size="12" value="" /></td>
			  </tr>
<%			next %>
			  <tr>
                <td width="100%" colspan="3" align="center"><br /><input type="submit" value="<%=yySubmit%>" /></td>
			  </tr>
			  <tr> 
                <td width="100%" colspan="3" align="center"><br /><a href="admin.asp"><strong><%=yyAdmHom%></strong></a><br />&nbsp;</td>
			  </tr>
            </table></td>
		  </form>
        </tr>
<% elseif request.form("posted")="1" AND request.form("act")="addnew" then
		if dropdown then
			allprodids=""
			sSQL = "SELECT pID FROM products LEFT JOIN pricebreaks ON products.pID=pricebreaks.pbProdID WHERE pbProdID IS NULL ORDER BY pID"
			rs.Open sSQL,cnn,0,1
			if NOT rs.EOF then
				allprodids=rs.getrows
			end if
			rs.Close
		end if
%>
        <tr>
		<form name="mainform" method="post" action="adminpricebreak.asp" onsubmit="return formvalidator(this)">
		  <td width="100%" align="center">
			<input type="hidden" name="posted" value="1" />
			<input type="hidden" name="act" value="doaddnew" />
			<input type="hidden" name="pg" value="<%=Request.Form("pg")%>" />
            <table width="320" border="0" cellspacing="0" cellpadding="1" bgcolor="">
			  <tr> 
                <td colspan="3" align="center"><strong><%=yyPBKAdm%></strong><br />&nbsp;</td>
			  </tr>
			  <tr>
				<td colspan="3" align="center" valign="top"><strong><%=yyPBFID%>:</strong> 
			<%	if dropdown then
					response.write "<select size=""1"" name=""pid""><option value="""">"&yySelect&"</option>"
					if IsArray(allprodids) then
						for index=0 to UBOUND(allprodids, 2)
							response.write "<option value="""&allprodids(0,index)&""">"&allprodids(0,index)&"</option>"&vbCrLf
						next
					end if
					response.write "</select>"
				else
					response.write "<input type=""text"" name=""pid"" size=""20"" />"
				end if %></td>
			  </tr>
			  <tr>
				<td align="center" valign="top"><strong><font size="1"><%=yyQuaFro%></font></strong></td>
				<td align="center" valign="top"><strong><font size="1"><%=yyPrPri%></font></strong></td>
				<td align="center" valign="top"><strong><font size="1"><%=yyWhoPri%></font></strong></td>
			  </tr>
<%			for index=1 to maxpricebreaks %>
			  <tr>
				<td align="center" valign="top"><input type="text" name="quant<%=index%>" size="12" value="" /></td>
				<td align="center" valign="top"><input type="text" name="price<%=index%>" size="12" value="" /></td>
				<td align="center" valign="top"><input type="text" name="wprice<%=index%>" size="12" value="" /></td>
			  </tr>
<%			next %>
			  <tr>
                <td width="100%" colspan="3" align="center"><br /><input type="submit" value="<%=yySubmit%>" /></td>
			  </tr>
			  <tr> 
                <td width="100%" colspan="3" align="center"><br /><a href="admin.asp"><strong><%=yyAdmHom%></strong></a><br />&nbsp;</td>
			  </tr>
            </table></td>
		  </form>
        </tr>
<% elseif request.form("posted")="1" AND success then %>
        <tr>
          <td width="100%">
			<table width="100%" border="0" cellspacing="0" cellpadding="3" bgcolor="">
			  <tr> 
                <td width="100%" colspan="2" align="center"><br /><strong><%=yyUpdSuc%></strong><br /><br /><%=yyNowFrd%><br /><br />
                        <%=yyNoAuto%><A href="adminpricebreak.asp"><strong><%=yyClkHer%></strong></a>.<br />
                        <br />
				<img src="../images/clearpixel.gif" width="300" height="3" alt="" />
                </td>
			  </tr>
			</table></td>
        </tr>
<% elseif request.form("posted")="1" then %>
        <tr>
          <td width="100%">
			<table width="100%" border="0" cellspacing="0" cellpadding="3" bgcolor="">
			  <tr> 
                <td width="100%" colspan="2" align="center"><br /><font color="#FF0000"><strong><%=yyOpFai%></strong></font><br /><br /><%=errmsg%><br /><br />
				<a href="javascript:history.go(-1)"><strong><%=yyClkBac%></strong></a></td>
			  </tr>
			</table></td>
        </tr>
<% else %>
<script language="javascript" type="text/javascript">
<!--
function checkcontrol(evt){
<%	netnav = false
	if instr(Request.ServerVariables("HTTP_USER_AGENT"), "Gecko") > 0 then netnav = true
	if netnav then %>
if(evt.ctrlKey || evt.altKey){
document.mainform.proddrop.checked=true;
}
<%	else %>
theevnt=window.event;
if(theevnt.ctrlKey){
document.mainform.proddrop.checked=true;
}
<%	end if %>
}
function mrk(id) {
	document.mainform.id.value = id;
	document.mainform.act.value = "modify";
	document.mainform.submit();
}
function newrec(evt) {
	checkcontrol(evt);
	document.mainform.act.value = "addnew";
	document.mainform.submit();
}
function crk(id, evt) {
	checkcontrol(evt);
	document.mainform.id.value = id;
	document.mainform.act.value = "clone";
	document.mainform.submit();
}
function drk(id) {
cmsg = "<%=yyConDel%>\n"
if (confirm(cmsg)) {
	document.mainform.id.value = id;
	document.mainform.act.value = "delete";
	document.mainform.submit();
}
}
// -->
</script>
        <tr>
		  <form name="mainform" method="post" action="adminpricebreak.asp">
		  <td width="100%">
			<input type="hidden" name="posted" value="1" />
			<input type="hidden" name="act" value="xxxxx" />
			<input type="hidden" name="id" value="xxxxx" />
			<input type="hidden" name="pg" value="<%=Request.QueryString("pg")%>" />
			<input type="hidden" name="selectedq" value="1" />
			<input type="hidden" name="newval" value="1" />
            <table width="100%" border="0" cellspacing="0" cellpadding="1" bgcolor="">
			  <tr> 
                <td width="100%" colspan="5" align="center"><strong><%=yyPBKAdm%></strong><br />&nbsp;</td>
			  </tr>
<%
Function writepagebar(CurPage, iNumPages)
	Dim sLink, i, sStr, startPage, endPage
	sLink = "<a href='adminpricebreak.asp?pg="
	startPage = vrmax(1,CInt(Int(CDbl(CurPage)/10.0)*10))
	endPage = vrmin(iNumPages,CInt(Int(CDbl(CurPage)/10.0)*10)+10)
	if CurPage > 1 then
		sStr = sLink & "1" & "'><strong><font face='Verdana'>&laquo;</font></strong></a> " & sLink & CurPage-1 & "'>Previous</a> | "
	else
		sStr = "<strong><font face='Verdana'>&laquo;</font></strong> Previous | "
	end if
	for i=startPage to endPage
		if i=CurPage then
			sStr = sStr & i & " | "
		else
			sStr = sStr & sLink & i & "'>"
			if i=startPage AND i > 1 then sStr=sStr&"..."
			sStr = sStr & i
			if i=endPage AND i < iNumPages then sStr=sStr&"..."
			sStr = sStr & "</a> | "
		end if
	next
	if CurPage < iNumPages then
		writepagebar = sStr & sLink & CurPage+1 & "'>Next</a> " & sLink & iNumPages & "'><strong><font face='Verdana'>&raquo;</font></strong></a>"
	else
		writepagebar = sStr & " Next <strong><font face='Verdana'>&raquo;</font></strong>"
	end if
End function
	If Request.QueryString("pg") = "" Then
		CurPage = 1
	Else
		CurPage = Int(Request.QueryString("pg"))
	End If
	sSQL = "SELECT DISTINCT pbProdID,pName FROM pricebreaks INNER JOIN products ON pricebreaks.pbProdID=products.pID ORDER BY pbProdID"
	rs2.CursorLocation = 3 ' adUseClient
	rs2.CacheSize = maxcatsperpage
	rs2.Open sSQL,cnn
	if NOT rs2.EOF then
		rs2.MoveFirst
		rs2.PageSize = maxcatsperpage
		rs2.AbsolutePage = CurPage
		islooping=false
		noproducts=false
		hascatinprodsection=false
		rowcounter=0
		totnumrows=rs2.RecordCount
		iNumOfPages = Int((totnumrows + (maxcatsperpage-1)) / maxcatsperpage)
		If iNumOfPages > 1 Then Response.Write "<tr><td align=""center"" colspan=""5"">" & writepagebar(CurPage, iNumOfPages) & "<br /><br /></td></tr>"
%>
			  <tr>
				<td align="left"><strong><%=yyPrId%></strong> <input type="checkbox" name="proddrop" value="OK" /></td>
				<td align="left"><strong><%=yyPrName%></strong></td>
				<td width="5%" align="center"><font size="1"><strong><%=yyClone%></strong></font></td>
				<td width="5%" align="center"><font size="1"><strong><%=yyModify%></strong></font></td>
				<td width="5%" align="center"><font size="1"><strong><%=yyDelete%></strong></font></td>
			  </tr>
<%		do while NOT rs2.EOF AND rowcounter < maxcatsperpage
			if bgcolor="#E7EAEF" then bgcolor="#FFFFFF" else bgcolor="#E7EAEF"%>
<tr bgcolor="<%=bgcolor%>">
<td><%=rs2("pbProdID")%></td>
<td><%=rs2("pName")%></td>
<td><input type="button" value="<%=yyClone%>" onclick="crk('<%=replace(rs2("pbProdID"),"'","\'")%>', event)" /></td>
<td><input type="button" value="<%=yyModify%>" onclick="mrk('<%=replace(rs2("pbProdID"),"'","\'")%>')" /></td>
<td><input type="button" value="<%=yyDelete%>" onclick="drk('<%=replace(rs2("pbProdID"),"'","\'")%>')" /></td>
</tr><%		rowcounter=rowcounter+1
			rs2.MoveNext
		loop
		If iNumOfPages > 1 Then Response.Write "<tr><td align=""center"" colspan=""5""><br />" & writepagebar(CurPage, iNumOfPages) & "</td></tr>"
	else
%>
			  <tr><td width="100%" colspan="5" align="center"><br /> <input type="checkbox" name="proddrop" value="OK" /> <strong><%=yyNoPBK%><br />&nbsp;</td></tr>
<%
	end if
	rs2.Close
%>
			  <tr> 
                <td width="100%" colspan="5" align="center"><br /><strong><%=yyPBKNew%></strong>&nbsp;&nbsp;<input type="button" value="<%=yyNewPBK%>" onclick="newrec(event)" /><br />&nbsp;</td>
			  </tr>
			  <tr> 
                <td width="100%" colspan="5" align="center"><br />
                          <a href="admin.asp"><strong><%=yyAdmHom%></strong></a><br />
				<img src="../images/clearpixel.gif" width="300" height="3" alt="" /></td>
			  </tr>
            </table></td>
		  </form>
        </tr>
<% end if
cnn.Close
set rs = nothing
set rs2 = nothing
set cnn = nothing
%>
      </table>