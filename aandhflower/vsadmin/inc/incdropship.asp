<%
if storesessionvalue="" then storesessionvalue="virtualstore"
if Session("loggedon") <> storesessionvalue OR disallowlogin=TRUE then response.end
Dim sSQL,rs,cnn,success,showaccount,addsuccess,alldata,index,allcountries,rowcounter,sd,ed,errmsg
addsuccess = true
success = true
showaccount = true
dorefresh = false
if dateadjust="" then dateadjust=0
thedate = DateAdd("h",dateadjust,Now())
thedate = DateSerial(year(thedate),month(thedate),day(thedate))
Set rs = Server.CreateObject("ADODB.RecordSet")
Set cnn=Server.CreateObject("ADODB.Connection")
cnn.open sDSN
alreadygotadmin = getadminsettings()
if request.form("act")="domodify" then
	sSQL = "UPDATE dropshipper SET dsEmail='"&replace(trim(request.form("Email")),"'","''") & "'," & _
		"dsName='"&replace(trim(request.form("Name")),"'","''") & "'," & _
		"dsAddress='"&replace(trim(request.form("Address")),"'","''") & "'," & _
		"dsCity='"&replace(trim(request.form("City")),"'","''") & "'," & _
		"dsState='"&replace(trim(request.form("State")),"'","''") & "'," & _
		"dsCountry='"&replace(trim(request.form("Country")),"'","''") & "'," & _
		"dsZip='"&replace(trim(request.form("Zip")),"'","''") & "'," & _
		"dsAction="&replace(trim(request.form("dsAction")),"'","''") & " " & _
		"WHERE dsID=" & replace(trim(request.form("dsID")),"'","")
	cnn.Execute(sSQL)
	dorefresh=true
elseif request.form("act")="doaddnew" then
	sSQL = "INSERT INTO dropshipper (dsEmail,dsName,dsAddress,dsCity,dsState,dsCountry,dsZip,dsAction) VALUES (" & _
		"'"&replace(trim(request.form("Email")),"'","''") & "'," & _
		"'"&replace(trim(request.form("Name")),"'","''") & "'," & _
		"'"&replace(trim(request.form("Address")),"'","''") & "'," & _
		"'"&replace(trim(request.form("City")),"'","''") & "'," & _
		"'"&replace(trim(request.form("State")),"'","''") & "'," & _
		"'"&replace(trim(request.form("Country")),"'","''") & "'," & _
		"'"&replace(trim(request.form("Zip")),"'","''") & "'," & _
		""&replace(trim(request.form("dsAction")),"'","''") & ")"
	cnn.Execute(sSQL)
	dorefresh=true
elseif request.form("act")="delete" then
	sSQL = "DELETE FROM dropshipper WHERE dsID=" & trim(request.form("id"))
	cnn.Execute(sSQL)
	dorefresh=true
end if
if dorefresh then
	response.write "<meta http-equiv=""refresh"" content=""2; url=admindropship.asp"">"
end if
if dorefresh then
%>
      <table border="0" cellspacing="0" cellpadding="0" width="100%" bgcolor="" align="center">
        <tr>
          <td width="100%">
			<table width="100%" border="0" cellspacing="0" cellpadding="3" bgcolor="">
			  <tr> 
                <td width="100%" colspan="2" align="center"><br /><strong><%=yyUpdSuc%></strong><br /><br /><%=yyNowFrd%><br /><br />
                        <%=yyNoAuto%> <a href="admindropship.asp"><strong><%=yyClkHer%></strong></a>.<br />
                        <br />
				<img src="../images/clearpixel.gif" width="300" height="3" alt="" />
                </td>
			  </tr>
			</table></td>
        </tr>
      </table>
<%
elseif trim(request.form("act"))="modify" OR trim(request.form("act"))="addnew" then
	if trim(request.form("act"))="modify" then
		dsID=trim(request.form("id"))
		sSQL = "SELECT dsName,dsAddress,dsCity,dsState,dsZip,dsCountry,dsEmail,dsAction FROM dropshipper WHERE dsID="&dsID
		rs.Open sSQL,cnn,0,1
		if NOT rs.EOF then
			dsName = rs("dsName")
			dsAddress = rs("dsAddress")
			dsCity = rs("dsCity")
			dsState = rs("dsState")
			dsZip = rs("dsZip")
			dsCountry = rs("dsCountry")
			dsEmail = rs("dsEmail")
			dsAction = rs("dsAction")
		end if
		rs.Close
	end if
%>
<script language="javascript" type="text/javascript">
<!--
function checkform(frm)
{
if(frm.name.value=="")
{
	alert("<%=yyPlsEntr%> \"<%=yyName%>\".");
	frm.name.focus();
	return (false);
}
if(frm.email.value=="")
{
	alert("<%=yyPlsEntr%> \"<%=yyEmail%>\".");
	frm.email.focus();
	return (false);
}
if(frm.address.value=="")
{
	alert("<%=yyPlsEntr%> \"<%=yyAddress%>\".");
	frm.address.focus();
	return (false);
}
if(frm.city.value=="")
{
	alert("<%=yyPlsEntr%> \"<%=yyCity%>\".");
	frm.city.focus();
	return (false);
}
if(frm.state.value=="")
{
	alert("<%=yyPlsEntr%> \"<%=yyState%>\".");
	frm.state.focus();
	return (false);
}
if(frm.zip.value=="")
{
	alert("<%=yyPlsEntr%> \"<%=yyZip%>\".");
	frm.zip.focus();
	return (false);
}
return (true);
}
//-->
</script>
	  <table border="0" cellspacing="0" cellpadding="0" width="100%" bgcolor="" align="center">
        <tr> 
          <td width="100%">
		    <form method="post" action="admindropship.asp" onsubmit="return checkform(this)">
		<%	if trim(request.form("act"))="modify" then %>
			<input type="hidden" name="act" value="domodify" />
		<%	else %>
			<input type="hidden" name="act" value="doaddnew" />
		<%	end if %>
			<input type="hidden" name="dsID" value="<%=dsID%>" />
			  <table width="100%" border="0" cellspacing="0" cellpadding="3" bgcolor="">
				<tr>
				  <td width="100%" align="center" colspan="4"><strong><%=yyDSAdm%></strong><br /></td>
				</tr>
				<tr>
				  <td width="20%" align="right"><strong><font color='#FF0000'>*</font><%=yyName%>:</strong></td>
				  <td width="30%" align="left"><input type="text" name="name" size="20" value="<%=dsName%>" /></td>
				  <td width="20%" align="right"><strong><font color='#FF0000'>*</font><%=yyEmail%>:</strong></td>
				  <td width="30%" align="left"><input type="text" name="email" size="25" value="<%=dsEmail%>" /></td>
				</tr>
				<tr>
				  <td align="right"><strong><font color='#FF0000'>*</font><%=yyAddress%>:</strong></td>
				  <td align="left"><input type="text" name="address" size="20" value="<%=dsAddress%>" /></td>
				  <td align="right"><strong><font color='#FF0000'>*</font><%=yyCity%>:</strong></td>
				  <td align="left"><input type="text" name="city" size="20" value="<%=dsCity%>" /></td>
				</tr>
				<tr>
				  <td align="right"><strong><font color='#FF0000'>*</font><%=yyState%>:</strong></td>
				  <td align="left"><input type="text" name="state" size="20" value="<%=dsState%>" /></td>
				  <td align="right"><strong><font color='#FF0000'>*</font><%=yyCountry%>:</strong></td>
				  <td align="left"><select name="country" size="1">
<%
Sub show_countries(tcountry)
	if NOT IsArray(allcountries) then
		sSQL = "SELECT countryName FROM countries ORDER BY countryOrder DESC, countryName"
		rs.Open sSQL,cnn,0,1
		allcountries=rs.getrows
		rs.Close
	end if
	for rowcounter=0 to UBOUND(allcountries,2)
		response.write "<option value='"&allcountries(0,rowcounter)&"'"
		if tcountry=allcountries(0,rowcounter) then
			response.write " selected"
		end if
		response.write ">"&allcountries(0,rowcounter)&"</option>"&vbCrLf
	next
End Sub
show_countries(dsCountry)
%>
					</select>
				  </td>
				</tr>
				<tr>
				  <td align="right"><strong><font color='#FF0000'>*</font><%=yyZip%>:</strong></td>
				  <td align="left"><input type="text" name="zip" size="10" value="<%=dsZip%>" /></td>
				  <td align="right"><strong><%=yyActns%>:</strong></td>
				  <td align="left"><select name="dsAction" size="1">
					<option value="0"><%=yyNoAct%></option>
					<option value="1"<% if dsAction=1 then response.write " selected"%>><%=yySendEM%></option>
					</select>
				  </td>
				</tr>
				<tr>
				  <td width="100%" colspan="4">&nbsp;<br />
					<font size="1"><ul><li><%=yyDSInf%></li></ul></font>
				  </td>
				</tr>
				<tr>
				  <td width="50%" align="center" colspan="4"><input type="submit" value="<%=yySubmit%>" /> <input type="reset" value="<%=yyReset%>" /> </td>
				</tr>
			  </table>
			</form>
		  </td>
        </tr>
      </table>
<%
else
	if Request("sd") = "" then
		sd=DateSerial(DatePart("yyyy",thedate),DatePart("m",thedate),1)
	else
		sd=Request("sd")
	end if
	if Request("ed") = "" then
		ed=thedate
	else
		ed=Request("ed")
	end if
	on error resume next
	sd = DateValue(sd)
	ed = Datevalue(ed)
	if err.number <> 0 then
		sd = DateSerial(DatePart("yyyy",thedate),DatePart("m",thedate),1)
		ed = thedate
		success=false
		errmsg=yyDatInv
	end if
	on error goto 0
	tdt = DateValue(sd)
	tdt2 = DateValue(ed)+1
	if mysqlserver=true then
		sSQL = "SELECT dsID,dsName,dsEmail,0 FROM dropshipper ORDER BY dsName"
	else
		sSQL = "SELECT dsID,dsName,dsEmail,(SELECT SUM(cartProdPrice*cartQuantity) FROM cart INNER JOIN products ON cart.cartProdID=products.pID WHERE dsID=pDropship AND cartCompleted=1 AND cartDateAdded BETWEEN "&datedelim&VSUSDate(tdt)&datedelim&" AND "&datedelim&VSUSDate(tdt2)&datedelim&") FROM dropshipper ORDER BY dsName"
	end if
	rs.Open sSQL,cnn,0,1
	if NOT rs.EOF then alldata=rs.GetRows()
	rs.Close
%>
<script language="javascript" type="text/javascript">
<!--
function modrec(id) {
	document.mainform.id.value = id;
	document.mainform.act.value = "modify";
	document.mainform.submit();
}
function newrec(id) {
	document.mainform.id.value = id;
	document.mainform.act.value = "addnew";
	document.mainform.submit();
}
function delrec(id) {
cmsg = "<%=yyConDel%>\n"
if (confirm(cmsg)) {
	document.mainform.id.value = id;
	document.mainform.act.value = "delete";
	document.mainform.submit();
}
}
// -->
</script>
	  <table border="0" cellspacing="0" cellpadding="0" width="100%" bgcolor="" align="center">
        <tr> 
          <td width="100%">
			<table width="100%" border="0" cellspacing="0" cellpadding="2" bgcolor="">
			  <tr>
				<td width="100%" align="center" colspan="6"><strong><%=yyDSAdm%></strong><br /></td>
			  </tr>
			  <form method="post" action="admindropship.asp">
			  <tr> 
                <td width="100%" colspan="6" align="center"><% if NOT success then response.write "<p><font color='#FF0000'>"&errmsg&"</font></p>" %><br /><strong><%=yyAffBet%>:</strong> <input type="text" size="12" name="sd" value="<%=sd%>" /> <strong><%=yyAnd%>:</strong> <input type="text" size="12" name="ed" value="<%=ed%>" /> <input type="submit" value="Go" /><br />&nbsp;</td>
			  </tr>
			  </form>
			  <form method="post" action="admindropship.asp">
			  <tr> 
                <td width="100%" colspan="6" align="center"><p><strong><%=yyAffFrm%>:</strong> <select name="sd" size="1"><%
					For rowcounter=0 to Day(thedate)-1
						response.write "<option value='"&thedate-rowcounter&"'"
						if thedate-rowcounter=sd then response.write " selected"
						response.write ">"&thedate-rowcounter&"</option>"&vbCrLf
						smonth=thedate-rowcounter
					Next
					For rowcounter=1 to 12
						response.write "<option value='"&DateAdd("m",0-rowcounter,smonth)&"'"
						if DateAdd("m",0-rowcounter,smonth)=sd then response.write " selected"
						response.write ">"&DateAdd("m",0-rowcounter,smonth)&"</option>"&vbCrLf
					Next
				%></select> <strong><%=yyTo%>:</strong> <select name="ed" size="1"><%
					For rowcounter=0 to Day(thedate)-1
						response.write "<option value='"&thedate-rowcounter&"'"
						if thedate-rowcounter=ed then response.write " selected"
						response.write ">"&thedate-rowcounter&"</option>"&vbCrLf
						smonth=thedate-rowcounter
					Next
					For rowcounter=1 to 12
						response.write "<option value='"&DateAdd("m",0-rowcounter,smonth)&"'"
						if DateAdd("m",0-rowcounter,smonth)=ed then response.write " selected"
						response.write ">"&DateAdd("m",0-rowcounter,smonth)&"</option>"&vbCrLf
					Next
				%></select> <input type="submit" value="Go" /><br />&nbsp;</p></td>
			  </tr>
			  </form>
			  <form name="mainform" method="post" action="admindropship.asp">
				<tr>
				  <td><strong><%=yyID%></strong></td>
				  <td><strong><%=yyName%></strong></td>
				  <td><strong><%=yyEmail%></strong></td>
				  <td align="right"><strong><%=yyTotSal%></strong></td>
				  <td align="center"><strong><%=yyModify%></strong></td>
				  <td align="center"><strong><%=yyDelete%></strong></td>
				</tr>
				<input type="hidden" name="id" value="xxx" />
				<input type="hidden" name="act" value="xxxxx" />
				<input type="hidden" name="ed" value="<%=DateValue(ed)%>" />
				<input type="hidden" name="sd" value="<%=DateValue(sd)%>" />
<%
	if NOT IsArray(alldata) then
%>
				<tr>
				  <td width="100%" align="center" colspan="6"><br />&nbsp;<br /><strong><%=yyNoAff%></strong><br />&nbsp;</td>
				</tr>
<%
	else
		for index=0 to UBOUND(alldata,2)
			if mysqlserver=true then
				sSQL = "SELECT SUM(cartProdPrice*cartQuantity) AS sumSale FROM cart INNER JOIN products ON cart.cartProdID=products.pID WHERE pDropship=" & alldata(0,index) & " AND cartCompleted=1 AND cartDateAdded BETWEEN "&datedelim&VSUSDate(tdt)&datedelim&" AND "&datedelim&VSUSDate(tdt2)&datedelim
				rs.Open sSQL,cnn,0,1
				if NOT rs.EOF then alldata(3,index)=rs("sumSale") else alldata(3,index)=0
				rs.Close
			end if
%>
				<tr>
				  <td><%=alldata(0,index)%></td>
				  <td><%=alldata(1,index)%></td>
				  <td><a href="mailto:<%=alldata(2,index)%>"><%=alldata(2,index)%></a></td>
				  <td align=right><%if NOT IsNumeric(alldata(3,index)) then response.write "-" else response.write FormatEuroCurrency(alldata(3,index))%></td>
				  <td align="center"><input type=button value="Modify" onclick="modrec('<%=alldata(0,index)%>')" /></td>
				  <td align="center"><input type=button value="Delete" onclick="delrec('<%=alldata(0,index)%>')" /></td>
				</tr><%
		next
	end if
%>
				<tr> 
				  <td width="100%" colspan="6" align="center"><br /><input type="button" value="<%=yyAddNew%>" onclick="newrec()" /><br />&nbsp;</td>
				</tr>
				<tr> 
				  <td width="100%" colspan="6" align="center"><br />
                          <a href="admin.asp"><strong><%=yyAdmHom%></strong></a><br />
				  <img src="../images/clearpixel.gif" width="300" height="3" alt="" /></td>
				</tr>
			  </form>
			</table>
		  </td>
        </tr>
      </table>
<%
end if
cnn.Close
set rs = nothing
set cnn = nothing
%>