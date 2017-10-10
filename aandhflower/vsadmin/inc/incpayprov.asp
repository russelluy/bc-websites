<%
if storesessionvalue="" then storesessionvalue="virtualstore"
if Session("loggedon") <> storesessionvalue OR disallowlogin=TRUE then response.end
Dim sSQL,rs,alldata,success,cnn,rowcounter,errmsg,data1name,data2name,isenabled,demomode,vsdetails
success=true
demomodeavailable=true
if maxloginlevels="" then maxloginlevels=5
Set rs = Server.CreateObject("ADODB.RecordSet")
Set cnn=Server.CreateObject("ADODB.Connection")
cnn.open sDSN
alreadygotadmin = getadminsettings()
if request.form("act")="domodify" then
	isenabled=0
	demomode=0
	if request.form("isenabled")="1" then isenabled=1
	if request.form("demomode")="1" then demomode=1
	sSQL = "UPDATE payprovider SET payProvShow='"&replace(request.form("showas"),"'","''")&"',payProvEnabled="&isenabled&",payProvDemo="&demomode&",payProvLevel="&request.form("payProvLevel")&","
	if Request.Form("id")="5" then ' WorldPay
		sSQL = sSQL & "payProvData1='"&replace(request.form("data1"),"'","''")&"',payProvData2='"&replace(request.form("data2"),"'","''")&"&"&replace(request.form("data3"),"'","''")&"'"
	elseif Request.Form("id")="7" then ' VeriSign
		sSQL = sSQL & "payProvData1='"&replace(request.form("data1"),"'","''")&"&"&replace(request.form("data2"),"'","''")&"&"&replace(request.form("data3"),"'","''")&"&"&replace(request.form("data4"),"'","''")&"'"
	elseif Request.Form("id")="9" then ' SECPay
		sSQL = sSQL & "payProvData1='"&replace(request.form("data1"),"'","''")&"',payProvData2='"&replace(request.form("data2")&"&"&server.urlencode(request.form("data3")),"'","''")&"'"
	elseif Request.Form("id")="10" then ' Capture Card
		data1 = ""
		for index=1 to 20
			if Request.Form("cardtype" & index)="X" then
				data1=data1&"X"
			else
				data1=data1&"O"
			end if
		next
		sSQL = sSQL & "payProvData1='"&data1&"'"
		if Trim(request.form("data2"))<>"" then
			sSQL2 = "UPDATE admin SET adminCert='"&replace(request.form("data2"),"'","''")&"' WHERE adminID=1"
			cnn.Execute(sSQL2)
		end if
	elseif Request.Form("id")="18" OR Request.Form("id")="19" then ' PayPal Pro
		sSQL = sSQL & "payProvData1='"&replace(request.form("data1"),"'","''")&"',payProvData2='"&replace(request.form("data2"),"'","''")&"',payProvData3='"&replace(request.form("data3"),"'","''")&"'"
	else
		thedata1 = trim(request.form("data1"))
		thedata2 = trim(request.form("data2"))
		if secretword<>"" AND (Request.Form("id")="3" OR Request.Form("id")="13") then
			thedata1 = upsencode(thedata1, secretword)
			thedata2 = upsencode(thedata2, secretword)
		end if
		sSQL = sSQL & "payProvData1='"&replace(thedata1,"'","''")&"',payProvData2='"&replace(thedata2,"'","''")&"'"
	end if
	for index=2 to adminlanguages+1
		if (adminlangsettings AND 128)=128 then
			sSQL = sSQL & ",payProvShow" & index & "='"&replace(request.form("showas" & index),"'","''")&"'"
		end if
	next
	if Trim(request.form("transtype"))<>"" then sSQL = sSQL & ",payProvMethod=" & Trim(request.form("transtype"))
	sSQL = sSQL & " WHERE payProvID="&Request.Form("id")
	cnn.Execute(sSQL)
	if Request.Form("id")="18" OR Request.Form("id")="19" then ' PayPal Pro
		sSQL = "UPDATE payprovider SET payProvDemo="&demomode&",payProvData1='"&replace(request.form("data1"),"'","''")&"',payProvData2='"&replace(request.form("data2"),"'","''")&"',payProvData3='"&replace(request.form("data3"),"'","''")&"'"
		if Request.Form("id")="18" then
			if isenabled=1 then sSQL = sSQL & ",payProvEnabled=1"
			sSQL = sSQL & " WHERE payProvID=19"
		end if
		if Request.Form("id")="19" then sSQL = sSQL & " WHERE payProvID=18"
		cnn.Execute(sSQL)
	end if
	response.write "<meta http-equiv=""refresh"" content=""3; url=adminpayprov.asp"">"
elseif request.form("act")="changepos" then
	currentorder = Int(Request.Form("selectedq"))
	neworder = Int(Request.Form("newval"))
	sSQL = "SELECT payProvID FROM payprovider ORDER BY payProvEnabled DESC,payProvOrder"
	rs.Open sSQL,cnn,0,1
	alldata=rs.getrows
	rs.Close
	FOR rowcounter=0 TO ubound(alldata,2)
		theorder = rowcounter+1
		if currentorder = theorder then
			theorder = neworder
		elseif (currentorder > theorder) AND (neworder <= theorder) then
			theorder = theorder + 1
		elseif (currentorder < theorder) AND (neworder >= theorder) then
			theorder = theorder - 1
		end if
		sSQL="UPDATE payprovider SET payProvOrder="&theorder&" WHERE payProvID="&alldata(0,rowcounter)
		cnn.Execute(sSQL)
	NEXT
	response.write "<meta http-equiv=""refresh"" content=""2; url=adminpayprov.asp"">"
end if
%>
<script language="javascript" type="text/javascript">
<!--
function modrec(id) {
	document.mainform.id.value = id;
	document.mainform.act.value = "modify";
	document.mainform.submit();
}
// -->
</script>
      <table border="0" cellspacing="0" cellpadding="0" width="100%" bgcolor="" align="center">
<%	if request.form("act")="domodify" AND success then %>
        <tr>
          <td width="100%">
			<table width="100%" border="0" cellspacing="0" cellpadding="2" bgcolor="">
			  <tr> 
                <td width="100%" colspan="2" align="center"><br /><strong><%=yyUpdSuc%></strong><br /><br /><%=yyNowFrd%><br /><br />
				<%=yyNoAuto%> <A href="adminpayprov.asp"><strong><%=yyClkHer%></strong></a>.<br /><br />
				<img src="../images/clearpixel.gif" width="300" height="3" alt="" />
                </td>
			  </tr>
			</table></td>
        </tr>
<%	elseif request.form("act")="domodify" then %>
        <tr>
          <td width="100%">
			<table width="100%" border="0" cellspacing="0" cellpadding="2" bgcolor="">
			  <tr> 
                <td width="100%" colspan="2" align="center"><br /><font color="#FF0000"><strong><%=yyOpFai%></strong></font><br /><br /><%=errmsg%><br /><br />
				<a href="javascript:history.go(-1)"><strong><%=yyClkBac%></strong></a></td>
			  </tr>
			</table></td>
        </tr>
<%	elseif request.form("act")="modify" then
		sSQL = "SELECT payProvID,payProvName,payProvShow,payProvDemo,payProvEnabled,payProvData1,payProvData2,payProvData3,payProvMethod,payProvLevel,payProvShow2,payProvShow3 FROM payprovider WHERE payProvAvailable=1"
		if Request.Form("id")<>"" then
			sSQL = sSQL & " AND payProvID=" & request.form("id")
		else
			sSQL = sSQL & " ORDER BY payProvEnabled DESC,payProvOrder"
		end if
		rs.Open sSQL,cnn,0,1
			payProvID=trim(rs("payProvID")&"")
			payProvName=trim(rs("payProvName")&"")
			payProvShow=trim(rs("payProvShow")&"")
			payProvDemo=rs("payProvDemo")
			payProvEnabled=rs("payProvEnabled")
			payProvData1=trim(rs("payProvData1")&"")
			payProvData2=trim(rs("payProvData2")&"")
			payProvData3=trim(rs("payProvData3")&"")
			payProvMethod=rs("payProvMethod")
			payProvLevel=rs("payProvLevel")
			payProvShow2=trim(rs("payProvShow2")&"")
			payProvShow3=trim(rs("payProvShow3")&"")
		rs.Close
		data2name=""
		if payProvID=1 then ' PayPal
			data1name=yyEmail
			data2name="Identity Token<br><font size='1'>(Only when using PDT)</font>"
			demomodeavailable=true
		elseif payProvID=2 then ' 2Checkout
			data1name=yyAccNum
			data2name=yyMD5H
			warning1=TRUE
		elseif payProvID=3 OR payProvID=13 then ' Authorize.net
			data1name=yyMercLID
			data2name=yyTrnKey
			if secretword<>"" then
				payProvData1 = upsdecode(payProvData1, secretword)
				payProvData2 = upsdecode(payProvData2, secretword)
			end if
		elseif payProvID=4 OR payProvID=17 then ' Email
			data1name=yyEAOrd
			demomodeavailable=false
		elseif payProvID=5 then ' World Pay
			data1name=yyAccNum
			data2name=yyMD5H
			warning1=TRUE
		elseif payProvID=6 then ' NOCHEX
			data1name=yyEmail
		elseif payProvID=8 then ' Payflow Link
			data1name=yyLogin
			data2name=yyPartner
		elseif payProvID=9 then ' secpay
			data1name=yyMercID
			data2name=yyMD5H
			warning1=TRUE
		elseif payProvID=10 then ' Capture Card
			demomodeavailable=false
		elseif payProvID=11 OR payProvID=12 then ' PSiGate
			data1name=yyMercID
		elseif payProvID=14 then ' Custom Payment Processor
			data1name="Data 1"
			data2name="Data 2"
		elseif payProvID=15 then ' Netbanx
			data1name=yyMercID
			demomodeavailable=false
		elseif payProvID=16 then ' Linkpoint
			data1name=yyNumSto
			data2name=yyOwnSit
		elseif payProvID=18 OR payProvID=19 then ' PayPal Payment Pro
			data1name="API Account Name"
			data2name="API Password.<br><font size=""1"">(NOT PayPal account password)</font>"
			data3name="Signature Hash.<br><font size=""1"">(Only when using 3-token authentication)</font>"
		elseif payProvID=20 then ' Google Checkout
			data1name="Merchant ID"
			data2name="Merchant Key"
		else
			data1name="Data 1"
		end if
%>
        <tr>
		  <form name="mainform" method="post" action="adminpayprov.asp">
          <td width="100%">
			<input type="hidden" name="posted" value="1" />
			<input type="hidden" name="act" value="domodify" />
			<input type="hidden" name="id" value="<%=payProvID%>" />
            <table width="100%" border="0" cellspacing="0" cellpadding="2" bgcolor="">
			  <tr> 
                <td width="100%" colspan="2" align="center"><strong><%=yyPPAdm%></strong><br />&nbsp;</td>
			  </tr>
			  <tr>
				<td width="50%" align="right"><strong><%=yyPPName%> : </strong></td>
				<td width="50%" align="left"><strong><%=payProvName%></strong></td>
			  </tr>
			  <tr>
				<td width="50%" align="right"><strong><%=yyShwAs%> : </strong></td>
				<td width="50%" align="left"><input type="text" name="showas" value="<%=payProvShow%>" size="25" /></td>
			  </tr>
<%	for index=2 to adminlanguages+1
		if index=2 then showthis=payProvShow2
		if index=3 then showthis=payProvShow3
		if (adminlangsettings AND 128)=128 then %>
			  <tr>
				<td width="50%" align="right"><strong><%=yyShwAs & " " & index%> : </strong></td>
				<td width="50%" align="left"><input type="text" name="showas<%=index%>" value="<%=showthis%>" size="25" /></td>
			  </tr>
<%		end if
	next
	if payProvID=7 then ' VeriSign PayFlo Pro %>
			  <tr>
				<td colspan="2" align="center"><%=yyPPExp%></td>
			  </tr>
<%	end if %>
			  <tr>
				<td width="50%" align="right"><strong><%=yyEnable%> : </strong></td>
				<td width="50%" align="left"><input type="checkbox" name="isenabled" value="1" <%if payProvEnabled=1 then response.write "checked"%> /></td>
			  </tr>
<%	if demomodeavailable then %>
			  <tr>
				<td width="50%" align="right"><strong><%=yyDemoMo%> : </strong></td>
				<td width="50%" align="left"><input type="checkbox" name="demomode" value="1" <%if payProvDemo=1 then response.write "checked"%> /></td>
			  </tr>
<%	end if
	if payProvID=7 then ' VeriSign PayFlo Pro
		Dim vs1,vs2,vs3,vs4
		if IsNull(payProvData1) then payProvData1=""
		vsdetails = Split(payProvData1, "&")
		if UBOUND(vsdetails) > 0 then
			vs1=vsdetails(0)
			vs2=vsdetails(1)
			vs3=vsdetails(2)
			vs4=vsdetails(3)
		end if
%>
			  <tr>
				<td width="50%" align="right"><strong><%=yyUserID%> : </strong></td>
				<td width="50%" align="left"><input type="text" name="data1" value="<%=vs1%>" size="25" /></td>
			  </tr>
			  <tr>
				<td width="50%" align="right"><strong><%=yyVendor%> : </strong></td>
				<td width="50%" align="left"><input type="text" name="data2" value="<%=vs2%>" size="25" /></td>
			  </tr>
			  <tr>
				<td width="50%" align="right"><strong><%=yyPartner%> : </strong></td>
				<td width="50%" align="left"><input type="text" name="data3" value="<%=vs3%>" size="25" /></td>
			  </tr>
			  <tr>
				<td width="50%" align="right"><strong><%=yyPass%> : </strong></td>
				<td width="50%" align="left"><input type="text" name="data4" value="<%=vs4%>" size="25" /></td>
			  </tr>
<%	elseif payProvID=10 then ' Capture Card %>
			  <tr>
				<td align="center" colspan="2"><hr width="50%" /><strong><%=yyAccCar%></strong><br />&nbsp;</td>
			  </tr>
			  <tr>
				<td align="right"><strong>Visa : </strong></td>
				<td align="left"><input type="checkbox" name="cardtype1" value="X" <% if Mid(payProvData1,1,1)="X" then response.write "checked" %> /></td>
			  </tr>
			  <tr>
				<td align="right"><strong>Mastercard : </strong></td>
				<td align="left"><input type="checkbox" name="cardtype2" value="X" <% if Mid(payProvData1,2,1)="X" then response.write "checked" %> /></td>
			  </tr>
			  <tr>
				<td align="right"><strong>American Express : </strong></td>
				<td align="left"><input type="checkbox" name="cardtype3" value="X" <% if Mid(payProvData1,3,1)="X" then response.write "checked" %> /></td>
			  </tr>
			  <tr>
				<td align="right"><strong>Diners Club : </strong></td>
				<td align="left"><input type="checkbox" name="cardtype4" value="X" <% if Mid(payProvData1,4,1)="X" then response.write "checked" %> /></td>
			  </tr>
			  <tr>
				<td align="right"><strong>Discover : </strong></td>
				<td align="left"><input type="checkbox" name="cardtype5" value="X" <% if Mid(payProvData1,5,1)="X" then response.write "checked" %> /></td>
			  </tr>
			  <tr>
				<td align="right"><strong>En Route : </strong></td>
				<td align="left"><input type="checkbox" name="cardtype6" value="X" <% if Mid(payProvData1,6,1)="X" then response.write "checked" %> /></td>
			  </tr>
			  <tr>
				<td align="right"><strong>JCB : </strong></td>
				<td align="left"><input type="checkbox" name="cardtype7" value="X" <% if Mid(payProvData1,7,1)="X" then response.write "checked" %> /></td>
			  </tr>
			  <tr>
				<td align="right"><strong>Switch/Solo : </strong></td>
				<td align="left"><input type="checkbox" name="cardtype8" value="X" <% if Mid(payProvData1,8,1)="X" then response.write "checked" %> /></td>
			  </tr>
			  <tr>
				<td align="right"><strong>Bankcard (AUS / NZ) : </strong></td>
				<td align="left"><input type="checkbox" name="cardtype9" value="X" <% if Mid(payProvData1,9,1)="X" then response.write "checked" %> /></td>
			  </tr>
			  <tr>
				<td align="center" colspan="2"><hr width="50%" /><strong><%=yyNewCer%></strong><br />&nbsp;</td>
			  </tr>
			  <tr>
				<td colspan="2" align="center"><textarea name="data2" rows="10" cols="82"></textarea></td>
			  </tr>
<%	else %>
			  <tr>
				<td width="50%" align="right"><strong><%=data1name%> : </strong></td>
				<td width="50%" align="left"><input type="text" name="data1" value="<%=payProvData1%>" size="25" /></td>
			  </tr>
<%	end if
	if payProvID=5 then
		data2arr = split(trim(payProvData2&""),"&",2)
		if UBOUND(data2arr) >= 0 then data2md5 = data2arr(0)
		if UBOUND(data2arr) > 0 then data2cbp = data2arr(1) else data2cbp = ""
%>
			  <tr>
				<td width="50%" align="right"><strong>MD5 Secret (Optional) : </strong></td>
				<td width="50%" align="left"><input type="text" name="data2" value="<%=data2md5%>" size="25" /></td>
			  </tr>
			  <tr>
				<td width="50%" align="right"><strong>Callback password (Optional) : </strong></td>
				<td width="50%" align="left"><input type="text" name="data3" value="<%=data2cbp%>" size="25" /></td>
			  </tr>
<%	elseif payProvID=9 then
		data2arr = split(trim(payProvData2&""),"&",2)
		if UBOUND(data2arr) >= 0 then data2md5 = data2arr(0)
		if UBOUND(data2arr) > 0 then data2template = urldecode(data2arr(1)) else data2template = ""
%>
			  <tr>
				<td width="50%" align="right"><strong><%=yyMD5H%> : </strong></td>
				<td width="50%" align="left"><input type="text" name="data2" value="<%=data2md5%>" size="25" /></td>
			  </tr>
			  <tr>
				<td width="50%" align="right"><strong>Payment Template (Optional) : </strong></td>
				<td width="50%" align="left"><input type="text" name="data3" value="<%=data2template%>" size="25" /></td>
			  </tr>
<%	elseif payProvID=16 then %>
			  <tr>
				<td width="50%" align="right"><strong><%=data2name%> : </strong></td>
				<td width="50%" align="left"><select name="data2" size="1"><option value="0"><%=yyLPSit%></option><option value="1" <% if payProvData2="1" then response.write "selected"%>><%=yyYesOS%></option></select></td>
			  </tr>
<%	elseif payProvID=18 OR payProvID=19 then
		data2arr = split(trim(payProvData2&""),"&",2)
		if UBOUND(data2arr) >= 0 then data2pwd = data2arr(0)
		data2hash = payProvData3
%>			  <tr>
				<td width="50%" align="right"><strong><%=data2name%> : </strong></td>
				<td width="50%" align="left"><input type="text" name="data2" value="<%=data2pwd%>" size="25" /></td>
			  </tr>
			  <tr>
				<td width="50%" align="right"><strong><%=data3name%> : </strong></td>
				<td width="50%" align="left"><input type="text" name="data3" value="<%=data2hash%>" size="45" /></td>
			  </tr>
<%	elseif data2name<>"" then %>
			  <tr>
				<td width="50%" align="right"><strong><%=data2name%> : </strong></td>
				<td width="50%" align="left"><input type="text" name="data2" value="<%=payProvData2%>" size="25" /></td>
			  </tr>
<%	end if
	if payProvID=3 OR payProvID=5 OR payProvID=7 OR payProvID=9 OR payProvID=11 OR payProvID=12 OR payProvID=13 OR payProvID=14 OR payProvID=16  OR payProvID=18 then ' Pay Providers we can set authorization type %>
			  <tr>
				<td width="50%" align="right"><strong><%=yyTrnTyp%> : </strong></td>
				<td width="50%" align="left"><select name="transtype" size="1"><option value="0"><%=yyAuthCp%></option><option value="1" <% if payProvMethod="1" then response.write "selected" %>><%=yyAuthOn%></option></select></td>
			  </tr>
<%	end if %>
			  <tr>
				<td width="50%" align="right"><strong><%=yyLiLev%> : </strong></td>
				<td width="50%" align="left"><select name="payProvLevel" size="1">
				<option value="0"><%=yyNoRes%></option>
				<%	for index=1 to maxloginlevels
						response.write "<option value="""&index&""""
						if payProvLevel=index then response.write " selected"
						response.write ">" & yyLiLev & " " & index & "</option>"
					next%></select></td>
			  </tr>
			  <tr>
				<td colspan="2">&nbsp;</td>
			  </tr>
			  <tr>
				<td width="50%" align="right"><input type="submit" value="<%=yySubmit%>" /></td>
				<td width="50%" align="left"><input type="reset" value="<%=yyReset%>" /></td>
			  </tr>
<%	if warning1=TRUE then %>
			  <tr>
				<td colspan="2">&nbsp;<br /><font size="1">Setting MD5 hash and callback password security features is optional. But if set, they will be checked so you must make sure they match with your payment processor.</font></td>
			  </tr>
<%	end if %>
			  <tr>
				<td colspan="2">&nbsp;</td>
			  </tr>
			</table>
		  </td>
		  </form>
		</tr>
<%	elseif request.form("act")="changepos" then %>
        <tr>
          <td width="100%" align="center">
			<p>&nbsp;</p>
			<p>&nbsp;</p>
			<p>&nbsp;</p>
			<p><strong><%=yyUpdat%> . . . . . . . </strong></font></p>
			<p>&nbsp;</p>
			<p><%=yyNoFor%> <a href="adminpayprov.asp"><%=yyClkHer%></a>.</p>
			<p>&nbsp;</p>
			<p>&nbsp;</p>
		  </td>
		</tr>
<%	else
function writeposition(currpos,maxpos)
	Dim reqtext,i
	reqtext="<select name='newpos" & currpos & "' size='1' onchange='javascript:validate_index("&currpos&");'>"
	for i = 1 to maxpos
		reqtext = reqtext & "<option value='"&i&"'"
		if currpos=i then reqtext=reqtext&" selected"
		reqtext = reqtext & ">"&i&"</option>"
	next
	writeposition = reqtext & "</select>"
end function
%>
<script language="javascript" type="text/javascript">
<!--
function validate_index(currindex)
{
	var i = eval("document.mainform.newpos"+currindex+".selectedIndex")+1;
	document.mainform.newval.value = i;
	document.mainform.selectedq.value = currindex;
	document.mainform.act.value = "changepos";
	if(i==document.mainform.selectedq.value){
		return (false);
	}
	document.mainform.submit();
}
//-->
</script>
        <tr>
		  <form name="mainform" method="post" action="adminpayprov.asp">
          <td width="100%" align="center">
			<input type="hidden" name="posted" value="1" />
			<input type="hidden" name="act" value="modify" />
			<input type="hidden" name="id" value="1" />
			<input type="hidden" name="selectedq" value="1" />
			<input type="hidden" name="newval" value="1" />
            <table width="80%" border="0" cellspacing="0" cellpadding="2" bgcolor="">
			  <tr> 
                <td width="100%" colspan="4" align="center"><strong><%=yyPPAdm%></strong><br />&nbsp;</td>
			  </tr>
			  <tr>
				<td width="8%" align="center"><strong><%=yyID%></strong></td>
				<td width="8%" align="center"><strong><%=yyOrder%></strong></td>
				<td width="42%" align="center"><strong><%=yyPPName%></strong></td>
				<td width="42%" align="center"><strong><%=yyConf%></strong></td>
			  </tr>
<%	showenabled=true
	for index=0 to 1
		sSQL = "SELECT payProvID,payProvName,payProvShow,payProvDemo,payProvEnabled,payProvData1,payProvData2,payProvMethod,payProvShow2,payProvShow3 FROM payprovider WHERE payProvAvailable=1"
		if showenabled then
			sSQL = sSQL & " AND payProvEnabled=1 ORDER BY payProvOrder"
		else
			sSQL = sSQL & " AND payProvEnabled=0 ORDER BY payProvName"
		end if
		rs.Open sSQL,cnn,0,1
		alldata = ""
		if NOT rs.EOF then alldata=rs.getrows
		rs.Close
		if IsArray(alldata) then
			if showenabled then enabledProv=UBOUND(alldata,2)+1 else enabledProv=0
			for rowcounter=0 to UBOUND(alldata,2) %>
			  <tr>
				<td align="center"><%=alldata(0,rowcounter)%></td>
				<td align="center"><%if alldata(4,rowcounter)=1 then response.write writeposition(rowcounter+1,enabledProv) else response.write "-" %></td>
				<td align="center"><%if alldata(3,rowcounter)=1 then response.write "<font color='#FF0000'>" %><%if alldata(4,rowcounter)=1 then response.write "<strong>" %><%=alldata(1,rowcounter)%><%if alldata(4,rowcounter)=1 then response.write "</strong>" %><%if alldata(3,rowcounter)=1 then response.write "</font>" %></td>
				<td align="center"><input type=button value="<%=yyModify%>" onclick="modrec('<%=alldata(0,rowcounter)%>')" /></td>
			  </tr>
<%			next
		end if
		showenabled=false
	next %>
			  <tr> 
                <td width="100%" colspan="4" align="center"><br /><%=yyPPEx1%><br />
				  <%=yyPPEx2%>&nbsp;</td>
			  </tr>
			  <tr> 
                <td width="100%" colspan="4" align="center"><br /><a href="admin.asp"><strong><%=yyAdmHom%></strong></a><br />&nbsp;</td>
			  </tr>
            </table></td>
		  </form>
        </tr>
<%
end if
cnn.Close
set rs = nothing
set cnn = nothing
%>
      </table>