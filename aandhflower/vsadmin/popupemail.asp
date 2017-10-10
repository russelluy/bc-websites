<%
Response.Buffer = True
%>
<!--#include file="inc/md5.asp"-->
<!--#include file="db_conn_open.asp"-->
<!--#include file="includes.asp"-->
<!--#include file="inc/languageadmin.asp"-->
<!--#include file="inc/languagefile.asp"-->
<!--#include file="inc/incemail.asp"-->
<!--#include file="inc/incfunctions.asp"-->
<%
if storesessionvalue="" then storesessionvalue="virtualstore"
if Session("loggedon") <> storesessionvalue OR disallowlogin=TRUE then response.redirect "login.asp"
isprinter=false
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html>
<head>
<title>Email Popup</title>
<link rel="stylesheet" type="text/css" href="adminstyle.css"/>
<meta http-equiv="Content-Type" content="text/html; charset=<%=adminencoding%>"/>
</head>
<body<%if request.querystring("prod")<>"" then response.write " onload=""updateopener()"""%>>
&nbsp;<br>
<div>
<form method="post" action="popupemail.asp">
<%
	Set rs = Server.CreateObject("ADODB.RecordSet")
	Set rs2 = Server.CreateObject("ADODB.RecordSet")
	Set cnn=Server.CreateObject("ADODB.Connection")
	cnn.open sDSN
	if request.form("posted")="1" then
		alreadygotadmin = getadminsettings()
		Call do_order_success(request.form("id"),emailAddr,FALSE,FALSE,request.form("customer")="1",request.form("affiliate")="1",IIfVr(request.form("manufacturer")="1",2,FALSE))
%>
<p align="center"><%=yyOpSuc%></p>
<p align="center"><a href="javascript:window.close()"><strong><%=xxClsWin%></strong></a></p>
<%	elseif request.form("posted")="2" then
		ordID = replace(request.form("oid"),"'","")
		alreadygotadmin = getadminsettings()
		sSQL = "SELECT ordTransID,ordPayProvider,ordAuthNumber,payProvData1,payProvData2,payProvDemo FROM orders INNER JOIN payprovider ON orders.ordPayProvider=payprovider.payProvID WHERE ordID=" & ordID
		rs.Open sSQL,cnn,0,1
		transid=rs("ordTransID")
		authcode=rs("ordAuthNumber")
		if InStr(authcode,"-") > 0 then authcode = Right(authcode,Len(authcode)-InStr(authcode,"-"))
		login = rs("payProvData1")
		trankey = rs("payProvData2")
		if secretword<>"" then
			login = upsdecode(login, secretword)
			trankey = upsdecode(trankey, secretword)
		end if
		demomode=(Int(rs("payProvDemo"))=1)
		rs.Close
		parmList = "x_version=3.1&x_delim_data=True&x_relay_response=False&x_delim_char=|"
		parmList = parmList & "&x_login="&login
		parmList = parmList & "&x_tran_key="&trankey
		parmList = parmList & "&x_trans_id="&transid
		parmList = parmList & "&x_auth_code="&authcode
		parmList = parmList & "&x_type=PRIOR_AUTH_CAPTURE"
		if demomode then parmList = parmList & "&x_test_request=TRUE"
		' response.write "paramlist is<br>" & replace(parmList,"&","<br>") & "<br>" & vbCrLf
		response.write "&nbsp;<br><p align=""center"" id=""process"">Processing. Please wait...</p>"
		response.flush
		set objHttp = Server.CreateObject("Msxml2.ServerXMLHTTP")
		objHttp.open "POST", "https://secure.authorize.net/gateway/transact.dll", false
		objHttp.Send parmList
		if err.number <> 0 OR objHttp.status <> 200 Then
			errormsg = "Error, couldn't connect to Authorize.net server"
		else
			varString = Split(objHttp.responseText, "|")
			' response.write "The response is " & objHttp.responseText & "<br>"
			vsRESULT=varString(0)
			vsRESPMSG=varString(3)
			success=FALSE
			if Int(vsRESULT)=1 then
				success=TRUE
				vsRESPMSG=yyOpSuc
				if capturedordstatus<>"" then
					sSQL="UPDATE orders SET ordStatus=" & capturedordstatus & " WHERE ordID=" & ordID
					cnn.Execute(sSQL)
				end if
			end if
		end if
		set objHttp = nothing
%>
<script language="javascript" type="text/javascript">
thestyle = document.getElementById('process').style;
thestyle.display = 'none';
</script>
<p align="center"><%=vsRESPMSG%></p>
<p align="center"><a href="javascript:window.close()"><strong><%=xxClsWin%></strong></a></p>
<%	elseif request.querystring("id")<>"" then %>
<input type="hidden" name="posted" value="1">
<input type="hidden" name="id" value="<%=request.querystring("id")%>">
<table width="100%" cellspacing="2" cellpadding="2">
<tr><td colspan="2" align="center"><strong><%=yySendFo%></strong></td></tr>
<tr><td align="right" width="60%"><%=yyCusto%>: </td><td><input type="checkbox" name="customer" value="1" checked></td></tr>
<tr><td align="right"><%=yyAffili%>: </td><td><input type="checkbox" name="affiliate" value="1"></td></tr>
<tr><td align="right"><%=yyManDes%>: </td><td><input type="checkbox" name="manufacturer" value="1"></td></tr>
<tr><td colspan="2" align="center"><input type="submit" value="<%=yySubmit%>" /></td></tr>
</table>
<%	elseif request.querystring("oid")<>"" then %>
&nbsp;<br>
<input type="hidden" name="posted" value="2">
<input type="hidden" name="oid" value="<%=request.querystring("oid")%>">
<table width="100%" cellspacing="2" cellpadding="2">
<tr><td colspan="2" align="center"><strong>Capture funds for order id <%=request.querystring("oid")%></strong><br>&nbsp;</td></tr>
<tr><td colspan="2" align="center"><input type="submit" value="<%=yySubmit%>" /></td></tr>
</table>
<%	elseif request.querystring("prod")<>"" then
		id = request.querystring("index")
		sSQL = "SELECT "&getlangid("pName",1)&",pPrice FROM products WHERE pID='"&replace(request.querystring("prod"),"'","''")&"'"
		rs.Open sSQL,cnn,0,1
		if rs.EOF then
			prodname="Not Found"
			prodprice=0
		else
			prodname=rs(getlangid("pName",1))
			prodprice=rs("pPrice")
		end if
		rs.Close
		response.write "<span id=""prodname"">"&prodname&"</span>"
		response.write "<span id=""prodprice"">"&prodprice&"</span>"
%>
<span id="bodytext"><%
sSQL = "SELECT poOptionGroup,optType,optFlags FROM prodoptions INNER JOIN optiongroup ON optiongroup.optGrpID=prodoptions.poOptionGroup WHERE poProdID='"&replace(request.querystring("prod"),"'","''")&"' ORDER BY poID"
prodoptions = ""
rs.Open sSQL,cnn,0,1
if NOT rs.EOF then
	prodoptions=rs.getrows
else
	response.write "-"
end if
rs.Close
if IsArray(prodoptions) then
	response.write "<table border='0' cellspacing='0' cellpadding='1' width='100%'>"
	for rowcounter=0 to UBOUND(prodoptions,2)
		index=0
		sSQL="SELECT optID,"&getlangid("optName",32)&","&getlangid("optGrpName",16)&","&OWSP&"optPriceDiff,optType,optFlags,optStock,optPriceDiff AS optDims FROM options INNER JOIN optiongroup ON options.optGroup=optiongroup.optGrpID WHERE optGroup="&prodoptions(0,rowcounter)&" ORDER BY optID"
		rs2.Open sSQL,cnn,0,1
		if NOT rs2.EOF then
			if Abs(Int(rs2("optType")))=3 then
				response.write "<tr><td align='right' width='30%'><strong>"&rs2(getlangid("optGrpName",16))&":</strong></td><td align=""left""> <input type='hidden' name='optn"&id&"_"&rowcounter&"' value='"&rs2("optID")&"' />"
				response.write "<textarea wrap='virtual' name='voptn"&id&"_"&rowcounter&"' id='voptn"&id&"_"&rowcounter&"' cols='30' rows='3'>"
				response.write rs2(getlangid("optName",32))&"</textarea>"
				response.write "</td></tr>"
			else
				response.write "<tr><td align='right' width='30%'><strong>"&rs2(getlangid("optGrpName",16))&":</strong></td><td align=""left""> <select class=""prodoption"" onchange=""dorecalc(true)"" name='optn"&id&"_"&rowcounter&"' id='optn"&id&"_"&rowcounter&"' size='1'>"
				if Int(rs2("optType"))>0 then response.write "<option value=''>"&xxPlsSel&"</option>"
				do while not rs2.EOF
					response.write "<option value='"&rs2("optID")&"|"&IIfVr((rs2("optFlags") AND 1) = 1,(prodprice*rs2("optPriceDiff"))/100.0,rs2("optPriceDiff"))&"'>"&rs2(getlangid("optName",32))
					if cDbl(rs2("optPriceDiff"))<>0 then
						response.write " "
						if cDbl(rs2("optPriceDiff")) > 0 then response.write "+"
						if (rs2("optFlags") AND 1) = 1 then
							response.write FormatNumber((prodprice*rs2("optPriceDiff"))/100.0,2)
						else
							response.write FormatNumber(rs2("optPriceDiff"),2)
						end if
					end if
					response.write "</option>"&vbCrLf
					rs2.MoveNext
				loop
				response.write "</select></td></tr>"
			end if
		end if
		rs2.Close
	next
	response.write "</table>"
end if
%>
</span>
<script language="javascript" type="text/javascript">
<!--
function updateopener(){
//alert(document.getElementById('prodname').innerHTML);
window.opener.document.getElementById('prodname<%=id%>').value = document.getElementById('prodname').innerHTML;
window.opener.document.getElementById('price<%=id%>').value = document.getElementById('prodprice').innerHTML;
window.opener.document.getElementById('optionsspan<%=id%>').innerHTML = document.getElementById('bodytext').innerHTML;
window.opener.document.getElementById('optdiffspan<%=id%>').value = 0;
window.close();
}
//-->
</script>
<%	end if
cnn.Close
%>
</form>
</div>
</body>
</html>
