<!--#include file="inc/md5.asp"-->
<!--#include file="db_conn_open.asp"-->
<!--#include file="includes.asp"-->
<!--#include file="inc/incemail.asp"-->
<!--#include file="inc/languagefile.asp"-->
<!--#include file="inc/incfunctions.asp"-->
<%	if wpconfirmnoheaders<>TRUE then %>
<html>
<head>
<title>Thanks for shopping with us</title>
<meta http-equiv="Content-Type" content="text/html; charset=<%=adminencoding%>">
<style type="text/css">
<!--
A:link {
	COLOR: #FFFFFF; TEXT-DECORATION: none
}
A:visited {
	COLOR: #FFFFFF; TEXT-DECORATION: none
}
A:active {
	COLOR: #FFFFFF; TEXT-DECORATION: none
}
A:hover {
	COLOR: #f39000; TEXT-DECORATION: underline
}
TD {
	FONT-FAMILY: Verdana; FONT-SIZE: 13px
}
P {
	FONT-FAMILY: Verdana; FONT-SIZE: 13px
}
-->
</style>
</head>
<%	end if
Dim rs,rs2,sSQL,orderText,custEmail,mailsystem,success,isworldpay,isauthnet,ordGrandTotal,ordID,ordAuthNumber
success=false
errtext=""
ordGrandTotal = 0
Session("couponapply") = ""
Set rs = Server.CreateObject("ADODB.RecordSet")
Set rs2 = Server.CreateObject("ADODB.RecordSet")
Set cnn=Server.CreateObject("ADODB.Connection")
cnn.open sDSN
alreadygotadmin = getadminsettings()
success = false
worldpaycallbackerror = false
isworldpay = false
isauthnet = false
isnetbanx = false
issecpay = false
if Trim(request.form("transStatus"))<>"" then ' WorldPay
	isworldpay = true
	transstatus = trim(request.form("transStatus"))
	data2cbp = ""
	if getpayprovdetails(5,acctno,data2,data3,demomode,ppmethod) then
		data2arr = split(data2,"&",2)
		if UBOUND(data2arr) >= 0 then data2md5 = data2arr(0)
		if UBOUND(data2arr) > 0 then data2cbp = data2arr(1)
		if data2cbp <> "" then
			if data2cbp <> request.form("callbackPW") then
				transstatus=""
				errormsg = "Callback password incorrect"
				worldpaycallbackerror = TRUE
			end if
		end if
		if transstatus="Y" then
			ordID = trim(replace(request.form("cartId"),"'",""))
			avscode = trim(request.form("AVS"))
			if trim(request.form("wafMerchMessage"))<>"" then avscode = trim(request.form("wafMerchMessage")) & "<br />" & avscode
			do_stock_management(ordID)
			cnn.Execute("UPDATE cart SET cartCompleted=1 WHERE cartOrderID="&ordID)
			cnn.Execute("UPDATE orders SET ordStatus=3,ordAVS='"&replace(avscode,"'","")&"',ordAuthNumber='"&replace(trim(request.form("transId")),"'","")&"' WHERE ordPayProvider=5 AND ordID="&ordID)
			Call do_order_success(ordID,emailAddr,sendEmail,FALSE,TRUE,TRUE,TRUE)
			success = true
		end if
	end if
elseif Trim(request.form("x_response_code"))<>"" then ' Authorize.net
	if getpayprovdetails(3,data1,data2,data3,demomode,ppmethod) then
		isauthnet = true
		ordID = trim(replace(request.form("x_ect_ordid"),"'",""))
		if trim(request.form("x_response_code"))="1" AND ordID<>"" then
			do_stock_management(ordID)
			cnn.Execute("UPDATE cart SET cartCompleted=1 WHERE cartOrderID="&ordID)
			cnn.Execute("UPDATE orders SET ordStatus=3,ordAVS='"&replace(trim(request.form("x_avs_code")),"'","")&"',ordCVV='"&replace(trim(request.form("x_cvv2_resp_code")),"'","")&"',ordAuthNumber='"&replace(trim(request.form("x_auth_code")),"'","")&"',ordTransID='"&replace(trim(request.form("x_trans_id")),"'","")&"' WHERE ordPayProvider=3 AND ordID="&ordID)
			Call do_order_success(ordID,emailAddr,sendEmail,FALSE,TRUE,TRUE,TRUE)
			success = true
		else
			errormsg = Trim(request.form("x_response_reason_text"))
		end if
	end if
elseif Trim(request.form("trans_id"))<>"" then ' Secpay
	if getpayprovdetails(9,data1,data2,data3,demomode,ppmethod) then
		issecpay = true
		data2arr = split(data2,"&",2)
		if UBOUND(data2arr) >= 0 then data2md5 = data2arr(0)
		callbacksuccess=TRUE
		if Trim(request.form("valid"))="true" AND Trim(request.form("auth_code"))<>"" then
			ordID = trim(replace(request.form("trans_id"),"'",""))
			if trim(data2md5) <> "" then
				thehash = calcmd5("trans_id=" & ordID & "&amount=" & trim(request.form("amount")) & "&callback=" & storeurl & "vsadmin/wpconfirm.asp&" & data2md5)
				if request.form("hash") <> thehash then callbacksuccess=FALSE
			end if
			if NOT callbacksuccess then
				errormsg = "Callback password incorrect"
			else
				do_stock_management(ordID)
				cnn.Execute("UPDATE cart SET cartCompleted=1 WHERE cartOrderID="&ordID)
				cnn.Execute("UPDATE orders SET ordStatus=3,ordAVS='" & replace(trim(request.form("cv2avs")),"'","") & "',ordAuthNumber='" & trim(replace(request.form("auth_code"),"'",""))&"' WHERE ordPayProvider=9 AND ordID="&ordID)
				Call do_order_success(ordID,emailAddr,sendEmail,FALSE,TRUE,TRUE,TRUE)
				success = true
			end if
		else
			errormsg = Trim(request.form("message"))
		end if
	end if
elseif Trim(request.form("netbanx_reference"))<>"" then ' Netbanx
	if getpayprovdetails(15,data1,data2,data3,demomode,ppmethod) then
		isnetbanx = true
		thereference = Trim(request.form("netbanx_reference"))
		if Trim(Request.ServerVariables("REMOTE_HOST"))<>"195.224.77.2" then
			errormsg = "Error: This transaction does not appear to have been initiated by Netbanx"
		elseif thereference<>"0" AND Trim(request.form("order_id"))<>"" then
			ordID = trim(replace(request.form("order_id"),"'",""))
			do_stock_management(ordID)
			cnn.Execute("UPDATE cart SET cartCompleted=1 WHERE cartOrderID="&ordID)
			allchecks = "X"
			if trim(request.form("houseno_auth"))="Matched" then
				allchecks = "Y"
			elseif trim(request.form("houseno_auth"))="Not matched" then
				allchecks = "N"
			end if
			if trim(request.form("postcode_auth"))="Matched" then
				allchecks = allchecks & "Y"
			elseif trim(request.form("postcode_auth"))="Not matched" then
				allchecks = allchecks & "N"
			else
				allchecks = allchecks & "X"
			end if
			cvv = "X"
			if trim(request.form("CV2_auth"))="Matched" then
				cvv = "Y"
			elseif trim(request.form("CV2_auth"))="Not matched" then
				cvv = "N"
			end if
			cnn.Execute("UPDATE orders SET ordStatus=3,ordAVS='"&allchecks&"',ordCVV='"&cvv&"',ordAuthNumber='" & thereference &"' WHERE ordPayProvider=15 AND ordID="&ordID)
			Call do_order_success(ordID,emailAddr,sendEmail,FALSE,TRUE,TRUE,TRUE)
			success = true
		else
			errormsg = "Transaction Declined"
		end if
	end if
end if
cnn.Close
set rs = nothing
set rs2 = nothing
set cnn = nothing
	if wpconfirmnoheaders<>TRUE then
%>
<body bgcolor="#FFFFFF" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<table width="100%" border="0" cellspacing="0" cellpadding="0" bgcolor="#F39900">
  <tr>
    <td>
      <table width="100%" border="1" cellspacing="1" cellpadding="3">
        <tr> 
          <td rowspan="4" bgcolor="#333333">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
          <td width="100%" bgcolor="#333333" align="center"><font color="#FFFFFF" face="Arial, Helvetica, sans-serif"><strong><% response.write xxInAssc&"&nbsp;"
		if isworldpay then
			response.write "WorldPay"
		elseif isauthnet then
			response.write "Authorize.Net"
		elseif isnetbanx then
			response.write "Netbanx"
		elseif issecpay then
			response.write "SECPay"
		else
			response.write "<a href=""http://www.ecommercetemplates.com"">EcommerceTemplates.com</a>"
		end if %></strong></font></td>
          <td rowspan="4" bgcolor="#333333">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
        </tr>
        <tr> 
          <td width="100%" bgcolor="#637BAD" align="center"><font color="#FFFFFF"><strong><font face="Verdana, Arial, Helvetica, sans-serif" size="3"><%=xxTnkStr%></font></strong></font></td>
        </tr>
        <tr> 
          <td width="100%" align="center" bgcolor="#F5F5F5">
<%	end if ' wpconfirmnoheaders %>
<%	if isworldpay then %>
			<p>&nbsp;</p>
			<p align="center"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><strong><%=xxTnkWit%> <WPDISPLAY ITEM=compName></strong></font></p>
<%		if worldpaycallbackerror then %>
			<table width="<%=innertablewidth%>" border="0" cellspacing="<%=innertablespacing%>" cellpadding="<%=innertablepadding%>" bgcolor="<%=innertablebg%>">
			  <tr> 
				<td width="100%" colspan="2" align="center"><%=xxThkErr%>
				<p>The error report returned by the server was:<br /><strong><%=errormsg%></strong></p>
				<a href="<%=storeUrl%>"><font color="#637BAD"><strong><%=xxCntShp%></strong></font></a><br />
				<p>&nbsp;</p>
				</td>
			  </tr>
			</table>
<%		end if %>
            <p><wpdisplay item="banner"></p>
<%		if NOT worldpaycallbackerror then
			if digidownloads=TRUE then
				response.write "<table width=95% cellpadding=3 cellspacing=0 border=0><tr><td><table width=100% cellspacing=0 cellpadding=3 border=0><tr><td>"
				noshowdigiordertext = TRUE
' WORLDPAY ONLY : To enable digital downloads, just add a "hash" back into the line below so it looks like this . . .
' <!--#include file="inc/digidownload.asp"-->
' If you apply an updater, you must repeat this step.
%>
<!--include file="inc/digidownload.asp"-->
<%				response.write "</td></tr></table></td></tr></table>"
			end if
%>
			<table width=95% cellpadding=3 cellspacing=0 border=0>
			<tr><td>
			<table width=100% cellspacing=0 cellpadding=3 border=0>
			<tr><td>
			<p align="left"><%response.write Replace(orderText,vbCrLf,"<br />")%></p>
			</td></tr></table>
			</td></tr></table>
<%		end if %>
			<p><font size="1"><strong><%=xxPlsNt1&" "&xxMerRef&" "&xxPlsNt2%></strong></font></p>
			<p>&nbsp;</p>
<%	elseif success then %>
		  <table border="0" cellspacing="<%=maintablespacing%>" cellpadding="<%=maintablepadding%>" width="<%=maintablewidth%>" bgcolor="<%=maintablebg%>" align="center">
			<tr>
			  <td width="100%" align="center">
				<table width="80%" border="0" cellspacing="<%=innertablespacing%>" cellpadding="<%=innertablepadding%>" bgcolor="<%=innertablebg%>">
				  <tr> 
					<td width="100%" align="center"><%=xxThkYou%>
					</td>
				  </tr>
<%		if digidownloads=TRUE then
			response.write "</table>"
			noshowdigiordertext = TRUE
' To enable digital downloads, just add a "hash" back into the line below so it looks like this . . .
' <!--#include file="inc/digidownload.asp"-->
' If you apply an updater, you must repeat this step.
%>
<!--include file="inc/digidownload.asp"-->
<%			response.write "<table width=""80%"" border=""0"" cellspacing="""&innertablespacing&""" cellpadding="""&innertablepadding&""" bgcolor="""&innertablebg&""">"
		end if
%>
				  <tr> 
					<td width="100%"><%response.write Replace(orderText,vbCrLf,"<br />")%>
					</td>
				  </tr>
				  <tr> 
					<td width="100%" align="center"><br /><br />
					<%=xxRecEml%><br /><br />
					<a href="<%=storeUrl%>"><font color="#637BAD"><strong><%=xxCntShp%></strong></font></a><br />
					<img src="images/clearpixel.gif" width="350" height="3">
					</td>
				  </tr>
				</table>
			  </td>
			</tr>
		  </table>
<%	else %>
		  <p>&nbsp;</p>
		  <table border="0" cellspacing="<%=maintablespacing%>" cellpadding="<%=maintablepadding%>" width="<%=maintablewidth%>" bgcolor="<%=maintablebg%>" align="center">
			<tr>
			  <td width="100%">
				<table width="<%=innertablewidth%>" border="0" cellspacing="<%=innertablespacing%>" cellpadding="<%=innertablepadding%>" bgcolor="<%=innertablebg%>">
				  <tr> 
					<td width="100%" colspan="2" align="center"><%=xxThkErr%>
					<p>The error report returned by the server was:<br /><strong><%=errormsg%></strong></p>
					<a href="<%=storeUrl%>"><font color="#637BAD"><strong><%=xxCntShp%></strong></font></a><br />
					<p>&nbsp;</p>
					</td>
				  </tr>
				</table>
			  </td>
			</tr>
		  </table>
<%	end if %>
<%	if wpconfirmnoheaders<>TRUE then %>
          </td>
        </tr>
        <tr> 
          <td width="100%" bgcolor="#333333" align="center"><font color="#FFFFFF"><strong><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><a href="<%=storeUrl%>"><%=xxClkBck%></a></font></strong></font></td>
        </tr>
      </table>
    </td>
  </tr>
</table>
</body>
</html>
<%	end if ' wpconfirmnoheaders %>
