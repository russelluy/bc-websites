<!--#include file="incemail.asp"-->
<!--#include file="md5.asp"-->
<%
Dim rs,rs2,sSQL,orderText,errtext,ordGrandTotal,ordTotal,ordID,ordAuthNumber
success=false
errtext=""
ordGrandTotal = 0
ordTotal = 0
Session("couponapply") = ""
Sub order_failed
%>
      <table border="0" cellspacing="<%=maintablespacing%>" cellpadding="<%=maintablepadding%>" width="<%=maintablewidth%>" bgcolor="<%=maintablebg%>" align="center">
        <tr>
          <td width="100%">
            <table width="<%=innertablewidth%>" border="0" cellspacing="<%=innertablespacing%>" cellpadding="<%=innertablepadding%>" bgcolor="<%=innertablebg%>">
			  <tr> 
                <td width="100%" colspan="2" align="center"><%=xxThkErr%>
				<% if errtext<>"" then response.write "<p><strong>" & errtext & "</strong></p>" %>
				<a href="<%=storeUrl%>"><strong><%=xxCntShp%></strong></a><br />
				<img src="images/clearpixel.gif" width="300" height="3" alt="" />
                </td>
			  </tr>
			</table>
		  </td>
        </tr>
      </table>
<%
End sub
Set rs = Server.CreateObject("ADODB.RecordSet")
Set rs2 = Server.CreateObject("ADODB.RecordSet")
Set cnn=Server.CreateObject("ADODB.Connection")
cnn.open sDSN
alreadygotadmin = getadminsettings()
sSQL="SELECT adminCert FROM admin WHERE adminID=1"
rs.Open sSQL,cnn,0,1
thecert=rs("adminCert")&""
rs.Close
if request.querystring("sig")<>"" AND request.querystring("tx")<>"" AND request.querystring("st")<>"" then
	success = getpayprovdetails(1,data1,data2,data3,demomode,ppmethod)
	if data2="" then
		errtext = "Identity token for PayPal Payment Data Transfer (PDT) not set."
		Call order_failed
	else
		set objHttp = Server.CreateObject("Msxml2.ServerXMLHTTP")
		objHttp.open "POST", "https://www." & IIfVr(demomode, "sandbox.", "") & "paypal.com/cgi-bin/webscr", false
		objHttp.setRequestHeader "Content-type", "application/x-www-form-urlencoded"
		objHttp.Send "&cmd=_notify-synch&tx="&request.querystring("tx")&"&at="&data2
		sQuerystring = objHttp.responseText
		if mid(sQuerystring,1,7) = "SUCCESS" Then
			sQuerystring = mid(sQuerystring,9)
			sParts = Split(sQuerystring, vbLf)
			iParts = UBound(sParts) - 1
			pending_reason = ""
			for i = 0 to iParts
				aParts = split(sParts(i), "=", 2)
				sKey = aParts(0)
				sValue = aParts(1)
				' response.write sKey & " : " & sValue & "<br>"
				select case sKey
				case "payment_status"
					payment_status = sValue
				case "pending_reason"
					pending_reason = sValue
				case "custom"
					ordID = replace(sValue,"'","")
				case "txn_id"
					txn_id = replace(sValue,"'","")
				end select
			next
			sSQL = "SELECT ordAuthNumber FROM orders WHERE ordPayProvider=1 AND ordStatus>=3 AND ordAuthNumber='"&txn_id&"' AND ordID=" & ordID
			success = (txn_id<>"")
			rs.Open sSQL,cnn,0,1
				if rs.EOF then success = FALSE
			rs.Close
			if success then
				Call do_order_success(ordID,emailAddr,FALSE,TRUE,FALSE,FALSE,FALSE)
			else
				cnn.Execute("UPDATE cart SET cartCompleted=2 WHERE cartCompleted=0 AND cartOrderID="&ordID)
				cnn.Execute("UPDATE orders SET ordAuthNumber='no ipn' WHERE ordAuthNumber='' AND ordPayProvider=1 AND ordID="&ordID)
				xxThkErr = ""
				if payment_status="Pending" then
					errtext = xxPPPend
				else
					errtext = xxNoCnf
				end if
				Call order_failed
			end if
		else
			errtext = sQuerystring
			Call order_failed
		end if
	end if
elseif Request.Form("custom")<>"" then ' PayPal
	ordID = trim(replace(request.form("custom"), "'", ""))
	txn_id = trim(replace(request.form("txn_id"), "'", ""))
	sSQL = "SELECT ordAuthNumber FROM orders WHERE ordPayProvider=1 AND ordStatus>=3 AND ordAuthNumber='"&txn_id&"' AND ordID=" & ordID
	success = (txn_id<>"")
	rs.Open sSQL,cnn,0,1
		if rs.EOF then success = FALSE
	rs.Close
	if success then
		Call do_order_success(ordID,emailAddr,FALSE,TRUE,FALSE,FALSE,FALSE)
	else
		cnn.Execute("UPDATE cart SET cartCompleted=2 WHERE cartCompleted=0 AND cartOrderID="&ordID)
		cnn.Execute("UPDATE orders SET ordAuthNumber='no ipn' WHERE ordAuthNumber='' AND ordPayProvider=1 AND ordID="&ordID)
		xxThkErr = ""
		if request.form("payment_status")="Pending" then
			errtext = xxPPPend
		else
			errtext = xxNoCnf
		end if
		Call order_failed
	end if
elseif request.form("method")="paypalexpress" AND Request.Form("token")<>"" then ' PayPal Express
	success = getpayprovdetails(18,username,data2pwd,data2hash,demomode,ppmethod)
	ordID = replace(trim(request.form("ordernumber")), "'", "")
	token = trim(request.form("token"))
	payerid = trim(request.form("payerid"))
	ordAuthNumber = "" : status = ""
	if demomode then sandbox = ".sandbox" else sandbox = ""
	sSQL = "SELECT ordShipping,ordStateTax,ordCountryTax,ordHandling,ordTotal,ordDiscount,ordAuthNumber,ordEmail FROM orders WHERE ordID=" & ordID
	rs.Open sSQL,cnn,0,1
	if NOT rs.EOF then
		session.LCID = 1033
		amount = formatnumber((rs("ordShipping")+rs("ordStateTax")+rs("ordCountryTax")+rs("ordTotal")+rs("ordHandling"))-rs("ordDiscount"),2,-1,0,0)
		session.LCID = saveLCID
		if rs("ordEmail")=trim(request.form("email")) then ordAuthNumber = trim(rs("ordAuthNumber")&"")
	else
		success = FALSE
	end if
	rs.Close
	if success then
		if ordAuthNumber="" then
			sXML = ppsoapheader(username, data2pwd, data2hash) & _
				"  <soap:Body>" & _
				"    <DoExpressCheckoutPaymentReq xmlns=""urn:ebay:api:PayPalAPI"">" & _
				"      <DoExpressCheckoutPaymentRequest>" & _
				"        <Version xmlns=""urn:ebay:apis:eBLBaseComponents"">1.00</Version>" & _
				"        <DoExpressCheckoutPaymentRequestDetails xmlns=""urn:ebay:apis:eBLBaseComponents"">" & _
				"          <PaymentAction>" & IIfVr(ppmethod=1, "Authorization", "Sale") & "</PaymentAction>" & _
				"          <Token>" & token & "</Token>" & _
				"          <PayerID>" & payerid & "</PayerID>" & _
				"          <PaymentDetails>" & _
				"            <OrderTotal currencyID=""" & countryCurrency & """>" & amount & "</OrderTotal>" & _
				"            <ButtonSource>ecommercetemplates.asp.ecommplus</ButtonSource>" & _
				"          </PaymentDetails>" & _
				"        </DoExpressCheckoutPaymentRequestDetails>" & _
				"      </DoExpressCheckoutPaymentRequest>" & _
				"    </DoExpressCheckoutPaymentReq>" & _
				"  </soap:Body>" & _
				"</soap:Envelope>"
			if callxmlfunction("https://api-aa" & IIfVr(sandbox="" AND data2hash<>"", "-3t", "") & sandbox & ".paypal.com/2.0/", sXML, res, IIfVr(data2hash<>"","",username), "WinHTTP.WinHTTPRequest.5.1", vsRESPMSG, FALSE) then
				set xmlDoc = Server.CreateObject("MSXML2.DOMDocument")
				xmlDoc.validateOnParse = False
				xmlDoc.loadXML (res)
				Set nodeList = xmlDoc.getElementsByTagName("SOAP-ENV:Body")
				Set n = nodeList.Item(0)
				for j = 0 to n.childNodes.length - 1
					Set e = n.childNodes.Item(i)
					if e.nodeName = "DoExpressCheckoutPaymentResponse" then
						for k = 0 To e.childNodes.length - 1
							Set t = e.childNodes.Item(k)
							if t.nodeName = "Token" then
								if t.firstChild.nodeValue = "Success" then success=TRUE
							elseif t.nodeName = "DoExpressCheckoutPaymentResponseDetails" then
								set ff = t.childNodes
								for kk = 0 to ff.length - 1
									set gg = ff.item(kk)
									if gg.nodeName = "PaymentInfo" then
										set hh = gg.childNodes
										for ll = 0 to hh.length - 1
											set ii = hh.item(ll)
											if ii.nodeName = "PaymentStatus" then
												status = ii.firstChild.nodeValue
											elseif ii.nodeName = "PendingReason" then
												pendingreason = ii.firstChild.nodeValue
											elseif ii.nodeName = "TransactionID" then
												txn_id = ii.firstChild.nodeValue
											end if
										next
									end if
								next
							elseif t.nodeName = "Errors" then
								set ff = t.childNodes
								for kk = 0 to ff.length - 1
									set gg = ff.item(kk)
									if gg.nodeName = "ShortMessage" then
										errormsg = gg.firstChild.nodeValue & "<br>" & errormsg
									elseif gg.nodeName = "LongMessage" then
										errormsg= errormsg & gg.firstChild.nodeValue
									elseif gg.nodeName = "ErrorCode" then
										errcode = gg.firstChild.nodeValue
									end if
								next
							end if
						next
					end if
				next
			else
				success = FALSE
			end if
		else
			status = "Refresh"
		end if
		if status = "Completed" OR status = "Pending" then
			if status = "Pending" AND pendingreason <> "" then txn_id = status & ": " & pendingreason & "<br>" & txn_id
			do_stock_management(ordID)
			cnn.Execute("UPDATE cart SET cartCompleted=1 WHERE cartOrderID=" & replace(ordID, "'", ""))
			cnn.Execute("UPDATE orders SET ordStatus=3,ordAuthNumber='" & txn_id & "' WHERE ordPayProvider=19 AND ordID=" & replace(ordID, "'", ""))
			call do_order_success(ordID,emailAddr,sendEmail,TRUE,TRUE,TRUE,TRUE)
		elseif status = "Refresh" then
			call do_order_success(ordID,emailAddr,sendEmail,FALSE,FALSE,FALSE,FALSE)
		else
			call order_failed()
		end if
	else
		call order_failed()
	end if
elseif Request.querystring("ncretval")<>"" AND Request.querystring("ncsessid")<>"" then ' NOCHEX
	ordID = trim(replace(request.querystring("ncretval"), "'", ""))
	ncsessid = trim(replace(request.querystring("ncsessid"), "'", ""))
	sSQL = "SELECT ordAuthNumber FROM orders WHERE ordPayProvider=6 AND ordStatus>=3 AND ordSessionID="&ncsessid&" AND ordID=" & ordID
	success = TRUE
	rs.Open sSQL,cnn,0,1
		if rs.EOF then success = FALSE
	rs.Close
	if success then
		Call do_order_success(ordID,emailAddr,FALSE,TRUE,FALSE,FALSE,FALSE)
	else
		cnn.Execute("UPDATE cart SET cartCompleted=2 WHERE cartCompleted=0 AND cartOrderID="&ordID)
		cnn.Execute("UPDATE orders SET ordAuthNumber='no apc' WHERE ordAuthNumber='' AND ordPayProvider=6 AND ordID="&ordID)
		errtext = xxNoCnf
		xxThkErr = ""
		Call order_failed
	end if
elseif Request.Form("xxpreauth")<>"" then
	ordID = trim(replace(Request.Form("xxpreauth"), "'", ""))
	thesessionid = trim(replace(request.form("thesessionid"),"'",""))
	themethod = trim(replace(request.form("xxpreauthmethod"),"'",""))
	success = getpayprovdetails(themethod,data1,data2,data3,demomode,ppmethod)
	if success then
		sSQL = "SELECT ordAuthNumber FROM orders WHERE ordSessionID="&thesessionid&" AND ordID=" & ordID
		rs.Open sSQL,cnn,0,1
			if rs.EOF then
				success = FALSE
			else
				success = (trim(rs("ordAuthNumber")&"")<>"")
			end if
		rs.Close
	end if
	if success then
		Call order_success(ordID,emailAddr,sendEmail)
	else
		Call order_failed
	end if
elseif Request.Form("cart_order_id") <> "" AND request.form("order_number")<>"" then ' 2Checkout Transaction
	if Trim(request.form("credit_card_processed"))="Y" then
		ordID = trim(replace(request.form("cart_order_id"),"'",""))
		success = getpayprovdetails(2,acctno,md5key,data3,demomode,ppmethod)
		keysmatch=TRUE
		if md5key<>"" then
			theirkey = Trim(Request.Form("key"))
			ourkey = Trim(UCase(calcMD5(md5key&acctno&IIfVr(demomode,"1",Request.Form("order_number"))&Request.Form("total"))))
			if ourkey=theirkey then keysmatch=true else keysmatch=false
		end if
		if success AND keysmatch then
			do_stock_management(ordID)
			cnn.Execute("UPDATE cart SET cartCompleted=1 WHERE cartOrderID="&ordID)
			cnn.Execute("UPDATE orders SET ordStatus=3,ordAuthNumber='"&trim(replace(request.form("order_number"),"'","''"))&"' WHERE ordPayProvider=2 AND ordID="&ordID)
			Call order_success(ordID,emailAddr,sendEmail)
		else
			Call order_failed
		end if
	else
		Call order_failed
	end if
elseif Request.Form("CUSTID") <> "" AND Request.Form("AUTHCODE") <> "" then ' PayFlow Link
	success = getpayprovdetails(8,data1,data2,data3,demomode,ppmethod)
	if success AND Trim(request.form("RESULT"))="0" then
		ordID = trim(replace(request.form("CUSTID"),"'",""))
		do_stock_management(ordID)
		cnn.Execute("UPDATE cart SET cartCompleted=1 WHERE cartOrderID="&ordID)
		cnn.Execute("UPDATE orders SET ordStatus=3,ordAVS='"&replace(trim(request.form("AVSDATA")),"'","''")&"',ordCVV='"&replace(trim(request.form("CSCMATCH")),"'","''")&"',ordAuthNumber='"&replace(trim(request.form("AUTHCODE")),"'","''")&"' WHERE ordPayProvider=8 AND ordID="&ordID)
		Call order_success(ordID,emailAddr,sendEmail)
	else
		Call order_failed
	end if
elseif Request.Form("emailorder")<>"" OR Request.Form("secondemailorder")<>"" then
	ordGndTot=1
	if emailorderstatus<>"" then ordStatus=emailorderstatus else ordStatus=3
	if Request.Form("emailorder")<>"" then
		ordID = trim(replace(request.form("emailorder"),"'",""))
		ppid = 4
	else
		ordID = trim(replace(request.form("secondemailorder"),"'",""))
		ppid = 17
	end if
	thesessionid = trim(replace(request.form("thesessionid"),"'",""))
	sSQL = "SELECT ordAuthNumber,((ordShipping+ordStateTax+ordCountryTax+ordTotal+ordHandling)-ordDiscount) AS ordGndTot FROM orders WHERE ordSessionID="&thesessionid&" AND ordID=" & ordID
	rs.Open sSQL,cnn,0,1
		if rs.EOF then success = FALSE else success = TRUE : ordGndTot = rs("ordGndTot")
	rs.Close
	session.LCID = 1033
	sSQL = "SELECT payProvShow FROM payprovider WHERE (payProvEnabled=1 OR "&ordGndTot&"=0) AND payProvID="&ppid
	session.LCID = saveLCID
	rs.Open sSQL,cnn,0,1
		if rs.EOF then success=FALSE else authnumber = rs("payProvShow")
	rs.Close
	if success then
		if ordStatus >= 3 then do_stock_management(ordID)
		sSQL="UPDATE cart SET cartCompleted=1 WHERE cartOrderID="&ordID
		cnn.Execute(sSQL)
		sSQL="UPDATE orders SET ordStatus="&ordStatus&",ordAuthNumber='"&left(replace(authnumber,"'","''"),48)&"' WHERE (ordPayProvider="&ppid&" OR (ordTotal-ordDiscount)<=0) AND ordID="&ordID
		cnn.Execute(sSQL)
		Call order_success(ordID,emailAddr,sendEmail)
	else
		Call order_failed
	end if
elseif Request.QueryString("OrdNo")<>"" AND Request.QueryString("RefNo")<>"" then ' PSiGate
	sSQL = "SELECT payProvID FROM payprovider WHERE payProvEnabled=1 AND payProvID=11 OR payProvID=12"
	rs.Open sSQL,cnn,0,1
	if NOT rs.EOF then success=TRUE else success=FALSE
	rs.Close
	if success then
		ordID = trim(replace(request.QueryString("OrdNo"),"'",""))
		do_stock_management(ordID)
		cnn.Execute("UPDATE cart SET cartCompleted=1 WHERE cartOrderID="&ordID)
		cnn.Execute("UPDATE orders SET ordStatus=3,ordAuthNumber='"&trim(replace(request.QueryString("RefNo"),"'","''"))&"' WHERE (ordPayProvider=11 OR ordPayProvider=12) AND ordID="&ordID)
		Call order_success(ordID,emailAddr,sendEmail)
	else
		Call order_failed
	end if
elseif trim(Request.Form("ponumber"))<>"" AND (trim(Request.Form("approval_code"))<>"" OR trim(request.form("failReason"))<>"") then ' Linkpoint
	if getpayprovdetails(16,data1,data2,data3,demomode,ppmethod) then
		ordID=trim(replace(request.form("ponumber"),"'",""))
		ordIDa=split(ordID,",")
		ordID=ordIDa(0)
		theauthcode=trim(replace(request.form("approval_code"),"'",""))
		thesuccess=lcase(trim(request.form("status")))
		if (thesuccess="approved" OR thesuccess="submitted") AND theauthcode<>"" then
			autharr = split(theauthcode,":")
			if autharr(0)="Y" AND UBOUND(autharr) >= 3 then
				theauthcode = autharr(1)
				theavscode = autharr(2)
				do_stock_management(ordID)
				cnn.Execute("UPDATE cart SET cartCompleted=1 WHERE cartOrderID="&ordID)
				cnn.Execute("UPDATE orders SET ordStatus=3,ordAVS='"&left(theavscode,3)&"',ordCVV='"&right(theavscode,1)&"',ordAuthNumber='"&left(theauthcode,6)&"',ordTransID='"&right(theauthcode,len(theauthcode)-6)&"' WHERE ordPayProvider=16 AND ordID="&ordID)
				Call order_success(ordID,emailAddr,sendEmail)
			else
				errtext = "Invalid auth code"
				Call order_failed
			end if
		else
			errtext = request.form("failReason")
			errtextarr = split(errtext, ":")
			if IsArray(errtextarr) then
				if UBOUND(errtextarr)>0 then errtext = errtextarr(1)
			end if
			Call order_failed
		end if
	else
		Call order_failed
	end if
elseif Request.Form("docapture")="vsprods" then
	success = getpayprovdetails(10,data1,data2,data3,demomode,ppmethod)
	if success then
		if capturecardorderstatus<>"" then ordStatus=capturecardorderstatus else ordStatus=3
		encryptmethod = LCase(encryptmethod&"")
		if encryptmethod<>"none" then
			thecert = Replace(thecert,"-----BEGIN CERTIFICATE-----","")
			thecert = Trim(Replace(thecert,"-----END CERTIFICATE-----",""))
			err.number = 0
			on error resume next
				Set CM = Server.CreateObject("Persits.CryptoManager")
				if err.number <> 0 then
					success=false
					errmsg = "Error ! Cannot invoke Persits ASPEncrypt"
				end if
				if success then
					Set Blob = CM.CreateBlob
					err.number = 0
						' Set Context = CM.OpenContextEx( "Microsoft Enhanced Cryptographic Provider v1.0", "mycontainer", True)
						Set Context = CM.OpenContext("", True)
						if err.number = 0 then
							Blob.Base64 = thecert
							Set Cert = CM.ImportCertFromBlob(Blob)
							if err.number <> 0 then
								success=false
								errmsg = "Error ! Cannot import certificate." & "<br />" & err.description
							end if
						else
							success=false
							errmsg = "Error ! Trying to invoke Microsoft Enhanced Cryptographic Provider v1.0" & "<br />" & err.description
						end if
				end if
			on error goto 0
		end if
	end if
	if success then
		ordID=trim(replace(request.form("ordernumber"),"'",""))
		if encryptmethod="none" then
			enctext = Trim(Replace(Request.Form("ACCT"),"'","")) & "&" & Trim(Replace(Request.Form("EXMON"),"'",""))&"/"&Trim(Replace(Request.Form("EXYEAR"),"'","")) & "&" & Trim(Replace(Request.Form("CVV2"),"'","")) & "&" & Trim(Replace(Request.Form("IssNum"),"'","")) & "&" & Server.URLEncode(Trim(Request.Form("cardname")))
		elseif encryptmethod="aspencrypt" OR encryptmethod="" then
			' response.write "Cert Version: " & Cert.Version & ":" & Cert.SerialNumber & ":" & Cert.NotAfter & ":" & Cert.NotBefore & "<br />"
			Set Msg = Context.CreateMessage(True)
			Msg.AddRecipientCert Cert
			enctext = Msg.EncryptText(Trim(Replace(Request.Form("ACCT"),"'","")) & "&" & Trim(Replace(Request.Form("EXMON"),"'",""))&"/"&Trim(Replace(Request.Form("EXYEAR"),"'","")) & "&" & Trim(Replace(Request.Form("CVV2"),"'","")) & "&" & Trim(Replace(Request.Form("IssNum"),"'","")) & "&" & Server.URLEncode(Trim(Request.Form("cardname"))))
		end if
		do_stock_management(ordID)
		cnn.Execute("UPDATE cart SET cartCompleted=1 WHERE cartOrderID="&ordID)
		cnn.Execute("UPDATE orders SET ordStatus="&ordStatus&",ordAuthNumber='Card Capture',ordCNum='"&enctext&"' WHERE ordPayProvider=10 AND ordID="&ordID)
		Call order_success(ordID,emailAddr,sendEmail)
	else
		response.write "<p>&nbsp;</p><p align='center'><strong><font color='#FF0000'>"&errmsg&"</font></strong></p>"
		Call order_failed
	end if
elseif Request.QueryString("OrdNo")<>"" AND Request.QueryString("ErrMsg") <> "" then ' PSiGate Error Reporting
	errtext = Request.QueryString("ErrMsg")
	Call order_failed
else
%>
<!--#include file="customppreturn.asp"-->
<%
end if
cnn.Close
set rs = nothing
set rs2 = nothing
set cnn = nothing
%>