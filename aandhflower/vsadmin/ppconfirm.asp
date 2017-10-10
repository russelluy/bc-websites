<%@LANGUAGE="VBScript"%>
<!--#include file="inc/md5.asp"-->
<!--#include file="db_conn_open.asp"-->
<!--#include file="includes.asp"-->
<!--#include file="inc/incemail.asp"-->
<!--#include file="inc/languagefile.asp"-->
<!--#include file="inc/incfunctions.asp"-->
<%
Dim str,Txn_id,Payment_status,objHttp
Set rs = Server.CreateObject("ADODB.RecordSet")
Set rs2 = Server.CreateObject("ADODB.RecordSet")
Set cnn=Server.CreateObject("ADODB.Connection")
cnn.open sDSN
alreadygotadmin = getadminsettings()
sSQL = "SELECT payProvDemo,payProvData1,payProvData2,payProvMethod FROM payprovider WHERE payProvID=1"
rs.Open sSQL,cnn,0,1
demomode=(rs("payProvDemo")="1")
data1=trim(rs("payProvData1")&"")
data2=trim(rs("payProvData2")&"")
ppmethod=Int(rs("payProvMethod"))
rs.Close
' read post from PayPal system and add 'cmd'
str = Request.Form
' post back to PayPal system to validate
str = str & "&cmd=_notify-validate"
set objHttp = Server.CreateObject("Msxml2.ServerXMLHTTP")
objHttp.open "POST", "https://www." & IIfVr(demomode, "sandbox.", "") & "paypal.com/cgi-bin/webscr", false
objHttp.Send str
' assign posted variables to local variables
Receiver_email = Request.Form("receiver_email")
Item_number = Request.Form("item_number")
Invoice = Request.Form("invoice")
Payment_status = Request.Form("payment_status")
Payment_gross = Request.Form("payment_gross")
Txn_id = Request.Form("txn_id")
ordID = trim(replace(request.form("custom"), "'", ""))
Payer_email = Request.Form("payer_email")
receipt_id = trim(request.form("receipt_id"))
address_status = lcase(trim(request.form("address_status")))
if address_status="confirmed" then
	avs = "Y"
elseif address_status="unconfirmed" then
	avs = "N"
else
	avs = "U"
end if
payer_status = lcase(trim(request.form("payer_status")))
if payer_status="verified" then
	cvv = "Y"
elseif payer_status="unverified" then
	cvv = "N"
else
	cvv = "U"
end if
if instr(ordID,":")=0 then
	' Check notification validation
	if (objHttp.status <> 200 ) then
	' HTTP error handling
	elseif (objHttp.responseText = "VERIFIED") AND (ordID<>"") then
		' check that Payment_status=Completed
		' check that Txn_id has not been previously processed
		' check that Receiver_email is an email address in your PayPal account
		' process payment
		if Payment_status="Completed" then
			do_stock_management(ordID)
			cnn.Execute("UPDATE cart SET cartCompleted=1 WHERE cartOrderID="&ordID)
			cnn.Execute("UPDATE orders SET ordAVS='"&avs&"',ordCVV='"&cvv&"',ordStatus=3,ordAuthNumber='"&Txn_id&"',ordTransID='"&receipt_id&"' WHERE ordID="&ordID)
			Call do_order_success(ordID,emailAddr,sendEmail,FALSE,TRUE,TRUE,TRUE)
		elseif Payment_status="Pending" then
			cnn.Execute("UPDATE cart SET cartCompleted=2 WHERE cartCompleted=0 AND cartOrderID="&ordID)
			cnn.Execute("UPDATE orders SET ordAVS='"&avs&"',ordCVV='"&cvv&"',ordAuthNumber='Pending: " & replace(Request.Form("pending_reason"),"'","''") & "' WHERE ordPayProvider=1 AND ordID="&ordID)
		end if
	elseif (objHttp.responseText = "INVALID") then
		' log for manual investigation
	else 
		if debugmode=TRUE then response.write objHttp.responseText ' error
	end if
end if
if debugmode=TRUE then
	if htmlemails=true then emlNl = "<br />" else emlNl=vbCrLf
	emailtxt = "Status: " & Payment_status & emlNl & "Txn ID: " & Txn_id & emlNl & "Response: " & objHttp.responseText & emlNl & "Ord ID: " & ordID & emlNl & "Pending Reason: " & Request.Form("pending_reason") & emlNl
	for each objItem In Request.Form
		emailtxt = emailtxt & objItem & " : " & Request.Form(objItem) & emlNl
	next
	Call DoSendEmailEO(emailAddr,emailAddr,"","ppconfirm.asp debug",emailtxt,emailObject,themailhost,theuser,thepass)
end if
set objHttp = nothing
%>