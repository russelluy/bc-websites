<%@LANGUAGE="VBScript"%>
<!--#include file="inc/md5.asp"-->
<!--#include file="db_conn_open.asp"-->
<!--#include file="includes.asp"-->
<!--#include file="inc/incemail.asp"-->
<!--#include file="inc/languagefile.asp"-->
<!--#include file="inc/incfunctions.asp"-->
<%
Dim str,Txn_id,Payment_status,objHttp
' read post from PayPal system and add 'cmd'
str = Request.Form
' assign posted variables to local variables
Receiver_email = Replace(Request.Form("to_email"),"'","")
Payment_gross = Replace(Request.Form("amount"),"'","")
Payer_email = Replace(Request.Form("from_email"),"'","")
ordID = trim(Replace(Request.Form("order_id"),"'",""))
Txn_id = Replace(Request.Form("transaction_id"),"'","")

' post back to NOCHEX system to validate
set objHttp = Server.CreateObject("Msxml2.ServerXMLHTTP")
objHttp.open "POST", "https://www.nochex.com/nochex.dll/apc/apc", false
objHttp.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
objHttp.Send str

' Check notification validation
if (objHttp.status <> 200 ) then
' HTTP error handling
elseif (objHttp.responseText = "AUTHORISED") then
	' check that Payment_status=Completed
	' check that Txn_id has not been previously processed
	' check that Receiver_email is an email address in your PayPal account
	' process payment
	Set rs = Server.CreateObject("ADODB.RecordSet")
	Set rs2 = Server.CreateObject("ADODB.RecordSet")
	Set cnn=Server.CreateObject("ADODB.Connection")
	cnn.open sDSN
	alreadygotadmin = getadminsettings()
	do_stock_management(ordID)
	sSQL="UPDATE cart SET cartCompleted=1 WHERE cartOrderID="&ordID
	cnn.Execute(sSQL)
	sSQL="UPDATE orders SET ordStatus=3,ordAuthNumber='"&Txn_id&"' WHERE ordID="&ordID
	cnn.Execute(sSQL)
	Call do_order_success(ordID,emailAddr,sendEmail,FALSE,TRUE,TRUE,TRUE)
elseif (objHttp.responseText = "DECLINED") then
' log for manual investigation
else 
' error
end if
set objHttp = nothing
%>
