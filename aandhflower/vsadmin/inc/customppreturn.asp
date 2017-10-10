<%
theorderid=Trim(replace(Request.Form("cart_order_id"),"'",""))
theauthcode=Trim(replace(Request.Form("order_number"),"'",""))
thesuccess=Trim(Request.Form("credit_card_processed"))
if theorderid<>"" AND theauthcode<>"" AND thesuccess="Y" then
	' You should not normally need to change the code below
	do_stock_management(theorderid)
	sSQL="UPDATE cart SET cartCompleted=1 WHERE cartOrderID="&theorderid
	cnn.Execute(sSQL)
	sSQL="UPDATE orders SET ordStatus=3,ordAuthNumber='"&theauthcode&"' WHERE ordPayProvider=14 AND ordID="&theorderid
	cnn.Execute(sSQL)
	Call order_success(theorderid,emailAddr,sendEmail)
else
	' Make sure you leave this condition here. It calls a failure routine if no match is found for any payment system.
	Call order_failed
end if
%>