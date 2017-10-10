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

Set rs = Server.CreateObject("ADODB.RecordSet")
Set rs2 = Server.CreateObject("ADODB.RecordSet")
Set cnn=Server.CreateObject("ADODB.Connection")
cnn.open sDSN
if request.querystring("gid")<>"" then
	ordID = replace(request.querystring("gid"),"'","")
	sSQL = "SELECT ordPayProvider,ordAuthNumber,payProvData1,payProvData2,payProvDemo FROM orders INNER JOIN payprovider ON orders.ordPayProvider=payprovider.payProvID WHERE ordID=" & ordID
	rs.Open sSQL,cnn,0,1
	authcode=rs("ordAuthNumber")
	googledata1=rs("payProvData1")
	googledata2=rs("payProvData2")
	googledemomode=rs("payProvDemo")
	rs.Close
	if request.querystring("act")="charge" then
		' First set the status to process-order
		set objHttp = Server.CreateObject("WinHTTP.WinHTTPRequest.5.1")
		theurl="https://"&IIfVr(googledemomode, "sandbox", "checkout")&".google.com/cws/v2/Merchant/"&googledata1&"/request"
		objHttp.open "POST", theurl, false
		objHttp.setRequestHeader "Authorization", "Basic " & vrbase64_encrypt(googledata1&":"&googledata2)
		objHttp.setRequestHeader "Content-Type", "application/xml"
		objHttp.setRequestHeader "Accept", "application/xml"
		objHttp.Send "<?xml version=""1.0"" encoding=""UTF-8""?>" & _
			"<process-order xmlns=""http://checkout.google.com/schema/2"" google-order-number="""&authcode&"""/>"
		set objHttp = nothing
		
		acttext = "<charge-order xmlns=""http://checkout.google.com/schema/2"" google-order-number="""&authcode&"""></charge-order>"
	elseif request.querystring("act")="cancel" then
		acttext = "<cancel-order xmlns=""http://checkout.google.com/schema/2"" google-order-number="""&authcode&"""><reason>Cancelled by store admin on " & Date() & ".</reason></cancel-order>"
	elseif request.querystring("act")="refund" then
		acttext = "<refund-order xmlns=""http://checkout.google.com/schema/2"" google-order-number="""&authcode&"""><reason>Refunded by store admin on " & Date() & ".</reason></refund-order>"
	elseif request.querystring("act")="ship" then
		' First set the status to process-order
		set objHttp = Server.CreateObject("WinHTTP.WinHTTPRequest.5.1")
		theurl="https://"&IIfVr(googledemomode, "sandbox", "checkout")&".google.com/cws/v2/Merchant/"&googledata1&"/request"
		objHttp.open "POST", theurl, false
		objHttp.setRequestHeader "Authorization", "Basic " & vrbase64_encrypt(googledata1&":"&googledata2)
		objHttp.setRequestHeader "Content-Type", "application/xml"
		objHttp.setRequestHeader "Accept", "application/xml"
		objHttp.Send "<?xml version=""1.0"" encoding=""UTF-8""?>" & _
			"<process-order xmlns=""http://checkout.google.com/schema/2"" google-order-number="""&authcode&"""/>"
		set objHttp = nothing
		
		acttext = "<deliver-order xmlns=""http://checkout.google.com/schema/2"" google-order-number="""&authcode&""">"
		if request.querystring("carrier")<>"" AND request.querystring("trackno")<>"" then
			sSQL = "UPDATE orders SET ordTrackNum='"&replace(request.querystring("trackno"),"'","")&"',ordShipCarrier="&replace(request.querystring("carrier"),"'","")&" WHERE ordID=" & ordID
			cnn.Execute(sSQL)
			acttext = acttext & "<tracking-data><carrier>"
			select case request.querystring("carrier")
				case "3"
					acttext = acttext & "USPS"
				case "4"
					acttext = acttext & "UPS"
				case "7"
					acttext = acttext & "FedEx"
				case "8"
					acttext = acttext & "DHL"
				case else
					acttext = acttext & "Other"
			end select
			acttext = acttext & "</carrier><tracking-number>"&trim(request.querystring("trackno"))&"</tracking-number></tracking-data>"
		end if
		acttext = acttext & "</deliver-order>"
	elseif request.querystring("act")="message" then
		' First set the status to process-order
		set objHttp = Server.CreateObject("WinHTTP.WinHTTPRequest.5.1")
		theurl="https://"&IIfVr(googledemomode, "sandbox", "checkout")&".google.com/cws/v2/Merchant/"&googledata1&"/request"
		objHttp.open "POST", theurl, false
		objHttp.setRequestHeader "Authorization", "Basic " & vrbase64_encrypt(googledata1&":"&googledata2)
		objHttp.setRequestHeader "Content-Type", "application/xml"
		objHttp.setRequestHeader "Accept", "application/xml"
		objHttp.Send "<?xml version=""1.0"" encoding=""UTF-8""?>" & _
			"<process-order xmlns=""http://checkout.google.com/schema/2"" google-order-number="""&authcode&"""/>"
		set objHttp = nothing
		
		acttext = "<send-buyer-message xmlns=""http://checkout.google.com/schema/2"" google-order-number="""&authcode&"""><message>"&request.form("googlemessage")&"</message><send-email>true</send-email></send-buyer-message>"
	end if
	set objHttp = Server.CreateObject("WinHTTP.WinHTTPRequest.5.1")
	theurl="https://"&IIfVr(googledemomode, "sandbox", "checkout")&".google.com/cws/v2/Merchant/"&googledata1&"/request"
	objHttp.open "POST", theurl, false
	objHttp.setRequestHeader "Authorization", "Basic " & vrbase64_encrypt(googledata1&":"&googledata2)
	objHttp.setRequestHeader "Content-Type", "application/xml"
	objHttp.setRequestHeader "Accept", "application/xml"
	'on error resume next
	err.number=0
	objHttp.Send "<?xml version=""1.0"" encoding=""UTF-8""?>" & acttext
	if err.number <> 0 OR (objHttp.status <> 200 AND objHttp.status <> 400)Then
		response.write "<font color=""#FF0000"">" & "Error, couldn't update order " & ordID & "</font><br/>"
	else
		set xmlDoc = Server.CreateObject("MSXML2.DOMDocument")
		xmlDoc.validateOnParse = False
		xmlDoc.loadXML (objHttp.responseText)
		Set errobj = xmlDoc.getElementsByTagName("error-message")
		if errobj.length > 0 then
			response.write "<font color=""#FF0000"">" & errobj.Item(0).firstChild.nodeValue & "</font><br/>"
		else
			response.write "Finished updating order " & ordID
		end if
		' response.write Replace(Replace(objHttp.responseText,"</","&lt;/"),"<","<br />&lt;")&"<HR>"
	end if
	on error goto 0
	set objHttp = nothing
end if
cnn.Close
%>