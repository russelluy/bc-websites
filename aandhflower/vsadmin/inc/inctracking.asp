<%
Set rs = Server.CreateObject("ADODB.RecordSet")
Set rs2 = Server.CreateObject("ADODB.RecordSet")
Set cnn=Server.CreateObject("ADODB.Connection")
cnn.open sDSN
alreadygotadmin = getadminsettings()
incupscopyright=false
incfedexcopyright=false
if request("carrier")<>"" then
	theshiptype=request("carrier")
else
	possshiptypes=0
	if defaulttrackingcarrier<>"" then theshiptype=defaulttrackingcarrier else theshiptype="ups"
	if shipType=3 OR alternateratesusps<>"" then
		theshiptype="usps"
		possshiptypes=possshiptypes+1
	end if
	if shipType=4 OR alternateratesups<>"" then
		theshiptype="ups"
		incupscopyright=true
		possshiptypes=possshiptypes+1
	end if
	if shipType=7 OR alternateratesfedex<>"" then
		theshiptype="fedex"
		incfedexcopyright=true
		possshiptypes=possshiptypes+1
	end if
	if possshiptypes>1 then theshiptype="undecided"
end if
%>
<script language="javascript" type="text/javascript">
<!--
function viewlicense()
{
	var prnttext = '<html><head><STYLE TYPE="text/css">A:link {COLOR: #333333; TEXT-DECORATION: none}A:visited {COLOR: #333333; TEXT-DECORATION: none}A:active {COLOR: #333333; TEXT-DECORATION: none}A:hover {COLOR: #f39000; TEXT-DECORATION: none}TD {FONT-FAMILY: Verdana;}P {FONT-FAMILY: Verdana;}HR {color: #637BAD;height: 1px;}</STYLE></head><body><table width="100%" border="0" cellspacing="1" cellpadding="3">\n';
	prnttext += '<tr><td colspan="2" align="center"><a href="javascript:window.close()"><strong>Close Window</strong></a></td></tr>';
	prnttext += '<tr><td width="40"><img src="images/LOGO_S.gif"  alt="UPS" /></td><td><p><font size="3" face="Verdana"><strong>UPS Tracking Terms and Conditions</strong></font></p></td></tr>';
	prnttext += '<tr><td colspan="2"><p><font size="2" face="Verdana">The UPS package tracking systems accessed via this Web Site (the &quot;Tracking Systems&quot;) and tracking information obtained through this Web Site (the &quot;Information&quot;) are the private property of UPS. UPS authorizes you to use the Tracking Systems solely to track shipments tendered by or for you to UPS for delivery and for no other purpose. Without limitation, you are not authorized to make the Information available on any web site or otherwise reproduce, distribute, copy, store, use or sell the Information for commercial gain without the express written consent of UPS. This is a personal service, thus your right to use the Tracking Systems or Information is non-assignable. Any access or use that is inconsistent with these terms is unauthorized and strictly prohibited.</font></p></td></tr>';
	prnttext += '<tr><td colspan="2"><hr /><font size="1" face="Verdana">Copyright&nbsp;&copy; 1994-2003 United Parcel Service of America, Inc. All rights reserved.</font></td></tr>';
	prnttext += '<tr><td colspan="2" align="center">&nbsp;<br /><a href="javascript:window.close()"><strong>Close Window</strong></a></td></tr>';
	prnttext += '</table></body></html>';
	var newwin = window.open("","viewlicense",'menubar=no, scrollbars=yes, width=500, height=400, directories=no,location=no,resizable=yes,status=no,toolbar=no');
	newwin.document.open();
	newwin.document.write(prnttext);
	newwin.document.close();
}
function checkaccept()
{
  if (document.trackform.agreeconds.checked == false)
  {
    alert("Please note: To track your package(s), you must accept the UPS Tracking Terms and Conditions by selecting the checkbox below.");
    return (false);
  }else{
	document.trackform.submit();
  }
  return (true);
}
//-->
</script>
<%
if theshiptype="ups" then
%>
&nbsp;<br />
      <table class="cobtbl" width="<%=maintablewidth%>" border="0" bordercolor="#B1B1B1" cellspacing="1" cellpadding="3" bgcolor="#B1B1B1">
		<tr>
		  <td class="cobll" bgcolor="#FFFFFF" colspan="2">
			<table border="0" cellspacing="0" cellpadding="0" width="100%" bgcolor="">
			  <tr>
				<td width="40"><img src="images/LOGO_S.gif" alt="UPS" /></td><td align="center">&nbsp;<br /><font size="4"><strong>UPS Tracking Tool</strong></font><br />&nbsp;</td><td width="40">&nbsp;</td>
			  </tr>
			</table>
		  </td>
		</tr>
<%
Function getAddress(t, byRef theAddress)
	signedby = ""
	For l = 0 To t.childNodes.length - 1
		Set u = t.childNodes.Item(l)
		if u.nodeName = "AddressLine1" then
			addressline1 = u.firstChild.nodeValue
		elseif u.nodeName = "AddressLine2" then
			addressline2 = u.firstChild.nodeValue
		elseif u.nodeName = "AddressLine3" then
			addressline3 = u.firstChild.nodeValue
		elseif u.nodeName = "City" then
			city = u.firstChild.nodeValue
		elseif u.nodeName = "StateProvinceCode" then
			statecode = u.firstChild.nodeValue
		elseif u.nodeName = "PostalCode" then
			postcode = u.firstChild.nodeValue
		elseif u.nodeName = "CountryCode" then
			sSQL = "SELECT countryName FROM countries WHERE countryCode='" & u.firstChild.nodeValue & "'"
			rs.Open sSQL,cnn,0,1
			if NOT rs.EOF then
				countrycode = rs("countryName")
			else
				countrycode = u.firstChild.nodeValue
			end if
			rs.Close
		end if
	next
	theAddress = ""
	if addressline1<>"" then theAddress = theAddress & addressline1 & "<br />"
	if addressline2<>"" then theAddress = theAddress & addressline2 & "<br />"
	if addressline3<>"" then theAddress = theAddress & addressline3 & "<br />"
	if city<>"" then theAddress = theAddress & city & "<br />"
	if statecode<>"" AND postcode<>"" then
		theAddress = theAddress & statecode & ", " & postcode & "<br />"
	else
		if statecode<>"" then theAddress = theAddress & statecode & "<br />"
		if postcode<>"" then theAddress = theAddress & postcode & "<br />"
	end if
	if countrycode<>"" then theAddress = theAddress & countrycode & "<br />"
End Function
Function ParseUPSTrackingOutput(sXML, byRef totActivity, byRef shipperNo, byRef serviceDesc, byRef shipperaddress, byRef shiptoaddress, byRef scheddeldate, byRef rescheddeldate, byRef errormsg, byRef activityList)
Dim noError, nodeList, packCost, xmlDoc, e, i, j, k, n, t, t2, index
	noError = True
	totalCost = 0
	packCost = 0
	index = 0
	errormsg = ""
	gotxml=false
	theaddress=""
	on error resume next
	err.number=0
	set xmlDoc = Server.CreateObject("MSXML2.DOMDocument")
	if err.number=0 then gotxml=true
	if NOT gotxml then
		err.number=0
		set xmlDoc = Server.CreateObject("MSXML.DOMDocument")
		if err.number=0 then gotxml=true
	end if
	on error goto 0
	xmlDoc.validateOnParse = False
	xmlDoc.loadXML (sXML)
	Set t2 = xmlDoc.getElementsByTagName("TrackResponse").Item(0)
	for j = 0 to t2.childNodes.length - 1
		Set n = t2.childNodes.Item(j)
		if n.nodename="Response" then
			For i = 0 To n.childNodes.length - 1
				Set e = n.childNodes.Item(i)
				if e.nodeName="ResponseStatusCode" then
					noError = Int(e.firstChild.nodeValue)=1
				end if
				if e.nodeName="Error" then
					errormsg = ""
					For k = 0 To e.childNodes.length - 1
						Set t = e.childNodes.Item(k)
						Select Case t.nodeName
							Case "ErrorSeverity"
								if t.firstChild.nodeValue="Transient" then errormsg = "This is a temporary error. Please wait a few moments then refresh this page.<br />" & errormsg
							Case "ErrorDescription"
								errormsg = errormsg & t.firstChild.nodeValue
						End Select
					Next
				end if
				' response.write "The Nodename is : " & e.nodeName & ":" & e.firstChild.nodeValue & "<br />"
			Next
		elseif n.nodename="Shipment" then
			For i = 0 To n.childNodes.length - 1
				Set e = n.childNodes.Item(i)
				' response.write "Nodename is : " & e.nodeName & "<br />"
				Select Case e.nodeName
					Case "Shipper"
						For k = 0 To e.childNodes.length - 1
							Set t = e.childNodes.Item(k)
							if t.nodeName = "ShipperNumber" then
								shipperNo = t.firstChild.nodeValue
							elseif t.nodeName = "Address" then
								call getAddress(t, shipperaddress)
							end if
						Next
					Case "ShipTo"
						For k = 0 To e.childNodes.length - 1
							Set t = e.childNodes.Item(k)
							if t.nodeName = "Address" then
								call getAddress(t, shiptoaddress)
							end if
						Next
					Case "ScheduledDeliveryDate"
						scheddeldate = e.firstChild.nodeValue
					Case "Service"
						For k = 0 To e.childNodes.length - 1
							Set t = e.childNodes.Item(k)
							if t.nodeName = "X_Code_X" then
								Select Case Int(t.firstChild.nodeValue)
									Case 1
										serviceDesc = "Next Day Air"
									Case 2
										serviceDesc = "2nd Day Air"
									Case 3
										serviceDesc = "Ground Service"
									Case 7
										serviceDesc = "Worldwide Express"
									Case 8
										serviceDesc = "Worldwide Expedited"
									Case 11
										serviceDesc = "Standard service"
									Case 12
										serviceDesc = "3-Day Select"
									Case 13
										serviceDesc = "Next Day Air Saver"
									Case 14
										serviceDesc = "Next Day Air Early AM"
									Case 54
										serviceDesc = "Worldwide Express Plus"
									Case 59
										serviceDesc = "2nd Day Air AM"
									Case 64
										serviceDesc = "UPS Express NA1"
									Case 65
										serviceDesc = "Express Saver"
								End Select
								' response.write "The service code is : " & t.nodeName & ":" & t.firstChild.nodeValue & "<br />"
							elseif t.nodeName = "Description" then
								serviceDesc = t.firstChild.nodeValue
							end if
						Next
					Case "Package"
						For k = 0 To e.childNodes.length - 1
							Set t = e.childNodes.Item(k)
							if t.nodeName = "RescheduledDeliveryDate" then
								rescheddeldate = t.firstChild.nodeValue
							elseif t.nodeName = "Activity" then
								For l = 0 To t.childNodes.length - 1
									Set u = t.childNodes.Item(l)
									if u.nodeName = "ActivityLocation" then
										For m = 0 To u.childNodes.length - 1
											Set v = u.childNodes.Item(m)
											if v.nodeName = "Address" then
												call getAddress(v, activityList(totActivity,0))
											elseif v.nodeName = "Description" then
												description = v.firstChild.nodeValue
											elseif v.nodeName = "SignedForByName" then
												activityList(totActivity,1) = v.firstChild.nodeValue
											end if
										Next
									elseif u.nodeName = "Status" then
										For m = 0 To u.childNodes.length - 1
											Set v = u.childNodes.Item(m)
											if v.nodeName = "StatusType" then
												For nn = 0 To v.childNodes.length - 1
													Set w = v.childNodes.Item(nn)
													if w.nodeName="Code" then
														activityList(totActivity,3)=w.firstChild.nodeValue
													elseif w.nodeName="Description" then
														activityList(totActivity,4)=w.firstChild.nodeValue
													end if
												next
											elseif v.nodeName = "StatusCode" then
												For nn = 0 To v.childNodes.length - 1
													Set w = v.childNodes.Item(nn)
													if w.nodeName="Code" then
														activityList(totActivity,5)=w.firstChild.nodeValue
													end if
												next
											end if
										Next
									else
										if u.nodeName="Date" then
											activityList(totActivity,6)=u.firstChild.nodeValue
										elseif u.nodeName="Time" then
											activityList(totActivity,7)=u.firstChild.nodeValue
										end if
									end if
								Next
								totActivity = totActivity + 1
								' response.write "<HR>"
							end if
						Next
				End select
			Next
		end if
	Next
	ParseUPSTrackingOutput = noError
end Function
function UPSTrack(trackNo)
	Dim objHttp, i, activityList(100,10),success,lastloc
	lastloc="xxxxxx"
	' ActivityList(0) = Address
	' ActivityList(1) = SignedForByName
	' ActivityList(2) = Not Used
	' ActivityList(3) = Activity -> Status -> StatusType -> Code
	' ActivityList(4) = Activity -> Status -> StatusType -> Description
	' ActivityList(5) = Activity -> Status -> StatusCode -> Code
	' ActivityList(6) = Activity -> Date
	' ActivityList(7) = Activity -> Time

	sXML = "<?xml version=""1.0""?><AccessRequest xml:lang=""en-US""><AccessLicenseNumber>"&upsAccess&"</AccessLicenseNumber><UserId>"&upsUser&"</UserId><Password>"&upsPw&"</Password></AccessRequest>"
	sXML = sXML & "<?xml version=""1.0""?><TrackRequest xml:lang=""en-US""><Request><TransactionReference><CustomerContext>Example 3</CustomerContext><XpciVersion>1.0001</XpciVersion></TransactionReference><RequestAction>Track</RequestAction><RequestOption>"
	if Trim(request.form("activity"))="LAST" then sXML = sXML & "none" else sXML = sXML & "activity"
	sXML = sXML & "</RequestOption></Request>"
	if false then
		sXML = sXML & "<ReferenceNumber><Value>"&trackNo&"</Value></ReferenceNumber>"
		sXML = sXML & "<ShipperNumber>116593</ShipperNumber></TrackRequest>"
	else
		sXML = sXML & "<TrackingNumber>"&trackNo&"</TrackingNumber></TrackRequest>"
	end if
	set objHttp = Server.CreateObject("Msxml2.ServerXMLHTTP")
	objHttp.open "POST", "https://www.ups.com/ups.app/xml/Track", false
	objHttp.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"

	' response.write Replace(Replace(sXML,"</","&lt;/"),"<","<br />&lt;")&"<HR>"
	on error resume next
	err.number=0
	objHttp.Send sXML
	on error goto 0
	If err.number <> 0 OR objHttp.status <> 200 Then
		errormsg = "Error, couldn't connect to UPS server"
		success = false
	Else
		saveLCID = Session.LCID
		Session.LCID = 1033
		totActivity = 0
		' response.write Replace(Replace(objHttp.responseText,"</","&lt;/"),"<","<br />&lt;")&"<HR>"
		success = ParseUPSTrackingOutput(objHttp.responseText, totActivity, shipperNo, serviceDesc, shipperaddress, shiptoaddress, scheduleddeliverydate, rescheddeliverydate, errormsg, activityList)
		Session.LCID = saveLCID
		if success then
			for index2=0 to totActivity-2
				for index=0 to totActivity-2
					if Int(activityList(index,6)&activityList(index,7))>Int(activityList(index+1,6)&activityList(index+1,7)) then
						for index3=0 to UBOUND(activityList,2)
							tempArr = activityList(index,index3)
							activityList(index,index3)=activityList(index+1,index3)
							activityList(index+1,index3)=tempArr
						next
					end if
				next
			next
			if Trim(shipperNo)<>"" then %>
		  <tr>
			<td class="cobhl" bgcolor="#EBEBEB" width="30%"><strong>Shipper Number</strong> </td>
			<td class="cobll" bgcolor="#FFFFFF"><%=shipperNo%></td>
		  </tr>
		<%	end if
			if Trim(serviceDesc)<>"" then %>
		  <tr>
			<td class="cobhl" bgcolor="#EBEBEB" width="30%"><strong>Service Description</strong> </td>
			<td class="cobll" bgcolor="#FFFFFF"><%=serviceDesc%></td>
		  </tr>
		<%	end if
			if Trim(shipperaddress)<>"" then %>
		  <tr>
			<td class="cobhl" bgcolor="#EBEBEB" width="30%" valign="top"><strong>Shipper Address</strong> </td>
			<td class="cobll" bgcolor="#FFFFFF"><%=shipperaddress%></td>
		  </tr>
		<%	end if
			if Trim(shiptoaddress)<>"" then %>
		  <tr>
			<td class="cobhl" bgcolor="#EBEBEB" width="30%" valign="top"><strong>Ship-To Address</strong> </td>
			<td class="cobll" bgcolor="#FFFFFF"><%=shiptoaddress%></td>
		  </tr>
		<%	end if
			if Trim(scheduleddeliverydate)<>"" then %>
		  <tr>
			<td class="cobhl" bgcolor="#EBEBEB" width="30%" valign="top"><strong>Sched. Delivery Date</strong> </td>
			<td class="cobll" bgcolor="#FFFFFF"><%=DateSerial(Left(scheduleddeliverydate,4),Mid(scheduleddeliverydate,5,2),Mid(scheduleddeliverydate,7,2)) %></td>
		  </tr>
		<%	end if
			if Trim(rescheddeliverydate)<>"" then %>
		  <tr>
			<td class="cobhl" bgcolor="#EBEBEB" width="30%" valign="top"><strong>ReSched. Delivery Date</strong> </td>
			<td class="cobll" bgcolor="#FFFFFF"><%=DateSerial(Left(rescheddeliverydate,4),Mid(rescheddeliverydate,5,2),Mid(rescheddeliverydate,7,2)) %></td>
		  </tr>
		  <tr>
			<td class="cobhl" bgcolor="#EBEBEB" width="30%" valign="top"><strong>Note</strong> </td>
			<td class="cobll" bgcolor="#FFFFFF">Your package is in the UPS system and has a rescheduled delivery date of <%=DateSerial(Left(rescheddeliverydate,4),Mid(rescheddeliverydate,5,2),Mid(rescheddeliverydate,7,2)) %></td>
		  </tr>
		<%	end if %>
		</table>
  &nbsp;
		<table class="cobtbl" width="<%=maintablewidth%>" border="0" bordercolor="#B1B1B1" cellspacing="1" cellpadding="3" bgcolor="#B1B1B1">
		  <tr>
			<td class="cobhl" bgcolor="#EBEBEB"><strong>Location</strong></td>
			<td class="cobhl" bgcolor="#EBEBEB"><strong>Description</strong></td>
			<td class="cobhl" bgcolor="#EBEBEB"><strong>Date&nbsp;/&nbsp;Time</strong></td>
		  </tr>
<% for index=0 to totActivity-1 
		if index MOD 2 = 0 then
			cellbg="class=""cobll"" bgcolor=""#FFFFFF"""
		else
			cellbg="class=""cobhl"" bgcolor=""#EBEBEB"""
		end if
%>
			  <tr>
			    <td <%=cellbg%>><font size="1"><%	if lastloc=activityList(index,0) then
										response.write "<p align='center'>""</p>"
									else
										response.write activityList(index,0)
										lastloc = activityList(index,0)
									end if %></font></td>
				<td <%=cellbg%>><font size="1"><%	response.write activityList(index,4)
									if activityList(index,1)<>"" then response.write "<br /><strong>Signed By :</strong> " & activityList(index,1) %></font></td>
				<td <%=cellbg%>><font size="1"><%=DateSerial(Left(activityList(index,6),4),Mid(activityList(index,6),5,2),Mid(activityList(index,6),7,2))%></font><br />
				<font size="1"><%=TimeSerial(Left(activityList(index,7),2),Mid(activityList(index,7),3,2),Mid(activityList(index,7),5,2))%></font></td>
			  </tr>
<% next %>
			</table>
	  <hr width="70%" />
	  <table class="cobtbl" width="<%=maintablewidth%>" border="0" bordercolor="#B1B1B1" cellspacing="1" cellpadding="3" bgcolor="#B1B1B1">
<%
		else
%>
			  <tr>
			    <td class="cobll" bgcolor="#FFFFFF" colspan="2" height="30" align="center"><strong>The tracking system returned the following error : <%=errormsg%></strong></td>
			  </tr>
			</table>
	  <hr width="70%" />
	  <table class="cobtbl" width="<%=maintablewidth%>" border="0" bordercolor="#B1B1B1" cellspacing="1" cellpadding="3" bgcolor="#B1B1B1">
<%
		end if
	End If
	UPSTrack = success
	set objHttp = nothing
end function
if Trim(Request.Form("trackno"))<>"" then
	UPSTrack(Trim(Request.Form("trackno")))
end if
%>
		<form method="post" name="trackform" action="tracking.asp">
		  <input type="hidden" name="carrier" value="ups">
		  <tr>
			<td class="cobhl" width="50%" bgcolor="#EBEBEB" align="right">Please enter your UPS Tracking Number : </td>
			<td class="cobll" width="50%" bgcolor="#FFFFFF"><input type="text" size="30" name="trackno" value="<%=Trim(Request("trackno"))%>" /></td>
		  </tr>
		  <tr>
			<td class="cobhl" width="50%" bgcolor="#EBEBEB" align="right">Show Activity : </td>
			<td class="cobll" width="50%" bgcolor="#FFFFFF"><select name="activity" size="1"><option value="LAST">Show Last Activity Only</option><option value="ALL"<% if Trim(Request.Form("activity")="ALL") then response.write " selected"%>>Show All Activity</option></select></td>
		  </tr>
		  <tr>
			<td class="cobll" bgcolor="#FFFFFF" colspan="2"><table width="100%" cellspacing="0" cellpadding="0" border="0">
				<tr>
				  <td class="cobll" bgcolor="#FFFFFF" width="17%" height="26">&nbsp;</td>
				  <td class="cobll" bgcolor="#FFFFFF" width="66%" align="center"><input type="button" onclick="viewlicense()" value="View License" /> <input type="button" value="Track Package" onclick="checkaccept()" /></td>
				  <td class="cobll" bgcolor="#FFFFFF" width="17%" height="26" align="right" valign="bottom"><img src="images/tablebr.gif" alt="" /></td>
				</tr>
			  </table></td>
		  </tr>
		  <tr>
			<td class="cobll" width="100%" bgcolor="#FFFFFF" height="30" colspan="2" align="center" valign="middle"><font size="1"><input type="checkbox" name="agreeconds" value="ON" <%if request.form("agreeconds")="ON" then response.write "checked"%> /> By selecting this box and the "Track Package" button, I agree to these <a href="javascript:viewlicense();"><strong>Terms and Conditions</strong></a>.</font></td>
		  </tr>
		</form>
	  </table>
	  <table border="0" cellspacing="0" cellpadding="0" width="100%" bgcolor="#FFFFFF" align="center">
        <tr>
          <td width="100%" align="center"><p>&nbsp;<br /><font size="1">UPS&reg;, UPS & Shield Design&reg; and UNITED PARCEL SERVICE&reg; 
				  are<br />registered trademarks of United Parcel Service of America, Inc.</font></p></td>
		</tr>
	  </table>
	  <br />
<%
elseif theshiptype="usps" then
%>
&nbsp;<br />
      <table class="cobtbl" width="<%=maintablewidth%>" border="0" bordercolor="#B1B1B1" cellspacing="1" cellpadding="3" bgcolor="#B1B1B1">
		<tr>
		  <td class="cobll" bgcolor="#FFFFFF" colspan="2">
			<table border="0" cellspacing="0" cellpadding="0" width="100%" bgcolor="">
			  <tr>
				<td width="40">&nbsp;</td><td align="center">&nbsp;<br /><font size="4"><strong>USPS Tracking Tool</strong></font><br />&nbsp;</td><td width="40">&nbsp;</td>
			  </tr>
			</table>
		  </td>
		</tr>
<%
Function ParseUSPSTrackingOutput(sXML, byRef totActivity, onlylast, byRef serviceDesc, byRef shipperaddress, byRef shiptoaddress, byRef scheddeldate, byRef rescheddeldate, byRef errormsg, byRef activityList)
Dim noError, nodeList, packCost, xmlDoc, e, i, j, k, n, t, t2, index
	noError = True
	totalCost = 0
	packCost = 0
	index = 0
	errormsg = ""
	gotxml=false
	theaddress=""
	on error resume next
	err.number=0
	set xmlDoc = Server.CreateObject("MSXML2.DOMDocument")
	if err.number=0 then gotxml=true
	if NOT gotxml then
		err.number=0
		set xmlDoc = Server.CreateObject("MSXML.DOMDocument")
		if err.number=0 then gotxml=true
	end if
	on error goto 0
	xmlDoc.validateOnParse = False
	xmlDoc.loadXML (sXML)
	If xmlDoc.documentElement.nodeName = "Error" then 'Top-level Error
		Set t2 = xmlDoc.getElementsByTagName("Error").Item(0)
		noError = FALSE
		for j = 0 to t2.childNodes.length - 1
			Set n = t2.childNodes.Item(j)
			if n.nodename="Description" then
				errormsg = n.firstChild.nodeValue
			end if
		next
	else
		Set t2 = xmlDoc.getElementsByTagName("TrackInfo").Item(0)
		for j = 0 to t2.childNodes.length - 1
			companyname= ""
			city=""
			statecode=""
			postcode=""
			countrycode=""
			Set n = t2.childNodes.Item(j)
			if n.nodename="Error" then
				For i = 0 To n.childNodes.length - 1
					Set e = n.childNodes.Item(i)
					if e.nodeName="Description" then
						errormsg = e.firstChild.nodeValue
						noError = FALSE
					end if
					' response.write "The Nodename is : " & e.nodeName & ":" & e.firstChild.nodeValue & "<br />"
				Next
			elseif n.nodename="TrackDetail" then
				if NOT onlylast then
					For i = 0 To n.childNodes.length - 1
						Set e = n.childNodes.Item(i)
						' response.write "Nodename is : " & e.nodeName & "<br />"
						Select Case e.nodeName
							Case "EventDate"
								activityList(totActivity,6)=e.firstChild.nodeValue
							Case "EventTime"
								activityList(totActivity,7)=e.firstChild.nodeValue
							Case "Event"
								activityList(totActivity,4)=e.firstChild.nodeValue
							Case "EventCity"
								if e.hasChildNodes then city = e.firstChild.nodeValue
							Case "EventState"
								if e.hasChildNodes then statecode = e.firstChild.nodeValue
							Case "EventZIPCode"
								if e.hasChildNodes then postcode = e.firstChild.nodeValue
							Case "EventCountry"
								if e.hasChildNodes then countrycode = e.firstChild.nodeValue
							Case "FirmName"
								if e.hasChildNodes then companyname = e.firstChild.nodeValue
						End select
					Next
					theAddress = ""
					if companyname<>"" then theAddress = theAddress & companyname & "<br />"
					if city<>"" then theAddress = theAddress & city & "<br />"
					if statecode<>"" AND postcode<>"" then
						theAddress = theAddress & statecode & ", " & postcode & "<br />"
					else
						if statecode<>"" then theAddress = theAddress & statecode & "<br />"
						if postcode<>"" then theAddress = theAddress & postcode & "<br />"
					end if
					if countrycode<>"" then theAddress = theAddress & countrycode & "<br />"
					activityList(totActivity,0) = theAddress
					totActivity = totActivity + 1
				end if
			elseif n.nodename="TrackSummary" then
				For i = 0 To n.childNodes.length - 1
					Set e = n.childNodes.Item(i)
					' response.write "Nodename is : " & e.nodeName & "<br />"
					Select Case e.nodeName
						Case "EventDate"
							scheddeldate=e.firstChild.nodeValue&scheddeldate
						Case "EventTime"
							scheddeldate=scheddeldate&" "&e.firstChild.nodeValue
						Case "Event"
							serviceDesc=e.firstChild.nodeValue
						Case "EventCity"
							if e.hasChildNodes then city = e.firstChild.nodeValue
						Case "EventState"
							if e.hasChildNodes then statecode = e.firstChild.nodeValue
						Case "EventZIPCode"
							if e.hasChildNodes then postcode = e.firstChild.nodeValue
						Case "EventCountry"
							if e.hasChildNodes then countrycode = e.firstChild.nodeValue
						Case "FirmName"
							if e.hasChildNodes then companyname = e.firstChild.nodeValue
					End select
				Next
				theAddress = ""
				if companyname<>"" then theAddress = theAddress & companyname & "<br />"
				if city<>"" then theAddress = theAddress & city & "<br />"
				if statecode<>"" AND postcode<>"" then
					theAddress = theAddress & statecode & ", " & postcode & "<br />"
				else
					if statecode<>"" then theAddress = theAddress & statecode & "<br />"
					if postcode<>"" then theAddress = theAddress & postcode & "<br />"
				end if
				if countrycode<>"" then theAddress = theAddress & countrycode & "<br />"
				shiptoaddress = theAddress
			end if
		Next
	end if
	ParseUSPSTrackingOutput = noError
end Function
function USPSTrack(trackNo)
	Dim objHttp, i, activityList(100,10),success,lastloc
	lastloc="xxxxxx"
	' ActivityList(0) = Address
	' ActivityList(1) = SignedForByName
	' ActivityList(2) = Not Used
	' ActivityList(3) = Activity -> Status -> StatusType -> Code
	' ActivityList(4) = Activity -> Status -> StatusType -> Description
	' ActivityList(5) = Activity -> Status -> StatusCode -> Code
	' ActivityList(6) = Activity -> Date
	' ActivityList(7) = Activity -> Time

	sXML = "<TrackFieldRequest USERID="""&uspsUser&"""><TrackID ID="""&trim(Request.Form("trackno"))&"""></TrackID></TrackFieldRequest>"
	set objHttp = Server.CreateObject("Msxml2.ServerXMLHTTP")
	objHttp.open "POST", "http://production.shippingapis.com/ShippingAPI.dll", false
	' objHttp.open "POST", "http://testing.shippingapis.com/ShippingAPITest.dll", false
	on error resume next
	err.number=0
	' response.write Replace(Replace(sXML,"</","&lt;/"),"<","<br />&lt;")&"<HR>"
	objHttp.Send "API=TrackV2&XML=" & Server.URLEncode(sXML)
	on error goto 0
	If err.number <> 0 OR objHttp.status <> 200 Then
		errormsg = "Error, couldn't connect to USPS server"
		success = false
	Else
		saveLCID = Session.LCID
		Session.LCID = 1033
		totActivity = 0
		' response.write Replace(Replace(objHttp.responseText,"</","&lt;/"),"<","<br />&lt;")&"<HR>"
		success = ParseUSPSTrackingOutput(objHttp.responseText, totActivity, trim(request.form("activity"))="LAST", serviceDesc, shipperaddress, shiptoaddress, scheduleddeliverydate, rescheddeliverydate, errormsg, activityList)
		Session.LCID = saveLCID
		if success then
			for index2=0 to totActivity-2
				for index=0 to totActivity-2
					if DateValue(activityList(index,6)&" "&activityList(index,7))>DateValue(activityList(index+1,6)&" "&activityList(index+1,7)) then
						for index3=0 to UBOUND(activityList,2)
							tempArr = activityList(index,index3)
							activityList(index,index3)=activityList(index+1,index3)
							activityList(index+1,index3)=tempArr
						next
					end if
				next
			next
			if Trim(serviceDesc)<>"" then %>
		  <tr>
			<td class="cobhl" bgcolor="#EBEBEB" width="30%"><strong>Event</strong> </td>
			<td class="cobll" bgcolor="#FFFFFF"><%=serviceDesc%></td>
		  </tr>
		<%	end if
			if Trim(shiptoaddress)<>"" then %>
		  <tr>
			<td class="cobhl" bgcolor="#EBEBEB" width="30%" valign="top"><strong>Address</strong> </td>
			<td class="cobll" bgcolor="#FFFFFF"><%=shiptoaddress%></td>
		  </tr>
		<%	end if
			if Trim(scheduleddeliverydate)<>"" then %>
		  <tr>
			<td class="cobhl" bgcolor="#EBEBEB" width="30%" valign="top"><strong>Event Date</strong> </td>
			<td class="cobll" bgcolor="#FFFFFF"><%=scheduleddeliverydate %></td>
		  </tr>
		<%	end if %>
		</table>
<%		if totActivity>0 then %>
  &nbsp;
		<table class="cobtbl" width="<%=maintablewidth%>" border="0" bordercolor="#B1B1B1" cellspacing="1" cellpadding="3" bgcolor="#B1B1B1">
		  <tr>
			<td class="cobhl" bgcolor="#EBEBEB"><strong>Location</strong></td>
			<td class="cobhl" bgcolor="#EBEBEB"><strong>Description</strong></td>
			<td class="cobhl" bgcolor="#EBEBEB"><strong>Date&nbsp;/&nbsp;Time</strong></td>
		  </tr>
<%			for index=0 to totActivity-1 
				if index MOD 2 = 0 then
					cellbg="class=""cobll"" bgcolor=""#FFFFFF"""
				else
					cellbg="class=""cobhl"" bgcolor=""#EBEBEB"""
				end if %>
			  <tr>
			    <td <%=cellbg%>><font size="1"><%
									if lastloc=activityList(index,0) then
										response.write "<p align='center'>""</p>"
									else
										response.write activityList(index,0)
										lastloc = activityList(index,0)
									end if %></font></td>
				<td <%=cellbg%>><font size="1"><%	response.write activityList(index,4)
									if activityList(index,1)<>"" then response.write "<br /><strong>Signed By :</strong> " & activityList(index,1) %></font></td>
				<td <%=cellbg%>><font size="1"><%=activityList(index,6)%></font><br />
				<font size="1"><%=activityList(index,7)%></font></td>
			  </tr>
<%			next %>
			</table>
<%		end if %>
	  <hr width="70%" />
	  <table class="cobtbl" width="<%=maintablewidth%>" border="0" bordercolor="#B1B1B1" cellspacing="1" cellpadding="3" bgcolor="#B1B1B1">
<%
		else
%>
			  <tr>
			    <td class="cobll" bgcolor="#FFFFFF" colspan="2" height="30" align="center"><strong>The tracking system returned the following error : <%=errormsg%></strong></td>
			  </tr>
			</table>
	  <hr width="70%" />
	  <table class="cobtbl" width="<%=maintablewidth%>" border="0" bordercolor="#B1B1B1" cellspacing="1" cellpadding="3" bgcolor="#B1B1B1">
<%
		end if
	End If
	USPSTrack = success
	set objHttp = nothing
end function
if Trim(Request.Form("trackno"))<>"" then
	USPSTrack(Trim(Request.Form("trackno")))
end if
%>
		<form method="post" name="trackform" action="tracking.asp">
		  <input type="hidden" name="carrier" value="usps">
		  <tr>
			<td class="cobhl" width="50%" bgcolor="#EBEBEB" align="right">Please enter your USPS Tracking Number : </td>
			<td class="cobll" width="50%" bgcolor="#FFFFFF"><input type="text" size="30" name="trackno" value="<%=Trim(Request("trackno"))%>" /></td>
		  </tr>
		  <tr>
			<td class="cobhl" width="50%" bgcolor="#EBEBEB" align="right">Show Activity : </td>
			<td class="cobll" width="50%" bgcolor="#FFFFFF"><select name="activity" size="1"><option value="LAST">Show Last Activity Only</option><option value="ALL"<% if Trim(Request.Form("activity")="ALL") OR Trim(Request.Form("activity")="") then response.write " selected"%>>Show All Activity</option></select></td>
		  </tr>
		  <tr>
			<td class="cobll" bgcolor="#FFFFFF" colspan="2"><table width="100%" cellspacing="0" cellpadding="0" border="0">
				<tr>
				  <td class="cobll" bgcolor="#FFFFFF" width="17%" height="26">&nbsp;</td>
				  <td class="cobll" bgcolor="#FFFFFF" width="66%" align="center"> <input type="submit" value="Track Package" /></td>
				  <td class="cobll" bgcolor="#FFFFFF" width="17%" height="26" align="right" valign="bottom"><img src="images/tablebr.gif" alt="" /></td>
				</tr>
			  </table></td>
		  </tr>
		</form>
	  </table>
	  <table border="0" cellspacing="0" cellpadding="0" width="100%" bgcolor="#FFFFFF" align="center">
        <tr>
          <td width="100%" align="center"><p>&nbsp;</p></td>
		</tr>
	  </table>
	  <br />
<%
elseif theshiptype="fedex" then
%>
&nbsp;<br />
      <table class="cobtbl" width="<%=maintablewidth%>" border="0" bordercolor="#B1B1B1" cellspacing="1" cellpadding="3" bgcolor="#B1B1B1">
		<tr>
		  <td class="cobll" bgcolor="#FFFFFF" colspan="2">
			<table border="0" cellspacing="0" cellpadding="0" width="100%" bgcolor="">
			  <tr>
				<td width="40"><img src="images/fedexsmall.gif" alt="FedEx" /></td><td align="center">&nbsp;<br /><font size="4"><strong>FedEx<small>&reg;</small> Tracking Tool</strong></font><br />&nbsp;</td><td width="40">&nbsp;</td>
			  </tr>
			</table>
		  </td>
		</tr>
<%
Function getAddress(t, byRef theAddress)
	signedby = ""
	For l = 0 To t.childNodes.length - 1
		Set u = t.childNodes.Item(l)
		if u.nodeName = "AddressLine1" then
			addressline1 = u.firstChild.nodeValue
		elseif u.nodeName = "AddressLine2" then
			addressline2 = u.firstChild.nodeValue
		elseif u.nodeName = "AddressLine3" then
			addressline3 = u.firstChild.nodeValue
		elseif u.nodeName = "City" then
			city = u.firstChild.nodeValue
		elseif u.nodeName = "StateOrProvinceCode" then
			statecode = u.firstChild.nodeValue
		elseif u.nodeName = "PostalCode" then
			postcode = u.firstChild.nodeValue
		elseif u.nodeName = "CountryCode" then
			sSQL = "SELECT countryName FROM countries WHERE countryCode='" & u.firstChild.nodeValue & "'"
			rs.Open sSQL,cnn,0,1
			if NOT rs.EOF then
				countrycode = rs("countryName")
			else
				countrycode = u.firstChild.nodeValue
			end if
			rs.Close
		end if
	next
	theAddress = ""
	if addressline1<>"" then theAddress = theAddress & addressline1 & "<br />"
	if addressline2<>"" then theAddress = theAddress & addressline2 & "<br />"
	if addressline3<>"" then theAddress = theAddress & addressline3 & "<br />"
	if city<>"" then theAddress = theAddress & city & "<br />"
	if statecode<>"" AND postcode<>"" then
		theAddress = theAddress & statecode & ", " & postcode & "<br />"
	else
		if statecode<>"" then theAddress = theAddress & statecode & "<br />"
		if postcode<>"" then theAddress = theAddress & postcode & "<br />"
	end if
	if countrycode<>"" then theAddress = theAddress & countrycode & "<br />"
End Function
Function ParseFedexTrackingOutput(sXML, byRef totActivity, byRef deliverydate, byRef serviceDesc, byRef packagecount, byRef shiptoaddress, byRef scheddeldate, byRef signedforby, byRef errormsg, byRef activityList)
Dim noError, nodeList, packCost, xmlDoc, e, i, j, k, n, t, t2, index
	noError = True
	totalCost = 0
	packCost = 0
	index = 0
	errormsg = ""
	gotxml=false
	theaddress=""
	on error resume next
	err.number=0
	set xmlDoc = Server.CreateObject("MSXML2.DOMDocument")
	if err.number=0 then gotxml=true
	if NOT gotxml then
		err.number=0
		set xmlDoc = Server.CreateObject("MSXML.DOMDocument")
		if err.number=0 then gotxml=true
	end if
	on error goto 0
	xmlDoc.validateOnParse = False
	xmlDoc.loadXML (sXML)
	Set t2 = xmlDoc.getElementsByTagName("FDXTrackReply").Item(0)
	for j = 0 to t2.childNodes.length - 1
		Set n = t2.childNodes.Item(j)
		if n.nodename="TrackProfile" then
			For i = 0 To n.childNodes.length - 1
				Set e = n.childNodes.Item(i)
				Select Case e.nodeName
					Case "SoftError"
						For k = 0 To e.childNodes.length - 1
							Set t = e.childNodes.Item(k)
							if t.nodeName = "Message" then
								noError = FALSE
								errormsg = t.firstChild.nodeValue
							end if
						Next
					Case "SignedForBy"
						signedforby = e.firstChild.nodeValue
					Case "DestinationAddress"
						call getAddress(e, shiptoaddress)
					Case "DeliveredDate"
						deliverydate = e.firstChild.nodeValue & deliverydate
					Case "DeliveredTime"
						deliverydate = deliverydate & " " & e.firstChild.nodeValue
					Case "Service"
						serviceDesc = e.firstChild.nodeValue
					Case "PackageCount"
						packagecount = e.firstChild.nodeValue
					Case "Scan"
						call getAddress(e, activityList(totActivity,0))
						For k = 0 To e.childNodes.length - 1
							Set t = e.childNodes.Item(k)
							if t.nodeName = "Date" then
								activityList(totActivity,6)=t.firstChild.nodeValue
							elseif t.nodeName = "Time" then
								activityList(totActivity,7)=t.firstChild.nodeValue
							elseif t.nodeName = "StatusExceptionCode" then
								activityList(totActivity,3)=t.firstChild.nodeValue
							elseif t.nodeName = "ScanDescription" OR t.nodeName = "StatusExceptionDescription" then
								if t.firstChild.nodeValue <> "Package status" then activityList(totActivity,4)=t.firstChild.nodeValue
							end if
						Next
						if activityList(totActivity,4)<>"" then totActivity = totActivity + 1
				End select
			Next
		end if
	Next
	ParseFedexTrackingOutput = noError
end Function
function FedexTrack(trackNo)
	Dim objHttp, i, activityList(100,10),success,lastloc
	lastloc="xxxxxx"
	' ActivityList(0) = Address
	' ActivityList(1) = SignedForByName
	' ActivityList(2) = Not Used
	' ActivityList(3) = Activity -> Status -> StatusType -> Code
	' ActivityList(4) = Activity -> Status -> StatusType -> Description
	' ActivityList(5) = Activity -> Status -> StatusCode -> Code
	' ActivityList(6) = Activity -> Date
	' ActivityList(7) = Activity -> Time
	sXML ="<?xml version=""1.0"" encoding=""UTF-8"" ?>"
	sXML = sXML & "<FDXTrackRequest xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xsi:noNamespaceSchemaLocation=""FDXTrackRequest.xsd"">"
	sXML = sXML & "<RequestHeader>"
	sXML = sXML & "<CustomerTransactionIdentifier>transidentifier</CustomerTransactionIdentifier>"
	sXML = sXML & "<AccountNumber>"&fedexaccount&"</AccountNumber>"
	sXML = sXML & "<MeterNumber>"&fedexmeter&"</MeterNumber>"
	sXML = sXML & "<CarrierCode></CarrierCode>"
	sXML = sXML & "</RequestHeader>"
	sXML = sXML & "<PackageIdentifier>"
	sXML = sXML & "<Value>"&trackNo&"</Value>"
	sXML = sXML & "<Type>TRACKING_NUMBER_OR_DOORTAG</Type>"
	sXML = sXML & "</PackageIdentifier>"
	if Trim(request.form("activity"))="LAST" then sXML = sXML & "<DetailScans>0</DetailScans>" else sXML = sXML & "<DetailScans>1</DetailScans>"
	sXML = sXML & "</FDXTrackRequest>"
	success = callxmlfunction("https://gateway.fedex.com:443/GatewayDC", sXML, xmlres, "", "Msxml2.ServerXMLHTTP", errormsg, FALSE)
	if success then
		Session.LCID = 1033
		totActivity = 0
		success = ParseFedexTrackingOutput(xmlres, totActivity, deliverydate, serviceDesc, packagecount, shiptoaddress, scheduleddeliverydate, signedforby, errormsg, activityList)
		Session.LCID = saveLCID
		if success then
			for index2=0 to totActivity-2
				for index=0 to totActivity-2
					if (activityList(index,6)&activityList(index,7))>(activityList(index+1,6)&activityList(index+1,7)) then
						for index3=0 to UBOUND(activityList,2)
							tempArr = activityList(index,index3)
							activityList(index,index3)=activityList(index+1,index3)
							activityList(index+1,index3)=tempArr
						next
					end if
				next
			next
			if Trim(serviceDesc)<>"" then %>
		  <tr>
			<td class="cobhl" bgcolor="#EBEBEB" width="30%"><strong>Service Description</strong> </td>
			<td class="cobll" bgcolor="#FFFFFF"><%=serviceDesc%></td>
		  </tr>
		<%	end if
			if Trim(packagecount)<>"" then %>
		  <tr>
			<td class="cobhl" bgcolor="#EBEBEB" width="30%" valign="top"><strong>Package Count</strong> </td>
			<td class="cobll" bgcolor="#FFFFFF"><%=packagecount%></td>
		  </tr>
		<%	end if
			if Trim(shiptoaddress)<>"" then %>
		  <tr>
			<td class="cobhl" bgcolor="#EBEBEB" width="30%" valign="top"><strong>Ship-To Address</strong> </td>
			<td class="cobll" bgcolor="#FFFFFF"><%=shiptoaddress%></td>
		  </tr>
		<%	end if
			if Trim(signedforby)<>"" then %>
		  <tr>
			<td class="cobhl" bgcolor="#EBEBEB" width="30%" valign="top"><strong>Signed For By</strong> </td>
			<td class="cobll" bgcolor="#FFFFFF"><%=signedforby %></td>
		  </tr>
		<%	end if
			if Trim(deliverydate)<>"" then %>
		  <tr>
			<td class="cobhl" bgcolor="#EBEBEB" width="30%"><strong>Delivery Date</strong> </td>
			<td class="cobll" bgcolor="#FFFFFF"><%=deliverydate%></td>
		  </tr>
		<%	end if %>
		</table>
  &nbsp;
		<table class="cobtbl" width="<%=maintablewidth%>" border="0" bordercolor="#B1B1B1" cellspacing="1" cellpadding="3" bgcolor="#B1B1B1">
		  <tr>
			<td class="cobhl" bgcolor="#EBEBEB"><strong>Location</strong></td>
			<td class="cobhl" bgcolor="#EBEBEB"><strong>Description</strong></td>
			<td class="cobhl" bgcolor="#EBEBEB"><strong>Date&nbsp;/&nbsp;Time</strong></td>
		  </tr>
<% for index=0 to totActivity-1 
		if index MOD 2 = 0 then
			cellbg="class=""cobll"" bgcolor=""#FFFFFF"""
		else
			cellbg="class=""cobhl"" bgcolor=""#EBEBEB"""
		end if
%>
			  <tr>
			    <td <%=cellbg%>><font size="1"><%	if lastloc=activityList(index,0) then
										response.write "<p align='center'>""</p>"
									else
										response.write activityList(index,0)
										lastloc = activityList(index,0)
									end if %></font></td>
				<td <%=cellbg%>><font size="1"><%	response.write activityList(index,4)
									if activityList(index,1)<>"" then response.write "<br /><strong>Signed By :</strong> " & activityList(index,1) %></font></td>
				<td <%=cellbg%>><font size="1"><%=DateValue(activityList(index,6))%></font><br />
				<font size="1"><%=activityList(index,7)%></font></td>
			  </tr>
<% next %>
			</table>
	  <hr width="70%" />
	  <table class="cobtbl" width="<%=maintablewidth%>" border="0" bordercolor="#B1B1B1" cellspacing="1" cellpadding="3" bgcolor="#B1B1B1">
<%
		else
%>
			  <tr>
			    <td class="cobll" bgcolor="#FFFFFF" colspan="2" height="30" align="center"><strong>The tracking system returned the following error : <%=errormsg%></strong></td>
			  </tr>
			</table>
	  <hr width="70%" />
	  <table class="cobtbl" width="<%=maintablewidth%>" border="0" bordercolor="#B1B1B1" cellspacing="1" cellpadding="3" bgcolor="#B1B1B1">
<%
		end if
	End If
	FedexTrack = success
	set objHttp = nothing
end function
if Trim(Request.Form("trackno"))<>"" then
	FedexTrack(Trim(Request.Form("trackno")))
end if
%>
		<form method="post" name="trackform" action="tracking.asp">
		  <input type="hidden" name="carrier" value="fedex">
		  <tr>
			<td class="cobhl" width="50%" bgcolor="#EBEBEB" align="right">Please enter your FedEx Tracking Number : </td>
			<td class="cobll" width="50%" bgcolor="#FFFFFF"><input type="text" size="30" name="trackno" value="<%=Trim(Request("trackno"))%>" /></td>
		  </tr>
		  <tr>
			<td class="cobhl" width="50%" bgcolor="#EBEBEB" align="right">Show Activity : </td>
			<td class="cobll" width="50%" bgcolor="#FFFFFF"><select name="activity" size="1"><option value="LAST">Show Last Activity Only</option><option value="ALL"<% if Trim(Request.Form("activity")="ALL") then response.write " selected"%>>Show All Activity</option></select></td>
		  </tr>
		  <tr>
			<td class="cobll" bgcolor="#FFFFFF" colspan="2"><table width="100%" cellspacing="0" cellpadding="0" border="0">
				<tr>
				  <td class="cobll" bgcolor="#FFFFFF" width="17%" height="26">&nbsp;</td>
				  <td class="cobll" bgcolor="#FFFFFF" width="66%" align="center"><input type="submit" value="Track Package" /></td>
				  <td class="cobll" bgcolor="#FFFFFF" width="17%" height="26" align="right" valign="bottom"><img src="images/tablebr.gif" alt="" /></td>
				</tr>
			  </table></td>
		  </tr>
		</form>
	  </table>
	  <table border="0" cellspacing="0" cellpadding="0" width="100%" bgcolor="#FFFFFF" align="center">
        <tr>
          <td width="100%" align="center"><p>&nbsp;<br /><font size="1">FedEx&reg; is a registered service mark of Federal Express Corporation.<br />
			FedEx logos used by permission. All rights reserved.</font></p></td>
		</tr>
	  </table>
	  <br />
<%
else ' undecided
%>
&nbsp;<br />
	  <table class="cobtbl" width="<%=maintablewidth%>" border="0" bordercolor="#B1B1B1" cellspacing="1" cellpadding="3" bgcolor="#B1B1B1">
		<tr>
		  <td class="cobll" bgcolor="#FFFFFF" colspan="2">
			<table border="0" cellspacing="0" cellpadding="0" width="100%" bgcolor="">
			  <tr>
				<td width="98" align="left"><% if incupscopyright then %><img src="images/LOGO_S.gif" alt="UPS" /><% else response.write "&nbsp;" end if %></td><td align="center">&nbsp;<br /><font size="4"><strong>Please select your shipping carrier.</strong></font><br />&nbsp;</td><td width="98"><% if incfedexcopyright then %><img src="images/fedexsmall.gif" alt="FedEx" /><% else response.write "&nbsp;" end if %></td>
			  </tr>
			</table>
		  </td>
		</tr>
<%		if shipType=4 OR alternateratesups<>"" then %>	
		<form method="post" action="tracking.asp">
		  <input type="hidden" name="carrier" value="ups">
		  <tr>
			<td class="cobhl" width="50%" bgcolor="#EBEBEB" align="right">Products shipped via UPS : </td>
			<td class="cobll" width="50%" bgcolor="#FFFFFF"><input type="submit" value="<%=xxGo%>" /></td>
		  </tr>
		</form>
<%		end if
		if shipType=3 OR alternateratesusps<>"" then %>
		<form method="post" action="tracking.asp">
		  <input type="hidden" name="carrier" value="usps">
		  <tr>
			<td class="cobhl" width="50%" bgcolor="#EBEBEB" align="right">Products shipped via USPS : </td>
			<td class="cobll" width="50%" bgcolor="#FFFFFF"><input type="submit" value="<%=xxGo%>" /></td>
		  </tr>
		</form>
<%		end if
		if shipType=7 OR alternateratesfedex<>"" then %>
		<form method="post" action="tracking.asp">
		  <input type="hidden" name="carrier" value="fedex">
		  <tr>
			<td class="cobhl" width="50%" bgcolor="#EBEBEB" align="right">Products shipped via FedEx : </td>
			<td class="cobll" width="50%" bgcolor="#FFFFFF"><input type="submit" value="<%=xxGo%>" /></td>
		  </tr>
		</form>
<%		end if %>
	  </table>
	  <table border="0" cellspacing="0" cellpadding="0" width="100%" bgcolor="#FFFFFF" align="center">
		<tr><td>&nbsp;</td></tr>
<%	if incupscopyright then %>
        <tr>
          <td width="100%" align="center"><p>&nbsp;<br /><font size="1">UPS&reg;, UPS & Shield Design&reg; and UNITED PARCEL SERVICE&reg; 
				  are<br />registered trademarks of United Parcel Service of America, Inc.</font></p></td>
		</tr>
<%	end if
	if incfedexcopyright then %>
        <tr>
          <td width="100%" align="center"><p>&nbsp;<br /><font size="1">FedEx&reg; is a registered service mark of Federal Express Corporation.<br />
			FedEx logos used by permission. All rights reserved.</font></p></td>
		</tr>
<%	end if %>
	  </table>
	  <br />
<%
end if
cnn.Close
set rs = nothing
set cnn = nothing
%>
