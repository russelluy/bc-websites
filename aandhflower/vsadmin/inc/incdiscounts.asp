<%
if storesessionvalue="" then storesessionvalue="virtualstore"
if Session("loggedon") <> storesessionvalue OR disallowlogin=TRUE then response.end
Dim sSQL,rs,alldata,alladmin,success,cnn,rowcounter,errmsg
success=true
Set rs = Server.CreateObject("ADODB.RecordSet")
Set cnn=Server.CreateObject("ADODB.Connection")
cnn.open sDSN
alreadygotadmin = getadminsettings()
Session.LCID = 1033
sSQL = ""
if request.form("posted")="1" then
	if request.form("act")="delete" then
		sSQL = "DELETE FROM cpnassign WHERE cpaCpnID=" & request.form("id")
		cnn.Execute(sSQL)
		sSQL = "DELETE FROM coupons WHERE cpnID=" & request.form("id")
		cnn.Execute(sSQL)
		response.write "<meta http-equiv=""refresh"" content=""3; url=admindiscounts.asp"">"
	elseif request.form("act")="domodify" then
		cpnName = trim(request.form("cpnName"))
		sSQL = "UPDATE coupons SET cpnName='" & replace(request.form("cpnName"),"'","''") & "'"
			for index=2 to adminlanguages+1
				if (adminlangsettings AND 1024)=1024 then sSQL = sSQL & ",cpnName"&index&"='" & replace(request.form("cpnName"&index),"'","''") & "'"
			next
			if trim(request.form("cpnWorkingName"))<>"" then
				sSQL = sSQL & ",cpnWorkingName='" & replace(request.form("cpnWorkingName"),"'","''")&"'"
			else
				sSQL = sSQL & ",cpnWorkingName='" & replace(cpnName,"'","''")&"'"
			end if
			if request.form("cpnIsCoupon")="0" then
				sSQL = sSQL & ",cpnNumber='',"
			else
				sSQL = sSQL & ",cpnNumber='" & replace(request.form("cpnNumber"),"'","''") & "',"
			end if
			sSQL = sSQL & "cpnType=" & request.form("cpnType") & ","
			numdays=0
			if IsNumeric(request.form("cpnEndDate")) then numdays = Int(request.form("cpnEndDate"))
			if numdays > 0 then
				tdt = Date() + Int(request.form("cpnEndDate"))
				sSQL = sSQL & "cpnEndDate=" & datedelim & VSUSDate(tdt) & datedelim & ","
			else
				sSQL = sSQL & "cpnEndDate=" & datedelim & VSUSDate(DateSerial(3000,1,1)) & datedelim & ","
			end if
			if IsNumeric(request.form("cpnDiscount")) AND request.form("cpnType")<>"0" then
				sSQL = sSQL & "cpnDiscount=" & request.form("cpnDiscount") & ","
			else
				sSQL = sSQL & "cpnDiscount=0,"
			end if
			if IsNumeric(request.form("cpnThreshold")) then
				sSQL = sSQL & "cpnThreshold=" & request.form("cpnThreshold") & ","
			else
				sSQL = sSQL & "cpnThreshold=0,"
			end if
			if IsNumeric(request.form("cpnThresholdMax")) then
				sSQL = sSQL & "cpnThresholdMax=" & request.form("cpnThresholdMax") & ","
			else
				sSQL = sSQL & "cpnThresholdMax=0,"
			end if
			if IsNumeric(request.form("cpnThresholdRepeat")) then
				sSQL = sSQL & "cpnThresholdRepeat=" & request.form("cpnThresholdRepeat") & ","
			else
				sSQL = sSQL & "cpnThresholdRepeat=0,"
			end if
			if IsNumeric(request.form("cpnQuantity")) then
				sSQL = sSQL & "cpnQuantity=" & request.form("cpnQuantity") & ","
			else
				sSQL = sSQL & "cpnQuantity=0,"
			end if
			if IsNumeric(request.form("cpnQuantityMax")) then
				sSQL = sSQL & "cpnQuantityMax=" & request.form("cpnQuantityMax") & ","
			else
				sSQL = sSQL & "cpnQuantityMax=0,"
			end if
			if IsNumeric(request.form("cpnQuantityRepeat")) then
				sSQL = sSQL & "cpnQuantityRepeat=" & request.form("cpnQuantityRepeat") & ","
			else
				sSQL = sSQL & "cpnQuantityRepeat=0,"
			end if
			if Trim(request.form("cpnNumAvail"))<>"" AND IsNumeric(request.form("cpnNumAvail")) then
				sSQL = sSQL & "cpnNumAvail=" & request.form("cpnNumAvail") & ","
			else
				sSQL = sSQL & "cpnNumAvail=30000000,"
			end if
			if request.form("cpnType")="0" then
				sSQL = sSQL & "cpnCntry=" & request.form("cpnCntry") &","
			else
				sSQL = sSQL & "cpnCntry=0,"
			end if
			sSQL = sSQL & "cpnIsCoupon=" & request.form("cpnIsCoupon") &","
			if request.form("cpnType")="0" then
				sSQL = sSQL & "cpnSitewide=1"
			else
				sSQL = sSQL & "cpnSitewide=" & request.form("cpnSitewide")
			end if
			sSQL = sSQL & " WHERE cpnID="&Request.Form("id")
		cnn.Execute(sSQL)
		response.write "<meta http-equiv=""refresh"" content=""3; url=admindiscounts.asp"">"
	elseif request.form("act")="doaddnew" then
		cpnName = trim(request.form("cpnName"))
		sSQL = "INSERT INTO coupons (cpnName,"
			for index=2 to adminlanguages+1
				if (adminlangsettings AND 1024)=1024 then sSQL = sSQL & "cpnName"&index&","
			next
			sSQL = sSQL & "cpnWorkingName,cpnNumber,cpnType,cpnEndDate,cpnDiscount,cpnThreshold,cpnThresholdMax,cpnThresholdRepeat,cpnQuantity,cpnQuantityMax,cpnQuantityRepeat,cpnNumAvail,cpnCntry,cpnIsCoupon,cpnSitewide) VALUES (" & _
			"'"&replace(cpnName,"'","''")&"',"
			for index=2 to adminlanguages+1
				if (adminlangsettings AND 1024)=1024 then sSQL = sSQL & "'"&replace(trim(request.form("cpnName"&index)),"'","''")&"',"
			next
			if trim(request.form("cpnWorkingName"))<>"" then
				sSQL = sSQL & "'"&replace(trim(request.form("cpnWorkingName")),"'","''")&"',"
			else
				sSQL = sSQL & "'"&replace(cpnName,"'","''")&"',"
			end if
			if request.form("cpnIsCoupon")="0" then
				sSQL = sSQL & "'',"
			else
				sSQL = sSQL & "'"&replace(request.form("cpnNumber"),"'","''")&"',"
			end if
			sSQL = sSQL & request.form("cpnType") & ","
			numdays=0
			if IsNumeric(request.form("cpnEndDate")) then numdays = Int(request.form("cpnEndDate"))
			if numdays > 0 then
				tdt = Date() + Int(request.form("cpnEndDate"))
				sSQL = sSQL & datedelim & VSUSDate(tdt) & datedelim & ","
			else
				sSQL = sSQL & datedelim & VSUSDate(DateSerial(3000,1,1)) & datedelim & ","
			end if
			if IsNumeric(request.form("cpnDiscount")) AND request.form("cpnType")<>"0" then
				sSQL = sSQL & request.form("cpnDiscount") & ","
			else
				sSQL = sSQL & "0,"
			end if
			if IsNumeric(request.form("cpnThreshold")) then
				sSQL = sSQL & request.form("cpnThreshold") & ","
			else
				sSQL = sSQL & "0,"
			end if
			if IsNumeric(request.form("cpnThresholdMax")) then
				sSQL = sSQL & request.form("cpnThresholdMax") & ","
			else
				sSQL = sSQL & "0,"
			end if
			if IsNumeric(request.form("cpnThresholdRepeat")) then
				sSQL = sSQL & request.form("cpnThresholdRepeat") & ","
			else
				sSQL = sSQL & "0,"
			end if
			if IsNumeric(request.form("cpnQuantity")) then
				sSQL = sSQL & request.form("cpnQuantity") & ","
			else
				sSQL = sSQL & "0,"
			end if
			if IsNumeric(request.form("cpnQuantityMax")) then
				sSQL = sSQL & request.form("cpnQuantityMax") & ","
			else
				sSQL = sSQL & "0,"
			end if
			if IsNumeric(request.form("cpnQuantityRepeat")) then
				sSQL = sSQL & request.form("cpnQuantityRepeat") & ","
			else
				sSQL = sSQL & "0,"
			end if
			if Trim(request.form("cpnNumAvail"))<>"" AND IsNumeric(request.form("cpnNumAvail")) then
				sSQL = sSQL & request.form("cpnNumAvail") & ","
			else
				sSQL = sSQL & "30000000,"
			end if
			if request.form("cpnType")="0" then
				sSQL = sSQL & request.form("cpnCntry") &","
			else
				sSQL = sSQL & "0,"
			end if
			sSQL = sSQL & request.form("cpnIsCoupon") &","
			if request.form("cpnType")="0" then
				sSQL = sSQL & "1)"
			else
				sSQL = sSQL & request.form("cpnSitewide") & ")"
			end if
		cnn.Execute(sSQL)
		response.write "<meta http-equiv=""refresh"" content=""3; url=admindiscounts.asp"">"
	end if
end if
%>
<script language="javascript" type="text/javascript">
<!--
var savebg, savebc, savecol;
function formvalidator(theForm)
{
  if(theForm.cpnName.value == "")
  {
    alert("<%=yyPlsEntr%> \"<%=yyDisTxt%>\".");
    theForm.cpnName.focus();
    return (false);
  }
  if(theForm.cpnName.value.length > 255)
  {
    alert("<%=yyMax255%> \"<%=yyDisTxt%>\".");
    theForm.cpnName.focus();
    return (false);
  }
  if(theForm.cpnType.selectedIndex!=0){
	if(theForm.cpnDiscount.value == "")
	{
	  alert("<%=yyPlsEntr%> \"<%=yyDscAmt%>\".");
	  theForm.cpnDiscount.focus();
	  return (false);
	}
	if(theForm.cpnType.selectedIndex==2){
	  if(theForm.cpnDiscount.value < 0 || theForm.cpnDiscount.value > 100){
		alert("<%=yyNum100%> \"<%=yyDscAmt%>\".");
		theForm.cpnDiscount.focus();
		return (false);
	  }
	}
  }
  if(theForm.cpnIsCoupon.selectedIndex==1){
	if(theForm.cpnNumber.value == "")
	{
	  alert("<%=yyPlsEntr%> \"<%=yyCpnCod%>\".");
	  theForm.cpnNumber.focus();
	  return (false);
	}
	var checkOK = "0123456789abcdefghijklmnopqrstuvwxyz-_";
	var checkStr = theForm.cpnNumber.value.toLowerCase();
	var allValid = true;
	for (i = 0;  i < checkStr.length;  i++)
	{
		ch = checkStr.charAt(i);
		for (j = 0;  j < checkOK.length;  j++)
			if (ch == checkOK.charAt(j))
				break;
		if (j == checkOK.length){
			allValid = false;
				break;
		}
	}
	if (!allValid)
	{
		alert("<%=yyAlpha2%> \"<%=yyCpnCod%>\".");
		theForm.cpnNumber.focus();
		return (false);
	}
  }
  var checkOK = "0123456789";
  var checkStr = theForm.cpnNumAvail.value;
  var allValid = true;
  for (i = 0;  i < checkStr.length;  i++)
  {
	ch = checkStr.charAt(i);
	for (j = 0;  j < checkOK.length;  j++)
		if (ch == checkOK.charAt(j))
			break;
	if (j == checkOK.length){
		allValid = false;
			break;
	}
  }
  if (!allValid)
  {
	alert("<%=yyOnlyNum%> \"<%=yyNumAvl%>\".");
	theForm.cpnNumAvail.focus();
	return (false);
  }
  if(theForm.cpnNumAvail.value != "" && theForm.cpnNumAvail.value > 1000000)
  {
    alert("<%=yyNumMil%> \"<%=yyNumAvl%>\"<%=yyOrBlank%>");
    theForm.cpnNumAvail.focus();
    return (false);
  }
  var checkOK = "0123456789";
  var checkStr = theForm.cpnEndDate.value;
  var allValid = true;
  for (i = 0;  i < checkStr.length;  i++)
  {
	ch = checkStr.charAt(i);
	for (j = 0;  j < checkOK.length;  j++)
		if (ch == checkOK.charAt(j))
			break;
	if (j == checkOK.length){
		allValid = false;
			break;
	}
  }
  if (!allValid)
  {
	alert("<%=yyOnlyNum%> \"<%=yyDaysAv%>\".");
	theForm.cpnEndDate.focus();
	return (false);
  }
  var checkOK = "0123456789.";
  var checkStr = theForm.cpnThreshold.value;
  var allValid = true;
  for (i = 0;  i < checkStr.length;  i++)
  {
	ch = checkStr.charAt(i);
	for (j = 0;  j < checkOK.length;  j++)
		if (ch == checkOK.charAt(j))
			break;
	if (j == checkOK.length){
		allValid = false;
			break;
	}
  }
  if (!allValid)
  {
	alert("<%=yyOnlyDec%> \"<%=yyMinPur%>\".");
	theForm.cpnThreshold.focus();
	return (false);
  }
  var checkOK = "0123456789.";
  var checkStr = theForm.cpnThresholdRepeat.value;
  var allValid = true;
  for (i = 0;  i < checkStr.length;  i++)
  {
	ch = checkStr.charAt(i);
	for (j = 0;  j < checkOK.length;  j++)
		if (ch == checkOK.charAt(j))
			break;
	if (j == checkOK.length){
		allValid = false;
			break;
	}
  }
  if (!allValid)
  {
	alert("<%=yyOnlyDec%> \"<%=yyRepEvy%>\".");
	theForm.cpnThresholdRepeat.focus();
	return (false);
  }
  var checkOK = "0123456789.";
  var checkStr = theForm.cpnThresholdMax.value;
  var allValid = true;
  for (i = 0;  i < checkStr.length;  i++)
  {
	ch = checkStr.charAt(i);
	for (j = 0;  j < checkOK.length;  j++)
		if (ch == checkOK.charAt(j))
			break;
	if (j == checkOK.length){
		allValid = false;
			break;
	}
  }
  if (!allValid)
  {
	alert("<%=yyOnlyDec%> \"<%=yyMaxPur%>\".");
	theForm.cpnThresholdMax.focus();
	return (false);
  }
  var checkOK = "0123456789";
  var checkStr = theForm.cpnQuantity.value;
  var allValid = true;
  for (i = 0;  i < checkStr.length;  i++)
  {
	ch = checkStr.charAt(i);
	for (j = 0;  j < checkOK.length;  j++)
		if (ch == checkOK.charAt(j))
			break;
	if (j == checkOK.length){
		allValid = false;
			break;
	}
  }
  if (!allValid)
  {
	alert("<%=yyOnlyNum%> \"<%=yyMinQua%>\".");
	theForm.cpnQuantity.focus();
	return (false);
  }
  var checkOK = "0123456789";
  var checkStr = theForm.cpnQuantityRepeat.value;
  var allValid = true;
  for (i = 0;  i < checkStr.length;  i++)
  {
	ch = checkStr.charAt(i);
	for (j = 0;  j < checkOK.length;  j++)
		if (ch == checkOK.charAt(j))
			break;
	if (j == checkOK.length){
		allValid = false;
			break;
	}
  }
  if (!allValid)
  {
	alert("<%=yyOnlyNum%> \"<%=yyRepEvy%>\".");
	theForm.cpnQuantityRepeat.focus();
	return (false);
  }
  var checkOK = "0123456789";
  var checkStr = theForm.cpnQuantityMax.value;
  var allValid = true;
  for (i = 0;  i < checkStr.length;  i++)
  {
	ch = checkStr.charAt(i);
	for (j = 0;  j < checkOK.length;  j++)
		if (ch == checkOK.charAt(j))
			break;
	if (j == checkOK.length){
		allValid = false;
			break;
	}
  }
  if (!allValid)
  {
	alert("<%=yyOnlyNum%> \"<%=yyMaxQua%>\".");
	theForm.cpnQuantityMax.focus();
	return (false);
  }
  var checkOK = "0123456789.";
  var checkStr = theForm.cpnDiscount.value;
  var allValid = true;
  for (i = 0;  i < checkStr.length;  i++)
  {
	ch = checkStr.charAt(i);
	for (j = 0;  j < checkOK.length;  j++)
		if (ch == checkOK.charAt(j))
			break;
	if (j == checkOK.length){
		allValid = false;
			break;
	}
  }
  if (!allValid)
  {
	alert("<%=yyOnlyDec%> \"<%=yyDscAmt%>\".");
	theForm.cpnDiscount.focus();
	return (false);
  }
  document.mainform.cpnNumber.disabled=false;
  document.mainform.cpnDiscount.disabled=false;
  document.mainform.cpnCntry.disabled=false;
  document.mainform.cpnSitewide.disabled=false;
  document.mainform.cpnThresholdRepeat.disabled=false;
  document.mainform.cpnQuantityRepeat.disabled=false;
  return (true);
}
function couponcodeactive(forceactive){
	if(document.mainform.cpnIsCoupon.selectedIndex==0){
		document.mainform.cpnNumber.style.backgroundColor="#DDDDDD";
		document.mainform.cpnNumber.style.borderColor="#aa3300";
		document.mainform.cpnNumber.style.color="#aa3300";
		document.mainform.cpnNumber.disabled=true;
	}
	else if(document.mainform.cpnIsCoupon.selectedIndex==1){
		document.mainform.cpnNumber.style.backgroundColor=savebg;
		document.mainform.cpnNumber.style.borderColor=savebc;
		document.mainform.cpnNumber.style.color=savecol;
		document.mainform.cpnNumber.disabled=false;
	}
}
function changecouponeffect(forceactive){
	if(document.mainform.cpnType.selectedIndex==0){
		document.mainform.cpnDiscount.style.backgroundColor="#DDDDDD";
		document.mainform.cpnDiscount.style.borderColor="#aa3300";
		document.mainform.cpnDiscount.style.color="#aa3300";
		document.mainform.cpnDiscount.disabled=true;

		document.mainform.cpnCntry.style.backgroundColor=savebg;
		document.mainform.cpnCntry.style.borderColor=savebc;
		document.mainform.cpnCntry.style.color=savecol;
		document.mainform.cpnCntry.disabled=false;

		document.mainform.cpnSitewide.style.backgroundColor="#DDDDDD";
		document.mainform.cpnSitewide.style.borderColor="#aa3300";
		document.mainform.cpnSitewide.style.color="#aa3300";
		document.mainform.cpnSitewide.disabled=true;
	}else{
		document.mainform.cpnDiscount.style.backgroundColor=savebg;
		document.mainform.cpnDiscount.style.borderColor=savebc;
		document.mainform.cpnDiscount.style.color=savecol;
		document.mainform.cpnDiscount.disabled=false;

		document.mainform.cpnCntry.style.backgroundColor="#DDDDDD";
		document.mainform.cpnCntry.style.borderColor="#aa3300";
		document.mainform.cpnCntry.style.color="#aa3300";
		document.mainform.cpnCntry.disabled=true;

		document.mainform.cpnSitewide.style.backgroundColor=savebg;
		document.mainform.cpnSitewide.style.borderColor=savebc;
		document.mainform.cpnSitewide.style.color=savecol;
		document.mainform.cpnSitewide.disabled=false;
	}
	if(document.mainform.cpnType.selectedIndex==1){
		document.mainform.cpnThresholdRepeat.style.backgroundColor=savebg;
		document.mainform.cpnThresholdRepeat.style.borderColor=savebc;
		document.mainform.cpnThresholdRepeat.style.color=savecol;
		document.mainform.cpnThresholdRepeat.disabled=false;

		document.mainform.cpnQuantityRepeat.style.backgroundColor=savebg;
		document.mainform.cpnQuantityRepeat.style.borderColor=savebc;
		document.mainform.cpnQuantityRepeat.style.color=savecol;
		document.mainform.cpnQuantityRepeat.disabled=false;
	}else{
		document.mainform.cpnThresholdRepeat.style.backgroundColor="#DDDDDD";
		document.mainform.cpnThresholdRepeat.style.borderColor="#aa3300";
		document.mainform.cpnThresholdRepeat.style.color="#aa3300";
		document.mainform.cpnThresholdRepeat.disabled=true;

		document.mainform.cpnQuantityRepeat.style.backgroundColor="#DDDDDD";
		document.mainform.cpnQuantityRepeat.style.borderColor="#aa3300";
		document.mainform.cpnQuantityRepeat.style.color="#aa3300";
		document.mainform.cpnQuantityRepeat.disabled=true;
	}
}
//-->
</script>
      <table border="0" cellspacing="<%=maintablespacing%>" cellpadding="<%=maintablepadding%>" width="<%=maintablewidth%>" bgcolor="<%=maintablebg%>" align="center">
<% if request.form("posted")="1" AND (request.form("act")="modify" OR request.form("act")="addnew") then
		if request.form("act")="modify" then
			sSQL = "SELECT cpnName,cpnName2,cpnName3,cpnWorkingName,cpnNumber,cpnType,cpnEndDate,cpnDiscount,cpnThreshold,cpnThresholdMax,cpnThresholdRepeat,cpnQuantity,cpnQuantityMax,cpnQuantityRepeat,cpnNumAvail,cpnCntry,cpnIsCoupon,cpnSitewide FROM coupons WHERE cpnID="&Request.Form("id")
			rs.Open sSQL,cnn,0,1
			cpnName = rs("cpnName")
			cpnName2 = rs("cpnName2")&""
			cpnName3 = rs("cpnName3")&""
			cpnWorkingName = rs("cpnWorkingName")
			cpnNumber = rs("cpnNumber")
			cpnType = rs("cpnType")
			cpnEndDate = rs("cpnEndDate")
			cpnDiscount = rs("cpnDiscount")
			cpnThreshold = rs("cpnThreshold")
			cpnThresholdMax = rs("cpnThresholdMax")
			cpnThresholdRepeat = rs("cpnThresholdRepeat")
			cpnQuantity = rs("cpnQuantity")
			cpnQuantityMax = rs("cpnQuantityMax")
			cpnQuantityRepeat = rs("cpnQuantityRepeat")
			cpnNumAvail = rs("cpnNumAvail")
			cpnCntry = rs("cpnCntry")
			cpnIsCoupon = rs("cpnIsCoupon")
			cpnSitewide = rs("cpnSitewide")
			rs.Close
		else
			cpnName = ""
			cpnName2 = ""
			cpnName3 = ""
			cpnWorkingName = ""
			cpnNumber = ""
			cpnType = 0
			cpnEndDate = DateSerial(3000,1,1)
			cpnDiscount = ""
			cpnThreshold = 0
			cpnThresholdMax = 0
			cpnThresholdRepeat = 0
			cpnQuantity = 0
			cpnQuantityMax = 0
			cpnQuantityRepeat = 0
			cpnNumAvail = 30000000
			cpnCntry = 0
			cpnIsCoupon = 0
			cpnSitewide = 0
		end if
%>
        <tr>
		<form name="mainform" method="post" action="admindiscounts.asp" onsubmit="return formvalidator(this)">
		  <td width="100%">
			<input type="hidden" name="posted" value="1" />
		<% if request.form("act")="modify" then %>
			<input type="hidden" name="act" value="domodify" />
			<input type="hidden" name="id" value="<%=Request.Form("id")%>" />
		<% else %>
			<input type="hidden" name="act" value="doaddnew" />
		<% end if %>
            <table width="100%" border="0" cellspacing="0" cellpadding="3" bgcolor="">
			  <tr> 
                <td width="100%" colspan="2" align="center"><strong><%=yyDscNew%></strong><br />&nbsp;</td>
			  </tr>
			  <tr>
				<td width="40%" align="right"><strong><%=yyCpnDsc%>:</td>
				<td width="60%"><select name="cpnIsCoupon" size="1" onchange="couponcodeactive(false);">
					<option value="0"><%=yyDisco%></option>
					<option value="1" <% if Int(cpnIsCoupon)=1 then response.write "selected" %>><%=yyCoupon%></option>
					</select></td>
			  </tr>
			  <tr>
				<td width="40%" align="right"><strong><%=yyDscEff%>:</td>
				<td width="60%"><select name="cpnType" size="1" onchange="changecouponeffect(false);">
					<option value="0"><%=yyFrSShp%></option>
					<option value="1" <% if Int(cpnType)=1 then response.write "selected" %>><%=yyFlatDs%></option>
					<option value="2" <% if Int(cpnType)=2 then response.write "selected" %>><%=yyPerDis%></option>
					</select></td>
			  </tr>
			  <tr>
				<td width="40%" align="right"><strong><%=yyDisTxt%>:</strong></td>
				<td width="60%"><input type="text" name="cpnName" size="30" value="<%=Replace(cpnName,"""","&quot;")%>" /></td>
			  </tr>
<%				for index=2 to adminlanguages+1
					if (adminlangsettings AND 1024)=1024 then
						if request.form("act")="modify" then execute ("cpnName = cpnName" & index) else cpnName = ""
			%><tr>
				<td width="40%" align="right"><strong><%=yyDisTxt & " " & index%>:</strong></td>
				<td width="60%"><input type="text" name="cpnName<%=index%>" size="30" value="<%=Replace(cpnName,"""","&quot;")%>" /></td>
			  </tr><%
					end if
				next %>
			  <tr>
				<td width="40%" align="right"><strong><%=yyWrkNam%>:</strong></td>
				<td width="60%"><input type="text" name="cpnWorkingName" size="30" value="<%=Replace(cpnWorkingName,"""","&quot;")%>" /></td>
			  </tr>
			  <tr>
				<td width="40%" align="right"><strong><%=yyCpnCod%>:</strong></td>
				<td width="60%"><input type="text" name="cpnNumber" size="30" value="<%=cpnNumber%>" /></td>
			  </tr>
			  <tr>
				<td width="40%" align="right"><strong><%=yyNumAvl%>:</strong></td>
				<td width="60%"><input type="text" name="cpnNumAvail" size="10" value="<% if Int(cpnNumAvail)<>30000000 then response.write cpnNumAvail%>" /></td>
			  </tr>
			  <tr>
				<td width="40%" align="right"><strong><%=yyDaysAv%>:</strong></td>
				<td width="60%"><input type="text" name="cpnEndDate" size="10" value="<%
				if cpnEndDate<>DateSerial(3000,1,1) then
					if cpnEndDate-Date() < 0 then response.write "Expired" else response.write cpnEndDate-Date()
				end if %>" /></td>
			  </tr>
			  <tr>
				<td width="40%" align="right"><strong><%=yyMinPur%>:</strong></td>
				<td width="60%"><input type="text" name="cpnThreshold" size="10" value="<% if Int(cpnThreshold)>0 then response.write cpnThreshold%>" /> <strong><%=yyRepEvy%>:</strong> <input type="text" name="cpnThresholdRepeat" size="10" value="<% if Int(cpnThresholdRepeat)>0 then response.write cpnThresholdRepeat%>" /></td>
			  </tr>
			  <tr>
				<td width="40%" align="right"><strong><%=yyMaxPur%>:</strong></td>
				<td width="60%"><input type="text" name="cpnThresholdMax" size="10" value="<% if Int(cpnThresholdMax)>0 then response.write cpnThresholdMax%>" /></td>
			  </tr>
			  <tr>
				<td width="40%" align="right"><strong><%=yyMinQua%>:</strong></td>
				<td width="60%"><input type="text" name="cpnQuantity" size="10" value="<% if Int(cpnQuantity)>0 then response.write cpnQuantity%>" /> <strong><%=yyRepEvy%>:</strong> <input type="text" name="cpnQuantityRepeat" size="10" value="<% if Int(cpnQuantityRepeat)>0 then response.write cpnQuantityRepeat%>" /></td>
			  </tr>
			  <tr>
				<td width="40%" align="right"><strong><%=yyMaxQua%>:</strong></td>
				<td width="60%"><input type="text" name="cpnQuantityMax" size="10" value="<% if Int(cpnQuantityMax)>0 then response.write cpnQuantityMax%>" /></td>
			  </tr>
			  <tr>
				<td width="40%" align="right"><strong><%=yyDscAmt%>:</strong></td>
				<td width="60%"><input type="text" name="cpnDiscount" size="10" value="<%=cpnDiscount%>" /></td>
			  </tr>
			  <tr>
				<td width="40%" align="right"><strong><%=yyScope%>:</strong></td>
				<td width="60%"><select name="cpnSitewide" size="1">
					<option value="0"><%=yyIndCat%></option>
					<option value="3" <% if Int(cpnSitewide)=3 then response.write "selected" %>><%=yyDsCaTo%></option>
					<option value="2" <% if Int(cpnSitewide)=2 then response.write "selected" %>><%=yyGlInPr%></option>
					<option value="1" <% if Int(cpnSitewide)=1 then response.write "selected" %>><%=yyGlPrTo%></option>
					</select></td>
			  </tr>
			  <tr>
				<td width="40%" align="right"><strong><%=yyRestr%>:</strong></td>
				<td width="60%"><select name="cpnCntry" size="1">
					<option value="0"><%=yyAppAll%></option>
					<option value="1" <% if Int(cpnCntry)=1 then response.write "selected" %>><%=yyYesRes%></option>
					</select></td>
			  </tr>
			  <tr>
                <td width="100%" colspan="2" align="center"><br /><input type="submit" value="<%=yySubmit%>" /><br />&nbsp;</td>
			  </tr>
			  <tr> 
                <td width="100%" colspan="2" align="center"><br />
                          <a href="admin.asp"><strong><%=yyAdmHom%></strong></a><br />
                          &nbsp;</td>
			  </tr>
            </table>
		  </td>
		</form>
        </tr>
<script language="javascript" type="text/javascript">
<!--
savebg=document.mainform.cpnNumber.style.backgroundColor;
savebc=document.mainform.cpnNumber.style.borderColor;
savecol=document.mainform.cpnNumber.style.color;
couponcodeactive(false);
changecouponeffect(false);
//-->
</script>
<% elseif request.form("posted")="1" AND success then %>
        <tr>
          <td width="100%">
			<table width="100%" border="0" cellspacing="0" cellpadding="3" bgcolor="">
			  <tr> 
                <td width="100%" colspan="2" align="center"><br /><strong><%=yyUpdSuc%></strong><br /><br /><%=yyNowFrd%><br /><br />
                        <%=yyNoAuto%> <A href="admindiscounts.asp"><strong><%=yyClkHer%></strong></a>.<br />
                        <br />
				<img src="../images/clearpixel.gif" width="300" height="3" alt="" />
                </td>
			  </tr>
			</table></td>
        </tr>
<% elseif request.form("posted")="1" then %>
        <tr>
          <td width="100%">
			<table width="100%" border="0" cellspacing="0" cellpadding="3" bgcolor="">
			  <tr> 
                <td width="100%" colspan="2" align="center"><br /><font color="#FF0000"><strong><%=yyOpFai%></strong></font><br /><br /><%=errmsg%><br /><br />
				<a href="javascript:history.go(-1)"><strong><%=yyClkBac%></strong></a></td>
			  </tr>
			</table></td>
        </tr>
<% else 
		sSQL = "SELECT cpnID,cpnWorkingName,cpnSitewide,cpnIsCoupon,cpnEndDate FROM coupons ORDER BY cpnIsCoupon,cpnWorkingName"
		rs.Open sSQL,cnn,0,1
		if NOT rs.EOF then alldata=rs.getrows
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
        <tr>
		  <form name="mainform" method="post" action="admindiscounts.asp">
		  <td width="100%">
			<input type="hidden" name="posted" value="1" />
			<input type="hidden" name="act" value="xxxxx" />
			<input type="hidden" name="id" value="xxxxx" />
			<input type="hidden" name="selectedq" value="1" />
			<input type="hidden" name="newval" value="1" />
            <table width="100%" border="0" cellspacing="0" cellpadding="1" bgcolor="">
			  <tr> 
                <td width="100%" colspan="6" align="center"><br /><strong><%=yyDscAdm%></strong><br />&nbsp;</td>
			  </tr>
			  <tr>
				<td width="40%" align="left"><strong><%=yyWrkNam%></strong></td>
				<td width="10%" align="center"><strong><%=yyType%></strong></td>
				<td width="20%" align="center"><strong><%=yyExpDat%></strong></td>
				<td width="10%" align="center"><strong><%=yyGlobal%></strong></td>
				<td width="10%" align="center"><strong><%=yyModify%></strong></td>
				<td width="10%" align="center"><strong><%=yyDelete%></strong></td>
			  </tr>
<%	Session.LCID = saveLCID
	if IsArray(alldata) then
		for rowcounter=0 to UBOUND(alldata,2)
			if bgcolor="#E7EAEF" then bgcolor="#FFFFFF" else bgcolor="#E7EAEF" %>
			  <tr bgcolor="<%=bgcolor%>">
				<td><%=alldata(1,rowcounter)%></td>
				<td align="center"><%	if alldata(3,rowcounter)=1 then response.write yyCoupon else response.write yyDisco%></td>
				<td align="center"><%	if alldata(4,rowcounter)=DateSerial(3000,1,1) then
											response.write yyNever
										elseif alldata(4,rowcounter)-Date() < 0 then
											response.write "<font color='#FF0000'>"&yyExpird&"</font>"
										else
											response.write alldata(4,rowcounter)
										end if %></td>
				<td align="center"><% if alldata(2,rowcounter)=1 OR alldata(2,rowcounter)=2 then response.write yyYes else response.write yyNo %></td>
				<td align="center"><input type=button value="<%=yyModify%>" onclick="modrec('<%=alldata(0,rowcounter)%>')" /></td>
				<td align="center"><input type=button value="<%=yyDelete%>" onclick="delrec('<%=alldata(0,rowcounter)%>')" /></td>
			  </tr>
<%		next
	else
%>
			  <tr> 
                <td width="100%" colspan="6" align="center"><br /><strong><%=yyNoDsc%></strong><br />&nbsp;</td>
			  </tr>
<%
	end if
%>
			  <tr> 
                <td width="100%" colspan="6" align="center"><br /><strong><%=yyPOClk%> </strong>&nbsp;&nbsp;<input type="button" value="<%=yyNewDsc%>" onclick="newrec()" /><br />&nbsp;</td>
			  </tr>
			  <tr> 
                <td width="100%" colspan="6" align="center"><br />
                          <a href="admin.asp"><strong><%=yyAdmHom%></strong></a><br />
				<img src="../images/clearpixel.gif" width="300" height="3" alt="" /></td>
			  </tr>
            </table></td>
		  </form>
        </tr>
<% end if
cnn.Close
set rs = nothing
set cnn = nothing
%>
      </table>