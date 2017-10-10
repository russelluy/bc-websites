<%
if storesessionvalue="" then storesessionvalue="virtualstore"
if Session("loggedon") <> storesessionvalue OR disallowlogin=TRUE then response.end
Dim sSQL,rs,alldata,success,cnn,rowcounter,alloptions,errmsg,index,zoneName,foundmatch,upperbound,isWeightBased,hasMultiShip,methodnames(10),hishipvals(10)
success=true
maxshippingmethods=5
alldata=""
Set rs = Server.CreateObject("ADODB.RecordSet")
Set cnn=Server.CreateObject("ADODB.Connection")
cnn.open sDSN
if request.form("posted")="1" then
	for index=1 to 200
		if request.form("id"&index)="1" then
			sSQL = "UPDATE postalzones SET pzName='"&replace(request.form("zon"&index),"'","''")&"' WHERE pzID="&index
			cnn.Execute(sSQL)
		end if
	next
	response.write "<meta http-equiv=""refresh"" content=""3; url=adminzones.asp"">"
elseif request.form("posted")="2" then
	numshipmethods=request.form("numshipmethods")
	zone = request.form("zone")
	cnn.Execute("DELETE FROM zonecharges WHERE zcZone="&request.form("zone"))
	if IsNumeric(Trim(request.form("highweight"))) then
		if cDbl(request.form("highweight")) > 0 then
			sSQL = "INSERT INTO zonecharges (zcZone,zcWeight,zcRate,zcRate2,zcRate3,zcRate4,zcRate5) VALUES ("&zone&","&cStr(0.0-cDbl(request.form("highweight")))
			for index=0 to maxshippingmethods-1
				if IsNumeric(Trim(request.form("highvalue"&index))) then
					sSQL = sSQL & "," & request.form("highvalue"&index)
				else
					sSQL = sSQL & ",0"
				end if
			next
			cnn.Execute(sSQL & ")")
		end if
	end if
	for index=0 to 59
		if IsNumeric(Trim(request.form("weight"&index))) then
			if cDbl(request.form("weight"&index)) > 0 then
				sSQL = "INSERT INTO zonecharges (zcZone,zcWeight,zcRate,zcRatePC,zcRate2,zcRatePC2,zcRate3,zcRatePC3,zcRate4,zcRatePC4,zcRate5,zcRatePC5) VALUES ("&zone&","&request.form("weight"&index)
				for index2=0 to maxshippingmethods-1
					thecharge = Trim(request.form("charge"&index2&"x"&index))
					if IsNumeric(replace(thecharge,"%","")) then
						sSQL = sSQL & "," & replace(thecharge,"%","")
					elseif LCase(thecharge)="x" then
						sSQL = sSQL & ",-99999.0"
					else
						sSQL = sSQL & ",0"
					end if
					if InStr(thecharge, "%") > 0 then sSQL = sSQL & ",1" else sSQL = sSQL & ",0"
				next
				cnn.Execute(sSQL & ")")
			end if
		end if
	next
	sSQL = "UPDATE postalzones SET "
	addcomma=""
	pzFSA = 0
	for index=0 to maxshippingmethods-1
		sSQL = sSQL & addcomma & "pzMethodName" & (index+1) & "='" & Trim(Replace(request.form("methodname"&index),"'","''")) & "'"
		if trim(request.form("methodfsa"&index))="ON" then pzFSA = (pzFSA OR (2 ^ index))
		addcomma=","
	next
	sSQL = sSQL & ",pzFSA=" & pzFSA
	cnn.Execute(sSQL & " WHERE pzID=" & zone)
	response.write "<meta http-equiv=""refresh"" content=""3; url=adminzones.asp"">"
elseif request.querystring("id")<>"" then
	if Trim(request.querystring("shippingmethods"))<>"" then
		sSQL = "UPDATE postalzones SET pzMultiShipping=" & request.querystring("shippingmethods") & " WHERE pzID=" & request.querystring("id")
		cnn.Execute(sSQL)
	end if
	sSQL = "SELECT pzName,pzMultiShipping,pzFSA,pzMethodName1,pzMethodName2,pzMethodName3,pzMethodName4,pzMethodName5 FROM postalzones WHERE pzID="&request.querystring("id")
	rs.Open sSQL,cnn,0,1
	zoneName=""
	if NOT rs.EOF then
		zoneName=rs("pzName")
		hasMultiShip=rs("pzMultiShipping")
		pzFSA=rs("pzFSA")
		for rowcounter=1 to maxshippingmethods
			methodnames(rowcounter-1)=rs("pzMethodName"&rowcounter)
		next
	end if
	rs.Close
	sSQL = "SELECT zcID,zcWeight,zcRate,zcRate2,zcRate3,zcRate4,zcRate5,zcRatePC,zcRatePC2,zcRatePC3,zcRatePC4,zcRatePC5 FROM zonecharges WHERE zcZone="&request.querystring("id")&" ORDER BY zcWeight"
	rs.Open sSQL,cnn,0,1
	if NOT rs.EOF then alldata=rs.getrows
	rs.Close
else
	if request.querystring("oneuszone")="yes" then
		sSQL = "UPDATE admin SET adminUSZones=0"
		cnn.Execute(sSQL)
	end if
	if request.querystring("oneuszone")="no" then
		sSQL = "UPDATE admin SET adminUSZones=1"
		cnn.Execute(sSQL)
	end if
	Application.Lock()
	Application("getadminsettings")=""
	Application.UnLock()
	sSQL = "SELECT pzID,pzName FROM postalzones ORDER BY pzID"
	rs.Open sSQL,cnn,0,1
	if NOT rs.EOF then alldata=rs.getrows
	rs.Close
end if
alreadygotadmin = getadminsettings()
Session.LCID=1033
isWeightBased = (shipType=2 OR shipType=5)
%>
      <table border="0" cellspacing="0" cellpadding="0" width="100%" bgcolor="" align="center">
<% if request.form("posted")="2" AND success then %>
        <tr>
          <td width="100%">
			<table width="100%" border="0" cellspacing="0" cellpadding="3" bgcolor="">
			  <tr> 
                <td width="100%" colspan="2" align="center"><br /><strong><%=yyUpdSuc%></strong><br /><br /><%=yyNowFrd%><br /><br />
                        <%=yyNoAuto%> <A href="adminzones.asp"><strong><%=yyClkHer%></strong></a>.<br />
                        <br />
				<img src="../images/clearpixel.gif" width="300" height="3" alt="" />
                </td>
			  </tr>
			</table></td>
        </tr>
<% elseif request.form("posted")="2" then %>
        <tr>
          <td width="100%">
			<table width="100%" border="0" cellspacing="0" cellpadding="3" bgcolor="">
			  <tr> 
                <td width="100%" colspan="2" align="center"><br /><font color="#FF0000"><strong><%=yyErrUpd%></strong></font><br /><br /><%=errmsg%><br /><br />
				<a href="javascript:history.go(-1)"><strong><%=yyClkBac%></strong></a></td>
			  </tr>
			</table></td>
        </tr>
<% elseif request.querystring("id")<>"" then %>
<script language="javascript" type="text/javascript">
<!--
function formvalidator(theForm)
{
	var emptyentries=false;
<% for index=0 to hasMultiShip %>
	if (theForm.methodname<%=index%>.value == ""){
		alert("<%=yyAllShp%>");
		theForm.methodname<%=index%>.focus();
		return (false);
	}
<% next %>
	var checkOK = "0123456789.";
	var checkStr = theForm.highweight.value;
	var allValid = true;
	for (i = 0;  i < checkStr.length;  i++){
		ch = checkStr.charAt(i);
		for (j = 0;  j < checkOK.length;  j++)
			if (ch == checkOK.charAt(j))
				break;
		if (j == checkOK.length){
			allValid = false;
				break;
		}
	}
	if (!allValid){
		alert("<%=yyDecFld%>");
		theForm.highweight.focus();
		return (false);
	}
	for(index=0; index<<%=maxshippingmethods%>;index++){
		var theobj = eval("theForm.highvalue"+index);
		var checkStr = theobj.value;
		var allValid = true;
		for (i = 0;  i < checkStr.length;  i++){
			ch = checkStr.charAt(i);
			for (j = 0;  j < checkOK.length;  j++)
				if (ch == checkOK.charAt(j))
					break;
			if (j == checkOK.length){
				allValid = false;
					break;
			}
		}
		if (!allValid){
			alert("<%=yyDecFld%>");
			theobj.focus();
			return (false);
		}
	}
	for(index=0;index<60;index++){
		var theobj = eval("theForm.weight"+index);
		var checkStr = theobj.value;
		var allValid = true;
		var hasweight = (theobj.value != "");
		for (i = 0;  i < checkStr.length;  i++){
			ch = checkStr.charAt(i);
			for (j = 0;  j < checkOK.length;  j++)
			  if (ch == checkOK.charAt(j))
				break;
			if (j == checkOK.length){
				allValid = false;
				break;
			}
		}
		if (!allValid){
			alert("<%=yyDecFld%>");
			theobj.focus();
			return (false);
		}
		for(index2=0; index2<=<%=hasMultiShip%>;index2++){
			var theobj = eval("theForm.charge"+index2+"x"+index);
			var checkOK = "0123456789.%";
			var checkStr = theobj.value;
			var allValid = true;
			if(hasweight && checkStr==""){
				emptyentries=true;
				emptyobj=theobj;
			}
			for (i = 0;  i < checkStr.length;  i++){
				ch = checkStr.charAt(i);
				for (j = 0;  j < checkOK.length;  j++)
					if (ch == checkOK.charAt(j))
						break;
				if (j == checkOK.length && checkStr.toLowerCase()!="x"){
					allValid = false;
					break;
				}
			}
			if (!allValid){
				alert("<%=yyDecFld%>");
				theobj.focus();
				return (false);
			}
		}
	}
	if(emptyentries){
		if(!confirm("<%=yyNoMeth%> <%if shipType=5 then response.write yyMaxPri else response.write yyMaxWei%><%=yyNoMet2%> <%=yyNoInt%>\n\n<%=yyOkCan%>")){
			emptyobj.focus();
			return(false);
		}
	}
	return (true);
}
function setnummethods(){
setto=document.forms.mainform.numshipmethods.selectedIndex;
document.location="adminzones.asp?shippingmethods="+setto+"&id=<%=request.querystring("id")%>";
}
//-->
</script>
        <tr>
		  <form name="mainform" method="post" action="adminzones.asp" onsubmit="return formvalidator(this)">
			<td width="100%" align="center">
			<input type="hidden" name="posted" value="2" />
			<input type="hidden" name="zone" value="<%=request.querystring("id")%>" />
            <table width="100%" border="0" cellspacing="0" cellpadding="3" bgcolor="">
			  <tr> 
                <td width="100%" align="center"><strong><%=yyModRul%> <%
				if zoneName<>"" then
					response.write chr(34)&zoneName&chr(34)
				else
					response.write "(unnamed)"
				end if%>.</strong><br />&nbsp;</td>
			  </tr>
			  <tr> 
                <td width="100%" align="center">
				  <font size="1"><%=yyZonUse%> 
					<select name="numshipmethods" size="1" onchange="setnummethods()"><% 
						for rowcounter=1 to 5
							response.write "<option value=""" & rowcounter & """"
							if rowcounter = (hasMultiShip+1) then response.write " selected"
							response.write ">" & rowcounter & "</option>"
						next %></select> <%=yyZonUs2%></font>
				</td>
			  </tr>
			  <tr> 
                <td width="100%" align="center">
				<table width="80%" cellspacing="2" cellpadding="0">
				  <tr>
					<td align="right" width="45%"><%=yyForEv%></td>
					<td width="10%"><input type=text name="highweight" value="<%
				foundmatch=0
				if IsArray(alldata) then
					for rowcounter=0 to UBOUND(alldata,2)
						if alldata(1,rowcounter) < 0 then
							foundmatch=abs(alldata(1,rowcounter))
							for index=0 to maxshippingmethods-1
								hishipvals(index)=alldata(2+index,rowcounter)
							next
						end if
					next
				end if
				response.write foundmatch
				%>" size="5" /></td>
					<td width="45%" align="left"><%=yyAbvHg & " "%> <%if shipType=5 then response.write yyPrice else response.write yyWeigh%>...</td>
				  </tr>
				<%	for index=0 to hasMultiShip %>
				  <tr>
					<td align="right"><%=yyAddExt%></td>
					<td><input type=text name="highvalue<%=index%>" value="<%=hishipvals(index) %>" size="5" /></td><td align="left"><%=yyFor%> <strong><% if methodnames(index)<>"" then response.write methodnames(index) else response.write yyShipMe & " " & index+1%></strong></td>
				  </tr>
				<%	next
					for index=hasMultiShip+1 to maxshippingmethods-1 %>
				  <input type="hidden" name="highvalue<%=index%>" value="<%=hishipvals(index) %>" />
				<%	next %>
				</table>
				</td>
			  </tr>
			  <tr> 
                <td width="100%" align="center">
                  <p><input type="submit" value="<%=yySubmit%>" />&nbsp;&nbsp;<input type="reset" value="<%=yyReset%>" /><br />&nbsp;</p>
                </td>
			  </tr>
			</table>
			<table width="120" border="0" cellspacing="0" cellpadding="1" bgcolor="">
			  <tr>
				<td width="<%=Int(100/(2+hasMultiShip))%>%" align="center">&nbsp;</td>
				<%	for index=0 to hasMultiShip
						response.write "<td width="""&Int(100/(2+hasMultiShip))&"%"" align=""center""><acronym title="""&yyFSApp&"""><strong>"&yyFSA&"</strong></acronym>: <input type=""checkbox"" value=""ON"" name=""methodfsa"&index&""" "&IIfVr((pzFSA AND (2 ^ index)) <> 0,"checked","")&" /></td>" & vbCrLf
					next
					for index=hasMultiShip+1 to maxshippingmethods-1
						response.write "<input type=""hidden"" name=""methodfsa"&index&""" value="""&IIfVr((pzFSA AND (2 ^ index)) <> 0,"ON","")&""" />" & vbCrLf
					next %>
			  </tr>
			  <tr>
				<td align="center"><strong><%if shipType=5 then response.write yyMaxPri else response.write yyMaxWgt%></strong></td>
				<%	for index=0 to hasMultiShip
						response.write "<td align=""center""><input class=""darkborder"" type=""text"" name=""methodname"&index&""" value="""&Replace(methodnames(index)&"","""","&quot;")&""" size=""14"" /></td>" & vbCrLf
					next
					for index=hasMultiShip+1 to maxshippingmethods-1
						response.write "<input type=""hidden"" name=""methodname"&index&""" value="""&Replace(methodnames(index)&"","""","&quot;")&""" />" & vbCrLf
					next %>
			  </tr>
<%
	rowcounter=0
	index=0
	if IsArray(alldata) then
		upperbound = UBOUND(alldata,2)
	else
		upperbound = -1
	end if
	do while index < 60
		if rowcounter <= upperbound then
			if alldata(1,rowcounter) > 0 then
%>
			  <tr>
				<td align="center"><input class="darkborder" type=text name="weight<%=index%>" value="<%=alldata(1,rowcounter)%>" size="10" /></td>
				<%	for index2=0 to maxshippingmethods-1
						if index2 <= hasMultiShip then
							response.write "<td align=""center""><input type=""text"" name=""charge"&index2&"x"&index&""" value="""&IIfVr(alldata(2+index2,rowcounter)<>-99999,alldata(2+index2,rowcounter)&IIfVr(cint(alldata(7+index2,rowcounter))<>0,"%",""),"x")&""" size=""14"" /></td>" & vbCrLf
						else
							response.write "<input type=""hidden"" name=""charge"&index2&"x"&index&""" value="""&alldata(2+index2,rowcounter)&""" />"
						end if
					next %>
			  </tr>
<%
				index=index+1
			end if
		else
%>
			  <tr>
				<td align="center"><input class="darkborder" type=text name="weight<%=index%>" value="" size="10" /></td>
				<%	for index2=0 to maxshippingmethods-1
						if index2 <= hasMultiShip then
							response.write "<td align=""center""><input type=""text"" name=""charge"&index2&"x"&index&""" size=""14"" /></td>" & vbCrLf
						end if
					next %>
			  </tr>
<%
			index=index+1
		end if
		rowcounter=rowcounter+1
	loop
%>
			  <tr> 
                <td width="100%" colspan="<%=2+hasMultiShip%>" align="center">
                  <p><input type="submit" value="<%=yySubmit%>" />&nbsp;&nbsp;<input type="reset" value="<%=yyReset%>" /><br />&nbsp;</p>
                </td>
			  </tr>
			  <tr> 
                <td width="100%" colspan="<%=2+hasMultiShip%>" align="center"><br />
                          <a href="admin.asp"><strong><%=yyAdmHom%></strong></a><br />
				<img src="../images/clearpixel.gif" width="300" height="3" alt="" /></td>
			  </tr>
            </table>
		  </td>
		  </form>
        </tr>
<% elseif request.form("posted")="1" AND success then %>
        <tr>
          <td width="100%">
			<table width="100%" border="0" cellspacing="0" cellpadding="3" bgcolor="">
			  <tr> 
                <td width="100%" colspan="2" align="center"><br /><strong><%=yyUpdSuc%></strong><br /><br /><%=yyNowFrd%><br /><br />
                        <%=yyNoAuto%> <A href="adminzones.asp"><strong><%=yyClkHer%></strong></a>.<br />
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
                <td width="100%" colspan="2" align="center"><br /><font color="#FF0000"><strong><%=yyErrUpd%></strong></font><br /><br /><%=errmsg%><br /><br />
				<a href="javascript:history.go(-1)"><strong><%=yyClkBac%></strong></a></td>
			  </tr>
			</table></td>
        </tr>
<% else %>
        <tr>
		  <form name="mainform" method="post" action="adminzones.asp">
		  <td width="100%">
			<input type="hidden" name="posted" value="1" />
            <table width="100%" border="0" cellspacing="0" cellpadding="3" bgcolor="">
			  <tr> 
                <td width="100%" <%if splitUSZones then response.write "colspan='2'"%> align="center"><strong><%=yyModPZo%></strong><br />&nbsp;</td>
			  </tr>
			  <tr> 
                <td width="100%" <%if splitUSZones then response.write "colspan='2'"%> align="left">
				  <ul>
				  <% if not isWeightBased then %>
					<li><font size="1"><%=yyPZEx1%> <a href="adminmain.asp"><strong><%=yyClkHer%></strong></a>.</font></li>
				  <% end if %>
				  <% if splitUSZones then %>
					<li><font size="1"><%=yyPZEx2%> <a href="adminzones.asp?oneuszone=yes"><strong><%=yyClkHer%></strong></a>.</font></li>
				  <% else %>
				    <li><font size="1"><%=yyPZEx3%> <a href="adminzones.asp?oneuszone=no"><strong><%=yyClkHer%></strong></a>.</font></li>
				  <% end if %>
					<li><font size="1"><%=yyPZEx4%></font></li>
				  </ul>
				</td>
			  </tr>
			  <tr>
				<td valign="top">
				  <table width="100%" cellspacing="1" cellpadding="1" border="0">
					<tr> 
					  <td width="100%" colspan="3" align="center"><strong><%=yyPZWor%></strong><br /><hr width="70%" /></td>
					</tr>
					 <tr>
					  <td width="40%" align=right>&nbsp;</td>
					  <td width="20%" align=center><strong><%=yyPZNam%></strong></td>
					  <td width="40%" align=left><strong><%=yyPZRul%></strong></td>
					</tr>
<%
	for rowcounter=0 to UBOUND(alldata,2)
		if alldata(0,rowcounter)<100 then ' First 100 are for world zones
%>
					<tr>
					  <td align=right><strong><%=alldata(0,rowcounter)%> : <input type="hidden" name="id<%=alldata(0,rowcounter)%>" value="1" /></strong></td>
					  <td align=center><input type=text name="zon<%=alldata(0,rowcounter)%>" value="<%=alldata(1,rowcounter)%>" size="20" /></td>
					  <td align=left><% if Trim(alldata(1,rowcounter)) <> "" then %><a href="adminzones.asp?id=<%=alldata(0,rowcounter)%>"><strong><%=yyEdRul%></strong></a><% else %>&nbsp;<% end if %></td>
					</tr>
<%
		end if
	next
%>
				  </table>
				</td>

<%
	if splitUSZones then
%>
				<td width="50%" valign="top">
				  <table width="100%" cellspacing="1" cellpadding="1" border="0">
					<tr> 
					  <td width="100%" colspan="3" align="center"><strong><%=yyPZSta%></strong><br /><hr width="70%" /></td>
					</tr>
					 <tr>
					  <td width="40%" align=right>&nbsp;</td>
					  <td width="20%" align=center><strong><%=yyPZNam%></strong></td>
					  <td width="40%" align=left><strong><%=yyPZRul%></strong></td>
					</tr>
<%
		index = 0
		for rowcounter=0 to UBOUND(alldata,2)
			if alldata(0,rowcounter)>100 then ' Next 100 are for world zones
				index=index+1
%>
					<tr>
					  <td align=right><strong><%=alldata(0,rowcounter)-100%> : <input type="hidden" name="id<%=alldata(0,rowcounter)%>" value="1" /></strong></td>
					  <td align=center><input type=text name="zon<%=alldata(0,rowcounter)%>" value="<%=alldata(1,rowcounter)%>" size="20" /></td>
					  <td align=left><% if Trim(alldata(1,rowcounter)) <> "" then %><a href="adminzones.asp?id=<%=alldata(0,rowcounter)%>"><strong><%=yyEdRul%></strong></a><% else %>&nbsp;<% end if %></td>
					</tr>
<%
			end if
		next
%>
				  </table>
				</td>
<%
	end if
%>			  </tr>
			  <tr> 
                <td width="100%" <%if splitUSZones then response.write "colspan='2'"%> align="center">
                  <p><input type="submit" value="<%=yySubmit%>" />&nbsp;&nbsp;<input type="reset" value="<%=yyReset%>" /><br />&nbsp;</p>
                </td>
			  </tr>
			  <tr> 
                <td width="100%" <%if splitUSZones then response.write "colspan='2'"%> align="center"><br />
                          <a href="admin.asp"><strong><%=yyAdmHom%></strong></a><br />
				<img src="../images/clearpixel.gif" width="300" height="3" alt="" /></td>
			  </tr>
            </table>
		  </td>
		  </form>
        </tr>
<% end if
cnn.Close
set rs = nothing
set cnn = nothing
%>
      </table>