<%
if storesessionvalue="" then storesessionvalue="virtualstore"
if Session("loggedon") <> storesessionvalue OR disallowlogin=TRUE then response.end
Dim sSQL,rs,alldata,success,cnn,rowcounter,netnav,errmsg,aOption,index,iID,bOption,fieldDims
success=true
Set rs = Server.CreateObject("ADODB.RecordSet")
Set cnn=Server.CreateObject("ADODB.Connection")
cnn.open sDSN
alreadygotadmin = getadminsettings()
Session.LCID = 1033
sSQL = ""
alldata=""
if request.form("posted")="1" then
	if request.form("act")="delete" then
		sSQL = "SELECT poID,poProdID FROM prodoptions WHERE poOptionGroup=" & request.form("id")
		rs.Open sSQL,cnn,0,1
		index=0
		if NOT rs.EOF then
			success=false
			errmsg = yyPOErr & "<br />"
			errmsg = errmsg & yyPOUse & "<br />(" & rs("poProdID") & ")"
		end if
		rs.Close
		if success then
			sSQL = "DELETE FROM options WHERE optGroup=" & request.form("id")
			cnn.Execute(sSQL)
			sSQL = "DELETE FROM optiongroup WHERE optGrpID=" & request.form("id")
			cnn.Execute(sSQL)
			response.write "<meta http-equiv=""refresh"" content=""2; url=adminprodopts.asp"">"
		end if
	elseif request.form("act")="domodify" OR request.form("act")="doaddnew" then
		sSQL = ""
		Redim aOption(10,maxprodopts)
		bOption=false
		optFlags = 0
		if Request.Form("pricepercent")="1" then optFlags=1
		if Request.Form("weightpercent")="1" then optFlags=optFlags + 2
		if Request.Form("singleline")="1" then optFlags=optFlags + 4
		if Request.Form("optdefault")<>"" then optDefault=int(Request.Form("optdefault")) else optDefault=-1
		for rowcounter=0 to maxprodopts-1
			if Trim(request.form("opt"&rowcounter))<>"" then bOption=true
			aOption(0,rowcounter)=Replace(Trim(request.form("opt"&rowcounter)),"'","''")
			for index=2 to adminlanguages+1
				if (adminlangsettings AND 32)=32 then aOption(5+index,rowcounter)=Replace(Trim(request.form("opl"&index&"x"&rowcounter)),"'","''")
			next
			if IsNumeric(Trim(request.form("pri"&rowcounter))) then
				aOption(1,rowcounter)=Trim(request.form("pri"&rowcounter))
			else
				aOption(1,rowcounter)=0
			end if
			if IsNumeric(Trim(request.form("wsp"&rowcounter))) then
				aOption(4,rowcounter)=Trim(request.form("wsp"&rowcounter))
			else
				aOption(4,rowcounter)=0
			end if
			if IsNumeric(Trim(request.form("wei"&rowcounter))) then
				aOption(2,rowcounter)=Trim(request.form("wei"&rowcounter))
			else
				aOption(2,rowcounter)=0
			end if
			if IsNumeric(Trim(request.form("optStock"&rowcounter))) then
				aOption(3,rowcounter)=Trim(request.form("optStock"&rowcounter))
			else
				aOption(3,rowcounter)=0
			end if
			aOption(5,rowcounter)=Replace(Trim(request.form("regexp"&rowcounter)),"'","''")
			aOption(6,rowcounter)=Trim(request.form("orig"&rowcounter))
		next
		if (trim(request.form("secname"))="" OR NOT bOption) AND Request.Form("optType")<>"3" then
			success=false
			errmsg = yyPOErr & "<br />"
			errmsg = errmsg & yyPOOne
		else
			if Request.Form("optType")="3" then ' Text option
				fieldDims = Trim(request.form("pri0"))&"."
				if Int(request.form("fieldheight")) < 10 then fieldDims = fieldDims & "0"
				fieldDims = fieldDims & Trim(request.form("fieldheight"))
				if request.form("act")="doaddnew" then
					rs.Open "optiongroup",cnn,1,3,&H0002
					rs.AddNew
					rs.Fields("optGrpName")	= trim(request.form("secname"))
					for index=2 to adminlanguages+1
						if (adminlangsettings AND 16)=16 then rs.Fields("optGrpName" & index) = trim(request.form("secname" & index))
					next
					if Request.Form("forceselec")="ON" then
						rs.Fields("optType") = Request.Form("optType")
					else
						rs.Fields("optType") = 0 - Int(Request.Form("optType"))
					end if
					if trim(request.form("workingname"))="" then
						rs.Fields("optGrpWorkingName") = trim(request.form("secname"))
					else
						rs.Fields("optGrpWorkingName") = trim(request.form("workingname"))
					end if
					rs.Fields("optFlags") = optFlags
					rs.Update
					if mysqlserver=true then
						rs.Close
						rs.Open "SELECT LAST_INSERT_ID() AS lstIns",cnn,0,1
						iID = rs("lstIns")
					else
						iID  = rs.Fields("optGrpID")
					end if
					rs.Close
					sSQL = "INSERT INTO options (optGroup,optName,optPriceDiff"
					for index=2 to adminlanguages+1
						if (adminlangsettings AND 16)=16 then sSQL = sSQL & ",optName" & index
					next
					sSQL = sSQL & ",optWeightDiff) VALUES ("&iID&",'"&Replace(Trim(request.form("opt0")),"'","''")&"',"&fieldDims
					for index=2 to adminlanguages+1
						if (adminlangsettings AND 16)=16 then sSQL = sSQL & ",'" & Replace(trim(request.form("opl" & index & "x0")),"'","''")&"'"
					next
					sSQL = sSQL & ",0)"
					cnn.Execute(sSQL)
				else
					iID = request.form("id")
					sSQL = "UPDATE optiongroup SET optGrpName='"&Replace(trim(request.form("secname")),"'","''")&"'"
					for index=2 to adminlanguages+1
						if (adminlangsettings AND 16)=16 then sSQL = sSQL & ",optGrpName" & index & "='"& Replace(trim(request.form("secname" & index)),"'","''")&"'"
					next
					if Request.Form("forceselec")="ON" then
						sSQL = sSQL & ",optType=" & Request.Form("optType")
					else
						sSQL = sSQL & ",optType=" & (0 - Int(Request.Form("optType")))
					end if
					sSQL = sSQL & ",optFlags=" & optFlags
					if trim(request.form("workingname"))="" then
						sSQL = sSQL & ",optGrpWorkingName='"& Replace(trim(request.form("secname")),"'","''")&"' "
					else
						sSQL = sSQL & ",optGrpWorkingName='"& Replace(trim(request.form("workingname")),"'","''")&"' "
					end if
					sSQL = sSQL & "WHERE optGrpID="&iID
					cnn.Execute(sSQL)
					sSQL = "UPDATE options SET optName='"&Replace(Trim(request.form("opt0")),"'","''")&"',optPriceDiff="&fieldDims
					for index=2 to adminlanguages+1
						if (adminlangsettings AND 16)=16 then sSQL = sSQL & ",optName" & index & "='"& Replace(trim(request.form("opl" & index & "x0")),"'","''")&"'"
					next
					sSQL = sSQL & " WHERE optGroup="&iID
					cnn.Execute(sSQL)
				end if
			else ' Non-text Option
				if request.form("act")="doaddnew" then
					rs.Open "optiongroup",cnn,1,3,&H0002
					rs.AddNew
					rs.Fields("optGrpName")	= trim(request.form("secname"))
					for index=2 to adminlanguages+1
						if (adminlangsettings AND 16)=16 then rs.Fields("optGrpName" & index) = trim(request.form("secname" & index))
					next
					if Request.Form("forceselec")="ON" then
						rs.Fields("optType") = Request.Form("optType")
					else
						rs.Fields("optType") = 0 - Int(Request.Form("optType"))
					end if
					if trim(request.form("workingname"))="" then
						rs.Fields("optGrpWorkingName") = trim(request.form("secname"))
					else
						rs.Fields("optGrpWorkingName") = trim(request.form("workingname"))
					end if
					rs.Fields("optFlags") = optFlags
					rs.Fields("optGrpSelect") = IIfVr(trim(request.form("optgrpselect"))="1",1,0)
					rs.Update
					if mysqlserver=true then
						rs.Close
						rs.Open "SELECT LAST_INSERT_ID() AS lstIns",cnn,0,1
						iID = rs("lstIns")
					else
						iID  = rs.Fields("optGrpID")
					end if
					rs.Close
				else
					iID = request.form("id")
					sSQL = "UPDATE optiongroup SET optGrpName='"& Replace(trim(request.form("secname")),"'","''")&"'"
					for index=2 to adminlanguages+1
						if (adminlangsettings AND 16)=16 then sSQL = sSQL & ",optGrpName" & index & "='"& Replace(trim(request.form("secname" & index)),"'","''")&"'"
					next
					if Request.Form("forceselec")="ON" then
						sSQL = sSQL & ",optType=" & Request.Form("optType")
					else
						sSQL = sSQL & ",optType=" & (0 - Int(Request.Form("optType")))
					end if
					sSQL = sSQL & ",optFlags=" & optFlags
					sSQL = sSQL & ",optGrpSelect=" & IIfVr(trim(request.form("optgrpselect"))="1",1,0)
					if trim(request.form("workingname"))="" then
						sSQL = sSQL & ",optGrpWorkingName='"& Replace(trim(request.form("secname")),"'","''")&"' "
					else
						sSQL = sSQL & ",optGrpWorkingName='"& Replace(trim(request.form("workingname")),"'","''")&"' "
					end if
					sSQL = sSQL & "WHERE optGrpID="&iID
					cnn.Execute(sSQL)
				end if
				for rowcounter=0 to UBOUND(aOption,2)
					if Trim(aOption(0,rowcounter)) <> "" then
						if aOption(6,rowcounter) <> "" then
							sSQL = "UPDATE options SET optName='"&aOption(0,rowcounter)&"',optRegExp='"&aOption(5,rowcounter)&"',optPriceDiff="&aOption(1,rowcounter)&",optWeightDiff="&aOption(2,rowcounter)&",optStock="&aOption(3,rowcounter)
							if wholesaleoptionpricediff=TRUE then sSQL = sSQL & ",optWholesalePriceDiff="&aOption(4,rowcounter)
							for index=2 to adminlanguages+1
								if (adminlangsettings AND 32)=32 then sSQL = sSQL & ",optName" & index & "='" & aOption(5+index,rowcounter) & "'"
							next
							sSQL = sSQL & ",optDefault=" & IIfVr(rowcounter=optDefault,"1","0")
							sSQL = sSQL & " WHERE optID=" & aOption(6,rowcounter)
							cnn.Execute(sSQL)
						else
							sSQL = "INSERT INTO options (optGroup,optName,optRegExp,optPriceDiff,optWeightDiff,optStock,optDefault"
							if wholesaleoptionpricediff=TRUE then sSQL = sSQL & ",optWholesalePriceDiff"
							for index=2 to adminlanguages+1
								if (adminlangsettings AND 32)=32 then sSQL = sSQL & ",optName" & index
							next
							sSQL = sSQL & ") VALUES ("&iID&",'"&aOption(0,rowcounter)&"','"&aOption(5,rowcounter)&"',"&aOption(1,rowcounter)&","&aOption(2,rowcounter)&","&aOption(3,rowcounter)&","&IIfVr(rowcounter=optDefault,"1","0")
							if wholesaleoptionpricediff=TRUE then sSQL = sSQL & "," & aOption(4,rowcounter)
							for index=2 to adminlanguages+1
								if (adminlangsettings AND 32)=32 then sSQL = sSQL & ",'" & aOption(5+index,rowcounter) & "'"
							next
							sSQL = sSQL & ")"
							cnn.Execute(sSQL)
						end if
					else
						if aOption(6,rowcounter) <> "" then
							cnn.Execute("DELETE FROM options WHERE optID=" & aOption(6,rowcounter))
						end if
					end if
				next
			end if
		end if
		response.write "<meta http-equiv=""refresh"" content=""2; url=adminprodopts.asp"">"
	end if
end if
%>
<script language="javascript" type="text/javascript">
<!--
function formvalidator(theForm){
  if (theForm.secname.value == "")
  {
    alert("<%=yyPlsEntr%> \"<%=yyPOName%>\".");
    theForm.secname.focus();
    return (false);
  }
  return (true);
}
function changeunits(){
	var nopercentchar="&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;";
	for(index=0;index<<%=maxprodopts%>;index++){
		wel = document.getElementById("wunitspan" + index);
		pel = document.getElementById("punitspan" + index);
<% if wholesaleoptionpricediff=TRUE then %>
		wspel = document.getElementById("pwspunitspan" + index);
		if(document.forms.mainform.pricepercent.checked){
			wspel.innerHTML='&nbsp;%&nbsp;';
		}else{
			wspel.innerHTML='&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;';
		}
<% end if %>
		if(document.forms.mainform.weightpercent.checked){
			wel.innerHTML='&nbsp;%&nbsp;';
		}else{
			wel.innerHTML=nopercentchar;
		}
		if(document.forms.mainform.pricepercent.checked){
			pel.innerHTML='&nbsp;%&nbsp;';
		}else{
			pel.innerHTML='&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;';
		}
	}
}
//-->
</script>
      <table border="0" cellspacing="<%=maintablespacing%>" cellpadding="<%=maintablepadding%>" width="<%=maintablewidth%>" bgcolor="<%=maintablebg%>" align="center">
<% if request.form("posted")="1" AND (request.form("act")="modify" OR request.form("act")="clone" OR request.form("act")="addnew") then
		iscloning = (request.form("act")="clone")
		if request.form("act")="modify" OR iscloning then
			doaddnew = false
			sSQL = "SELECT optID,optName,optGrpName,optGrpWorkingName,optPriceDiff,optType,optWeightDiff,optFlags,optStock,optWholesalePriceDiff,optRegExp,optName2,optName3,optGrpName2,optGrpName3,optDefault,optGrpSelect FROM options INNER JOIN optiongroup ON optiongroup.optGrpID=options.optGroup WHERE optGroup="&Request.Form("id")&" ORDER BY optID"
			rs.Open sSQL,cnn,0,1
			alldata=rs.getrows
			rs.Close
			optName = alldata(1,0)
			optGrpName = alldata(2,0)
			optGrpWorkingName = alldata(3,0)
			optPriceDiff = alldata(4,0)
			optType = alldata(5,0)
			optWeightDiff = alldata(6,0)
			optFlags = alldata(7,0)
			optStock = alldata(8,0)
			optWholesalePriceDiff = alldata(9,0)
			optName2 = alldata(11,0)
			optName3 = alldata(12,0)
			optGrpName2 = alldata(13,0)
			optGrpName3 = alldata(14,0)
			optDefault = alldata(15,0)
			optGrpSelect = alldata(16,0)
			maxoptnumber = UBOUND(alldata,2)
		else
			doaddnew = true
			optName = ""
			optGrpName = ""
			optGrpWorkingName = ""
			optPriceDiff = 15
			optType = Int(Request.Form("optType"))
			optWeightDiff = ""
			optFlags = 0
			optStock = ""
			optWholesalePriceDiff = ""
			optName2 = ""
			optName3 = ""
			optGrpName2 = ""
			optGrpName3 = ""
			optDefault = ""
			optGrpSelect = 1
			maxoptnumber = -1
		end if
%>
        <tr>
		  <form name="mainform" method="post" action="adminprodopts.asp" onsubmit="return formvalidator(this)">
			<td width="100%" align="center">
			<input type="hidden" name="posted" value="1" />
			<%	if iscloning OR request.form("act")="addnew" then %>
			<input type="hidden" name="act" value="doaddnew" />
			<%	else %>
			<input type="hidden" name="act" value="domodify" />
			<input type="hidden" name="id" value="<%=Request.Form("id")%>" />
			<%	end if
				if abs(optType)=3 then response.write "<input type=""hidden"" name=""optType"" value=""3"" />" %>
            <table width="100%" border="0" cellspacing="0" cellpadding="3" bgcolor="">
<% if abs(optType)=3 then ' Text option %>
			  <tr> 
                <td width="100%" colspan="3" align="center"><strong><%=yyPOAdm%></strong><br />&nbsp;</td>
			  </tr>
			  <tr>
				<td width="30%" align="center"><p><strong><%=yyPOName%></strong><br />
				  <input type="text" name="secname" size="30" value="<%=Replace(optGrpName,"""","&quot;")%>" /><br /><%
				for index=2 to adminlanguages+1
					if (adminlangsettings AND 16)=16 then
			%><strong><%=yyPOName & " " & index%></strong><br />
				  <input type="text" name="secname<%=index%>" size="30" value="<% execute("response.write Replace(optGrpName" & index & "&"""",chr(34),""&quot;"")")%>" /><br /><%
					end if
				next %></p>
				  <p><strong><%=yyWrkNam%></strong><br />
				  <input type="text" name="workingname" size="30" value="<%=Replace(optGrpWorkingName,"""","&quot;")%>" /></p>
                </td>
				<td width="30%" align="center"><p><strong><%=yyDefTxt%></strong><br />
				<input type="text" name="opt0" size="25" value="<%=Replace(optName,"""","&quot;")%>" /><br /><%
				for index=2 to adminlanguages+1
					if (adminlangsettings AND 16)=16 then
			%><strong><%=yyDefTxt & " " & index%></strong><br />
				  <input type="text" name="opl<%=index%>x0" size="25" value="<% execute("response.write Replace(optName" & index & "&"""",chr(34),""&quot;"")")%>" /><br /><%
					end if
				next %></p>
				<p>&nbsp;<br /><input type="checkbox" name="forceselec" value="ON" <% if optType > 0 then response.write "checked"%> /> <strong><%=yyForSel%></strong></p>
                </td>
				<td width="40%" align="center"><p><strong><%=yyFldWdt%></strong><br />
				<select name="pri0" size="1">
				<%
					for rowcounter=1 to 35
						response.write "<option value='"&rowcounter&"'"
						if rowcounter=Int(optPriceDiff) then response.write " selected"
						response.write ">&nbsp; "&rowcounter&" </option>"&vbCrLf
					next
				%>
				</select></p>
				<p><strong><%=yyFldHgt%></strong><br />
				<select name="fieldheight" size="1">
				<%
					Dim fieldHeight
					fieldHeight = cInt((cDbl(optPriceDiff)-Int(optPriceDiff))*100.0)
					for rowcounter=1 to 15
						response.write "<option value='"&rowcounter&"'"
						if rowcounter=fieldHeight then response.write " selected"
						response.write ">&nbsp; "&rowcounter&" </option>"&vbCrLf
					next
				%>
				</select></p>
				</td>
			  </tr>
			  <tr>
				<td colspan="3" align="left">
				  <ul>
				  <li><font size="1"><%=yyPOEx1%></li>
				  <li><font size="1"><%=yyPOEx2%></li>
				  <li><font size="1"><%=yyPOEx3%></li>
				  </ul>
                </td>
			  </tr>
<% else %>
			  <tr>
				<td width="30%" align="center">
				  <table border="0" cellspacing="0" cellpadding="3" bgcolor="">
				  <tr><td align="right"><strong><%=replace(yyPOName," ","&nbsp;")%></strong></td><td align="left" colspan="3">
				  <input type="text" name="secname" size="30" value="<%=Replace(optGrpName,"""","&quot;")%>" /></td></tr>
<%				for index=2 to adminlanguages+1
					if (adminlangsettings AND 16)=16 then
			%><tr><td align="right"><strong><%=replace(yyPOName & " " & index," ","&nbsp;")%></strong></td><td align="left" colspan="3">
				  <input type="text" name="secname<%=index%>" size="30" value="<% execute("response.write Replace(optGrpName" & index & "&"""",chr(34),""&quot;"")")%>" /></td></tr><%
					end if
				next %>
				  <tr><td align="right"><strong><%=replace(yyWrkNam," ","&nbsp;")%></strong></td><td align="left" colspan="3"><input type="text" name="workingname" size="30" value="<%=Replace(optGrpWorkingName,"""","&quot;")%>" /></td></tr>
				  <tr><td align="right"><strong><%=replace(yyOptSty," ","&nbsp;")%></strong></td><td align="left" colspan="3"><select name="optType" size="1"><option value="2">Drop down menu</option><option value="1"<% if abs(optType)=1 then response.write " selected"%>>Radio Buttons</option></select></td></tr>
				  <tr><td align="right"><strong><%=replace(yyForSel," ","&nbsp;")%></strong></td><td align="left"><input type="checkbox" name="forceselec" value="ON" <% if optType > 0 then response.write "checked"%> />&nbsp;</td><td align="right">&nbsp;<input type="radio" name="optdefault" value="" /></td></td><td align="left"><strong><%=replace(yyNoDefa," ","&nbsp;")%></strong></tr>
				  <tr><td align="right"><strong><%=replace(yySinLin," ","&nbsp;")%></strong></td><td align="left"><input type="checkbox" name="singleline" value="1" <% if (optFlags AND 4) = 4 then response.write "checked"%> /></td><td align="right"><input type="checkbox" name="optgrpselect" value="1" <% if cint(optGrpSelect)<>0 then response.write "checked"%> /></td><td align="left"><strong><%=replace(yyPlsSLi," ","&nbsp;")%></strong></td></tr>
				  </table>
                </td>
				<td colspan="2" align="left">
				  <p align="center"><strong><%=yyPOAdm%></strong></p>
				  <ul>
				  <li><font size="1"><%=yyPOEx1%></font></li>
				  <li><font size="1"><%=yyPOEx4%></font></li>
				  <li><font size="1"><%=yyPOEx5%></font></li>
				  <% if useStockManagement then %>
				  <li><font size="1"><%=yyPOEx6%></font></li>
				  <% end if %>
				  </ul>
                </td>
			  </tr>
			</table>
			<table width="500" border="0" cellspacing="0" cellpadding="3" bgcolor="">
			  <tr>
				<td><strong><%=yyDefaul%></strong></td>
				<td width="3%" align="center">&nbsp;</td>
				<td align="center"><strong><%=yyPOOpts%></strong></td>
				<td width="3%" align="center">&nbsp;</td>
<%				for index=2 to adminlanguages+1
					if (adminlangsettings AND 32)=32 then
			%><td align="center"><strong><%=yyPOOpts & " " & index%></strong></td>
				<td width="3%" align="center">&nbsp;</td><%
					end if
				next %>
				<td align="center" nowrap><strong><%=yyPOPrDf%>&nbsp;%<input class="noborder" type="checkbox" name="pricepercent" value="1" onclick="javascript:changeunits();" <% if (optFlags AND 1) = 1 then response.write "checked"%> /></strong></td>
				<td width="3%" align="center">&nbsp;</td>
				<% if wholesaleoptionpricediff=TRUE then %>
				<td align="center" nowrap><strong><%=yyWhoPri%></strong></td>
				<td width="3%" align="center">&nbsp;</td>
				<% end if %>
				<td align="center" nowrap><strong><%=yyPOWtDf%>&nbsp;%<input class="noborder" type="checkbox" name="weightpercent" value="1" onclick="javascript:changeunits();" <% if (optFlags AND 2) = 2 then response.write "checked"%> /></strong></td>
				<% if useStockManagement then %>
				<td width="3%" align="center">&nbsp;</td>
				<td align="center" nowrap><strong><%=yyStkLvl%></strong></td>
				<% end if %>
				<td width="3%" align="center">&nbsp;</td>
				<td align="center" nowrap><strong>Alt Prod ID</strong></td>
			  </tr>
<%			if (optFlags AND 1) = 1 then pdUnits="&nbsp;%&nbsp;" else pdUnits="&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
			if (optFlags AND 2) = 2 then wdUnits="&nbsp;%&nbsp;" else wdUnits="&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
			for rowcounter=0 to maxprodopts-1 %>
			  <tr>
				<td><input type="radio" name="optdefault" value="<%=rowcounter%>" <% if rowcounter <= maxoptnumber then if cint(alldata(15,rowcounter))<>0 then response.write "checked"%> /></td>
				<td align="center"><strong>&raquo;</strong></td>
				<td align="center"><%
					if rowcounter <= maxoptnumber AND NOT iscloning then response.write "<input type=""hidden"" name=""orig" & rowcounter & """ value=""" & alldata(0,rowcounter) & """>"
					response.write "<input type=""text"" name=""opt"&rowcounter&""" size=""20"" value="""
					if rowcounter <= maxoptnumber then response.write Replace(alldata(1,rowcounter)&"","""", "&quot;")
					response.write """ /><br />"&vbCrLf
				%></td>
				<td align="center"><strong>&raquo;</strong></td>

<%				for index=2 to adminlanguages+1
					if (adminlangsettings AND 32)=32 then
			%><td align="center"><%
					response.write "<input type=""text"" name=""opl"&index&"x"&rowcounter&""" size=""20"" value="""
					if rowcounter <= maxoptnumber then response.write Replace(alldata(9+index,rowcounter)&"","""", "&quot;")
					response.write """ /><br />"&vbCrLf
				%></td>
				<td align="center"><strong>&raquo;</strong></td><%
					end if
				next %>

				<td align="center"><%
					response.write "&nbsp;&nbsp;&nbsp;&nbsp;<input type='text' name='pri"&rowcounter&"' size='5' value='"
					if rowcounter <= maxoptnumber then response.write alldata(4,rowcounter)
					response.write "' /><span name=""punitspan"&rowcounter&""" id=""punitspan"&rowcounter&""">"&pdUnits&"</span><br />"&vbCrLf
				%></td>
				<td align="center"><strong>&raquo;</strong></td>
				<% if wholesaleoptionpricediff=TRUE then %>
				<td align="center"><%
					response.write "&nbsp;&nbsp;&nbsp;&nbsp;<input type='text' name='wsp"&rowcounter&"' size='5' value='"
					if rowcounter <= maxoptnumber then response.write alldata(9,rowcounter)
					response.write "' /><span name=""pwspunitspan"&rowcounter&""" id=""pwspunitspan"&rowcounter&""">"&pdUnits&"</span><br />"&vbCrLf
				%></td>
				<td align="center"><strong>&raquo;</strong></td>
				<% end if %>
				<td align="center" nowrap><%
					response.write "&nbsp;&nbsp;&nbsp;&nbsp;<input type='text' name='wei"&rowcounter&"' size='5' value='"
					if rowcounter <= maxoptnumber then response.write alldata(6,rowcounter)
					response.write "' /><span name=""wunitspan"&rowcounter&""" id=""wunitspan"&rowcounter&""">"&wdUnits&"</span><br />"&vbCrLf
				%></td>
				<%	if useStockManagement then %>
				<td align="center"><strong>&raquo;</strong></td>
				<td align="center"><input type="text" name="optStock<%=rowcounter%>" size="4" value="<% if rowcounter <= maxoptnumber then response.write alldata(8,rowcounter) %>" /></td>
				<%	else
						if rowcounter <= maxoptnumber then %><input type="hidden" name="optStock<%=rowcounter%>" value="<%=alldata(8,rowcounter)%>" /><% end if
					end if %>
				<td align="center"><strong>&raquo;</strong></td>
				<td align="center"><input type="text" name="regexp<%=rowcounter%>" size="12" value="<% if rowcounter <= maxoptnumber then response.write alldata(10,rowcounter) %>" /></td>
			  </tr>
<%			next %>
			</table>
			<table width="100%" border="0" cellspacing="0" cellpadding="3" bgcolor="">
<% end if %>
			  <tr>
                <td width="100%" colspan="3" align="center"><br /><input type="submit" value="<%=yySubmit%>" /><br />&nbsp;</td>
			  </tr>
			  <tr> 
                <td width="100%" colspan="3" align="center"><br />
                          <a href="admin.asp"><strong><%=yyAdmHom%></strong></a><br />
                          &nbsp;</td>
			  </tr>
            </table></td>
		  </form>
        </tr>
<% elseif request.form("posted")="1" AND success then %>
        <tr>
          <td width="100%">
			<table width="100%" border="0" cellspacing="0" cellpadding="3" bgcolor="">
			  <tr> 
                <td width="100%" colspan="2" align="center"><br /><strong><%=yyUpdSuc%></strong><br /><br /><%=yyNowFrd%><br /><br />
                        <%=yyNoAuto%> <A href="adminprodopts.asp"><strong><%=yyClkHer%></strong></a>.<br />
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
		sSQL = "SELECT optGrpID,optGrpName,optGrpWorkingName FROM optiongroup ORDER BY optGrpName,optGrpWorkingName"
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
function clone(id) {
	document.mainform.id.value = id;
	document.mainform.act.value = "clone";
	document.mainform.submit();
}
function newtextrec(id) {
	document.mainform.id.value = id;
	document.mainform.act.value = "addnew";
	document.mainform.optType.value = "3";
	document.mainform.submit();
}
function newrec(id) {
	document.mainform.id.value = id;
	document.mainform.act.value = "addnew";
	document.mainform.optType.value = "2";
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
		<form name="mainform" method="post" action="adminprodopts.asp">
		  <td width="100%">
			<input type="hidden" name="posted" value="1" />
			<input type="hidden" name="act" value="xxxxx" />
			<input type="hidden" name="id" value="xxxxx" />
			<input type="hidden" name="optType" value="xxxxx" />
            <table width="100%" border="0" cellspacing="0" cellpadding="1" bgcolor="">
			  <tr> 
                <td width="100%" colspan="5" align="center"><strong><%=yyPOAdm%></strong><br />&nbsp;</td>
			  </tr>
			  <tr>
				<td width="32%"><strong><%=yyPOName%></strong></td>
				<td width="50%"><strong><%=yyWrkNam%></strong></td>
				<td width="6%" align="center"><strong><%=yyClone%></strong></td>
				<td width="6%" align="center"><strong><%=yyModify%></strong></td>
				<td width="6%" align="center"><strong><%=yyDelete%></strong></td>
			  </tr>
<%	if IsArray(alldata) then
		for rowcounter=0 to UBOUND(alldata,2)
			if bgcolor="#E7EAEF" then bgcolor="#FFFFFF" else bgcolor="#E7EAEF" %>
			  <tr bgcolor="<%=bgcolor%>">
				<td><%=alldata(1,rowcounter)%></td>
				<td><%=alldata(2,rowcounter)%></td>
				<td align="center"><input type=button value="<%=yyClone%>" onclick="clone('<%=alldata(0,rowcounter)%>')" /></td>
				<td align="center"><input type=button value="<%=yyModify%>" onclick="modrec('<%=alldata(0,rowcounter)%>')" /></td>
				<td align="center"><input type=button value="<%=yyDelete%>" onclick="delrec('<%=alldata(0,rowcounter)%>')" /></td>
			  </tr>
<%		next
	else
%>
			  <tr>
                <td width="100%" colspan="5" align="center"><br /><%=yyPONon%><br />&nbsp;</td>
			  </tr>
<%	end if %>
			  <tr>
                <td width="100%" colspan="5" align="center"><br /><strong><%=yyPOClk%> </strong>&nbsp;&nbsp;<input type="button" value="<%=yyPONew%>" onclick="newrec()" />&nbsp;<strong><%=yyOr%></strong>&nbsp;<input type="button" value="<%=yyPONewT%>" onclick="newtextrec()" /><br />&nbsp;</td>
			  </tr>
<%	if useStockManagement then
		sSQL = "SELECT DISTINCT optGrpID,optGrpName,optGrpWorkingName FROM optiongroup INNER JOIN (options INNER JOIN (prodoptions INNER JOIN products ON prodoptions.poProdID=products.pID) ON options.optGroup=prodoptions.poOptionGroup) ON optiongroup.optGrpID=options.optGroup WHERE options.optStock<=0 AND products.pStockByOpts<>0 AND optType IN (-2,-1,1,2) ORDER BY optGrpName,optGrpWorkingName"
		rs.Open sSQL,cnn,0,1
		if NOT rs.EOF then response.write "<tr><td colspan=""5"" align=""center""><strong>The following options contain at least 1 item that is out of stock</strong></td></tr>"
		do while NOT rs.EOF
			if bgcolor="#E7EAEF" then bgcolor="#FFFFFF" else bgcolor="#E7EAEF" %>
			  <tr bgcolor="<%=bgcolor%>">
				<td><%=rs("optGrpName")%></td>
				<td><%=rs("optGrpWorkingName")%></td>
				<td align="center">&nbsp;</td>
				<td align="center"><input type=button value="<%=yyModify%>" onclick="modrec('<%=rs("optGrpID")%>')" /></td>
				<td align="center"><input type=button value="<%=yyDelete%>" onclick="delrec('<%=rs("optGrpID")%>')" /></td>
			  </tr><%
			rs.MoveNext
		loop
		rs.Close
	end if %>
			  <tr>
                <td width="100%" colspan="5" align="center"><br />
                          <a href="admin.asp"><strong><%=yyAdmHom%></strong></a><br />
				<img src="../images/clearpixel.gif" width="300" height="3" alt="" /></td>
			  </tr>
            </table></td>
		  </form>
        </tr>
<%
end if
cnn.Close
set rs = nothing
set cnn = nothing
%>
      </table>