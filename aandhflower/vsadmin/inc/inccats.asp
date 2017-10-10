<%
if storesessionvalue="" then storesessionvalue="virtualstore"
if Session("loggedon") <> storesessionvalue OR disallowlogin=TRUE then response.end
Dim sSQL,rs,alldata,alladmin,success,cnn,rowcounter,errmsg
success=true
maxcatsperpage = 100
if maxloginlevels="" then maxloginlevels=5
Set rs = Server.CreateObject("ADODB.RecordSet")
Set rsCats = Server.CreateObject("ADODB.RecordSet")
Set cnn=Server.CreateObject("ADODB.Connection")
cnn.open sDSN
alreadygotadmin = getadminsettings()
sSQL = ""
if request.form("act")="changepos" then
	currentorder = Int(Request.Form("selectedq"))
	neworder = Int(Request.Form("newval"))
	sSQL = "SELECT sectionID FROM sections ORDER BY sectionOrder"
	rs.Open sSQL,cnn,0,1
	alldata=rs.getrows
	rs.Close
	FOR rowcounter=0 TO ubound(alldata,2)
		theorder = rowcounter+1
		if currentorder = theorder then
			theorder = neworder
		elseif (currentorder > theorder) AND (neworder <= theorder) then
			theorder = theorder + 1
		elseif (currentorder < theorder) AND (neworder >= theorder) then
			theorder = theorder - 1
		end if
		sSQL="UPDATE sections SET sectionOrder="&theorder&" WHERE sectionID="&alldata(0,rowcounter)
		cnn.Execute(sSQL)
	NEXT
	response.write "<meta http-equiv=""refresh"" content=""1; url=admincats.asp?pg="& request.form("pg") &""">"
elseif request.form("posted")="1" then
	if request.form("act")="delete" then
		sSQL = "DELETE FROM cpnassign WHERE cpaType=1 AND cpaAssignment='"&request.form("id")&"'"
		cnn.Execute(sSQL)
		sSQL = "DELETE FROM sections WHERE sectionID=" & request.form("id")
		cnn.Execute(sSQL)
		sSQL = "DELETE FROM multisections WHERE pSection=" & request.form("id")
		cnn.Execute(sSQL)
		response.write "<meta http-equiv=""refresh"" content=""2; url=admincats.asp?pg="& request.form("pg") &""">"
	elseif request.form("act")="domodify" then
		sSQL = "UPDATE sections SET sectionName='"&Trim(replace(request.form("secname"),"'","''"))&"',sectionDescription='"&replace(request.form("secdesc"),"'","''")&"',sectionImage='"&replace(request.form("secimage"),"'","''")&"',topSection="&request.form("tsTopSection")&",rootSection="&request.form("catfunction")
		workname = Trim(replace(request.form("secworkname"),"'","''"))
		if workname<>"" then
			sSQL = sSQL & ",sectionWorkingName='"&workname&"'"
		else
			sSQL = sSQL & ",sectionWorkingName='"&Trim(replace(request.form("secname"),"'","''"))&"'"
		end if
		for index=2 to adminlanguages+1
			if (adminlangsettings AND 256)=256 then
				sSQL = sSQL & ",sectionName" & index & "='"&Trim(replace(request.form("secname" & index),"'","''"))&"'"
			end if
			if (adminlangsettings AND 512)=512 then
				sSQL = sSQL & ",sectionDescription" & index & "='"&Trim(replace(request.form("secdesc" & index),"'","''"))&"'"
			end if
		next
		sSQL = sSQL & ",sectionDisabled=" & trim(request.form("sectionDisabled"))
		sSQL = sSQL & ",sectionurl='" & replace(trim(request.form("sectionurl")),"'","''") & "'"
		sSQL = sSQL & " WHERE sectionID="&Request.Form("id")
		cnn.Execute(sSQL)
		response.write "<meta http-equiv=""refresh"" content=""2; url=admincats.asp?pg="& request.form("pg") &""">"
	elseif request.form("act")="doaddnew" then
		sSQL = "SELECT MAX(sectionOrder) AS mxOrder FROM sections"
		rs.Open sSQL,cnn,0,1
		mxOrder = rs("mxOrder")
		rs.Close
		if IsNull(mxOrder) OR mxOrder="" then mxOrder=1 else mxOrder=mxOrder+1
		sSQL = "INSERT INTO sections (sectionName,sectionDescription,sectionImage,sectionOrder,topSection,rootSection,sectionWorkingName"
		for index=2 to adminlanguages+1
			if (adminlangsettings AND 256)=256 then
				sSQL = sSQL & ",sectionName" & index
			end if
			if (adminlangsettings AND 512)=512 then
				sSQL = sSQL & ",sectionDescription" & index
			end if
		next
		sSQL = sSQL & ",sectionDisabled,sectionurl) VALUES ('"&Trim(replace(request.form("secname"),"'","''"))&"','"&replace(request.form("secdesc"),"'","''")&"','"&replace(request.form("secimage"),"'","''")&"',"&mxOrder&","&request.form("tsTopSection")&","&request.form("catfunction")
		workname = Trim(replace(request.form("secworkname"),"'","''"))
		if workname<>"" then
			sSQL = sSQL & ",'"&workname&"'"
		else
			sSQL = sSQL & ",'"&Trim(replace(request.form("secname"),"'","''"))&"'"
		end if
		for index=2 to adminlanguages+1
			if (adminlangsettings AND 256)=256 then
				sSQL = sSQL & ",'"&Trim(replace(request.form("secname" & index),"'","''"))&"'"
			end if
			if (adminlangsettings AND 512)=512 then
				sSQL = sSQL & ",'"&Trim(replace(request.form("secdesc" & index),"'","''"))&"'"
			end if
		next
		sSQL = sSQL & "," & trim(request.form("sectionDisabled"))
		sSQL = sSQL & ",'" & replace(trim(request.form("sectionurl")),"'","''") & "')"
		cnn.Execute(sSQL)
		response.write "<meta http-equiv=""refresh"" content=""2; url=admincats.asp?pg="& request.form("pg") &""">"
	elseif request.form("act")="dodiscounts" then
		sSQL = "INSERT INTO cpnassign (cpaCpnID,cpaType,cpaAssignment) VALUES ("&request.form("assdisc")&",1,'"&request.form("id")&"')"
		cnn.Execute(sSQL)
		response.write "<meta http-equiv=""refresh"" content=""2; url=admincats.asp?pg="& request.form("pg") &""">"
	elseif request.form("act")="deletedisc" then
		sSQL = "DELETE FROM cpnassign WHERE cpaID="&request.form("id")
		cnn.Execute(sSQL)
		response.write "<meta http-equiv=""refresh"" content=""2; url=admincats.asp?pg="& request.form("pg") &""">"
	end if
end if
%>
<script language="javascript" type="text/javascript">
<!--
function formvalidator(theForm)
{
  if (theForm.secname.value == "")
  {
    alert("<%=yyPlsEntr%> \"<%=yyCatNam%>\".");
    theForm.secname.focus();
    return (false);
  }
  if (theForm.tsTopSection[theForm.tsTopSection.selectedIndex].value == "")
  {
    alert("<%=yyPlsSel%> \"<%=yyCatSub%>\".");
    theForm.tsTopSection.focus();
    return (false);
  }
  return (true);
}
//-->
</script>
      <table border="0" cellspacing="0" cellpadding="0" width="100%" bgcolor="" align="center">
<% if request.form("posted")="1" AND (request.form("act")="modify" OR request.form("act")="addnew") then
		alltopsections=""
		if request.form("act")="modify" then
			sSQL = "SELECT sectionID,sectionName,sectionName2,sectionName3,rootSection,sectionImage,sectionWorkingName,topSection,sectionDisabled,sectionurl,sectionDescription FROM sections WHERE sectionID="&Request.Form("id")
			rs.Open sSQL,cnn,0,1
			sectionID = rs("sectionID")
			sectionName = rs("sectionName")
			sectionName2 = rs("sectionName2")
			sectionName3 = rs("sectionName3")
			rootSection = rs("rootSection")
			sectionImage = rs("sectionImage")
			sectionWorkingName = rs("sectionWorkingName")
			topSection = rs("topSection")
			sectionDisabled = rs("sectionDisabled")
			sectionurl = rs("sectionurl")&""
			sectionDescription = rs("sectionDescription")
			rs.Close
		else
			sectionID = ""
			sectionName = ""
			sectionName2 = ""
			sectionName3 = ""
			rootSection = 1
			sectionImage = ""
			sectionWorkingName = ""
			topSection = 0
			sectionDisabled = 0
			sectionurl = ""
			sectionDescription = ""
		end if
		sSQL = "SELECT sectionID,sectionWorkingName FROM sections WHERE rootSection=0 ORDER BY sectionWorkingName"
		rs.Open sSQL,cnn,0,1
		if NOT rs.EOF then
			alltopsections=rs.getrows
		end if
		rs.Close
%>
        <tr>
		<form name="mainform" method="post" action="admincats.asp" onsubmit="return formvalidator(this)">
		  <td width="100%">
			<input type="hidden" name="posted" value="1" />
			<% if request.form("act")="modify" then %>
			<input type="hidden" name="act" value="domodify" />
			<% else %>
			<input type="hidden" name="act" value="doaddnew" />
			<% end if %>
			<input type="hidden" name="id" value="<%=Request.Form("id")%>" />
			<input type="hidden" name="pg" value="<%=Request.Form("pg")%>" />
            <table width="100%" border="0" cellspacing="0" cellpadding="3" bgcolor="">
			  <tr> 
                <td width="100%" colspan="2" align="center"><strong><%=yyCatAdm%></strong><br />&nbsp;</td>
			  </tr>
			  <tr>
				<td width="40%" align="center" valign="top"><strong><%=yyCatNam%></strong><br /><input type="text" name="secname" size="30" value="<%=replace(sectionName,"""","&quot;")%>" /><br />
<%			for index=2 to adminlanguages+1
				if (adminlangsettings AND 256)=256 then %>
				<strong><%=yyCatNam & " " & index %></strong><br />
				<input type="text" name="secname<%=index%>" size="30" value="<% execute("response.write replace(sectionName" & index & "&"""",chr(34),""&quot;"")")%>" /><br />
<%				end if
			next %></td>
				<td width="60%" rowspan="9" align="center" valign="top"><strong><%=yyCatDes%></strong><br /><textarea name="secdesc" cols="38" rows="8" wrap=virtual><%=sectionDescription%></textarea><br />
<%			for index=2 to adminlanguages+1
				if (adminlangsettings AND 512)=512 then
					if request.form("act")="modify" then
						sSQL = "SELECT sectionDescription" & index & " FROM sections WHERE sectionID="&Request.Form("id")
						rs.Open sSQL,cnn,0,1
						sectionDescription = rs("sectionDescription" & index)
						rs.Close
					else
						sectionDescription = ""
					end if
%>
				<strong><%=yyCatDes & " " & index %></strong><br />
				<textarea name="secdesc<%=index%>" cols="38" rows="8" wrap=virtual><%=sectionDescription%></textarea><br />
<%				end if
			next %>
				&nbsp;<br /><select name="sectionDisabled" size="1">
				<option value="0"><%=yyNoRes%></option>
				<%	for index=1 to maxloginlevels
						response.write "<option value="""&index&""""
						if sectionDisabled=index then response.write " selected"
						response.write ">" & yyLiLev & " " & index & "</option>"
					next%>
				<option value="127"<% if sectionDisabled=127 then response.write " selected"%>><%=yyDisCat%></option>
				</select><br />
				&nbsp;<br /><strong>Category URL (Optional)</strong><br />
				<input type="text" name="sectionurl" size="40" value="<%=replace(sectionurl,"""","&quot;")%>" />
                </td>
			  </tr>
			  <tr>
				<td align="center" valign="top"><strong><%=yyCatWrNa%></strong></td>
			  </tr>
			  <tr>
				<td align="center" valign="top"><input type="text" name="secworkname" size="30" value="<%=replace(sectionWorkingName&"","""","&quot;")%>" /></td>
			  </tr>
			  <tr>
				<td align="center" valign="top"><strong><%=yyCatSub%></strong></td>
			  </tr>
			  <tr>
				<td align="center" valign="top"><select name="tsTopSection" size="1"><option value="0"><%=yyCatHom%></option>
				<%	foundcat=(topSection=0)
					if IsArray(alltopsections) then
						for index=0 to UBOUND(alltopsections,2)
							if alltopsections(0, index)<>sectionID then
								response.write "<option value=""" & alltopsections(0, index) & """"
								if topSection=alltopsections(0, index) then
									response.write " selected"
									foundcat=true
								end if
								response.write ">" & alltopsections(1, index) & "</option>" & vbCrLf
							end if
						next
					end if
					if NOT foundcat then response.write "<option value="""" selected>**undefined**</option>"
					%></select>
                </td>
			  </tr>
			  <tr>
				<td align="center" valign="top"><strong><%=yyCatFn%></strong></td>
			  </tr>
			  <tr>
				<td align="center" valign="top"><select name="catfunction" size="1">
				  <option value="1"><%=yyCatPrd%></option>
				  <option value="0" <% if rootSection=0 then response.write "selected"%>><%=yyCatCat%></option>
				  </select></td>
			  </tr>
			  <tr>
				<td align="center" valign="top"><strong><%=yyCatImg%></strong></td>
			  </tr>
			  <tr>
				<td align="center" valign="top"><input type="text" name="secimage" size="30" value="<%=replace(sectionImage&"","""","&quot;")%>" /></td>
			  </tr>
			  <tr> 
                <td width="100%" colspan="2" align="center"><br /><input type="submit" value="<%=yySubmit%>" /></td>
			  </tr>
			  <tr> 
                <td width="100%" colspan="2"><br /><ul>
				  <li><%=yyCatEx1%></li>
				  <li><%=yyCatEx2%></li>
				  </ul></td>
			  </tr>
			  <tr> 
                <td width="100%" colspan="2" align="center"><br />
                          <a href="admin.asp"><strong><%=yyAdmHom%></strong></a><br />
                          &nbsp;</td>
			  </tr>
            </table></td>
		  </form>
        </tr>
<% elseif request.form("act")="discounts" then 
		sSQL = "SELECT sectionName FROM sections WHERE sectionID="&request.form("id")
		rs.Open sSQL,cnn,0,1
		thisname=rs("sectionName")
		rs.Close
		alldata=""
		sSQL = "SELECT cpaID,cpaCpnID,cpnWorkingName,cpnSitewide,cpnEndDate,cpnType FROM cpnassign INNER JOIN coupons ON cpnassign.cpaCpnID=coupons.cpnID WHERE cpaType=1 AND cpaAssignment='" & request.form("id") & "'"
		rs.Open sSQL,cnn,0,1
		if NOT rs.EOF then alldata=rs.GetRows
		rs.Close
		alldata2=""
		tdt = Date()
		sSQL = "SELECT cpnID,cpnWorkingName,cpnSitewide FROM coupons WHERE (cpnSitewide=0 OR cpnSitewide=3) AND cpnEndDate >="&datedelim&VSUSDate(tdt)&datedelim
		rs.Open sSQL,cnn,0,1
		if NOT rs.EOF then alldata2=rs.GetRows
		rs.Close
%>
<script language="javascript" type="text/javascript">
<!--
function delrec(id) {
cmsg = "<%=yyConAss%>\n"
if (confirm(cmsg)) {
	document.mainform.id.value = id;
	document.mainform.act.value = "deletedisc";
	document.mainform.submit();
}
}
// -->
</script>
        <tr>
		<form name="mainform" method="post" action="admincats.asp">
		  <td width="100%">
			<input type="hidden" name="posted" value="1" />
			<input type="hidden" name="act" value="dodiscounts" />
			<input type="hidden" name="id" value="<%=request.form("id")%>" />
			<input type="hidden" name="pg" value="<%=Request.Form("pg")%>" />
            <table width="100%" border="0" cellspacing="0" cellpadding="3" bgcolor="">
			  <tr> 
                <td width="100%" colspan="4" align="center"><strong><%=yyAssDis%> &quot;<%=thisname%>&quot;.</strong><br />&nbsp;</td>
			  </tr>
<%	gotone=false
	if IsArray(alldata2) then
		thestr = "<tr><td colspan='4' align='center'>"&yyAsDsCp&": <select name='assdisc' size='1'>"
		for index=0 to UBOUND(alldata2,2)
			alreadyassign=false
			if IsArray(alldata) then
				for index2=0 to UBOUND(alldata,2)
					if alldata2(0,index)=alldata(1,index2) then alreadyassign=true
				next
			end if
			if NOT alreadyassign then
				thestr = thestr & "<option value='"&alldata2(0,index)&"'>"&alldata2(1,index)&"</option>" & vbCrLf
				gotone=true
			end if
		next
		thestr = thestr & "</select> <input type='submit' value='Go' /></td></tr>"
	end if
	if gotone then
		response.write thestr
	else
%>
			  <tr> 
                <td width="100%" colspan="4" align="center"><br /><strong><%=yyNoDis%></td>
			  </tr>
<%
	end if
	if IsArray(alldata) then
%>
			  <tr> 
                <td width="100%" colspan="4" align="center"><br /><strong><%=yyCurDis%> &quot;<%=thisname%>&quot;.</strong><br />&nbsp;</td>
			  </tr>
			  <tr> 
                <td><strong><%=yyWrkNam%></strong></td>
				<td><strong><%=yyDisTyp%></strong></td>
				<td><strong><%=yyExpire%></strong></td>
				<td align="center"><strong><%=yyDelete%></strong></td>
			  </tr>
<%
		for index=0 to UBOUND(alldata,2)
			prefont = ""
			postfont = ""
			if alldata(3,index)=1 OR alldata(4,index)-Date() < 0 then
				prefont = "<font color=""#FF0000"">"
				postfont = "</font>"
			end if
%>
			  <tr> 
                <td><%=prefont & alldata(2,index) & postfont %></td>
				<td><%	if alldata(5,index)=0 then
							response.write prefont & yyFrSShp & postfont
						elseif alldata(5,index)=1 then
							response.write prefont & yyFlatDs & postfont
						elseif alldata(5,index)=2 then
							response.write prefont & yyPerDis & postfont
						end if %></td>
				<td><%	if alldata(4,index)=DateSerial(3000,1,1) then
							response.write yyNever
						elseif alldata(4,index)-Date() < 0 then
							response.write "<font color='#FF0000'>"&yyExpird&"</font>"
						else
							response.write prefont & alldata(4,index) & postfont
						end if %></td>
				<td align="center"><input type="button" name="discount" value="Delete Assignment" onclick="delrec('<%=alldata(0,index)%>')" /></td>
			  </tr>
<%
		next
	else
%>
			  <tr> 
                <td width="100%" colspan="4" align="center"><br /><strong><%=yyNoAss%></td>
			  </tr>
<%
	end if
%>
			  <tr>
                <td width="100%" colspan="4" align="center"><br />&nbsp;</td>
			  </tr>
			  <tr> 
                <td width="100%" colspan="4" align="center"><br />
                          <a href="admin.asp"><strong><%=yyAdmHom%></strong></a><br />
                          &nbsp;</td>
			  </tr>
            </table></td>
		  </form>
        </tr>
<% elseif request.form("act")="changepos" then %>
        <tr>
          <td width="100%" align="center">
			<p>&nbsp;</p>
			<p>&nbsp;</p>
			<p>&nbsp;</p>
			<p><strong><%=yyUpdat%> . . . . . . . </strong></font></p>
			<p>&nbsp;</p>
			<p><%=yyNoFor%> <a href="admincats.asp?pg=<%=request.form("pg")%>"><%=yyClkHer%></a>.</p>
			<p>&nbsp;</p>
			<p>&nbsp;</p>
		  </td>
		</tr>
<% elseif request.form("posted")="1" AND success then %>
        <tr>
          <td width="100%">
			<table width="100%" border="0" cellspacing="0" cellpadding="3" bgcolor="">
			  <tr> 
                <td width="100%" colspan="2" align="center"><br /><strong><%=yyUpdSuc%></strong><br /><br /><%=yyNowFrd%><br /><br />
                        <%=yyNoAuto%><A href="admincats.asp"><strong><%=yyClkHer%></strong></a>.<br />
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
function writeposition(currpos,maxpos)
	Dim reqtext,i
	reqtext="<select name='newpos" & currpos & "' onchange='chi("&currpos&");'>"
	for i = 1 to maxpos
		reqtext = reqtext & "<option" ' value='"&i&"'"
		if currpos=i then reqtext=reqtext&" selected"
		reqtext = reqtext & ">"&i ' &"</option>"
		if i >= 10 AND i < (maxpos-15) AND ABS(currpos-i) > 40 then i = i + 9
	next
	writeposition = reqtext & "</select>"
end function
allcoupon=""
sSQL = "SELECT DISTINCT cpaAssignment FROM cpnassign WHERE cpaType=1"
rs.Open sSQL,cnn,0,1
if NOT rs.EOF then allcoupon=rs.getrows
rs.Close
%>
<script language="javascript" type="text/javascript">
<!--
function chi(currindex) {
	var i = eval("document.mainform.newpos"+currindex+".selectedIndex");
	document.mainform.newval.value = eval("document.mainform.newpos"+currindex+".options[i].text");
	document.mainform.selectedq.value = currindex;
	document.mainform.act.value = "changepos";
	document.mainform.submit();
}
function mrk(id) {
	document.mainform.id.value = id;
	document.mainform.act.value = "modify";
	document.mainform.submit();
}
function newrec(id) {
	document.mainform.id.value = id;
	document.mainform.act.value = "addnew";
	document.mainform.submit();
}
function dsk(id) {
	document.mainform.id.value = id;
	document.mainform.act.value = "discounts";
	document.mainform.submit();
}
function drk(id) {
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
		  <form name="mainform" method="post" action="admincats.asp">
		  <td width="100%">
			<input type="hidden" name="posted" value="1" />
			<input type="hidden" name="act" value="xxxxx" />
			<input type="hidden" name="id" value="xxxxx" />
			<input type="hidden" name="pg" value="<%=Request.QueryString("pg")%>" />
			<input type="hidden" name="selectedq" value="1" />
			<input type="hidden" name="newval" value="1" />
            <table width="100%" border="0" cellspacing="0" cellpadding="1" bgcolor="">
			  <tr> 
                <td width="100%" colspan="6" align="center"><strong><%=yyCatAdm%></strong><br />&nbsp;</td>
			  </tr>
<%
Function writepagebar(CurPage, iNumPages)
	Dim sLink, i, sStr, startPage, endPage
	sLink = "<a href='admincats.asp?pg="
	startPage = vrmax(1,CInt(Int(CDbl(CurPage)/10.0)*10))
	endPage = vrmin(iNumPages,CInt(Int(CDbl(CurPage)/10.0)*10)+10)
	if CurPage > 1 then
		sStr = sLink & "1" & "'><strong><font face='Verdana'>&laquo;</font></strong></a> " & sLink & CurPage-1 & "'>Previous</a> | "
	else
		sStr = "<strong><font face='Verdana'>&laquo;</font></strong> Previous | "
	end if
	for i=startPage to endPage
		if i=CurPage then
			sStr = sStr & i & " | "
		else
			sStr = sStr & sLink & i & "'>"
			if i=startPage AND i > 1 then sStr=sStr&"..."
			sStr = sStr & i
			if i=endPage AND i < iNumPages then sStr=sStr&"..."
			sStr = sStr & "</a> | "
		end if
	next
	if CurPage < iNumPages then
		writepagebar = sStr & sLink & CurPage+1 & "'>Next</a> " & sLink & iNumPages & "'><strong><font face='Verdana'>&raquo;</font></strong></a>"
	else
		writepagebar = sStr & " Next <strong><font face='Verdana'>&raquo;</font></strong>"
	end if
End function
	If Request.QueryString("pg") = "" Then
		CurPage = 1
	Else
		CurPage = Int(Request.QueryString("pg"))
	End If
	sSQL = "SELECT sectionID,sectionWorkingName,sectionDescription,topSection,rootSection,sectionDisabled FROM sections ORDER BY sectionOrder"
	rsCats.CursorLocation = 3 ' adUseClient
	rsCats.CacheSize = maxcatsperpage
	rsCats.Open sSQL,cnn
	if NOT rsCats.EOF then
		rsCats.MoveFirst
		rsCats.PageSize = maxcatsperpage
		rsCats.AbsolutePage = CurPage
		islooping=false
		noproducts=false
		hascatinprodsection=false
		rowcounter=0
		totnumrows=rsCats.RecordCount
		iNumOfPages = Int((totnumrows + (maxcatsperpage-1)) / maxcatsperpage)
		If iNumOfPages > 1 Then Response.Write "<tr><td align=""center"" colspan=""6"">" & writepagebar(CurPage, iNumOfPages) & "<br /><br /></td></tr>"
%>
			  <tr>
				<td width="5%"><strong><%=yyOrder%></strong></td>
				<td align="left"><strong><%=yyCatPat%></strong></td>
				<td align="left"><strong><%=yyCatNam%></strong></td>
				<td width="5%" align="center"><font size="1"><strong><%=yyDiscnt%></strong></font></td>
				<td width="5%" align="center"><font size="1"><strong><%=yyModify%></strong></font></td>
				<td width="5%" align="center"><font size="1"><strong><%=yyDelete%></strong></font></td>
			  </tr>
<%		do while NOT rsCats.EOF AND rowcounter < maxcatsperpage
			if bgcolor="#E7EAEF" then bgcolor="#FFFFFF" else bgcolor="#E7EAEF"%>
<tr bgcolor="<%=bgcolor%>"><td><%=writeposition((maxcatsperpage*(CurPage-1))+rowcounter+1,totnumrows)%></td>
<td><%
tslist = ""
thetopts = rsCats("topSection")
for index=0 to 10
	if thetopts=0 then
		tslist = yyHome & tslist
		exit for
	elseif index=10 then
		tslist = "<strong><font color='#FF0000'>"&yyLoop&"</font></strong>" & tslist
		islooping=true
	else
		sSQL = "SELECT sectionID,topSection,sectionWorkingName,rootSection FROM sections WHERE sectionID=" & thetopts
		rs.Open sSQL,cnn,0,1
		if NOT rs.EOF then
			errstart = ""
			errend = ""
			if rs("rootSection")=1 then
				errstart = "<strong><font color='#FF0000'>"
				errend = "</font></strong>"
				hascatinprodsection=true
			end if
			tslist = " &raquo; " & errstart & rs("sectionWorkingName") & errend & tslist
			thetopts = rs("topSection")
		else
			tslist = "<strong><font color='#FF0000'>"&yyTopDel&"</font></strong>" & tslist
			rs.Close
			exit for
		end if
		rs.Close
	end if
next
response.write "<font size=""1"">" & tslist & "</font></td><td>"
if Int(CStr(rsCats("rootSection")))=1 then response.write "<strong>"
if Int(CStr(rsCats("sectionDisabled")))=127 then response.write "<strike><font color=""#FF0000"">"
response.write rsCats("sectionWorkingName") & " (" & rsCats("sectionID") & ")"
if Int(CStr(rsCats("sectionDisabled")))=127 then response.write "</font></strike>"
if Int(CStr(rsCats("rootSection")))=1 then response.write "</strong>"
response.write "</td><td><input"
if IsArray(allcoupon) then
	for index=0 to UBOUND(allcoupon,2)
		if Int(allcoupon(0,index))=rsCats("sectionID") then
			response.write " style=""color: #FF0000"""
			exit for
		end if
	next
end if
%> type="button" value="<%=yyAssign%>" onclick="dsk('<%=rsCats("sectionID")%>')"></td>
<td><input type="button" value="<%=yyModify%>" onclick="mrk('<%=rsCats("sectionID")%>')" /></td>
<td><input type="button" value="<%=yyDelete%>" onclick="drk('<%=rsCats("sectionID")%>')" /></td>
</tr><%		rowcounter=rowcounter+1
			rsCats.MoveNext
		loop
		If iNumOfPages > 1 Then Response.Write "<tr><td align=""center"" colspan=""6""><br />" & writepagebar(CurPage, iNumOfPages) & "</td></tr>"
		if islooping then
%>
			  <tr><td width="100%" colspan="6"><br /><strong><font color='#FF0000'>** </font></strong><%=yyCatEx3%></td></tr>
<%
		end if
		if hascatinprodsection then
%>
			  <tr><td width="100%" colspan="6"><br /><ul><li><%=yyCPErr%></li></ul></td></tr>
<%
		end if
%>
			  <tr><td width="100%" colspan="6"><br /><ul><li><%=yyCatEx4%></li></ul></td></tr>
<%
	else
%>
			  <tr><td width="100%" colspan="6" align="center"><br /><strong><%=yyCatEx5%><br />&nbsp;</td></tr>
<%
	end if
	rsCats.Close
%>
			  <tr> 
                <td width="100%" colspan="6" align="center"><br /><strong><%=yyCatNew%></strong>&nbsp;&nbsp;<input type="button" value="<%=yyNewCat%>" onclick="newrec()" /><br />&nbsp;</td>
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
set rsCats = nothing
set cnn = nothing
%>
      </table>