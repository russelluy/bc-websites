<%
if storesessionvalue="" then storesessionvalue="virtualstore"
if Session("loggedon") <> storesessionvalue OR disallowlogin=TRUE then response.end
success=true
Set rs = Server.CreateObject("ADODB.RecordSet")
Set cnn=Server.CreateObject("ADODB.Connection")
cnn.open sDSN
alreadygotadmin = getadminsettings()
if request.form("act")="domodify" then
	for index=0 to 30
		statusid=Trim(Request.Form("statusid" & index))
		if statusid<>"" then
			statPrivate = Trim(replace(request.form("privstatus" & index),"'","''"))
			statPublic = Trim(replace(request.form("pubstatus" & index),"'","''"))
			if statPublic="" then statPublic = statPrivate
			sSQL = "UPDATE orderstatus SET statPrivate='" & statPrivate & "',statPublic='" & statPublic & "'"
			for index2=2 to adminlanguages+1
				if (adminlangsettings AND 64)=64 then sSQL = sSQL & ",statPublic" & index2 & " ='" & trim(replace(request.form("pubstatus" & index & "x" & index2),"'","''")) & "'"
			next
			sSQL = sSQL & " WHERE statID="&statusid
			cnn.Execute(sSQL)
		end if
	next
	response.write "<meta http-equiv=""refresh"" content=""3; url=admin.asp"">"
else
	sSQL = "SELECT statID,statPrivate,statPublic,statPublic2,statPublic3 FROM orderstatus ORDER BY statID"
	rs.Open sSQL,cnn,0,1
	alldata=rs.getrows
	rs.Close
end if
%>
<script language="javascript" type="text/javascript">
<!--
function formvalidator(theForm){
for(index=0;index<=3;index++){
theelm=eval('theForm.privstatus'+index);
if(theelm.value == ""){
alert("Please enter a value in the field \"Private Text (Status " + (index+1) + ")\".");
theelm.focus();
return (false);
}
}
return (true);
}
//-->
</script>
      <table border="0" cellspacing="0" cellpadding="0" width="100%" bgcolor="" align="center">
<% if request.form("act")="domodify" AND success then %>
        <tr>
          <td width="100%">
			<table width="100%" border="0" cellspacing="0" cellpadding="3" bgcolor="">
			  <tr> 
                <td width="100%" colspan="2" align="center"><br /><strong><%=yyUpdSuc%></strong><br /><br /><%=yyNowFrd%><br /><br />
				<%=yyNoAuto%> <A href="admin.asp"><strong><%=yyClkHer%></strong></a>.<br /><br />
				<img src="../images/clearpixel.gif" width="300" height="3" alt="" />
                </td>
			  </tr>
			</table></td>
        </tr>
<% elseif request.form("act")="domodify" then %>
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
	if (adminlangsettings AND 64)<>64 then numcols=5 else numcols=5+adminlanguages
%>
        <tr>
		  <form name="mainform" method="post" action="adminordstatus.asp" onsubmit="return formvalidator(this)">
          <td width="100%" align="center">
			<input type="hidden" name="posted" value="1" />
			<input type="hidden" name="act" value="domodify" />
            <table width="500" border="0" cellspacing="0" cellpadding="3" bgcolor="">
			  <tr> 
                <td width="100%" colspan="<%=numcols%>" align="center"><strong><%=yyOSAdm%></strong><br />&nbsp;</td>
			  </tr>
			  <tr>
				<td align="center" valign="top" width="50"><strong>&nbsp;</strong></td>
				<td align="center" valign="top"><strong>&nbsp;</strong></td>
				<td align="center" valign="top"><strong><%=yyPrTxt%></strong></td>
				<td align="center" valign="top"><strong><%=yyPubTxt%></strong></td>
<%	for index2=2 to adminlanguages+1
		if (adminlangsettings AND 64)=64 then response.write "<td align=""center"" valign=""top""><strong>" & yyPubTxt & " " & index2 & "</strong></td>"
	next %>
				<td align="center" valign="top" width="50"><strong>&nbsp;</strong></td>
			  </tr>
<%	for rowcounter=0 to UBOUND(alldata,2)
		if alldata(0,rowcounter)=4 then %>
			  <tr> 
                <td width="100%" colspan="<%=numcols%>" align="center"><font size="1"><%=yyOSExp1%></font></td>
			  </tr>
<%		end if %>
			  <tr>
				<td align="center" valign="top"><strong>&nbsp;&nbsp;&nbsp;&nbsp;</strong></td>
				<td align="right"><input type="hidden" name="statusid<%=rowcounter%>" value="<%=alldata(0,rowcounter) %>" /><%=yyStatus%>&nbsp;<%=rowcounter%>:</td>
				<td align="center"><input type="text" size="20" name="privstatus<%=rowcounter%>" value="<%=replace(trim(alldata(1,rowcounter)&""),"""","&quot;") %>" /></td>
				<td align="center"><input type="text" size="20" name="pubstatus<%=rowcounter%>" value="<%=replace(trim(alldata(2,rowcounter)&""),"""","&quot;") %>" /></td>
<%	for index2=2 to adminlanguages+1
		if (adminlangsettings AND 64)=64 then response.write "<td align=""center""><input type=""text"" size=""20"" name=""pubstatus" & rowcounter & "x" & index2 & """ value=""" & replace(trim(alldata(1 + index2,rowcounter)&""),"""","&quot;") & """ /></td>"
	next %>
				<td align="center" valign="top"><strong>&nbsp;&nbsp;&nbsp;&nbsp;</strong></td>
			  </tr>
<%	next %>
			  <tr> 
                <td width="100%" colspan="<%=numcols%>" align="center"><input type="submit" value="<%=yySubmit%>" /></td>
			  </tr>
			  <tr> 
                <td width="100%" colspan="<%=numcols%>" align="center"><br /><a href="admin.asp"><strong><%=yyAdmHom%></strong></a><br />&nbsp;</td>
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