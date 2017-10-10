<%
if storesessionvalue="" then storesessionvalue="virtualstore"
if Session("loggedon") <> storesessionvalue OR disallowlogin=TRUE then response.end
success=TRUE
Set rs = Server.CreateObject("ADODB.RecordSet")
Set cnn=Server.CreateObject("ADODB.Connection")
cnn.open sDSN
if request.form("posted")="1" then
	for each objItem In request.form
		if left(objItem,4) = "idxx" then
			ip1 = ip2long(request.form(objItem))
			if trim(request.form(replace(objItem,"xx","yy")))<>"" then
				ip2 = ip2long(trim(request.form(replace(objItem,"xx","yy"))))
			else
				ip2 = 0
			end if
			if ip1 <> -1 AND ip2 <> -1 then
				sSQL = "UPDATE ipblocking SET dcip1=" & ip1 & ",dcip2=" & ip2 & " WHERE dcid=" & mid(objItem,5)
				cnn.Execute(sSQL)
			end if
		elseif left(objItem,7) = "newidxx" then
			ip1 = ip2long(request.form(objItem))
			if trim(request.form(replace(objItem,"xx","yy")))<>"" then
				ip2 = ip2long(trim(request.form(replace(objItem,"xx","yy"))))
			else
				ip2 = 0
			end if
			if ip1 <> -1 AND ip2 <> -1 then
				sSQL = "INSERT INTO ipblocking (dcip1,dcip2) VALUES (" & ip1 & "," & ip2 & ")"
				cnn.Execute(sSQL)
			end if
		elseif left(objItem,5) = "delip" then
			sSQL = "DELETE FROM ipblocking WHERE dcid=" & mid(objItem,6)
			cnn.Execute(sSQL)
		elseif left(objItem,5) = "delss" then
			sSQL = "DELETE FROM multibuyblock WHERE ssdenyid=" & mid(objItem,6)
			cnn.Execute(sSQL)
		end if
	next
	if success then response.write "<meta http-equiv=""refresh"" content=""2; url=adminipblock.asp"">"
end if
%>
      <table border="0" cellspacing="0" cellpadding="0" width="100%" bgcolor="" align="center">
<%	if request.form("posted") = "1" AND success then %>
        <tr>
          <td width="100%">
			<table width="100%" border="0" cellspacing="0" cellpadding="3" bgcolor="">
			  <tr> 
                <td width="100%" colspan="2" align="center"><br /><strong><%=yyUpdSuc%></strong><br /><br /><%=yyNowFrd%><br /><br />
                        <%=yyNoAuto%> <A href="adminipblock.asp"><strong><%=yyClkHer%></strong></a>.<br />
                        <br />
				<img src="../images/clearpixel.gif" width="300" height="3" alt="" />
                </td>
			  </tr>
			</table></td>
        </tr>
<%	elseif request.form("posted") = "1" then %>
        <tr>
          <td width="100%">
			<table width="100%" border="0" cellspacing="0" cellpadding="3" bgcolor="">
			  <tr> 
                <td width="100%" colspan="2" align="center"><br /><font color="#FF0000"><strong><%=yyErrUpd%></strong></font><br /><br /><%=errmsg%><br /><br />
				<a href="javascript:history.go(-1)"><strong><%=yyClkBac%></strong></a></td>
			  </tr>
			</table></td>
        </tr>
<%	else %>
<script language="javascript" type="text/javascript">
<!--

//-->
</script>
        <tr>
		  <form name="mainform" method="post" action="adminipblock.asp">
		  <td width="100%">
			<input type="hidden" name="posted" value="1" />
            <table width="100%" border="0" cellspacing="0" cellpadding="1" bgcolor="">
			  <tr> 
                <td width="100%" colspan="5" align="center"><strong><%=yyUsIPBl%></strong><br />&nbsp;
				</td>
			  </tr>
			  <tr>
				<td align=center><strong><%=yySinIP%></strong></td>
				<td align=center><strong><%=yyLasIP%></strong></td>
				<td align=center><strong><%=yyDelete%></strong></td>
			  </tr><%
	bgcolor="#FFFFFF"
	sSQL = "SELECT dcid,dcip1,dcip2 FROM ipblocking ORDER BY dcip1"
	rs.Open sSQL,cnn,0,1
	if rs.EOF then
		response.write "<tr><td colspan=""3"" align=""center"">" & yyNoIPBl & "</td></tr>"
	else
		do while NOT rs.EOF
			if bgcolor="#E7EAEF" then bgcolor="#FFFFFF" else bgcolor="#E7EAEF" %>
<tr bgcolor="<%=bgcolor%>">
<td align="center"><input type="text" size="15" name="idxx<%=rs("dcid")%>" value="<% response.write long2ip(int(rs("dcip1")))%>" /></td>
<td align="center"><input type="text" size="15" name="idyy<%=rs("dcid")%>" value="<% if rs("dcip2") <> 0 then response.write long2ip(int(rs("dcip2")))%>" /></td>
<td align="center"><input type="checkbox" name="delip<%=rs("dcid")%>"></td>
</tr>
<%			rs.movenext
		loop
	end if
	rs.Close
	for index=0 to 4
		if bgcolor="#E7EAEF" then bgcolor="#FFFFFF" else bgcolor="#E7EAEF" %>
<tr bgcolor="<%=bgcolor%>">
<td align="center"><input type="text" size="15" name="newidxx<%=index%>" /></td>
<td align="center"><input type="text" size="15" name="newidyy<%=index%>" /></td>
<td align="center">n/a</td>
</tr>
<%
	next
	if blockmultipurchase <> "" then %>
<tr bgcolor="#FFFFFF"><td colspan="3" align="center">&nbsp;<br><strong><%=yyFolIPB%></strong><br>&nbsp;</td></tr>
<tr><td align="center"><strong>IP Address</strong></td>
<td align="center"><strong>Checkout Attempts</strong></td>
<td align="center"><strong>Delete</strong></td></tr>
<%		sSQL = "SELECT ssdenyid,ssdenyip,sstimesaccess,lastaccess FROM multibuyblock WHERE sstimesaccess>=" & blockmultipurchase & " ORDER BY ssdenyip"
		rs.Open sSQL,cnn,0,1
		if rs.EOF then
			response.write "<tr><td colspan=""3"" align=""center"">" & yyNoIPBl & "</td></tr>"
		else
			do while NOT rs.EOF
				if bgcolor="#E7EAEF" then bgcolor="#FFFFFF" else bgcolor="#E7EAEF" %>
<tr bgcolor="<%=bgcolor%>">
<td align="center"><%=rs("ssdenyip")%></td>
<td align="center"><%=(rs("sstimesaccess")+1)%></td>
<td align="center"><input type="checkbox" name="delss<%=rs("ssdenyid")%>"></td>
</tr>
<%				rs.movenext
			loop
		end if
	end if
%>			  <tr> 
                <td width="100%" colspan="5" align="center">
                  <p><br /><input type="submit" value="<%=yySubmit%>" />&nbsp;&nbsp;<input type="reset" value="<%=yyReset%>" /><br />&nbsp;</p>
                </td>
			  </tr>
			  <tr> 
                <td width="100%" colspan="5" align="center"><br />
                          <a href="admin.asp"><strong><%=yyAdmHom%></strong></a><br />
				<img src="../images/clearpixel.gif" width="300" height="3" alt="" /></td>
			  </tr>
            </table>
		  </td>
		  </form>
        </tr>
<%
end if
%>
      </table>