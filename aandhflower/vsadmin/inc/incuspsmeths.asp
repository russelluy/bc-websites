<%
if storesessionvalue="" then storesessionvalue="virtualstore"
if Session("loggedon") <> storesessionvalue OR disallowlogin=TRUE then response.end
Dim sSQL,rs,success,cnn,errmsg,index,allcountries
success=true
Set rs = Server.CreateObject("ADODB.RecordSet")
Set cnn=Server.CreateObject("ADODB.Connection")
cnn.open sDSN
alreadygotadmin = getadminsettings()
method=trim(request("method"))
if method<>"" then shipType=int(method)
shipmet = "USPS"
if shipType=4 then shipmet = "UPS"
if shipType=6 then shipmet = yyCanPos
if shipType=7 then shipmet = "FedEx"
if request.form("posted")="1" then
	if shipType=3 then
		for index=1 to 25
			if Trim(request.form("methodshow"&index))<>"" then
				sSQL = "UPDATE uspsmethods SET uspsShowAs='"&replace(request.form("methodshow"&index),"'","''")&"',"
				if request.form("methodfsa"&index)="ON" then
					sSQL = sSQL & "uspsFSA=1,"
				else
					sSQL = sSQL & "uspsFSA=0,"
				end if
				if request.form("methoduse"&index)="ON" then
					sSQL = sSQL & "uspsUseMethod=1 WHERE uspsID="&index
				else
					sSQL = sSQL & "uspsUseMethod=0 WHERE uspsID="&index
				end if
				cnn.Execute(sSQL)
			end if
		next
	elseif shipType=4 OR shipType=6 OR shipType=7 then
		indexadd=0
		if shipType=6 then
			indexadd=100
		elseif shipType=7 then
			indexadd=200
		end if
		for index=100+indexadd to 150+indexadd
			if Trim(request.form("methodshow"&index))<>"" then
				sSQL = "UPDATE uspsmethods SET "
				if request.form("methodfsa"&index)="ON" then
					sSQL = sSQL & "uspsFSA=1,"
				else
					sSQL = sSQL & "uspsFSA=0,"
				end if
				if request.form("methoduse"&index)="ON" then
					sSQL = sSQL & "uspsUseMethod=1 WHERE uspsID="&index
				else
					sSQL = sSQL & "uspsUseMethod=0 WHERE uspsID="&index
				end if
				cnn.Execute(sSQL)
			end if
		next
	end if
	response.write "<meta http-equiv=""refresh"" content=""3; url=admin.asp"">"
else
	sSQL = "SELECT uspsID,uspsMethod,uspsShowAs,uspsUseMethod,uspsLocal,uspsFSA FROM uspsmethods "
	if shipType=3 then
		sSQL = sSQL & " WHERE uspsID < 100"
	elseif shipType=4 then
		sSQL = sSQL & " WHERE uspsID > 100 AND uspsID < 200"
	elseif shipType=6 then
		sSQL = sSQL & " WHERE uspsID > 200 AND uspsID < 300"
	elseif shipType=7 then
		sSQL = sSQL & " WHERE uspsID > 300 AND uspsID < 400"
	end if
	sSQL = sSQL & " ORDER BY uspsLocal DESC, uspsID"
	rs.Open sSQL,cnn,0,1
	allmethods=rs.getrows
	rs.Close
end if
cnn.Close
set rs = nothing
set cnn = nothing
%>
      <table border="0" cellspacing="0" cellpadding="0" width="100%" bgcolor="" align="center">
<% if method="" then %>
        <tr>
          <td width="100%" align="center">
            <table width="80%" border="0" cellspacing="1" cellpadding="2" bgcolor="">
			  <tr>
                <td colspan="3" align="center"><strong><%=yyUsUpd & " " & yyShpMet%>.</strong><br />&nbsp;</td>
			  </tr>
			  <tr bgcolor="#E7EAEF">
				<td align="left">&nbsp;&nbsp;<a href="adminuspsmeths.asp?method=3"><strong><%=yyEdit & " USPS " & yyShpMet%></strong></a> </td>
				<td>&nbsp; </td>
				<td><input type="button" value="<%=yyEdit&" "&yyShpMet%>" onclick="javascript:document.location='adminuspsmeths.asp?method=3'"></td>
			  </tr>
			  <tr bgcolor="#FFFFFF">
				<td align="left">&nbsp;&nbsp;<a href="adminuspsmeths.asp?method=4"><strong><%=yyEdit & " UPS " & yyShpMet%></strong></a> </td>
				<td><input type="button" value="<%=yyRegUPS%>" onclick="javascript:document.location='adminupslicense.asp'"></td>
				<td><input type="button" value="<%=yyEdit&" "&yyShpMet%>" onclick="javascript:document.location='adminuspsmeths.asp?method=4'"></td>
			  </tr>
			  <tr bgcolor="#E7EAEF">
				<td align="left">&nbsp;&nbsp;<a href="adminuspsmeths.asp?method=6"><strong><%=yyEdit & " " & yyCanPos & " " & yyShpMet%></strong></a> </td>
				<td>&nbsp; </td>
				<td><input type="button" value="<%=yyEdit&" "&yyShpMet%>" onclick="javascript:document.location='adminuspsmeths.asp?method=6'"></td>
			  </tr>
			  <tr bgcolor="#FFFFFF"> 
				<td align="left">&nbsp;&nbsp;<a href="adminuspsmeths.asp?method=7"><strong><%=yyEdit & " FedEx " & yyShpMet%></strong></a> </td>
				<td><input type="button" value="<%=replace(yyRegUPS,"UPS","FedEx")%>" onclick="javascript:document.location='adminfedexlicense.asp'"></td>
				<td><input type="button" value="<%=yyEdit&" "&yyShpMet%>" onclick="javascript:document.location='adminuspsmeths.asp?method=7'"></td>
			  </tr>
			  <tr bgcolor="#FFFFFF">
                <td colspan="3" align="center"><br />&nbsp;<br />&nbsp;</td>
			  </tr>
			</table></td>
        </tr>
<% elseif request.form("posted")="1" AND success then %>
        <tr>
          <td width="100%">
            <table width="100%" border="0" cellspacing="0" cellpadding="2" bgcolor="">
			  <tr> 
                <td width="100%" colspan="2" align="center"><br /><strong><%=yyUpdSuc%></strong><br /><br /><%=yyNowFrd%><br /><br />
                        <%=yyNoAuto%> <A href="admin.asp"><strong><%=yyClkHer%></strong></a>.<br />
                        <br />
				<img src="../images/clearpixel.gif" width="300" height="1" alt="" />
                </td>
			  </tr>
			</table></td>
        </tr>
<% else %>
        <tr>
		  <form method="post" action="adminuspsmeths.asp">
			<td width="100%">
			<input type="hidden" name="posted" value="1" />
			<input type="hidden" name="method" value="<%=method%>" />
            <table width="100%" border="0" cellspacing="0" cellpadding="0" bgcolor="">
			  <tr> 
                <td width="100%" colspan="5" align="center"><strong><%=yyUsUpd & " " & shipmet & " " & yyShpMet%>.</strong><br />&nbsp;</td>
			  </tr>
<% if not success then %>
			  <tr> 
                <td width="100%" colspan="5" align="center"><br /><font color="#FF0000"><%=errmsg%></font>
                </td>
			  </tr>
<% end if %>
<% if shipType=3 then %>
			  <tr>
				<td colspan="5"><ul><li><font size="1"><%=yyUSS1%></font></li>
				<li><font size="1"><%=yyUSS2%> <a href="http://www.usps.com">http://www.usps.com</a>.</font></li></ul></td>
			  </tr>
<% for index=0 to UBOUND(allmethods,2) %>
			  <tr>
			    <td align="right"><%=yyUSPSMe%>:</td>
				<td align="left"><font size="1"><strong><%=allmethods(1,index)%></strong></font></td>
				<td align="center"><%=yyUseMet%></td>
				<td align="center"><acronym title="<%=yyFSApp%>"><%=yyFSA%></acronym></td>
				<td align="center"><%=yyType%></td>
			  </tr>
			  <tr>
				<td align="right"><%=yyShwAs%>:</td>
			    <td align="left"><input type="text" name="methodshow<%=allmethods(0,index)%>" value="<%=allmethods(2,index)%>" size="36" /></td>
				<td align="center"><input type="checkbox" name="methoduse<%=allmethods(0,index)%>" value="ON" <% if Int(allmethods(3,index))=1 then response.write "checked"%> /></td>
				<td align="center"><input type="checkbox" name="methodfsa<%=allmethods(0,index)%>" value="ON" <% if Int(allmethods(5,index))=1 then response.write "checked"%> /></td>
				<td align="center"><%if Int(allmethods(4,index))=1 then response.write "<font color=""#FF0000"">Domestic</font>" else response.write "<font color=""#0000FF"">Internat.</font>"%></td>
			  </tr>
			  <tr>
				<td colspan="5" align="center"><hr width="80%" /></td>
			  </tr>
<% next
   else
		if shipType=4 then %>
			  <tr>
				<td colspan="5"><ul><li><font size="1"><%=yyUSS3%></li>
				<li><font size="1"><%=replace(yyUSS2,"USPS","UPS")%> <a href="http://www.ups.com">http://www.ups.com</a>.</font></li></ul></td>
			  </tr>
<%		else %>
			  <tr>
				<td colspan="5"><ul><li><font size="1">You can use this page to set which <%=shipmet%> shipping methods qualify for free shipping discounts by checking the FSA (Free Shipping Available) checkbox.</li>
				<li><font size="1"><%
			response.write replace(yyUSS2,"USPS",shipmet)
			if shipType=6 then %>
				<a href="http://www.canadapost.ca">http://www.canadapost.ca</a>.
<%			else %>
				<a href="http://www.fedex.com">http://www.fedex.com</a>.
<%			end if %>
				</font></li>
				</ul></td>
			  </tr>
<%		end if
	for index=0 to UBOUND(allmethods,2) %>
			  <tr>
				<input type="hidden" name="methodshow<%=allmethods(0,index)%>" value="1" />
			    <td align="right"><strong><%=yyShipMe%>:</strong></td>
				<td align="left"> <%=allmethods(2,index)%></td>
				<td align="center"><strong><%=IIfVr(shipType=4 OR shipType=7,yyUseMet,"&nbsp;")%></strong></td>
				<td align="center"><acronym title="<%=yyFSApp%>"><%=yyFSA%></acronym></td>
				<td>&nbsp;</td>
			  </tr>
			  <tr>
				<td colspan="2">&nbsp;</td>
				<td align="center"><input type="<%=IIfVr(shipType=4 OR shipType=7,"checkbox","hidden")%>" name="methoduse<%=allmethods(0,index)%>" value="ON" <% if Int(allmethods(3,index))=1 then response.write "checked"%> /></td>
				<td align="center"><input type="checkbox" name="methodfsa<%=allmethods(0,index)%>" value="ON" <% if Int(allmethods(5,index))=1 then response.write "checked"%> /></td>
				<td>&nbsp;</td>
			  </tr>
			  <tr>
				<td colspan="5" align="center"><hr width="80%" /></td>
			  </tr>
<%	next %>
<% end if %>
			  <tr> 
                <td width="100%" colspan="5" align="center"><br /><input type="submit" value="<%=yySubmit%>" /><br />&nbsp;</td>
			  </tr>
            </table></td>
		  </form>
        </tr>
<% end if %>
      </table>