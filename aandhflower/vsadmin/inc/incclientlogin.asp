<%
Dim sSQL,rs,alldata,success,cnn,errmsg
success=true
if enableclientlogin<>true then
	success=false
	errmsg="Client login not enabled"
end if
Set rs = Server.CreateObject("ADODB.RecordSet")
Set cnn=Server.CreateObject("ADODB.Connection")
cnn.open sDSN
if success AND request.form("posted")="1" then
	theuser = Trim(replace(request.form("user"), "'", ""))
	thepass = Trim(replace(request.form("pass"), "'", ""))
	sSQL = "SELECT clientUser,clientActions,clientLoginLevel,clientPercentDiscount FROM clientlogin WHERE clientUser='"&theuser&"' AND clientPW='"&thepass&"'"
	rs.Open sSQL,cnn,0,1
	if NOT rs.EOF then
		cnn.Execute("DELETE FROM cart WHERE cartCompleted=0 AND cartSessionID="&Session.SessionID)
		Session("clientUser")=theuser
		Session("clientActions")=rs("clientActions")
		Session("clientLoginLevel")=rs("clientLoginLevel")
		Session("clientPercentDiscount")=(100.0-cDbl(rs("clientPercentDiscount")))/100.0
		response.write "<script src='vsadmin/savecookie.asp?WRITECLL="&theuser&"&WRITECLP="&thepass
		if request.form("cook")="ON" then response.write "&permanent=Y"
		response.write "'></script>"
	else
		success=false
		errmsg=xxNoLog
	end if
	rs.Close
	execute("theref = clientloginref" & Session("clientLoginLevel"))
	if theref<>"" then
		if LCase(theref) = "referer" then
			if Trim(request.form("refurl"))<>"" then refURL = Trim(request.form("refurl")) else refURL = xxHomeURL
		else
			refURL = theref
		end if
	elseif clientloginref<>"" then
		if clientloginref="referer" then
			if Trim(request.form("refurl"))<>"" then refURL = Trim(request.form("refurl")) else refURL = xxHomeURL
		else
			refURL = clientloginref
		end if
	else
		refURL = xxHomeURL
	end if
	if success then response.write "<meta http-equiv=""refresh"" content=""3; url="&refURL&""">"
end if
cnn.Close
set rs = nothing
set cnn = nothing
%>
      &nbsp;<br />
	  <table border="0" cellspacing="0" cellpadding="0" width="<%=maintablewidth%>" bgcolor="#B1B1B1" align="center">
<%	if request.querystring("action")="logout" then
		Session.abandon
		response.write "<script src='vsadmin/savecookie.asp?DELCLL=true'></script>"
		if clientlogoutref <> "" then refURL = clientlogoutref else refURL = xxHomeURL
		response.write "<meta http-equiv=""refresh"" content=""3; url="&refURL&""">"
%>
        <tr>
          <td width="100%">
            <table width="100%" border="0" bordercolor="#B1B1B1" cellspacing="1" cellpadding="3" bgcolor="#B1B1B1">
			  <tr>
				<td colspan="2" bgcolor="#FFFFFF">
				  <table width="100%" border="0" cellspacing="0" cellpadding="0" bgcolor="">
					<tr> 
					  <td width="100%" colspan="2" align="center"><br /><strong><%=xxLOSuc%></strong><br /><br /><%=xxAutFo%><br /><br />
                        <%=xxForAut%> <A href="<%=refURL%>"><strong><%=xxClkHere%></strong></a>.<br />
                        <br />
						<img src="../images/clearpixel.gif" width="300" height="3" alt="" />
					  </td>
					</tr>
				  </table>
                </td>
			  </tr>
			</table>
		  </td>
        </tr>
<%	else
		if request.form("posted")="1" AND success then %>
        <tr>
          <td width="100%">
            <table width="100%" border="0" bordercolor="#B1B1B1" cellspacing="1" cellpadding="3" bgcolor="#B1B1B1">
			  <tr>
				<td colspan="2" bgcolor="#FFFFFF">
				  <table width="100%" border="0" cellspacing="0" cellpadding="0" bgcolor="">
					<tr> 
					  <td width="100%" colspan="2" align="center"><br /><strong><%=xxLISuc%></strong><br /><br /><%=xxAutFo%><br /><br />
                        <%=xxForAut%> <A href="<%=refURL%>"><strong><%=xxClkHere%></strong></a>.<br />
                        <br />
						<img src="../images/clearpixel.gif" width="300" height="3" alt="" />
					  </td>
					</tr>
				  </table>
                </td>
			  </tr>
			</table>
		  </td>
        </tr>
<%		else %>
        <tr>
		  <form method="post" action="clientlogin.asp">
		  <td width="100%">
			<input type="hidden" name="posted" value="1" />
			<input type="hidden" name="refurl" value="<%
			refurl = request("refurl")
			if refurl="" then refurl = request.servervariables("HTTP_REFERER")
			response.write refurl %>" />
            <table class="cobtbl" width="100%" border="0" bordercolor="#B1B1B1" cellspacing="1" cellpadding="3" bgcolor="#B1B1B1">
			  <tr>
				<td class="cobll" colspan="2" bgcolor="#FFFFFF" height="34">
				  <table width="100%" border="0" cellspacing="0" cellpadding="0" bgcolor="">
					<tr>
					  <td width="14%" align="center"><img src="images/minipadlock.gif" alt="<%=xxMLLIS%>" /></td><td width="72%" align="center"><font size="4"><strong><%=xxPlEnt%></strong></font></td><td width="14%" align="center" height="30"><img src="images/minipadlock.gif" alt="<%=xxMLLIS%>" /></td>
					</tr>
				  </table>
				</td>
			  </tr>
<%			if not success then %>
			  <tr> 
                <td class="cobll" width="100%" bgcolor="#FFFFFF" height="34" colspan="2" align="center"><font color="#FF0000"><%=errmsg%></font></td>
			  </tr>
<%			end if %>
              <tr> 
                <td class="cobhl" width="40%" bgcolor="#EBEBEB" align="right" height="34"><strong><%=xxLogin%>: </strong></td>
				<td class="cobll" align="left" bgcolor="#FFFFFF" height="34"><input type="text" name="user" size="20" value="<%=request.form("user")%>" /> </td>
			  </tr>
			  <tr> 
                <td class="cobhl" bgcolor="#EBEBEB" align="right" height="34"><strong><%=xxPwd%>: </strong></td>
				<td class="cobll" align="left" bgcolor="#FFFFFF" height="34"><input type="password" name="pass" size="20" value="<%=request.form("pass")%>" /> </td>
			  </tr>
			  <tr> 
                <td class="cobll" align="center" colspan="2" bgcolor="#FFFFFF" height="34"><input type="checkbox" name="cook" value="ON"<% if request.form("cook")="ON" then response.write "checked"%> /> <font size="1"><%=xxWrCk%></font></td>
			  </tr>
			  <tr> 
                <td class="cobll" width="100%" colspan="2" align="center" bgcolor="#FFFFFF" height="34"><input type="submit" value="<%=xxSubmt%>" /><br />
				<img src="../images/clearpixel.gif" width="300" height="3" alt="" /></td>
			  </tr>
            </table>
		  </td>
		  </form>
        </tr>
<%		end if
	end if %>
      </table><br />&nbsp;