<%
if storesessionvalue="" then storesessionvalue="virtualstore"
if Session("loggedon") <> storesessionvalue OR disallowlogin=TRUE then response.end
Dim sSQL,rs,alldata,alladmin,success,cnn,errmsg
success=true
Set rs = Server.CreateObject("ADODB.RecordSet")
Set cnn=Server.CreateObject("ADODB.Connection")
cnn.open sDSN
sSQL = "SELECT adminEmail, adminStoreURL, adminUser, adminPassword FROM admin WHERE adminID=1"
rs.Open sSQL,cnn,0,1
alladmin=rs.getrows
rs.Close
if request.form("posted")="1" then
	if Trim(request.form("pass")) <> Trim(request.form("pass2")) then
		success = false
		errmsg=yyNoMat
	else
		sSQL = "UPDATE admin SET adminUser='"&request.form("user")&"',adminPassword='"&request.form("pass")&"' WHERE adminID=1"
		on error resume next
		cnn.Execute(sSQL)
		if err.number<>0 then
			success=false
			errmsg = "There was an error writing to the database.<br />"
			if err.number = -2147467259 then
				errmsg = errmsg & "Your database does not have write permissions."
			else
				errmsg = errmsg & err.description
			end if
		else
			response.write "<meta http-equiv=""refresh"" content=""3; url=admin.asp"">"
		end if
		on error goto 0
	end if
end if
cnn.Close
set rs = nothing
set cnn = nothing
%>
      <table border="0" cellspacing="0" cellpadding="0" width="100%" bgcolor="" align="center">
<% if request.form("posted")="1" AND success then %>
        <tr>
          <td width="100%">
            <table width="100%" border="0" cellspacing="0" cellpadding="3" bgcolor="">
			  <tr> 
                <td width="100%" colspan="2" align="center"><br /><strong><%=yyUpdSuc%></strong><br /><br /><%=yyNowFrd%><br /><br />
                        <%=yyNoAuto%><A href="admin.asp"><strong><%=yyClkHer%></strong></a>.<br />
                        <br />
				<img src="../images/clearpixel.gif" width="300" height="3" alt="" />
                </td>
			  </tr>
			</table></td>
        </tr>
<% else %>
        <tr>
		        <form method="post" action="adminlogin.asp">
                  <td width="100%">
			<input type="hidden" name="posted" value="1" />
            <table width="100%" border="0" cellspacing="0" cellpadding="3" bgcolor="">
			  <tr> 
                <td width="100%" colspan="2" align="center"><br /><strong><%=yyNewUN%></strong>
                </td>
			  </tr>
<% if not success then %>
			  <tr> 
                <td width="100%" colspan="2" align="center"><br /><font color="#FF0000"><%=errmsg%></font>
                </td>
			  </tr>
<% end if %>
              <tr> 
                <td width="50%" align="right"><strong><%=yyUname%>: </strong>
                </td>
				<td width="50%" align="left"><input type="text" name="user" size="20" value="<%=alladmin(2,0)%>" /> 
                </td>
			  </tr>
			  <tr> 
                <td width="50%" align="right"><strong><%=yyPass%>: </strong>
                </td>
				<td width="50%" align="left"><input type="password" name="pass" size="20" value="<%=alladmin(3,0)%>" /> 
                </td>
			  </tr>
			  <tr> 
                <td width="50%" align="right"><strong><%=yyPassCo%>: </strong>
                </td>
				<td width="50%" align="left"><input type="password" name="pass2" size="20" value="<%=alladmin(3,0)%>" /> 
                </td>
			  </tr>
			  <tr> 
                <td width="100%" colspan="2" align="center"><br /><input type="submit" value="<%=yySubmit%>" /><br />
				<img src="../images/clearpixel.gif" width="300" height="3" alt="" /></td>
			  </tr>
            </table></td>
		  </form>
        </tr>
<% end if %>
      </table>