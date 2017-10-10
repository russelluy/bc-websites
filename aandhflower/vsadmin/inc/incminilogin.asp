<%
if alreadygotadmin<>TRUE then
	Set rs = Server.CreateObject("ADODB.RecordSet")
	Set cnn=Server.CreateObject("ADODB.Connection")
	cnn.open sDSN
	alreadygotadmin = getadminsettings()
	cnn.Close
	set rs = nothing
	set cnn = nothing
end if
%>
      <table width="130" bgcolor="#FFFFFF">
        <tr> 
          <td class="mincart" bgcolor="#F0F0F0" align="center"><img src="images/minipadlock.gif" align="top" alt="<%=xxMLLIS%>" /> 
            &nbsp;<strong><%=xxMLLIS%></strong></td>
        </tr>
	<% if enableclientlogin<>true then %>
		<tr>
		  <td class="mincart" bgcolor="#F0F0F0" align="center">
		  <p>Client login not enabled</p>
		  </td>
		</tr>
	<% elseif Session("clientUser")<>"" then %>
		<tr>
		  <td class="mincart" bgcolor="#F0F0F0" align="center">
		  <p><%=xxMLLIA%><strong><%=Session("clientUser")%></strong></p>
		  </td>
		</tr>
		<tr> 
          <td class="mincart" bgcolor="#F0F0F0" align="center"><font face='Verdana'>&raquo;</font> <a href="<%=storeurl%>clientlogin.asp?action=logout"><strong>Logout</strong></a></td>
        </tr>
	<% else %>
		<tr>
		  <td class="mincart" bgcolor="#F0F0F0" align="center">
		  <p><%=xxMLNLI%></p>
		  </td>
		</tr>
		<tr> 
          <td class="mincart" bgcolor="#F0F0F0" align="center"><font face='Verdana'>&raquo;</font> <a href="<%=storeurl%>clientlogin.asp"><strong>Login</strong></a></td>
        </tr>
	<% end if %>
      </table>