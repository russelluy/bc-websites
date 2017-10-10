<!--#include file="incemail.asp"-->
<!--#include file="md5.asp"-->
<%
Dim netnav, success
success = true
digidownloads=false
Set rs = Server.CreateObject("ADODB.RecordSet")
Set rs2 = Server.CreateObject("ADODB.RecordSet")
Set cnn=Server.CreateObject("ADODB.Connection")
cnn.open sDSN
alreadygotadmin = getadminsettings()
if Request.Form("posted")="1" then
	email = Trim(Replace(Request.form("email"), "'", ""))
	ordid = Trim(Replace(Request.form("ordid"), "'", ""))
	if NOT IsNumeric(ordid) then
		success = false
		errormsg = xxStaEr1
	elseif email<>"" AND ordid<>"" then
		sSQL = "SELECT ordStatus,ordStatusDate,"&getlangid("statPublic",64)&",ordTrackNum,ordAuthNumber,ordStatusInfo FROM orders INNER JOIN orderstatus ON orders.ordStatus=orderstatus.statID WHERE ordID=" & ordid & " AND ordEmail='" & email & "'"
		rs.Open sSQL,cnn,0,1
			if NOT rs.EOF then
				ordStatus = rs("ordStatus")
				ordStatusDate = rs("ordStatusDate")
				statPublic = rs(getlangid("statPublic",64))
				ordAuthNumber = trim(rs("ordAuthNumber")&"")
				ordStatusInfo = trim(rs("ordStatusInfo")&"")
				ordTrackNum = trim(rs("ordTrackNum")&"")
				if trackingnumtext = "" then trackingnumtext=xxTrackT
				if ordTrackNum <> "" then trackingnum=replace(trackingnumtext, "%s", ordTrackNum) else trackingnum=""
				trackingnum = replace(trackingnum, "%nl%", "<br>")
				' if dateadjust<>"" then ordStatusDate = DateAdd("h",dateadjust,ordStatusDate)
			else
				success = false
				errormsg = xxStaEr2
			end if
		rs.Close
	else
		success = false
		errormsg = xxStaEnt
	end if
end if
%>
<br />
		<form method="post" name="statusform" action="orderstatus.asp">
		  <input type="hidden" name="posted" value="1" />
<%		if Request.Form("posted")="1" AND success then %>
			<table class="cobtbl" width="<%=maintablewidth%>" border="0" bordercolor="#B1B1B1" cellspacing="1" cellpadding="3" bgcolor="#B1B1B1">
			  <tr>
				<td class="cobhl" colspan="2" bgcolor="#EBEBEB" height="34" align="center"><font size="4"><strong><%=xxStaVw%></strong></font></td>
			  </tr>
			  <tr>
			    <td class="cobhl" bgcolor="#EBEBEB" height="34" align="center" colspan="2"><strong><%=xxStaCur & " " & ordid%></strong></td>
			  </tr>
			  <tr>
			    <td class="cobhl" bgcolor="#EBEBEB" height="34" align="right" width="40%"><strong><%=xxStatus%> : </strong></td>
				<td class="cobll" bgcolor="#FFFFFF" height="34" ><%=statPublic%></td>
			  </tr>
			  <tr>
			    <td class="cobhl" bgcolor="#EBEBEB" height="34" align="right" width="40%"><strong><%=xxDate%> : </strong></td>
				<td class="cobll" bgcolor="#FFFFFF" height="34" ><%=FormatDateTime(ordStatusDate, 1)%></td>
			  </tr>
			  <tr>
			    <td class="cobhl" bgcolor="#EBEBEB" height="34" align="right" width="40%"><strong><%=xxTime%> : </strong></td>
				<td class="cobll" bgcolor="#FFFFFF" height="34" ><%=FormatDateTime(ordStatusDate, 4)%></td>
			  </tr>
<%			if trackingnum<>"" then %>
			  <tr>
			    <td class="cobhl" bgcolor="#EBEBEB" height="34" align="right" width="40%"><strong><%=xxTraNum%> : </strong></td>
				<td class="cobll" bgcolor="#FFFFFF" height="34" ><%=trackingnum%></td>
			  </tr>
<%			end if
			if ordStatusInfo<>"" then %>
			  <tr>
			    <td class="cobhl" bgcolor="#EBEBEB" height="34" align="right" width="40%"><strong><%=xxAddInf%> : </strong></td>
				<td class="cobll" bgcolor="#FFFFFF" height="34" ><%=ordStatusInfo%></td>
			  </tr>
<%			end if 
			if ordAuthNumber>"" then %>
			  <tr>
				<td class="cobll" bgcolor="#FFFFFF" colspan="2" align="center"><%
					xxThkYou=""
					xxRecEml=""
					call do_order_success(ordid,"",FALSE,TRUE,FALSE,FALSE,FALSE) %></td>
			  </tr>
<%			end if
		else %>
            <table class="cobtbl" width="<%=maintablewidth%>" border="0" bordercolor="#B1B1B1" cellspacing="1" cellpadding="3" bgcolor="#B1B1B1">
			  <tr>
				<td class="cobhl" colspan="2" bgcolor="#EBEBEB" align="center">&nbsp;<br /><font size="4"><strong><%=xxStaVw%></strong></font><br />&nbsp;</td>
			  </tr>
<%		end if %>
			  <tr>
			    <td class="cobhl" colspan="2" bgcolor="#EBEBEB" height="34" align="center"><strong><%=xxStaEnt%></strong></td>
			  </tr>
<%		if success=false then %>
			  <tr>
			    <td class="cobhl" width="40%" bgcolor="#EBEBEB" height="34" align="right"><strong><%=xxStaErr%> : </strong></td>
				<td class="cobll" bgcolor="#FFFFFF" height="34"><font color="#FF0000"><%=errormsg%></font></td>
			  </tr>
<%		end if %>
			  <tr>
			    <td class="cobhl" width="40%" bgcolor="#EBEBEB" height="34" align="right"><strong><%=xxOrdId%> : </strong></td>
				<td class="cobll" bgcolor="#FFFFFF" height="34"><input type="text" size="20" name="ordid" value="<%=Trim(Request.Form("ordid"))%>" /></td>
			  </tr>
			  <tr>
			    <td class="cobhl" width="40%" bgcolor="#EBEBEB" height="34" align="right"><strong><%=xxEmail%> : </strong></td>
				<td class="cobll" bgcolor="#FFFFFF" height="34"><input type="text" size="30" name="email" value="<%=Trim(Request.Form("email"))%>" /></td>
			  </tr>
			  <tr>
				<td class="cobll" bgcolor="#FFFFFF" height="34" colspan="2" align="center" valign="middle"><input type="submit" value="<%=xxStaVw%>" /></td>
			  </tr>
			</form>
			</table>
		  <br />&nbsp;
<%
set rs = nothing
set cnn = nothing
%>