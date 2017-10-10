<%
if storesessionvalue="" then storesessionvalue="virtualstore"
if Session("loggedon") <> storesessionvalue OR disallowlogin=TRUE then response.end
Dim sSQL,rs,alldata,allzones,success,cnn,rowcounter,alloptions,errmsg,index,cena,tax
success=true
Set rs = Server.CreateObject("ADODB.RecordSet")
Set cnn=Server.CreateObject("ADODB.Connection")
cnn.open sDSN
alreadygotadmin = getadminsettings()
Session.LCID=1033
simpleShipping = (adminTweaks AND 1)=1
if request.form("posted")="1" then
	for index=1 to 150
		if request.form("id"&index)<>"" then
			cena=0
			if request.form("ena"&index)<>"" then cena=1
			fsa=0
			if request.form("fsa"&index)<>"" then fsa=1
			tax = request.form("tax"&index)
			if NOT IsNumeric(tax) then
				success=false
				errmsg = yyNum100 & " """ & yyTax & """."
			elseif tax > 100 OR tax < 0 then
				success=false
				errmsg = yyNum100 & " """ & yyTax & """."
			else
				if splitUSZones then
					sSQL = "UPDATE states SET stateEnabled="&cena&",stateTax="&tax&",stateFreeShip="&fsa&",stateZone="&request.form("zon"&index)&" WHERE stateID="&index
				else
					sSQL = "UPDATE states SET stateEnabled="&cena&",stateTax="&tax&",stateFreeShip="&fsa&" WHERE stateID="&index
				end if
				cnn.Execute(sSQL)
			end if
		end if
	next
	if success then
		response.write "<meta http-equiv=""refresh"" content=""3; url=admin.asp"">"
	end if
else
	sSQL = "SELECT stateID,stateName,stateEnabled,stateTax,stateZone,stateFreeShip FROM states ORDER BY stateName"
	rs.Open sSQL,cnn,0,1
	alldata=rs.getrows
	rs.Close
	sSQL = "SELECT pzID,pzName FROM postalzones WHERE pzName<>'' AND pzID>100"
	rs.Open sSQL,cnn,0,1
	allzones=""
	if NOT rs.EOF then allzones=rs.getrows
	rs.Close
end if
%>
      <table border="0" cellspacing="0" cellpadding="0" width="100%" bgcolor="" align="center">
<% if request.form("posted")="1" AND success then %>
        <tr>
          <td width="100%">
			<table width="100%" border="0" cellspacing="0" cellpadding="3" bgcolor="">
			  <tr> 
                <td width="100%" colspan="2" align="center"><br /><strong><%=yyUpdSuc%></strong><br /><br /><%=yyNowFrd%><br /><br />
                        <%=yyNoAuto%> <A href="admin.asp"><strong><%=yyClkHer%></strong></a>.<br />
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
                <td width="100%" colspan="2" align="center"><br /><font color="#FF0000"><strong>Some records could not be updated.</strong></font><br /><br /><%=errmsg%><br /><br />
				<a href="javascript:history.go(-1)"><strong><%=yyClkBac%></strong></a></td>
			  </tr>
			</table></td>
        </tr>
<% else %>
<script language="javascript" type="text/javascript">
<!--
function writezone(id,zone) {
var foundzone=false;
document.write('<select name="zon'+id+'" size="1">');
<%
	if IsArray(allzones) then
		for index=0 to UBOUND(allzones,2)
			response.write "document.write('<option value="""&allzones(0,index)&"""');"&vbCrLf
			response.write "if(zone=="&allzones(0,index)&"){document.write(' selected');foundzone=true;}"&vbCrLf
			response.write "document.write('>"&allzones(1,index)&"</option>');"&vbCrLf
		next
	end if
%>
if(!foundzone)document.write('<option value="0" selected><%=yyUndef%></option>');
document.write('</select>');
}
//-->
</script>
        <tr>
		  <form name="mainform" method="post" action="adminstate.asp">
		  <td width="100%">
			<input type="hidden" name="posted" value="1" />
            <table width="100%" border="0" cellspacing="0" cellpadding="1" bgcolor="">
			  <tr> 
                <td width="100%" colspan="5" align="center"><strong><%=yyStaAdm%></strong><br /><br />
				<font size="1"><%=yyFSANot%><br />&nbsp;</font>
				</td>
			  </tr>
			  <tr>
				<td><strong><%=yyStaNam%></strong></td>
				<td align=center><strong><%=yyEnable%></strong></td>
				<td align=center><strong><%=yyTax%></strong></td>
				<td align=center><strong><acronym title="<%=yyFSApp%>"><%=yyFSA%></acronym></strong></td>
				<td align=center><strong><% if splitUSZones then
										response.write yyPZone
									else
										response.write "&nbsp;"
									end if %></strong></td>
			  </tr><%
	if simpleShipping then
		for rowcounter=0 to UBOUND(alldata,2)
			if bgcolor="#E7EAEF" then bgcolor="#FFFFFF" else bgcolor="#E7EAEF"
			%><tr align=center bgcolor="<%=bgcolor%>">
<td align=left><strong><%=alldata(1,rowcounter)%></strong><input type="hidden" name="id<%=alldata(0,rowcounter)%>" value="1" /></td>
<td><input type=checkbox name="ena<%=alldata(0,rowcounter)%>" <% if Int(alldata(2,rowcounter))=1 then response.write "checked" %> /></td>
<td><input type=text name="tax<%=alldata(0,rowcounter)%>" value="<%=alldata(3,rowcounter)%>" size="4" />%</td>
<td><input type=checkbox name="fsa<%=alldata(0,rowcounter)%>"<% if alldata(5,rowcounter)=1 then response.write " checked"%> /></td>
	<td><%	if splitUSZones then
				foundzone=false
				response.write "<select name=""zon"&alldata(0,rowcounter)&""" size=""1"">"
				if IsArray(allzones) then
					for index=0 to UBOUND(allzones,2)
						response.write "<option value="""&allzones(0,index)&""""
						if alldata(4,rowcounter)=allzones(0,index) then
							response.write " selected"
							foundzone=true
						end if
						response.write ">"&allzones(1,index)&"</option>"
					next
				end if
				if NOT foundzone then response.write "<option value=""0"" selected>"&yyUndef&"</option>"
				response.write "</select>"
			else
				response.write "&nbsp;"
			end if %></td></tr>
<%
		next
	else
		for rowcounter=0 to UBOUND(alldata,2)
			if bgcolor="#E7EAEF" then bgcolor="#FFFFFF" else bgcolor="#E7EAEF"
			%><tr align=center bgcolor="<%=bgcolor%>">
<td align=left><strong><%=alldata(1,rowcounter)%></strong><input type="hidden" name="id<%=alldata(0,rowcounter)%>" value="1" /></td>
<td><input type=checkbox name="ena<%=alldata(0,rowcounter)%>" <% if Int(alldata(2,rowcounter))=1 then response.write "checked" %> /></td>
<td><input type=text name="tax<%=alldata(0,rowcounter)%>" value="<%=alldata(3,rowcounter)%>" size="4" />%</td>
<td><input type=checkbox name="fsa<%=alldata(0,rowcounter)%>"<% if alldata(5,rowcounter)=1 then response.write " checked"%> /></td>
<td><%	if splitUSZones then
			response.write "<script type=""text/javascript"">writezone("&alldata(0,rowcounter)&","&alldata(4,rowcounter)&");</script>"
		else
			response.write "&nbsp;"
		end if %></td>
</tr><%
		next
	end if
%>			  <tr> 
                <td width="100%" colspan="5" align="center">
                  <p><input type="submit" value="<%=yySubmit%>" />&nbsp;&nbsp;<input type="reset" value="<%=yyReset%>" /><br />&nbsp;</p>
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
<% end if
cnn.Close
set rs = nothing
set cnn = nothing
%>
      </table>