<%
if storesessionvalue="" then storesessionvalue="virtualstore"
if Session("loggedon") <> storesessionvalue OR disallowlogin=TRUE then response.end
Dim sSQL,rs,alldata,allzones,success,cnn,rowcounter,alloptions,errmsg,index,cena,tax
sub writepos(id,pos)
	response.write "<select name='pos"&id&"' size='1'>"
	response.write "<option value='0'>" & yyAlphab & "</option>"
	response.write "<option value='1'"
	if pos=1 then response.write " selected"
	response.write ">"&yyTop&"</option>"
	response.write "<option value='2'"
	if pos=2 then response.write " selected"
	response.write ">"&yyTopTop&"</option></select>"
end sub
sub writezone(id,zone)
	if IsArray(allzones) then
		response.write "<select name='zon"&id&"' size='1'>"
		for index=0 to UBOUND(allzones,2)
			response.write "<option value='"&allzones(0,index)&"'"
			if zone=allzones(0,index) then response.write " selected"
			response.write ">"&allzones(1,index)&"</option>"
		next
		response.write "</select>"
	else
		response.write "No Zones"
	end if
end sub
success=true
Set rs = Server.CreateObject("ADODB.RecordSet")
Set cnn=Server.CreateObject("ADODB.Connection")
cnn.open sDSN
alreadygotadmin = getadminsettings()
Session.LCID=1033
simpleShipping = (adminTweaks AND 1)=1
editzones = (shipType=2 OR shipType=5 OR adminIntShipping=2 OR adminIntShipping=5 OR alternateratesweightbased <> "")
if request.form("posted")="1" then
	for index=1 to 300
		if request.form("pos"&index)<>"" then
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
				if editzones then
					sSQL = "UPDATE countries SET countryEnabled="&cena&",countryTax="&tax&",countryFreeShip="&fsa&",countryOrder="&request.form("pos"&index)&",countryZone="&request.form("zon"&index)&" WHERE countryID="&index
				else
					sSQL = "UPDATE countries SET countryEnabled="&cena&",countryTax="&tax&",countryFreeShip="&fsa&",countryOrder="&request.form("pos"&index)&" WHERE countryID="&index
				end if
				cnn.Execute(sSQL)
			end if
		end if
	next
	if success then
		response.write "<meta http-equiv=""refresh"" content=""3; url=admin.asp"">"
	end if
else
	sSQL = "SELECT countryID,countryName,countryEnabled,countryTax,countryOrder,countryZone,countryFreeShip FROM countries ORDER BY countryOrder DESC,countryName"
	rs.Open sSQL,cnn,0,1
	alldata=rs.getrows
	rs.Close
	sSQL = "SELECT pzID,pzName FROM postalzones WHERE pzName<>'' AND pzID<100"
	rs.Open sSQL,cnn,0,1
	allzones=""
	if NOT rs.EOF then allzones=rs.getrows
	rs.Close
end if
   if request.form("posted")="1" AND success then %>
			<table width="100%" border="0" cellspacing="0" cellpadding="3" bgcolor="">
			  <tr> 
                <td width="100%" colspan="2" align="center"><br /><strong><%=yyUpdSuc%></strong><br /><br /><%=yyNowFrd%><br /><br />
                        <%=yyNoAuto%> <A href="admin.asp"><strong><%=yyClkHer%></strong></a>.<br />
                        <br />
				<img src="../images/clearpixel.gif" width="300" height="3" alt="" />
                </td>
			  </tr>
			</table>
<% elseif request.form("posted")="1" then %>
			<table width="100%" border="0" cellspacing="0" cellpadding="3" bgcolor="">
			  <tr> 
                <td width="100%" colspan="2" align="center"><br /><font color="#FF0000"><strong>Some records could not be updated.</strong></font><br /><br /><%=errmsg%><br /><br />
				<a href="javascript:history.go(-1)"><strong><%=yyClkBac%></strong></a></td>
			  </tr>
			</table>
<% else %>
<script language="javascript" type="text/javascript">
<!--
function wposzon(id,pos,zone,fsa) {
// fsa
document.write('<td><input type=checkbox name="fsa'+id+'"');
if(fsa==1) document.write(' checked');
document.write(' /></td>');
// pos
document.write('<td><select name="pos'+id+'" size="1">');
document.write('<option value="0"><%=yyAlphab%></option>');
document.write('<option value="1"');
if(pos==1)document.write(' selected');
document.write('><%=replace(yyTop,"'","\'")%></option>');
document.write('<option value="2"');
if(pos==2)document.write(' selected');
document.write('><%=replace(yyTopTop,"'","\'")%></option></select></td>');
// zone
<%	if editzones then %>
var foundzone=false;
document.write('<td><select name="zon'+id+'" size="1">');
<%
	if IsArray(allzones) then
		for index=0 to UBOUND(allzones,2)
			response.write "document.write('<option value="""&allzones(0,index)&"""');"&vbCrLf
			response.write "if(zone=="&allzones(0,index)&"){document.write(' selected');foundzone=true;}"&vbCrLf
			response.write "document.write('>"&replace(allzones(1,index),"'","\'")&"</option>');"&vbCrLf
		next
	end if
%>
if(!foundzone)document.write('<option value="0" selected><%=replace(yyUndef,"'","\'")%></option>');
document.write('</select></td>');
<%	else %>
document.write('<td>&nbsp;</td>');
<%	end if %>
}
//-->
</script>
		  <form name="mainform" method="post" action="admincountry.asp">
			<input type="hidden" name="posted" value="1" />
            <table width="100%" border="0" cellspacing="0" cellpadding="1" bgcolor="">
			  <tr> 
                <td width="100%" colspan="6" align="center"><strong><%=yyCntAdm%></strong><br />&nbsp;</td>
			  </tr>
			  <tr> 
                <td width="100%" colspan="6"><ul><li><%=yyHomCou%></li></ul></td>
			  </tr>
			  <tr>
				<td><strong><%=yyCntNam%></strong></td>
				<td align=center><strong><%=yyEnable%></strong></td>
				<td align=center><strong><%=yyTax%></strong></td>
				<td align=center><strong><acronym title="<%=yyFSApp%>"><%=yyFSA%></acronym></strong></td>
				<td align=center><strong><%=yyPosit%></strong></td>
				<td align=center><strong><% if editzones then
											response.write yyPZone
									   else
											response.write "&nbsp;"
									   end if %></strong></td>
			  </tr><%
	if simpleShipping then
		for rowcounter=0 to UBOUND(alldata,2)
			if bgcolor="#E7EAEF" then bgcolor="#FFFFFF" else bgcolor="#E7EAEF"
			%><tr align=center bgcolor="<%=bgcolor%>">
<td align=left><strong><%=alldata(1,rowcounter)%></strong></td>
<td><input type=checkbox name="ena<%=alldata(0,rowcounter)%>"<% if Int(alldata(2,rowcounter))=1 then response.write " checked" %> /></td>
<td><input type=text name="tax<%=alldata(0,rowcounter)%>" value="<%=alldata(3,rowcounter)%>" size="4" />%</td>
<td><input type=checkbox name="fsa<%=alldata(0,rowcounter)%>"<% if alldata(6,rowcounter)=1 then response.write " checked"%> /></td>
<td><select name="pos<%=alldata(0,rowcounter)%>" size="1">
<option value="0"><%=yyAlphab%></option>
<option value="1"<% if alldata(4,rowcounter)=1 then response.write " selected" %>><%=yyTop%></option>
<option value="2"<% if alldata(4,rowcounter)=2 then response.write " selected" %>><%=yyTopTop%></option></select></td><%
			if editzones then
				foundzone=false
				response.write "<td><select name=""zon"&alldata(0,rowcounter)&""" size=""1"">"
				if IsArray(allzones) then
					for index=0 to UBOUND(allzones,2)
						response.write "<option value="""&allzones(0,index)&""""
						if alldata(5,rowcounter)=allzones(0,index) then
							response.write " selected"
							foundzone=true
						end if
						response.write ">"&allzones(1,index)&"</option>"&vbCrLf
					next
				end if
				if NOT foundzone then response.write "<option value=""0"" selected>"&yyUndef&"</option>"
				response.write "</select></td>"
			else
				response.write "<td>&nbsp;</td>"
			end if
			response.write "</tr>"
		next
	else
		for rowcounter=0 to UBOUND(alldata,2)
			if bgcolor="#E7EAEF" then bgcolor="#FFFFFF" else bgcolor="#E7EAEF"
			%><tr align=center bgcolor="<%=bgcolor%>">
<td align=left><strong><%=alldata(1,rowcounter)%></strong></td>
<td><input type=checkbox name="ena<%=alldata(0,rowcounter)%>" <% if Int(alldata(2,rowcounter))=1 then response.write "checked" %> /></td>
<td><input type=text name="tax<%=alldata(0,rowcounter)%>" value="<%=alldata(3,rowcounter)%>" size="4" />%</td>
<%	response.write "<script type=""text/javascript"">wposzon("&alldata(0,rowcounter)&","&alldata(4,rowcounter)&","&alldata(5,rowcounter)&","&alldata(6,rowcounter)&");</script>" %>
</tr><%
		next
	end if
%>			  <tr> 
                <td width="100%" colspan="6" align="center">
                  <p><input type="submit" value="<%=yySubmit%>" />&nbsp;&nbsp;<input type="reset" value="<%=yyReset%>" /><br />&nbsp;</p>
                </td>
			  </tr>
			  <tr> 
                <td width="100%" colspan="6" align="center"><br />
                          <a href="admin.asp"><strong><%=yyAdmHom%></strong></a><br />
				<img src="../images/clearpixel.gif" width="300" height="3" alt="" /></td>
			  </tr>
            </table>
		  </form>
<% end if
cnn.Close
set rs = nothing
set cnn = nothing
%>