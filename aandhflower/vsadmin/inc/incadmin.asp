<%
Dim sSQL,rs,alldata,storeVersion,success,cnn,adoversion,errtext
if storesessionvalue="" then storesessionvalue="virtualstore"
success=0
Set rs = Server.CreateObject("ADODB.RecordSet")
Set cnn=Server.CreateObject("ADODB.Connection")
cnn.open sDSN
adoversion = cnn.Version
if adoversion < 2.5 then
	errtext = errtext & "Your ADO version is less than that required. (2.5).<br />You can update this at <a href='http://www.microsoft.com/data/'>http://www.microsoft.com/data/</a><br />" & vbCrLf
	success = -1
end if
on error resume next
sSQL = "UPDATE admin SET adminID=1 WHERE adminID=1"
cnn.Execute(sSQL)
if err.number<>0 then
	errtext = errtext & "Your database is not writeable. This probably means you just need to set the permissions on the directory the database is in to allow writing.<br />Your host can help you with this.<br />" & vbCrLf
	success = -1
end if
on error goto 0
if debugmode=TRUE then
	errtext = errtext & yyDebug & "<br />" & vbCrLf
	success = -1
end if
if Session("loggedon") <> storesessionvalue AND Trim(request.cookies("WRITECKL"))<>"" then
	sSQL="SELECT adminID FROM admin WHERE adminUser='" & Replace(request.cookies("WRITECKL"),"'","''") & "' AND adminPassword='" & Replace(request.cookies("WRITECKP"),"'","''") & "' AND adminID=1"
	rs.Open sSQL,cnn,0,1
	if NOT (rs.EOF OR rs.BOF) then
		Session("loggedon") = storesessionvalue
	else
		success=2
	end if
	rs.Close
end if
if (Session("loggedon") <> storesessionvalue AND success<>2) OR disallowlogin=TRUE then response.end
sSQL = "SELECT adminShipping,adminVersion,adminUser,adminPassword FROM admin WHERE adminID=1"
rs.Open sSQL,cnn,0,1
shipType = Int(rs("adminShipping"))
storeVersion = rs("adminVersion")
adminUser = rs("adminUser")
adminPassword = rs("adminPassword")
rs.Close
cnn.Close
set rs = nothing
set cnn = nothing
if Trim(request.querystring("writeck"))="yes" then
	response.write "<script src='savecookie.asp?WRITECKL=" & adminUser & "&WRITECKP=" & adminPassword & "'></script>"
	response.write "<meta http-equiv=""Refresh"" content=""3; URL=admin.asp"">"
	success=1
elseif Trim(request.querystring("writeck"))="no" then
	response.write "<script src='savecookie.asp?DELCK=yes'></script>"
	response.write "<meta http-equiv=""Refresh"" content=""3; URL=admin.asp"">"
	success=1
end if
%>
      <table border="0" cellspacing="0" cellpadding="0" width="100%" bgcolor="" align="center">
        <tr> 
          <td width="100%">
            <table width="100%" border="0" cellspacing="0" cellpadding="3" bgcolor="">
			  <tr> 
			    <td colspan="2"><span id="testspanid"></span></td>
			  </tr>
<% if success < 0 then %>
			  <tr> 
			    <td colspan="2"> 
				  <p><font size="2" color="#DD0000"><strong><%=yySorOut%></strong></font> <br />
				  <%=errtext%></p>
			    </td>
			  </tr>
<% end if %>
<% if success=1 then %>
			  <tr> 
				<td colspan="2" width="100%" align="center"><p>&nbsp;</p><p>&nbsp;</p>
				  <p><strong><%=yyOpSuc%></strong></p><p>&nbsp;</p>
				  <p><font size="1"><%=yyNowFrd%><br /><br /><%=yyNoAuto%> <a href="admin.asp"><%=yyClkHer%></a>.</font></td>
			  </tr>
<% elseif success=2 then %>
			  <tr> 
				<td colspan="2" width="100%" align="center"><p>&nbsp;</p><p>&nbsp;</p>
				  <p><strong><%=yyOpFai%></strong></p><p>&nbsp;</p>
				  <p><%=yyCorCoo%> <%=yyCorLI%> <a href="login.asp"><%=yyClkHer%></a>.</p></td>
			  </tr>
<% else %>
			  <tr> 
                <td colspan="2" width="100%" align="center"><strong><%=yyChsLst%></strong><br /><font size="1">(<%=yyVers%>: <%=storeVersion%>)</font><br />&nbsp;</td>
			  </tr>
			  <tr> 
				<td valign="top" width="50%" align="left">&nbsp;&nbsp;<a href="adminorders.asp"><strong><%=yyVwOrd%></strong></a><br />
                        &nbsp;
                </td>
				<td valign="top" width="50%"><a href="<%=helpbaseurl%>help.asp#orders" target="ttshelp"></a></td>
			  </tr>
			  <tr> 
				<td width="50%" align="left">&nbsp;&nbsp;<a href="adminlogin.asp"><strong><%=yyCngPw%></strong></a> </td>
				<td width="50%"><a href="<%=helpbaseurl%>help.asp#uname" target="ttshelp"></a></td>
			  </tr>
              <tr> 
				<td width="50%" align="left">&nbsp;&nbsp;<a href="adminmain.asp"><strong><%=yyEdAdm%></strong></a> </td>
				<td width="50%"><a href="<%=helpbaseurl%>help.asp#admin" target="ttshelp"></a></td>
			  </tr>
			  <tr> 
				<td width="50%" align="left">&nbsp;&nbsp;<a href="adminaffil.asp"><strong><%=yyVwAff%></strong></a> </td>
				<td width="50%"><a href="<%=helpbaseurl%>help.asp#affiliate" target="ttshelp"></a></td>
			  </tr>
			  <tr> 
				<td width="50%" align="left">&nbsp;&nbsp;<a href="adminprods.asp"><strong><%=yyEdPrd%></strong></a> </td>
				<td width="50%"><a href="<%=helpbaseurl%>help.asp#prods" target="ttshelp"></a></td>
			  </tr>
			  <tr> 
				<td width="50%" align="left">&nbsp;&nbsp;<a href="adminprodopts.asp"><strong><%=yyEdOpt%></strong></a> </td>
				<td width="50%"><a href="<%=helpbaseurl%>help.asp#prodopt" target="ttshelp"></a></td>
			  </tr>
			  <tr> 
				<td width="50%" align="left">&nbsp;&nbsp;<a href="adminpricebreak.asp"><strong><%=yyEdPrBk%></strong></a> </td>
				<td width="50%"><a href="<%=helpbaseurl%>help.asp#pricebreak" target="ttshelp"></a></td>
			  </tr>
			  <tr> 
				<td width="50%" align="left">&nbsp;&nbsp;<a href="admincats.asp"><strong><%=yyEdCat%></strong></a> </td>
				<td width="50%"><a href="<%=helpbaseurl%>help.asp#cats" target="ttshelp"></a></td>
			  </tr>
			  <tr> 
				<td width="50%" align="left">&nbsp;&nbsp;<a href="admindiscounts.asp"><strong><%=yyDisCou%></strong></a> </td>
				<td width="50%"><a href="<%=helpbaseurl%>help.asp#discounts" target="ttshelp"></a></td>
			  </tr>
			  <tr> 
				<td width="50%" align="left">&nbsp;&nbsp;<a href="adminclientlog.asp"><strong><%=yyCliLog%></strong></a> </td>
				<td width="50%"><a href="<%=helpbaseurl%>help.asp#clientlogin" target="ttshelp"></a></td>
			  </tr>
			  <tr> 
				<td width="50%" align="left">&nbsp;&nbsp;<a href="adminstate.asp"><strong><%=yyEdSta%></strong></a> </td>
				<td width="50%"><a href="<%=helpbaseurl%>help.asp#state" target="ttshelp"></a></td>
			  </tr>
			  <tr> 
				<td width="50%" align="left">&nbsp;&nbsp;<a href="admincountry.asp"><strong><%=yyEdCnt%></strong></a> </td>
				<td width="50%"><a href="<%=helpbaseurl%>help.asp#country" target="ttshelp"></a></td>
			  </tr>
			  <tr> 
				<td width="50%" align="left">&nbsp;&nbsp;<a href="adminzones.asp"><strong><%=yyEdPzon%></strong></a> </td>
				<td width="50%"><a href="<%=helpbaseurl%>help.asp#pzone" target="ttshelp"></a></td>
			  </tr>
			  <tr> 
				<td width="50%" align="left">&nbsp;&nbsp;<a href="adminuspsmeths.asp"><strong><%=yyShmReg%></strong></a> </td>
				<td width="50%"><a href="<%=helpbaseurl%>help.asp#shipmeth" target="ttshelp"></a></td>
			  </tr>
			  <tr> 
                <td width="50%" align="left">&nbsp;&nbsp;<a href="adminpayprov.asp"><strong><%=yyEdPPro%></strong></a></td>
				<td width="50%"><a href="<%=helpbaseurl%>help.asp#payprov" target="ttshelp"></a></td>
			  </tr>
			  <tr> 
                <td width="50%" align="left">&nbsp;&nbsp;<a href="adminordstatus.asp"><strong><%=yyEdOSta%></strong></a></td>
				<td width="50%"><a href="<%=helpbaseurl%>help.asp#ordstat" target="ttshelp"></a></td>
			  </tr>
			  <tr> 
                <td width="50%" align="left">&nbsp;&nbsp;<a href="admindropship.asp"><strong><%=yyEdDrSp%></strong></a></td>
				<td width="50%"><a href="<%=helpbaseurl%>help.asp#droshp" target="ttshelp"></a></td>
			  </tr>
			  <tr> 
                <td width="50%" align="left">&nbsp;&nbsp;<a href="admincsv.asp"><strong><%=yyCSVUp%></strong></a></td>
				<td width="50%"><a href="<%=helpbaseurl%>help.asp#csv" target="ttshelp"></a></td>
			  </tr>
			  <tr> 
                <td width="50%" align="left">&nbsp;&nbsp;<a href="adminipblock.asp"><strong><%=yyIPBlock%></strong></a></td>
				<td width="50%"><a href="<%=helpbaseurl%>help.asp#ipblock" target="ttshelp"></a></td>
			  </tr>
			  <tr> 
				<td colspan="2" width="100%" align="left">&nbsp;&nbsp;<a href="logout.asp"><strong><%=yyLOut%></strong></a> </td>
			  </tr>
			  <tr> 
				<td colspan="2" width="100%" align="center"><p>&nbsp;</p>
				<%	if Trim(request.cookies("WRITECKL"))<>"" then %>
					<a href="admin.asp?writeck=no"><%=yyDelCoo%></a><br />
				<%	else %>
					<a href="admin.asp?writeck=yes"><%=yyWrCoo%></a><br /><font size="1"><%=yyNoRec%></font>
				<%	end if %>
				</td>
			  </tr>
<%		if nocheckdatabasedownload<>TRUE then %>
<script language="javascript" type="text/javascript">
function getwarnmessage(){
	return('<p><STRONG><font color="#FF0000">WARNING!!</font></STRONG> It may be that your database is downloadable. This may mean that someone could download your database and gain access to your admin username and password.<a href="http://www.beancastle.com>http://www.beancastle.com</a></p>');
}
function checkstatechange(){
	if(ckAJAX.readyState==4){
		if(ckAJAX.status==200){
			document.getElementById("testspanid").innerHTML=getwarnmessage();
		}
	}
}
function checkstatechange2(){
	if(ckAJAX2.readyState==4){
		if(ckAJAX2.status==200){
			document.getElementById("testspanid").innerHTML=getwarnmessage();
		}
	}
}
if(window.XMLHttpRequest){
	ckAJAX = new XMLHttpRequest();
	ckAJAX2 = new XMLHttpRequest();
}else{
	ckAJAX = new ActiveXObject("MSXML2.XMLHTTP");
	ckAJAX2 = new ActiveXObject("MSXML2.XMLHTTP");
}
ckAJAX.onreadystatechange = checkstatechange;
ckAJAX.open("GET", "../fpdb/vsproducts.mdb", true);
ckAJAX.send(null);
setTimeout('ckAJAX.abort();',1000);
ckAJAX2.onreadystatechange = checkstatechange2;
ckAJAX2.open("GET", "/fpdb/vsproducts.mdb", true);
ckAJAX2.send(null);
setTimeout('ckAJAX2.abort();',1100);
</script>
<%		end if %>
<% end if %>
			  <tr> 
                <td colspan="2" width="100%" align="left"><img src="../images/clearpixel.gif" width="300" height="5">
                </td>
			  </tr>
            </table>
          </td>
        </tr>
      </table>