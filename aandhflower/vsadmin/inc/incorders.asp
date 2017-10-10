<%
Dim sSQL,rs,alldata,success,cnn,errmsg,rowcounter,startfont,endfont,sd,ed,smonth,allorders,addcomma,delStr,delOptions,ordAddInfo,ordCNum
if storesessionvalue="" then storesessionvalue="virtualstore"
netnav = false
if htmlemails=true then emlNl = "<br />" else emlNl=vbCrLf
if instr(Request.ServerVariables("HTTP_USER_AGENT"), "Gecko") > 0 then netnav = true
lisuccess=0
if dateadjust="" then dateadjust=0
thedate = DateAdd("h",dateadjust,Now())
thedate = DateSerial(year(thedate),month(thedate),day(thedate))
if request.querystring("doedit")="true" then doedit=TRUE else doedit=FALSE
function editfunc(data,col,size)
	if doedit then editfunc = "<input type=""text"" id="""&col&""" name="""&col&""" value="""&replace(data&"","""","&quot;")&""" size="""&size&""">" else editfunc = data
end function
function editnumeric(data,col,size)
	if doedit then editnumeric = "<input type=""text"" id="""&col&""" name="""&col&""" value="""&replace(FormatNumber(data,2),",","")&""" size="""&size&""">" else editnumeric = FormatEuroCurrency(data)
end function
function getNumericField(fldname)
	fldval = Trim(Request.Form(fldname))
	if NOT IsNumeric(fldval) then getNumericField=0.0 else getNumericField=cDbl(fldval)
end function
Set rs = Server.CreateObject("ADODB.RecordSet")
Set rs2 = Server.CreateObject("ADODB.RecordSet")
Set rsl = Server.CreateObject("ADODB.RecordSet")
Set cnn=Server.CreateObject("ADODB.Connection")
cnn.open sDSN
if Session("loggedon") <> storesessionvalue AND Trim(request.cookies("WRITECKL"))<>"" then
	sSQL="SELECT adminID FROM admin WHERE adminUser='" & Replace(request.cookies("WRITECKL"),"'","''") & "' AND adminPassword='" & Replace(request.cookies("WRITECKP"),"'","''") & "' AND adminID=1"
	rs.Open sSQL,cnn,0,1
	if NOT (rs.EOF OR rs.BOF) then
		Session("loggedon") = storesessionvalue
	else
		lisuccess=2
	end if
	rs.Close
end if
if (Session("loggedon") <> storesessionvalue AND lisuccess<>2) OR disallowlogin=TRUE then response.end
Sub release_stock(smOrdId)
	if stockManage <> 0 then
		sSQL="SELECT cartID,cartProdID,cartQuantity,pStockByOpts FROM cart INNER JOIN products ON cart.cartProdID=products.pID WHERE cartCompleted=1 AND cartOrderID=" & smOrdId
		rsl.Open sSQL,cnn,0,1
		do while NOT rsl.EOF
			if cint(rsl("pStockByOpts")) <> 0 then
				sSQL = "SELECT coOptID FROM cartoptions INNER JOIN (options INNER JOIN optiongroup ON options.optGroup=optiongroup.optGrpID) ON cartoptions.coOptID=options.optID WHERE optType IN (-2,-1,1,2) AND coCartID=" & rsl("cartID")
				rs.Open sSQL,cnn,0,1
				do while NOT rs.EOF
					sSQL = "UPDATE options SET optStock=optStock+"&rsl("cartQuantity")&" WHERE optID="&rs("coOptID")
					cnn.Execute(sSQL)
					rs.MoveNext
				loop
				rs.Close
			else
				sSQL = "UPDATE products SET pInStock=pInStock+"&rsl("cartQuantity")&" WHERE pID='"&rsl("cartProdID")&"'"
				cnn.Execute(sSQL)
			end if
			rsl.MoveNext
		loop
		rsl.Close
	end if
End Sub
if lisuccess=2 then
%>
	  <table border="0" cellspacing="0" cellpadding="0" width="100%" bgcolor="" align="center">
        <tr>
          <td width="100%">
            <table width="100%" border="0" cellspacing="0" cellpadding="3" bgcolor="">
			  <tr> 
                <td width="100%" colspan="4" align="center"><p>&nbsp;</p><p>&nbsp;</p>
				  <p><strong><%=yyOpFai%></strong></p><p>&nbsp;</p>
				  <p><%=yyCorCoo%> <%=yyCorLI%> <a href="login.asp"><%=yyClkHer%></a>.</p>
				</td>
			  </tr>
			</table>
		  </td>
		</tr>
	  </table>
<%
else
success=true
alreadygotadmin = getadminsettings()
if Request.Form("updatestatus")="1" then
	cnn.Execute("UPDATE orders SET ordTrackNum='"&Replace(request.form("ordTrackNum"),"'","''")&"',ordStatusInfo='"&Replace(request.form("ordStatusInfo"),"'","''")&"',ordInvoice='"&Replace(request.form("ordInvoice"),"'","''")&"' WHERE ordID="&Request.Form("orderid"))
elseif Request.QueryString("id")<>"" then
	if Request.Form("delccdets")<>"" then
		sSQL = "UPDATE orders SET ordCNum='' WHERE ordID="&Request.QueryString("id")
		cnn.Execute(sSQL)
	end if
	sSQL = "SELECT cartProdId,cartProdName,cartProdPrice,cartQuantity,cartID FROM cart WHERE cartOrderID="&Request.QueryString("id")
	rs.Open sSQL,cnn,0,1
	allorders = ""
	if NOT rs.EOF then allorders=rs.getrows
	rs.Close
else
	if delccafter<>0 then
		tdt = thedate-delccafter
		sSQL = "UPDATE orders SET ordCNum='' WHERE ordDate<"&datedelim & VSUSDate(tdt) & datedelim
		cnn.Execute(sSQL)
	end if
	if delAfter<>0 then
		tdt = thedate-delAfter
		sSQL = "SELECT cartOrderID,cartID FROM cart WHERE cartCompleted=0 AND cartDateAdded<"&datedelim & VSUSDate(tdt) & datedelim
		rs.Open sSQL,cnn,0,1
		if NOT rs.EOF then
			delOptions=""
			do while NOT rs.EOF
				delOptions = delOptions & addcomma & rs("cartID")
				addcomma = ","
				rs.MoveNext
			loop
			cnn.Execute("DELETE FROM cartoptions WHERE coCartID IN ("&delOptions&")")
			cnn.Execute("DELETE FROM cart WHERE cartID IN ("&delOptions&")")
		end if
		rs.Close
		cnn.Execute("DELETE FROM orders WHERE ordAuthNumber='' AND ordDate<" & datedelim & VSUSDate(tdt) & datedelim & " AND ordStatus=2")
	else
		tdt = thedate - 3
		sSQL = "SELECT cartOrderID,cartID FROM cart WHERE cartCompleted=0 AND cartOrderID=0 AND cartDateAdded<"&datedelim & VSUSDate(tdt) & datedelim
		rs.Open sSQL,cnn,0,1
		if NOT rs.EOF then
			delStr=""
			delOptions=""
			do while NOT rs.EOF
				delStr = delStr & addcomma & rs("cartOrderID")
				delOptions = delOptions & addcomma & rs("cartID")
				addcomma = ","
				rs.MoveNext
			loop
			cnn.Execute("DELETE FROM cartoptions WHERE coCartID IN ("&delOptions&")")
			cnn.Execute("DELETE FROM cart WHERE cartID IN ("&delOptions&")")
		end if
		rs.Close
	end if
	sSQL = "SELECT statID,statPrivate FROM orderstatus WHERE statPrivate<>'' ORDER BY statID"
	rs.Open sSQL,cnn,0,1
		allstatus=rs.GetRows
	rs.Close
end if
if Request.Form("updatestatus")="1" then
%>
<script language="javascript" type="text/javascript">
<!--
setTimeout("history.go(-2);",1100);
// -->
</script>
	  <table border="0" cellspacing="0" cellpadding="0" width="100%" bgcolor="" align="center">
        <tr>
          <td width="100%">
            <table width="100%" border="0" cellspacing="0" cellpadding="2" bgcolor="">
			  <tr> 
                <td width="100%" colspan="4" align="center"><br /><strong><%=yyUpdSuc%></strong><br /><br /><%=yyNowFrd%><br /><br />
                        <%=yyNoAuto%> <a href="javascript:history.go(-2)"><strong><%=yyClkHer%></strong></a>.<br /><br />
						<img src="../images/clearpixel.gif" width="300" height="3" alt="" /></td>
			  </tr>
			</table>
		  </td>
		</tr>
	  </table>
<%
elseif Request.Form("doedit")="true" then
	session.LCID = 1033
	OWSP = ""
	if mysqlserver then rs.CursorLocation = 3
	rs.Open "SELECT * FROM orders WHERE ordID="&request.form("orderid"),cnn,1,3,&H0001
	thesessionid = rs.Fields("ordSessionID")
	rs.Fields("ordName")		= trim(request.form("name"))
	rs.Fields("ordAddress")		= trim(request.form("address"))
	if useaddressline2=TRUE then rs.Fields("ordAddress2") = trim(request.form("address2"))
	rs.Fields("ordCity")		= trim(request.form("city"))
	rs.Fields("ordState")		= trim(request.form("state"))
	rs.Fields("ordZip")			= trim(request.form("zip"))
	rs.Fields("ordCountry")		= trim(request.form("country"))
	rs.Fields("ordEmail")		= trim(request.form("email"))
	rs.Fields("ordPhone")		= trim(request.form("phone"))
	rs.Fields("ordShipName")	= trim(request.form("sname"))
	rs.Fields("ordShipAddress")	= trim(request.form("saddress"))
	if useaddressline2=TRUE then rs.Fields("ordShipAddress2") = trim(request.form("saddress2"))
	rs.Fields("ordShipCity")	= trim(request.form("scity"))
	rs.Fields("ordShipState")	= trim(request.form("sstate"))
	rs.Fields("ordShipZip")		= trim(request.form("szip"))
	rs.Fields("ordShipCountry")	= trim(request.form("scountry"))
	rs.Fields("ordShipType")	= trim(request.form("shipmethod"))
	rs.Fields("ordShipCarrier")	= trim(request.form("shipcarrier"))
	rs.Fields("ordIP")			= trim(request.form("ipaddress"))
	ordComLoc=0
	if Trim(request.form("commercialloc"))="Y" then ordComLoc = 1
	if Trim(request.form("wantinsurance"))="Y" then ordComLoc = ordComLoc + 2
	if Trim(request.form("saturdaydelivery"))="Y" then ordComLoc = ordComLoc + 4
	if Trim(request.form("signaturerelease"))="Y" then ordComLoc = ordComLoc + 8
	if Trim(request.form("insidedelivery"))="Y" then ordComLoc = ordComLoc + 16
	rs.Fields("ordComLoc")		= ordComLoc
	rs.Fields("ordAffiliate")	= trim(Request.Form("PARTNER"))
	rs.Fields("ordAddInfo")		= trim(Request.Form("ordAddInfo"))
	rs.Fields("ordStatusInfo")	= trim(Request.Form("ordStatusInfo"))
	rs.Fields("ordTrackNum")	= trim(Request.Form("ordTrackNum"))
	discounttext = replace(trim(Request.Form("discounttext")), vbCrLf, "<br />")
	discounttext = replace(discounttext, vbCr, "<br />")
	rs.Fields("ordDiscountText")= replace(discounttext, vbLf, "<br />")
	rs.Fields("ordInvoice")		= Trim(Request.Form("ordInvoice"))
	rs.Fields("ordExtra1")		= Trim(Request.Form("ordextra1"))
	rs.Fields("ordExtra2")		= Trim(Request.Form("ordextra2"))
	rs.Fields("ordExtra3")		= Trim(Request.Form("ordextra3"))
	rs.Fields("ordShipping")	= getNumericField("ordShipping")
	if canadataxsystem=true then rs.Fields("ordHSTTax") = getNumericField("ordHSTTax")
	rs.Fields("ordStateTax")	= getNumericField("ordStateTax")
	rs.Fields("ordCountryTax")	= getNumericField("ordCountryTax")
	rs.Fields("ordDiscount")	= getNumericField("ordDiscount")
	rs.Fields("ordHandling")	= getNumericField("ordHandling")
	rs.Fields("ordAuthNumber")	= Trim(Request.Form("ordAuthNumber"))
	rs.Fields("ordTransID")		= Trim(Request.Form("ordTransID"))
	rs.Fields("ordTotal")		= getNumericField("ordtotal")
	rs.Update
	rs.Close
	Dim forminorder()
	redim forminorder(100)
	formitemcnt = 0
	for jj = 1 to Request.Form.Count
		for each objElem in Request.Form
			if Request.Form(objElem) is Request.Form(jj) then
				if Left(objElem,6)="prodid" OR Left(objElem,4)="optn" then
					forminorder(formitemcnt) = objElem
					formitemcnt = formitemcnt + 1
					forminorderubound = UBOUND(forminorder)
					if formitemcnt > forminorderubound then redim preserve forminorder(forminorderubound+100)
				end if
				exit for
			end if
		next
	next
	for jj = 0 to formitemcnt-1
		objForm = forminorder(jj)
		' response.write objForm & " : " & Request.Form(objForm) & "<br>"
		if Left(objForm,6)="prodid" then
			idno = trim(right(objForm, Len(objForm)-6))
			cartid = trim(request.form("cartid"&idno))
			prodid = trim(request.form("prodid"&idno))
			quant = trim(request.form("quant"&idno))
			theprice = trim(request.form("price"&idno))
			prodname = trim(request.form("prodname"&idno))
			delitem = trim(request.form("del_"&idno))
			if delitem="yes" then
				cnn.Execute("DELETE FROM cart WHERE cartID=" & cartid)
				cnn.Execute("DELETE FROM cartoptions WHERE coCartID=" & cartid)
				cartid = ""
			elseif cartid<>"" then
				Session.LCID = 1033
				sSQL = "UPDATE cart SET cartProdID='"&replace(prodid,"'","''")&"',cartProdPrice="&theprice&",cartProdName='"&replace(prodname,"'","''")&"',cartQuantity="&quant&" WHERE cartID="&cartid
				cnn.Execute(sSQL)
				Session.LCID = saveLCID
				cnn.Execute("DELETE FROM cartoptions WHERE coCartID=" & cartid)
			else
				rs.Open "cart",cnn,1,3,&H0002
				rs.AddNew
				rs.Fields("cartSessionID")		= thesessionid
				rs.Fields("cartProdID")			= prodid
				rs.Fields("cartQuantity")		= quant
				rs.Fields("cartCompleted")		= 1
				rs.Fields("cartProdName")		= prodname
				rs.Fields("cartProdPrice")		= theprice
				rs.Fields("cartDateAdded")		= DateAdd("h",dateadjust,Now())
				rs.Fields("cartOrderID")		= request.form("orderid")
				rs.Update
				if mysqlserver=true then
					rs.Close
					rs.Open "SELECT LAST_INSERT_ID() AS lstIns",cnn,0,1
					cartid = rs("lstIns")
				else
					cartid = rs.Fields("cartID")
				end if
				rs.Close
			end if
			if cartid<>"" then
				optprefix = "optn"&idno&"_"
				prefixlen = len(optprefix)
				for kk = 0 to formitemcnt-1
					objForm = forminorder(kk)
					if Left(objForm,prefixlen)=optprefix AND trim(Request.Form(objForm))<>"" then
						optidarr = split(Request.Form(objForm),"|")
						optid = optidarr(0)
						if Trim(Request.Form("v"&objForm))="" then
							sSQL="SELECT optID,"&getlangid("optGrpName",16)&","&getlangid("optName",32)&","&OWSP&"optPriceDiff,optWeightDiff,optType,optFlags FROM options INNER JOIN optiongroup ON options.optGroup=optiongroup.optGrpID WHERE optID="&Replace(optid,"'","")
							rs.Open sSQL,cnn,0,1
							if abs(rs("optType"))<> 3 then
								sSQL = "INSERT INTO cartoptions (coCartID,coOptID,coOptGroup,coCartOption,coPriceDiff,coWeightDiff) VALUES ("&cartID&","&rs("optID")&",'"&Replace(rs(getlangid("optGrpName",16))&"","'","''")&"','"&Replace(rs(getlangid("optName",32))&"","'","''")&"',"
								sSQL = sSQL & optidarr(1) & ",0)"
							else
								sSQL = "INSERT INTO cartoptions (coCartID,coOptID,coOptGroup,coCartOption,coPriceDiff,coWeightDiff) VALUES ("&cartID&","&rs("optID")&",'"&Replace(rs(getlangid("optGrpName",16))&"","'","''")&"','',0,0)"
							end if
							rs.Close
							cnn.Execute(sSQL)
						else
							sSQL="SELECT optID,"&getlangid("optGrpName",16)&","&getlangid("optName",32)&" FROM options INNER JOIN optiongroup ON options.optGroup=optiongroup.optGrpID WHERE optID="&replace(optid,"'","")
							rs.Open sSQL,cnn,0,1
							sSQL = "INSERT INTO cartoptions (coCartID,coOptID,coOptGroup,coCartOption,coPriceDiff,coWeightDiff) VALUES ("&cartID&","&rs("optID")&",'"&Replace(rs(getlangid("optGrpName",16))&"","'","''")&"','"&replace(trim(Request.Form("v"&objForm)),"'","''")&"',0,0)"
							cnn.Execute(sSQL)
							rs.Close
						end if
					end if
				next
			end if
		end if
	next
	session.LCID = saveLCID
%>
<script language="javascript" type="text/javascript">
<!--
setTimeout("history.go(-2);",1100);
// -->
</script>
	  <table border="0" cellspacing="0" cellpadding="0" width="100%" bgcolor="" align="center">
        <tr>
          <td width="100%">
            <table width="100%" border="0" cellspacing="0" cellpadding="2" bgcolor="">
			  <tr> 
                <td width="100%" colspan="4" align="center"><br /><strong><%=yyUpdSuc%></strong><br /><br /><%=yyNowFrd%><br /><br />
                        <%=yyNoAuto%> <a href="javascript:history.go(-2)"><strong><%=yyClkHer%></strong></a>.<br /><br />
						<img src="../images/clearpixel.gif" width="300" height="3" alt="" /></td>
			  </tr>
			</table>
		  </td>
		</tr>
	  </table>
<%
elseif Request.QueryString("id")<>"" then
	statetaxrate=0
	countrytaxrate=0
	hsttaxrate=0
	countryorder=0
	sSQL = "SELECT ordID,ordName,ordAddress,ordAddress2,ordCity,ordState,ordZip,ordCountry,ordEmail,ordPhone,ordShipName,ordShipAddress,ordShipAddress2,ordShipCity,ordShipState,ordShipZip,ordShipCountry,ordPayProvider,ordAuthNumber,ordTransID,ordTotal,ordDate,ordStateTax,ordCountryTax,ordShipping,ordShipType,ordShipCarrier,ordIP,ordAffiliate,ordDiscount,ordDiscountText,ordHandling,ordComLoc,ordExtra1,ordExtra2,ordExtra3,ordHSTTax,ordTrackNum,ordInvoice,ordAddInfo FROM orders INNER JOIN payprovider ON payprovider.payProvID=orders.ordPayProvider WHERE ordID="&Request.QueryString("id")
	rs.Open sSQL,cnn,0,1
	if doedit then
		Session.LCID = 1033
		response.write "<form method=""post"" name=""editform"" action=""adminorders.asp"" onsubmit=""return confirmedit()""><input type=""hidden"" name=""orderid"" value="""&Request.QueryString("id")&""" /><input type=""hidden"" name=""doedit"" value=""true"" />"
		overridecurrency=TRUE
		orcsymbol=""
		orcdecplaces=2
		orcpreamount=true
	end if
%>
<script language="javascript" type="text/javascript">
<!--
var newwin="";
var plinecnt=0;
function openemailpopup(id) {
  popupWin = window.open('popupemail.asp?'+id,'emailpopup','menubar=no, scrollbars=no, width=300, height=250, directories=no,location=no,resizable=yes,status=no,toolbar=no')
}
function updateoptions(id){
	prodid = document.getElementById('prodid'+id).value;
	if(prodid != ''){
		newwin = window.open('popupemail.asp?prod='+prodid+'&index='+id,'updateopts','menubar=no, scrollbars=no, width=50, height=40, directories=no,location=no,resizable=yes,status=no,toolbar=no');
	}
	return(false);
}
function extraproduct(plusminus){
var productspan=document.getElementById('productspan');
if(plusminus=='+'){
productspan.innerHTML=productspan.innerHTML.replace(/<!--NEXTPRODUCTCOMMENT-->/,'<!--PLINE'+plinecnt+'--><tr><td valign="top"><input type="button" value="..." onclick="updateoptions('+(plinecnt+1000)+')">&nbsp;<input name="prodid'+(plinecnt+1000)+'" size="18" id="prodid'+(plinecnt+1000)+'"></td><td valign="top"><input type="text" id="prodname'+(plinecnt+1000)+'" name="prodname'+(plinecnt+1000)+'" size="24"></td><td><span id="optionsspan'+(plinecnt+1000)+'">-</span></td><td valign="top"><input type="text" id="quant'+(plinecnt+1000)+'" name="quant'+(plinecnt+1000)+'" size="5" value="1"></td><td valign="top"><input type="text" id="price'+(plinecnt+1000)+'" name="price'+(plinecnt+1000)+'" value="0" size="7"><br /><input type="hidden" id="optdiffspan'+(plinecnt+1000)+'" value="0"></td><td>&nbsp;</td></tr><!--PLINEEND'+plinecnt+'--><!--NEXTPRODUCTCOMMENT-->');
plinecnt++;
}else{
if(plinecnt>0){
plinecnt--;
var restr = '<!--PLINE'+plinecnt+'-->(.|\\n)+<!--PLINEEND'+plinecnt+'-->';
//alert(restr);
var re = new RegExp(restr);
productspan.innerHTML=productspan.innerHTML.replace(re,'');
}
}
}
function confirmedit(){
if(confirm('<%=replace(yyChkRec,"'","\'")%>'))
	return(true);
return(false);
}
function dorecalc(onlytotal){
var thetotal=0,totoptdiff=0;
for(var i in document.forms.editform){
if(i.substr(0,5)=="quant"){
	theid = i.substr(5);
	totopts=0;
	delbutton = document.getElementById("del_"+theid);
	if(delbutton==null)
		isdeleted=false;
	else
		isdeleted=delbutton.checked;
	if(! isdeleted){
	for(var ii in document.forms.editform){
		var opttext="optn"+theid+"_";
		if(ii.substr(0,opttext.length)==opttext){
			theitem = document.getElementById(ii);
			if(document.getElementById('v'+ii)==null){
				thevalue = theitem[theitem.selectedIndex].value;
				if(thevalue.indexOf('|')>0){
					totopts += parseFloat(thevalue.substr(thevalue.indexOf('|')+1));
				}
			}
		}
	}
	thequant = parseInt(document.getElementById(i).value);
	if(isNaN(thequant)) thequant=0;
	theprice = parseFloat(document.getElementById("price"+theid).value);
	if(isNaN(theprice)) theprice=0;
	document.getElementById("optdiffspan"+theid).value=totopts;
	optdiff = parseFloat(document.getElementById("optdiffspan"+theid).value);
	if(isNaN(optdiff)) optdiff=0;
	thetotal += thequant * (theprice + optdiff);
	totoptdiff += thequant * optdiff;
	}
}
}
document.getElementById("optdiffspan").innerHTML=totoptdiff.toFixed(2);
document.getElementById("ordtotal").value = thetotal.toFixed(2);
if(onlytotal==true) return;
statetaxrate = parseFloat(document.getElementById("staterate").value);
if(isNaN(statetaxrate)) statetaxrate=0;
countrytaxrate = parseFloat(document.getElementById("countryrate").value);
if(isNaN(countrytaxrate)) countrytaxrate=0;
discount = parseFloat(document.getElementById("ordDiscount").value);
if(isNaN(discount)){
	discount=0;
	document.getElementById("ordDiscount").value=0;
}
statetaxtotal = (statetaxrate * (thetotal-discount)) / 100.0;
countrytaxtotal = (countrytaxrate * (thetotal-discount)) / 100.0;
shipping = parseFloat(document.getElementById("ordShipping").value);
if(isNaN(shipping)){
	shipping=0;
	document.getElementById("ordShipping").value=0;
}
handling = parseFloat(document.getElementById("ordHandling").value);
if(isNaN(handling)){
	handling=0;
	document.getElementById("ordHandling").value=0;
}
<%	if taxShipping=2 then %>
statetaxtotal += (statetaxrate * shipping) / 100.0;
countrytaxtotal += (countrytaxrate * shipping) / 100.0;
<%	end if
	if taxHandling=2 then %>
statetaxtotal += (statetaxrate * handling) / 100.0;
countrytaxtotal += (countrytaxrate * handling) / 100.0;
<%	end if %>
document.getElementById("ordStateTax").value = statetaxtotal.toFixed(2);
document.getElementById("ordCountryTax").value = countrytaxtotal.toFixed(2);
hstobj = document.getElementById("ordHSTTax");
hsttax=0;
if(! (hstobj==null)){
	hsttax = parseFloat(hstobj.value);
}
grandtotal = (thetotal + shipping + handling + statetaxtotal + countrytaxtotal + hsttax) - discount;
document.getElementById("grandtotalspan").innerHTML = grandtotal.toFixed(2);
}
function ajaxcallback() {
	if(ajaxobj.readyState==4){
		document.getElementById("googleupdatespan").innerHTML = ajaxobj.responseText;
	}
}
function updategoogleorder(theact,ordid){
	if(confirm('Inform Google of change to order id ' + ordid + "?")){
		document.getElementById("googleupdatespan").innerHTML = '';
		if(window.XMLHttpRequest){
			ajaxobj = new XMLHttpRequest();
		}else{
			ajaxobj = new ActiveXObject("MSXML2.XMLHTTP");
		}
		ajaxobj.onreadystatechange = ajaxcallback;
		extraparams='';
		if(theact=='ship'){
			shipcar = document.getElementById("shipcarrier");
			if(shipcar!= null){
				trackno=document.getElementById("ordTrackNum").value
				if(trackno!='' && confirm('Include tracking and carrier info?')){
					extraparams='&carrier='+(shipcar.options[shipcar.selectedIndex].value)+'&trackno='+document.getElementById("ordTrackNum").value;
				}
			}
		}
		document.getElementById("googleupdatespan").innerHTML = 'Connecting...';
		ajaxobj.open("GET", "ajaxservice.asp?gid="+ordid+"&act="+theact+extraparams, true);
		ajaxobj.send(null);
	}
}
function updategooglestatus(theact,ordid){
	if(confirm('Update Google account status and inform customer of this status change?')){
		document.getElementById("googleupdatespan").innerHTML = '';
		if(window.XMLHttpRequest){
			ajaxobj = new XMLHttpRequest();
		}else{
			ajaxobj = new ActiveXObject("MSXML2.XMLHTTP");
		}
		ajaxobj.onreadystatechange = ajaxcallback;
		themessage="googlemessage=" + encodeURI(document.getElementById("ordStatusInfo").value);
		document.getElementById("googleupdatespan").innerHTML = 'Connecting...';
		ajaxobj.open("POST", "ajaxservice.asp?gid="+ordid+"&act="+theact, true);
		ajaxobj.setRequestHeader("Content-type", "application/x-www-form-urlencoded");
		ajaxobj.setRequestHeader('Content-Length', themessage.length);
		ajaxobj.send(themessage);
	}
}
//-->
</script>
	  <table border="0" cellspacing="0" cellpadding="0" width="100%" bgcolor="" align="center">
        <tr>
          <td width="100%">
            <table width="100%" border="0" cellspacing="0" cellpadding="2" bgcolor="">
<%			if isprinter AND invoiceheader<>"" then %>
			  <tr> 
                <td width="100%" colspan="4"><%=invoiceheader%></td>
			  </tr>
<%			end if %>
			  <tr> 
                <td width="100%" colspan="4" align="center"><strong><%=xxOrdNum & " " & rs("ordID") & "<br /><br />" & FormatDateTime(rs("ordDate"), 1) & " " & FormatDateTime(rs("ordDate"), 4) %></strong></td>
			  </tr>
<%			if isprinter AND invoiceaddress<>"" then %>
			  <tr> 
                <td width="100%" colspan="4"><%=invoiceaddress%></td>
			  </tr>
<%			end if
			if Trim(extraorderfield1)<>"" then %>
			<tr>
			  <td width="20%" align="right"><strong><%=extraorderfield1 %>:</strong></td>
			  <td align="left" colspan="3"><%=editfunc(rs("ordExtra1"),"ordExtra1",25)%></td>
			</tr>
<%			end if %>
			<tr>
			  <td width="20%" align="right"><strong><%=xxName%>:</strong></td>
			  <td width="30%" align="left"><%=editfunc(rs("ordName"),"name",25)%></td>
			  <td width="20%" align="right"><% if NOT isprinter AND (rs("ordAuthNumber")&"") <> "" AND NOT doedit then response.write "<input type=""button"" value=""Resend"" onclick=""javascript:openemailpopup('id=" & rs("ordID") & "')"" />"%>
			  <strong><%=xxEmail%>:</strong></td>
			  <td width="30%" align="left"><%
				if isprinter OR doedit then response.write editfunc(rs("ordEmail"),"email",25) else response.write "<a href=""mailto:"&rs("ordEmail")&""">"&rs("ordEmail")&"</a>" %></td>
			</tr>
			<tr>
			  <td align="right"><strong><%=xxAddress%>:</strong></td>
			  <td align="left"<% if useaddressline2=TRUE then response.write " colspan=""3"""%>><%=editfunc(rs("ordAddress"),"address",25)%></td>
<%	if useaddressline2=TRUE then %>
			</tr>
			<tr>
			  <td align="right"><strong><%=xxAddress2%>:</strong></td>
			  <td align="left"><%=editfunc(rs("ordAddress2"),"address2",25)%></td>
<%	end if %>
			  <td align="right"><strong><%=xxCity%>:</strong></td>
			  <td align="left"><%=editfunc(rs("ordCity"),"city",25)%></td>
			</tr>
			<tr>
			  <td align="right"><strong><%=xxAllSta%>:</strong></td>
			  <td align="left"><%=editfunc(rs("ordState"),"state",25)%></td>
			  <td align="right"><strong><%=xxCountry%>:</strong></td>
			  <td align="left"><%
			if doedit then
				foundmatch=FALSE
				response.write "<select name=""country"" size=""1"">"
				sSQL = "SELECT countryName,countryTax,countryOrder FROM countries ORDER BY countryOrder DESC, countryName"
				rs2.Open sSQL,cnn,0,1
				do while not rs2.EOF
					response.write "<option value="""&Replace(rs2("countryName"),"""","&quot;")&""""
					if rs("ordCountry")=rs2("countryName") then
						response.write " selected"
						foundmatch=TRUE
						countrytaxrate=rs2("countryTax")
						countryorder=rs2("countryOrder")
					end if
					response.write ">"&rs2("countryName")&"</option>"&vbCrLf
					rs2.MoveNext
				loop
				rs2.Close
				if NOT foundmatch then response.write "<option value="""&Replace(rs("ordCountry"),"""","&quot;")&""" selected>"&rs("ordCountry")&"</option>"&vbCrLf
				response.write "</select>"
				if countryorder=2 then
					sSQL = "SELECT stateTax FROM states WHERE stateName='"&replace(rs("ordState"),"'","''")&"'"
					rs2.Open sSQL,cnn,0,1
					if NOT rs2.EOF then statetaxrate = rs2("stateTax")
					rs2.Close
				end if
			else
				response.write rs("ordCountry")
			end if %></td>
			</tr>
			<tr>
			  <td align="right"><strong><%=xxZip%>:</strong></td>
			  <td align="left"><%=editfunc(rs("ordZip"),"zip",15)%></td>
			  <td align="right"><strong><%=xxPhone%>:</strong></td>
			  <td align="left"><%=editfunc(rs("ordPhone"),"phone",25)%></td>
			</tr>
			<%	if Trim(extraorderfield2)<>"" then %>
			<tr>
			  <td align="right"><strong><% response.write extraorderfield2 %>:</strong></td>
			  <td align="left" colspan="3"><%=editfunc(rs("ordExtra2"),"ordextra2",25)%></td>
			</tr>
			<%	end if %>
<% if NOT isprinter then %>
			<tr>
			  <td align="right"><strong>IP Address:</strong></td>
			  <td align="left"><%=editfunc(rs("ordIP"),"ipaddress",15)%></td>
			  <td align="right"><strong><%=yyAffili%>:</strong></td>
			  <td align="left"><%=editfunc(rs("ordAffiliate"),"PARTNER",15)%></td>
			</tr>
<% end if
   if Trim(rs("ordDiscountText"))<>"" OR doedit then %>
			<tr>
			  <td align="right" valign="top"><strong><%=xxAppDs%>:</strong></td>
			  <td align="left" colspan="3"><% if doedit then response.write "<textarea name=""discounttext"" cols=""50"" rows=""2"" wrap=virtual>" & replace(rs("ordDiscountText")&"","<br />",vbNewLine) & "</textarea>" else response.write rs("ordDiscountText") %></td>
			</tr>
<% end if
   if trim(rs("ordShipName")&"")<>"" OR trim(rs("ordShipAddress")&"")<>"" OR trim(rs("ordShipCity")&"")<>"" OR trim(rs("ordExtra3")&"")<>"" OR doedit then %>
			<tr>
			  <td width="100%" align="center" colspan="4"><strong><%=xxShpDet%>.</strong></td>
			</tr>
<%			if (trim(extraorderfield3)<>"" AND (trim(rs("ordExtra3")&"")<>"" OR doedit)) then %>
			<tr>
			  <td width="20%" align="right"><strong><%=extraorderfield3 %>:</strong></td>
			  <td align="left" colspan="3"><%=editfunc(rs("ordExtra3"),"ordExtra3",25)%></td>
			</tr>
<%			end if %>
			<tr>
			  <td align="right"><strong><%=xxName%>:</strong></td>
			  <td align="left" colspan="3"><%=editfunc(rs("ordShipName"),"sname",25)%></td>
			</tr>
			<tr>
			  <td align="right"><strong><%=xxAddress%>:</strong></td>
			  <td align="left"<% if useaddressline2=TRUE then response.write " colspan=""3"""%>><%=editfunc(rs("ordShipAddress"),"saddress",25)%></td>
<%	if useaddressline2=TRUE then %>
			</tr>
			<tr>
			  <td align="right"><strong><%=xxAddress2%>:</strong></td>
			  <td align="left"><%=editfunc(rs("ordShipAddress2"),"saddress2",25)%></td>
<%	end if %>
			  <td align="right"><strong><%=xxCity%>:</strong></td>
			  <td align="left"><%=editfunc(rs("ordShipCity"),"scity",25)%></td>
			</tr>
			<tr>
			  <td align="right"><strong><%=xxAllSta%>:</strong></td>
			  <td align="left"><%=editfunc(rs("ordShipState"),"sstate",25)%></td>
			  <td align="right"><strong><%=xxCountry%>:</strong></td>
			  <td align="left"><%
			if doedit then
				if trim(rs("ordShipName")&"")<>"" OR trim(rs("ordShipAddress")&"")<>"" then usingshipcountry=TRUE else usingshipcountry=FALSE
				foundmatch=FALSE
				response.write "<select name=""scountry"" size=""1"">"
				sSQL = "SELECT countryName,countryTax,countryOrder FROM countries ORDER BY countryOrder DESC, countryName"
				rs2.Open sSQL,cnn,0,1
				do while not rs2.EOF
					response.write "<option value="""&Replace(rs2("countryName"),"""","&quot;")&""""
					if rs("ordShipCountry")=rs2("countryName") then
						response.write " selected"
						foundmatch=TRUE
						if usingshipcountry then countrytaxrate=rs2("countryTax")
						countryorder=rs2("countryOrder")
					end if
					response.write ">"&rs2("countryName")&"</option>"&vbCrLf
					rs2.MoveNext
				loop
				rs2.Close
				if NOT foundmatch then response.write "<option value="""&replace(trim(rs("ordShipCountry")&""),"""","&quot;")&""" selected>"&rs("ordShipCountry")&"</option>"&vbCrLf
				response.write "</select>"
				if countryorder=2 AND usingshipcountry then
					sSQL = "SELECT stateTax FROM states WHERE stateName='"&replace(trim(rs("ordShipState")&""),"'","''")&"'"
					rs2.Open sSQL,cnn,0,1
					if NOT rs2.EOF then statetaxrate = rs2("stateTax")
					rs2.Close
				end if
			else
				response.write rs("ordShipCountry")
			end if %></td>
			</tr>
			<tr>
			  <td align="right"><strong><%=xxZip%>:</strong></td>
			  <td align="left" colspan="3"><%=editfunc(rs("ordShipZip"),"szip",15)%></td>
			</tr>
<% end if
	if rs("ordShipCarrier")<>0 OR rs("ordShipType")<>"" OR doedit then %>
			<tr>
			  <td align="right"><strong><%=xxShpMet%>:</strong></td>
			  <td align="left"><%	if NOT isprinter then %>
					<select name="shipcarrier" id="shipcarrier" size="1">
					<option value="0"><%=yyOther%></option>
					<option value="3" <%if Int(rs("ordShipCarrier"))=3 then response.write "selected"%>>USPS</option>
					<option value="4" <%if Int(rs("ordShipCarrier"))=4 then response.write "selected"%>>UPS</option>
					<option value="6" <%if Int(rs("ordShipCarrier"))=6 then response.write "selected"%>>CanPos</option>
					<option value="7" <%if Int(rs("ordShipCarrier"))=7 then response.write "selected"%>>FedEx</option>
					<option value="8" <%if Int(rs("ordShipCarrier"))=8 then response.write "selected"%>>DHL</option>
					</select> <%	end if
									response.write editfunc(rs("ordShipType"),"shipmethod",25) %></td>
			  <td align="right"><strong><% if doedit then response.write xxCLoc & ":"%></strong></td>
			  <td align="left"><%	if doedit then
										response.write "<select name=""commercialloc"" size=""1"">"
										response.write "<option value=""N"">"&yyNo&"</option>"
										response.write "<option value=""Y"""&IIfVr((rs("ordComLoc") AND 1)=1," selected","")&">"&yyYes&"</option>"
										response.write "</select>"
									end if %></td>
			</tr>
<%		if doedit then %>
			<tr>
			  <td align="right"><strong><%=xxShpIns%>:</strong></td>
			  <td align="left"><%	response.write "<select name=""wantinsurance"" size=""1"">"
									response.write "<option value=""N"">"&yyNo&"</option>"
									response.write "<option value=""Y"""&IIfVr((rs("ordComLoc") AND 2)=2," selected","")&">"&yyYes&"</option>"
									response.write "</select>" %></td>
			  <td align="right"><strong><%=xxSatDe2%>:</strong></td>
			  <td align="left"><%	response.write "<select name=""saturdaydelivery"" size=""1"">"
									response.write "<option value=""N"">"&yyNo&"</option>"
									response.write "<option value=""Y"""&IIfVr((rs("ordComLoc") AND 4)=4," selected","")&">"&yyYes&"</option>"
									response.write "</select>" %></td>
			</tr>
			<tr>
			  <td align="right"><strong><%=xxSigRe2%>:</strong></td>
			  <td align="left"><%	response.write "<select name=""signaturerelease"" size=""1"">"
									response.write "<option value=""N"">"&yyNo&"</option>"
									response.write "<option value=""Y"""&IIfVr((rs("ordComLoc") AND 8)=8," selected","")&">"&yyYes&"</option>"
									response.write "</select>" %></td>
			  <td align="right"><strong><%=xxInsDe2%>:</strong></td>
			  <td align="left"><%	response.write "<select name=""insidedelivery"" size=""1"">"
									response.write "<option value=""N"">"&yyNo&"</option>"
									response.write "<option value=""Y"""&IIfVr((rs("ordComLoc") AND 16)=16," selected","")&">"&yyYes&"</option>"
									response.write "</select>" %></td>
			</tr>
<%		elseif rs("ordComLoc")>0 then
			shipopts="<strong>Shipping options:</strong>"
			if (rs("ordComLoc") AND 1)=1 then response.write "<tr><td align=""right"">"&shipopts&"</td><td align=""left"" colspan=""3"">"&xxCerCLo&"</td></tr>" : shipopts=""
			if (rs("ordComLoc") AND 2)=2 then response.write "<tr><td align=""right"">"&shipopts&"</td><td align=""left"" colspan=""3"">"&xxShiInI&"</td></tr>" : shipopts=""
			if (rs("ordComLoc") AND 4)=4 then response.write "<tr><td align=""right"">"&shipopts&"</td><td align=""left"" colspan=""3"">"&xxSatDeR&"</td></tr>" : shipopts=""
			if (rs("ordComLoc") AND 8)=8 then response.write "<tr><td align=""right"">"&shipopts&"</td><td align=""left"" colspan=""3"">"&xxSigRe2&"</td></tr>" : shipopts=""
			if (rs("ordComLoc") AND 16)=16 then response.write "<tr><td align=""right"">"&shipopts&"</td><td align=""left"" colspan=""3"">"&xxInsDe2&"</td></tr>" : shipopts=""
		end if
	end if
	ordAuthNumber = trim(rs("ordAuthNumber")&"")
	ordTransID = trim(rs("ordTransID")&"")
	if NOT isprinter AND (ordAuthNumber<>"" OR ordTransID<>"" OR doedit) then %>
			<tr>
			  <td align="right"><strong><%=yyAutCod%>:</strong></td>
			  <td align="left"><%=editfunc(ordAuthNumber,"ordAuthNumber",15) %></td>
			  <td align="right"><strong><%=yyTranID%>:</strong></td>
			  <td align="left"><%=editfunc(ordTransID,"ordTransID",15) %></td>
			</tr>
<%	end if
	ordAddInfo = Trim(rs("ordAddInfo"))
	if ordAddInfo <> "" OR doedit then %>
			<tr>
			  <td align="right" valign="top"><strong><%=xxAddInf%>:</strong></td>
			  <td align="left" colspan="3"><%
			if doedit then
				response.write "<textarea name=""ordAddInfo"" cols=""50"" rows=""4"" wrap=virtual>" & ordAddInfo & "</textarea>"
			else
				response.write replace(ordAddInfo,vbNewLine,"<br />")
			end if %></td>
			</tr>
<%	end if
	if NOT isprinter then
		rs2.Open "SELECT ordStatusInfo FROM orders WHERE ordID="&Request.QueryString("id"),cnn,0,1
		ordStatusInfo = rs2("ordStatusInfo")
		rs2.Close
		if Int(rs("ordPayProvider"))=20 then ' Google Checkout
			sSQL = "SELECT ordCNum FROM orders WHERE ordID="&Request.QueryString("id")
			rs2.Open sSQL,cnn,0,1
			ordCNum = trim(rs2("ordCNum")&"")
			rs2.Close
			if ordCNum<>"" then %>
				<tr>
				  <td align="right"><strong>Partial CC Number:</strong></td>
				  <td align="left" colspan="3">-<%=ordCNum %></td>
				</tr>
<%			end if
		end if
		if NOT doedit then response.write "<form method=""post"" action=""adminorders.asp""><input type=""hidden"" name=""updatestatus"" value=""1"" /><input type=""hidden"" name=""orderid"" value="""&Request.QueryString("id")&""" />"
%>			<tr>
			  <td align="right" valign="top"><strong><%=yyTraNum%>:</strong></td>
			  <td align="left"><input type="text" name="ordTrackNum" id="ordTrackNum" size="25" value="<%=rs("ordTrackNum")%>"></td>
			  <td align="right" valign="top"><strong><%=yyInvNum%>:</strong></td>
			  <td align="left"><input type="text" name="ordInvoice" size="25" value="<%=rs("ordInvoice")%>"></td>
			</tr>
			<tr>
			  <td align="right" valign="top"><strong><%=yyStaInf%>:</strong></td>
			  <td align="left" colspan="3"><textarea name="ordStatusInfo" id="ordStatusInfo" cols="50" rows="4" wrap=virtual><%=ordStatusInfo%></textarea>
<%		if NOT doedit then response.write "<input type=""submit"" value="""&yyUpdate&""" " & IIfVr(rs("ordPayProvider")=20, "onclick=""updategooglestatus('message',"&Request.QueryString("id")&")"" ", "") & "/>"%></td>
			</tr>
<%		if (rs("ordPayProvider")=3 OR rs("ordPayProvider")=13 OR rs("ordPayProvider")=20) AND rs("ordAuthNumber")<>"" AND NOT doedit then
			if rs("ordPayProvider")=20 then %>
			<tr>
			  <td width="50%" align="center" colspan="4">
				<strong>Update Google Account Status:</strong> <span id="googleupdatespan"></span>
			  </td>
			</tr>
			<tr>
			  <td width="50%" align="center" colspan="4">
				<input type="button" value="Charge Order" onclick="updategoogleorder('charge',<%=rs("ordID")%>)" />
				<input type="button" value="Cancel Order" onclick="updategoogleorder('cancel',<%=rs("ordID")%>)" />
				<input type="button" value="Refund Order" onclick="updategoogleorder('refund',<%=rs("ordID")%>)" />
				<input type="button" value="Ship Order" onclick="updategoogleorder('ship',<%=rs("ordID")%>)" />
			  </td>
			</tr>
<%			else %>
				<tr><td width="50%" align="center" colspan="4"><input type="button" value="Capture Funds" onclick="javascript:openemailpopup('oid=<%=rs("ordID")%>')" /></td></tr>
<%			end if
		end if
		if NOT doedit then response.write "</form>"
	else
		if trim(rs("ordInvoice")&"")<>"" then %>
			<tr>
			  <td align="right" valign="top"><strong><%=yyInvNum%>:</strong></td>
			  <td align="left" colspan="3"><%=editfunc(rs("ordInvoice"),"ordInvoice",15)%></td>
			</tr>
<%		end if
	end if %>
<%
if NOT isprinter AND NOT doedit then
	if Int(rs("ordPayProvider"))=10 then %>
			<tr>
			  <td width="50%" align="center" colspan="4"><hr width="50%" /></td>
			</tr>
<%		if request.servervariables("HTTPS")<>"on" AND (Request.ServerVariables("SERVER_PORT_SECURE") <> "1") AND nochecksslserver<>true then %>
			<tr>
			  <td width="50%" align="center" colspan="4"><strong><font color="#FF0000">You do not appear to be viewing this page on a secure (https) connection. Credit card information cannot be shown.</strong></td>
			</tr>
<%		else
			sSQL = "SELECT ordCNum FROM orders WHERE ordID="&Request.QueryString("id")
			rs2.Open sSQL,cnn,0,1
			ordCNum = rs2("ordCNum")
			rs2.Close
			if encryptmethod="aspencrypt" OR encryptmethod="" then %>
<OBJECT classid="CLSID:F9463571-87CB-4A90-A1AC-2284B7F5AF4E" 
	codeBase="https://www.beancastle.com" 
	id="XEncrypt">
</OBJECT>
<%			end if
			if ordCNum<>"" then
				if encryptmethod="none" then
					cnumarr = Split(ordCNum, "&")
				elseif encryptmethod="aspencrypt" OR encryptmethod="" then
%>
<SCRIPT LANGUAGE="VBScript">
function URLDecodeHex(match, hex_digits, pos, source)
	URLDecodeHex = chr("&H" & hex_digits)
end function
function URLDecode(decstr)
	set re = new RegExp
	decstr = Replace(decstr, "+", " ")
	re.Pattern = "%([0-9a-fA-F]{2})"
	re.Global = True
	URLDecode = re.Replace(decstr, GetRef("URLDecodeHex"))
end function
	' Set Context = XEncrypt.OpenContextEx("Microsoft Enhanced Cryptographic Provider v1.0", "mycontainer", False)
	Set Context = XEncrypt.OpenContext("mycontainer", False)
	Set Msg = Context.CreateMessage(True) ' use 3DES
	on error resume next
		err.number=0
		cnum = Msg.DecryptText("<%=Replace(ordCNum,vbNewLine,"")%>", "")
		If err.number = 0 then
			cnumarr = Split(cnum, "&")
		else
			Document.Write err.description
		end if
	on error goto 0
</SCRIPT>
<%				end if
			end if %>
			<tr>
			  <td width="50%" align="right" colspan="2"><strong><%=xxCCName%>:</strong></td>
			  <td width="50%" align="left" colspan="2"><%
			if encryptmethod="none" then
				if IsArray(cnumarr) then
					if UBOUND(cnumarr)>=4 then response.write URLDecode(cnumarr(4))
				end if
			elseif encryptmethod="aspencrypt" OR encryptmethod="" then
%><SCRIPT LANGUAGE="VBScript">
	if IsArray(cnumarr) then
		if UBOUND(cnumarr)>=4 then Document.Write URLDecode(cnumarr(4))
	end if
</SCRIPT><%
			end if %></td>
			</tr>
			<tr>
			  <td width="50%" align="right" colspan="2"><strong><%=yyCarNum%>:</strong></td>
			  <td width="50%" align="left" colspan="2"><%
			if ordCNum<>"" then
				if encryptmethod="none" then
					if IsArray(cnumarr) then response.write cnumarr(0)
				elseif encryptmethod="aspencrypt" OR encryptmethod="" then
%><SCRIPT LANGUAGE="VBScript">
	if IsArray(cnumarr) then Document.Write cnumarr(0)
</SCRIPT><%				end if
			else
				response.write "(no data)"
			end if %></td>
			</tr>
			<tr>
			  <td width="50%" align="right" colspan="2"><strong><%=yyExpDat%>:</strong></td>
			  <td width="50%" align="left" colspan="2"><%
			if encryptmethod="none" then
				if IsArray(cnumarr) then
					if UBOUND(cnumarr)>=1 then response.write cnumarr(1)
				end if
			elseif encryptmethod="aspencrypt" OR encryptmethod="" then
%><SCRIPT LANGUAGE="VBScript">
	if IsArray(cnumarr) then
		if UBOUND(cnumarr)>=1 then Document.Write cnumarr(1)
	end if
</SCRIPT><%
			end if %></td>
			</tr>
			<tr>
			  <td width="50%" align="right" colspan="2"><strong>CVV Code:</strong></td>
			  <td width="50%" align="left" colspan="2"><%
			if encryptmethod="none" then
				if IsArray(cnumarr) then
					if UBOUND(cnumarr)>=2 then response.write cnumarr(2)
				end if
			elseif encryptmethod="aspencrypt" OR encryptmethod="" then
%><SCRIPT LANGUAGE="VBScript">
	if IsArray(cnumarr) then
		if UBOUND(cnumarr)>=2 then Document.Write cnumarr(2)
	end if
</SCRIPT><%
			end if %></td>
			</tr>
			<tr>
			  <td width="50%" align="right" colspan="2"><strong>Issue Number:</strong></td>
			  <td width="50%" align="left" colspan="2"><%
			if encryptmethod="none" then
				if IsArray(cnumarr) then
					if UBOUND(cnumarr)>=3 then response.write cnumarr(3)
				end if
			elseif encryptmethod="aspencrypt" OR encryptmethod="" then
%><SCRIPT LANGUAGE="VBScript">
	if IsArray(cnumarr) then
		if UBOUND(cnumarr)>=3 then Document.Write cnumarr(3)
	end if
</SCRIPT><%
			end if %></td>
			</tr>
<%		end if
		if ordCNum<>"" AND NOT doedit then %>
		  <form method="post" action="adminorders.asp?id=<%=Request.QueryString("id")%>">
			<input type="hidden" name="delccdets" value="<%=Request.QueryString("id")%>" />
			<tr>
			  <td width="100%" align="center" colspan="4"><input type=submit value="<%=yyDelCC%>" /></td>
			</tr>
		  </form>
<%		end if
	end if
end if ' isprinter %>
			<tr>
			  <td width="100%" align="center" colspan="4">&nbsp;<br /></td>
			</tr>
		  </table>
<span id="productspan">
		  <table width="100%" border="1" cellspacing="0" cellpadding="4" bordercolor="#E7EAEF" bgcolor="">
			<tr>
			  <td><strong><%=xxPrId%></strong></td>
			  <td><strong><%=xxPrNm%></strong></td>
			  <td><strong><%=xxPrOpts%></strong></td>
			  <td><strong><%=xxQuant%></strong></td>
			  <td><strong><% if doedit then response.write xxUnitPr else response.write xxPrice%></strong></td>
<%	if doedit then response.write "<td align=""center""><strong>DEL</strong></td>" %>
			</tr>
<%
	if IsArray(allorders) then
		totoptpricediff = 0
		for rowcounter=0 to UBOUND(allorders,2)
			optpricediff = 0
%>
			<tr>
			  <td valign="top" nowrap><% if doedit then response.write "<input type=""button"" value=""..."" onclick=""updateoptions("&rowcounter&")"">&nbsp;<input type=""hidden"" name=""cartid"&rowcounter&""" value="""&replace(allorders(4,rowcounter),"""","&quot;")&""" />"%><strong><%=editfunc(allorders(0,rowcounter),"prodid"&rowcounter,18)%></strong></td>
			  <td valign="top"><%=editfunc(allorders(1,rowcounter),"prodname"&rowcounter,24)%></td>
			  <td valign="top"><%
			if doedit then response.write "<span id=""optionsspan"&rowcounter&""">"
			sSQL = "SELECT coOptGroup,coCartOption,coPriceDiff,coOptID,optGroup FROM cartoptions LEFT JOIN options ON cartoptions.coOptID=options.optID WHERE coCartID="&allorders(4,rowcounter) & " ORDER BY coID"
			rs2.Open sSQL,cnn,0,1
			if NOT rs2.EOF then
				if doedit then response.write "<table border=""0"" cellspacing=""0"" cellpadding=""1"" width=""100%"">"
				do while NOT rs2.EOF
					if doedit then
						response.write "<tr><td align=""right""><strong>" & rs2("coOptGroup") & ":</strong></td><td>"
						if IsNull(rs2("optGroup")) then
							response.write "xxxxxx"
						else
							sSQL="SELECT optID,"&getlangid("optName",32)&",optPriceDiff,optType,optFlags,optStock,optPriceDiff AS optDims FROM options INNER JOIN optiongroup ON options.optGroup=optiongroup.optGrpID WHERE optGroup=" & rs2("optGroup")
							rsl.Open sSQL,cnn,0,1
							if NOT rsl.EOF then
								if abs(rsl("optType"))=1 OR abs(rsl("optType"))=2 then
									response.write "<select onchange=""dorecalc(true)"" name=""optn"&rowcounter&"_"&rs2("coOptID")&""" id=""optn"&rowcounter&"_"&rs2("coOptID")&""" size=""1"">"
									do while NOT rsl.EOF
										response.write "<option value="""&rsl("optID")&"|"&IIfVr((rsl("optFlags") AND 1) = 1,(allorders(3,rowcounter)*rsl("optPriceDiff"))/100.0,rsl("optPriceDiff"))&""""
										if rsl("optID")=rs2("coOptID") then response.write " selected"
										response.write ">"&rsl(getlangid("optName",32))
										if cDbl(rsl("optPriceDiff"))<>0 then
											response.write " "
											if cDbl(rsl("optPriceDiff")) > 0 then response.write "+"
											if (rsl("optFlags") AND 1) = 1 then
												response.write FormatNumber((allorders(3,rowcounter)*rsl("optPriceDiff"))/100.0,2)
											else
												response.write FormatNumber(rsl("optPriceDiff"),2)
											end if
										end if
										response.write "</option>"
										rsl.MoveNext
									loop
									response.write "</select>"
								else
									response.write "<input type='hidden' name='optn"&rowcounter&"_"&rs2("coOptID")&"' value='"&rsl("optID")&"' /><textarea wrap='virtual' name='voptn"&rowcounter&"_"&rs2("coOptID")&"' id='voptn"&rowcounter&"_"&rs2("coOptID")&"' cols='30' rows='3'>"
									response.write rs2("coCartOption")&"</textarea>"
								end if
							end if
							rsl.Close
						end if
						response.write "</td></tr>"
					else
						response.write "<strong>" & rs2("coOptGroup") & ":</strong> " & replace(rs2("coCartOption")&"", vbCrLf, "<br>") & "<br />"
					end if
					if doedit then
						optpricediff = optpricediff + rs2("coPriceDiff")
					else
						allorders(2,rowcounter) = allorders(2,rowcounter) + rs2("coPriceDiff")
					end if
					rs2.MoveNext
				loop
				if doedit then response.write "</table>"
			else
				response.write " - "
			end if
			rs2.Close
			if doedit then response.write "</span>" %></td>
			  <td valign="top"><%=editfunc(allorders(3,rowcounter),"quant"&rowcounter&""" onchange=""dorecalc(true)",5)%></td>
			  <td valign="top"><%if doedit then response.write editnumeric(allorders(2,rowcounter),"price"&rowcounter&""" onchange=""dorecalc(true)",7) else response.write FormatEuroCurrency(allorders(2,rowcounter)*allorders(3,rowcounter))%>
			<%		if doedit then
						response.write "<input type=""hidden"" id=""optdiffspan"&rowcounter&""" value="""&optpricediff&""">"
						totoptpricediff = totoptpricediff + (optpricediff*allorders(3,rowcounter))
					end if
			%></td>
<%			if doedit then response.write "<td align=""center""><input type=""checkbox"" name=""del_"&rowcounter&""" id=""del_"&rowcounter&""" value=""yes"" /></td>" %>
			</tr>
<%
		next
	end if
%>
<!--NEXTPRODUCTCOMMENT-->
<%	if doedit then %>
			<tr>
			  <td align="right" colspan="4">
				<table width="100%" border="0" cellspacing="0" cellpadding="0">
				  <tr>
					<td align="center"><% if doedit then response.write "<input style=""width:30px;"" type=""button"" value=""-"" onclick=""extraproduct('-')""> "&yyMoProd&" <input style=""width:30px;"" type=""button"" value=""+"" onclick=""extraproduct('+')""> &nbsp; <input type=""button"" value="""&yyRecal&""" onclick=""dorecalc(false)"">"%></td>
					<td align="right"><strong>Options Total:</strong></td>
				  </tr>
				</table></td>
			  <td align="left" colspan="2"><span id="optdiffspan"><%=FormatNumber(totoptpricediff, 2)%></span></td>
			</tr>
<%	end if %>
			<tr>
			  <td align="right" colspan="4"><strong><%=xxOrdTot%>:</strong></td>
			  <td align="left"><%=editnumeric(rs("ordTotal"),"ordtotal",7)%></td>
<%		if doedit then response.write "<td align=""center"">&nbsp;</td>" %>
			</tr>
<%	if isprinter AND combineshippinghandling=TRUE then %>
			<tr>
			  <td align="right" colspan="4"><strong><%=xxShipHa%>:</strong></td>
			  <td align="left"><%=FormatEuroCurrency(rs("ordShipping")+rs("ordHandling"))%></td>
			</tr>
<%	else
		if rs("ordShipping") > 0 OR doedit then %>
			<tr>
			  <td align="right" colspan="4"><strong><%=xxShippg%>:</strong></td>
			  <td align="left"><%=editnumeric(rs("ordShipping"),"ordShipping",7)%></td>
<%		if doedit then response.write "<td align=""center"">&nbsp;</td>" %>
			</tr>
<%		end if
		if cDbl(rs("ordHandling"))<>0.0 OR doedit then %>
			<tr>
			  <td align="right" colspan="4"><strong><%=xxHndlg%>:</strong></td>
			  <td align="left"><%=editnumeric(rs("ordHandling"),"ordHandling",7)%></td>
<%		if doedit then response.write "<td align=""center"">&nbsp;</td>" %>
			</tr>
<%		end if
	end if
	if cDbl(rs("ordDiscount"))<>0.0 OR doedit then %>
			<tr>
			  <td align="right" colspan="4"><strong><%=xxDscnts%>:</strong></td>
			  <td align="left"><font color="#FF0000"><%=editnumeric(rs("ordDiscount"),"ordDiscount",7)%></font></td>
<%		if doedit then response.write "<td align=""center"">&nbsp;</td>" %>
			</tr>
<%	end if
	if rs("ordStateTax") > 0 OR doedit  then %>
			<tr>
			  <td align="right" colspan="4"><strong><%=xxStaTax%>:</strong></td>
			  <td align="left"><%=editnumeric(rs("ordStateTax"),"ordStateTax",7)%></td>
<%		if doedit then response.write "<td align=""center"" nowrap><input type=""text"" name=""staterate"" id=""staterate"" size=""1"" value="""&statetaxrate&""">%</td>" %>
			</tr>
<%	end if
	if rs("ordCountryTax") > 0 OR doedit then %>
			<tr>
			  <td align="right" colspan="4"><strong><%=xxCntTax%>:</strong></td>
			  <td align="left"><%=editnumeric(rs("ordCountryTax"),"ordCountryTax",7)%></td>
<%		if doedit then response.write "<td align=""center"" nowrap><input type=""text"" name=""countryrate"" id=""countryrate"" size=""1"" value="""&countrytaxrate&""">%</td>" %>
			</tr>
<%	end if
	if rs("ordHSTTax") > 0 OR (doedit AND canadataxsystem) then %>
			<tr>
			  <td align="right" colspan="4"><strong><%=xxHST%>:</strong></td>
			  <td align="left"><%=editnumeric(rs("ordHSTTax"),"ordHSTTax",7)%></td>
<%		if doedit then response.write "<td align=""center"" nowrap><input type=""text"" name=""hstrate"" id=""hstrate"" size=""1"" value="""&hsttaxrate&""">%</td>" %>
			</tr>
<%	end if %>
			<tr>
			  <td align="right" colspan="4"><strong><%=xxGndTot%>:</strong></td>
			  <td align="left"><span id="grandtotalspan"><%=FormatEuroCurrency((rs("ordTotal")+rs("ordStateTax")+rs("ordCountryTax")+rs("ordShipping")+rs("ordHSTTax")+rs("ordHandling"))-rs("ordDiscount"))%></span></td>
<%		if doedit then response.write "<td align=""center"">&nbsp;</td>" %>
			</tr>
			</table>
</span>
		  </td>
		</tr>
<%	if isprinter AND invoicefooter<>"" then %>
		<tr> 
          <td width="100%"><%=invoicefooter%></td>
		</tr>
<%	elseif doedit then %>
		<tr> 
          <td align="center" width="100%">&nbsp;<br /><input type="submit" value="<%=yyUpdate%>" /><br />&nbsp;</td>
		</tr>
<%	end if %>
	  </table>
<%	if doedit then response.write "</form>"
	rs.Close
else
	sSQL = "SELECT ordID FROM orders WHERE ordStatus=1"
	if request.form("act")<>"purge" then sSQL = sSQL & " AND ordStatusDate<"&datedelim & VSUSDate(thedate - 3) & datedelim
	rs.Open sSQL,cnn,0,1
	do while NOT rs.EOF
		theid = rs("ordID")
		addcomma = ""
		delOptions = ""
		sSQL = "SELECT cartID FROM cart WHERE cartOrderID="&theid
		rsl.Open sSQL,cnn,0,1
		do while NOT rsl.EOF
			delOptions = delOptions & addcomma & rsl("cartID")
			addcomma = ","
			rsl.MoveNext
		loop
		rsl.Close
		if delOptions<>"" then cnn.Execute("DELETE FROM cartoptions WHERE coCartID IN ("&delOptions&")")
		cnn.Execute("DELETE FROM cart WHERE cartOrderID="&theid)
		cnn.Execute("DELETE FROM orders WHERE ordID="&theid)
		rs.MoveNext
	loop
	rs.Close
	if request.form("act")="authorize" then
		do_stock_management(trim(request.form("id")))
		if Trim(request.form("authcode"))<>"" then
			sSQL = "UPDATE orders set ordAuthNumber='"&replace(Trim(request.form("authcode")),"'","''")&"',ordStatus=3 WHERE ordID="&request.form("id")
		else
			sSQL = "UPDATE orders set ordAuthNumber='"&replace(yyManAut,"'","''")&"',ordStatus=3 WHERE ordID="&request.form("id")
		end if
		cnn.Execute(sSQL)
		cnn.Execute("UPDATE cart SET cartCompleted=1 WHERE cartOrderID="&request.form("id"))
	elseif request.form("act")="status" then
		maxitems=Int(request.form("maxitems"))
		orderstatusxml=""
		call getpayprovdetails(20,googledata1,googledata2,googledata3,googledemomode,ppmethod)
		for index=0 to maxitems-1
			iordid = Trim(request.form("ordid" & index))
			ordstatus = Trim(request.form("ordstatus" & index))
			ordauthno = ""
			oldordstatus=999
			payprovider=0
			rs.Open "SELECT ordStatus,ordAuthNumber,ordEmail,ordDate,"&getlangid("statPublic",64)&",ordStatusInfo,ordName,ordTrackNum,ordPayProvider FROM orders INNER JOIN orderstatus ON orders.ordStatus=orderstatus.statID WHERE ordID="&iordid,cnn,0,1
			if NOT rs.EOF then
				oldordstatus=rs("ordStatus")
				ordauthno=rs("ordAuthNumber")
				ordemail=rs("ordEmail")
				orddate=rs("ordDate")
				oldstattext=rs(getlangid("statPublic",64))&""
				ordstatinfo=rs("ordStatusInfo")&""
				ordername=rs("ordName")
				if trackingnumtext = "" then trackingnumtext=yyTrackT
				if trim(rs("ordTrackNum")) <> "" then trackingnum=replace(trackingnumtext, "%s", rs("ordTrackNum")) else trackingnum=""
				payprovider=rs("ordPayProvider")
			end if
			rs.Close
			if payprovider<>20 then
				if NOT oldordstatus=999 AND (oldordstatus < 3 AND ordstatus >=3) then
					' This is to force stock management
					cnn.Execute("UPDATE cart SET cartCompleted=0 WHERE cartOrderID="&iordid)
					do_stock_management(iordid)
					cnn.Execute("UPDATE cart SET cartCompleted=1 WHERE cartOrderID="&iordid)
					if ordauthno="" then cnn.Execute("UPDATE orders SET ordAuthNumber='"&replace(yyManAut,"'","''")&"' WHERE ordID=" & iordid)
				end if
				if NOT oldordstatus=999 AND (oldordstatus >=3 AND ordstatus < 3) then release_stock(iordid)
				if iordid<>"" AND ordstatus<>"" then
					if oldordstatus<>Int(ordstatus) then
						if request.form("emailstat")="1" then
							rs.Open "SELECT "&getlangid("statPublic",64)&" FROM orderstatus WHERE statID=" & ordstatus,cnn,0,1
							if NOT rs.EOF then newstattext = rs(getlangid("statPublic",64))&""
							rs.Close
							if orderstatussubject<>"" then emailsubject=orderstatussubject else emailsubject = "Order status updated"
							ose = orderstatusemail
							ose = replace(ose, "%orderid%", iordid)
							ose = replace(ose, "%orderdate%", FormatDateTime(orddate, 1) & " " & FormatDateTime(orddate, 4))
							ose = replace(ose, "%oldstatus%", oldstattext)
							ose = replace(ose, "%newstatus%", newstattext)
							ose = replace(ose, "%date%", FormatDateTime(DateAdd("h",dateadjust,Now()), 1) & " " & FormatDateTime(DateAdd("h",dateadjust,Now()), 4))
							ose = replace(ose, "%statusinfo%", ordstatinfo)
							ose = replace(ose, "%ordername%", ordername)
							ose = replace(ose, "%trackingnum%", trackingnum)
							ose = replace(ose, "%nl%", emlNl)
							Call DoSendEmailEO(ordemail,emailAddr,"",emailsubject,ose,emailObject,themailhost,theuser,thepass)
						end if
						if payprovider=20 AND noupdategooglestatus<>TRUE then
							if Int(ordstatus)=0 then
								orderstatusxml=orderstatusxml&"<cancel-order xmlns=""http://checkout.google.com/schema/2"" google-order-number="""&ordauthno&"""><reason>Cancelled by store admin on " & Date() & ".</reason></cancel-order>"
							end if
							if FALSE AND orderstatusxml<>"" AND noupdategooglestatus<>TRUE then
								' response.write Replace(Replace(orderstatusxml,"</","&lt;/"),"<","<br />&lt;")&"<HR>"
								set objHttp = Server.CreateObject("WinHTTP.WinHTTPRequest.5.1")
								theurl="https://"&IIfVr(googledemomode, "sandbox", "checkout")&".google.com/cws/v2/Merchant/"&googledata1&"/request"
								objHttp.open "POST", theurl, false
								objHttp.setRequestHeader "Authorization", "Basic " & vrbase64_encrypt(googledata1&":"&googledata2)
								objHttp.setRequestHeader "Content-Type", "application/xml"
								objHttp.setRequestHeader "Accept", "application/xml"
								on error resume next
								err.number=0
								objHttp.Send "<?xml version=""1.0"" encoding=""UTF-8""?>" & orderstatusxml
								if err.number <> 0 OR objHttp.status <> 200 Then
									response.write "<font color=""#FF0000"">" & "Error, couldn't change status of order " & iordid & "</font><br/>"
								else
									res = objHttp.responseText
									' response.write Replace(Replace(objHttp.responseText,"</","&lt;/"),"<","<br />&lt;")&"<HR>"
								end if
								ordstatus=oldordstatus
								on error goto 0
								set objHttp = nothing
								orderstatusxml=""
							end if
						end if
					end if
					if oldordstatus<>Int(ordstatus) then cnn.Execute("UPDATE orders SET ordStatus=" & ordstatus & ",ordStatusDate=" & datedelim & VSUSDateTime(DateAdd("h",dateadjust,Now())) & datedelim & " WHERE ordID=" & iordid)
				end if
			end if
		next
	end if
	if Request("sd") = "" then sd=thedate else sd=Request("sd")
	if Request("ed") = "" then ed=thedate else ed=Request("ed")
	on error resume next
	sd = DateValue(sd)
	ed = DateValue(ed)
	if err.number <> 0 then
		sd = thedate
		ed = thedate
		success=false
		errmsg=yyDatInv
	end if
	on error goto 0
	if ed < sd then ed = sd
	if request.form("powersearch")="1" then
		sSQL = "SELECT ordID,ordName,payProvName,ordAuthNumber,ordDate,ordStatus,(ordTotal-ordDiscount),ordTransID,ordAVS,ordCVV,ordPayProvider FROM orders INNER JOIN payprovider ON payprovider.payProvID=orders.ordPayProvider WHERE ordStatus>=0 "
		fromdate = Trim(request.form("fromdate"))
		todate = Trim(request.form("todate"))
		ordid = Trim(Replace(Replace(request.form("ordid"),"'",""),"""",""))
		origsearchtext = Trim(Replace(request.form("searchtext"),"""","&quot;"))
		searchtext = Trim(Replace(request.form("searchtext"),"'","''"))
		ordstatus = Trim(request.form("ordstatus"))
		if ordid<>"" then
			if IsNumeric(ordid) then
				sSQL = sSQL & " AND ordID=" & ordid
			else
				success=false
				errmsg="The order id you specified seems to be invalid - " & ordid
				sSQL = sSQL & " AND ordID=0"
			end if
		else
			if fromdate<>"" then
				if IsNumeric(fromdate) then
					thefromdate = (thedate-fromdate)
				else
					err.number=0
					on error resume next
					thefromdate = DateValue(fromdate)
					if err.number <> 0 then
						thefromdate = thedate
						success=false
						errmsg=yyDatInv & " - " & fromdate
					end if
					on error goto 0
				end if
				if todate="" then
					thetodate = thefromdate
				elseif IsNumeric(todate) then
					thetodate = (thedate-todate)
				else
					err.number=0
					on error resume next
					thetodate = DateValue(todate)
					if err.number <> 0 then
						thetodate = thedate
						success=false
						errmsg=yyDatInv & " - " & todate
					end if
					on error goto 0
				end if
				if thefromdate > thetodate then
					tmpdate = thetodate
					thetodate = thefromdate
					thefromdate = tmpdate
				end if
				sd = thefromdate
				ed = thetodate
				sSQL = sSQL & " AND ordDate BETWEEN " & datedelim & VSUSDate(thefromdate) & datedelim & " AND " & datedelim & VSUSDate(thetodate+1) & datedelim
			end if
			if ordstatus<>"" AND NOT InStr(ordstatus,"9999")>0 then sSQL = sSQL & " AND ordStatus IN (" & ordstatus & ")"
			if searchtext<>"" then sSQL = sSQL & " AND (ordTransID LIKE '%"&searchtext&"%' OR ordAuthNumber LIKE '%"&searchtext&"%' OR ordName LIKE '%"&searchtext&"%' OR ordEmail LIKE '%"&searchtext&"%' OR ordAddress LIKE '%"&searchtext&"%' OR ordCity LIKE '%"&searchtext&"%' OR ordState LIKE '%"&searchtext&"%' OR ordZip LIKE '%"&searchtext&"%' OR ordPhone LIKE '%"&searchtext&"%' OR ordInvoice LIKE '%"&searchtext&"%' OR ordAffiliate='"&searchtext&"')"
		end if
		sSQL = sSQL & " ORDER BY ordID"
	else
		sSQL = "SELECT ordID,ordName,payProvName,ordAuthNumber,ordDate,ordStatus,(ordTotal-ordDiscount),ordTransID,ordAVS,ordCVV,ordPayProvider FROM orders INNER JOIN payprovider ON payprovider.payProvID=orders.ordPayProvider WHERE ordStatus<>1 AND ordDate BETWEEN "&datedelim & VSUSDate(DateValue(sd)) & datedelim&" AND "&datedelim & VSUSDate(DateValue(ed)+1) & datedelim&" ORDER BY ordID"
	end if
	rs.Open sSQL,cnn,0,1
	alldata = ""
	if NOT rs.EOF then alldata=rs.getrows
	rs.Close
	hasdeleted=false
	sSQL = "SELECT COUNT(*) AS NumDeleted FROM orders WHERE ordStatus=1"
	rs.Open sSQL,cnn,0,1
		if rs("NumDeleted") > 0 then hasdeleted=true
	rs.Close
%>
<script language="javascript" type="text/javascript" src="popcalendar.js">
</script>
<script language="javascript" type="text/javascript">
<!--
function delrec(id) {
cmsg = "<%=yyConDel%>\n"
if (confirm(cmsg)) {
	document.mainform.id.value = id;
	document.mainform.act.value = "delete";
	document.mainform.sd.value="<%=DateValue(sd)%>";
	document.mainform.ed.value="<%=DateValue(ed)%>";
	document.mainform.submit();
}
}
function authrec(id) {
var aucode;
cmsg = "<%=yyEntAuth%>"
if ((aucode=prompt(cmsg,'<%=yyManAut%>'))!=null) {
	document.mainform.id.value = id;
	document.mainform.act.value = "authorize";
	document.mainform.authcode.value = aucode;
	document.mainform.sd.value="<%=DateValue(sd)%>";
	document.mainform.ed.value="<%=DateValue(ed)%>";
	document.mainform.submit();
}
}
function checkcontrol(tt,evt){
<% if netnav then %>
theevnt = evt;
return;
<% else %>
theevnt=window.event;
<% end if %>
if(theevnt.ctrlKey){
	maxitems=document.mainform.maxitems.value;
	for(index=0;index<maxitems;index++){
		isdisabled = eval('document.mainform.ordstatus'+index+'.disabled');
		if(! isdisabled){
			if(eval('document.mainform.ordstatus'+index+'.length') > tt.selectedIndex){
				eval('document.mainform.ordstatus'+index+'.selectedIndex='+tt.selectedIndex);
				eval('document.mainform.ordstatus'+index+'.options['+tt.selectedIndex+'].selected=true');
			}
		}
	}
}
}
function displaysearch(){
thestyle = document.getElementById('searchspan').style;
if(thestyle.display=='none')
	thestyle.display = 'block';
else
	thestyle.display = 'none';
}
function checkprinter(tt,evt){
<% if netnav then %>
if(evt.ctrlKey || evt.altKey || document.mainform.ctrlmod[document.mainform.ctrlmod.selectedIndex].value=="1"){
	tt.href += "&printer=true";
	window.location.href = tt.href;
}
if(document.mainform.ctrlmod[document.mainform.ctrlmod.selectedIndex].value=="2"){
	tt.href += "&doedit=true";
	window.location.href = tt.href;
}
<% else %>
theevnt=window.event;
if(theevnt.ctrlKey || document.mainform.ctrlmod[document.mainform.ctrlmod.selectedIndex].value=="1")tt.href += "&printer=true";
if(document.mainform.ctrlmod[document.mainform.ctrlmod.selectedIndex].value=="2")tt.href += "&doedit=true";
<% end if %>
return(true);
}
function setdumpformat(){
formatindex = document.forms.dumpform.filedump[document.forms.dumpform.filedump.selectedIndex].value;
if(formatindex==1)
	document.dumpform.act.value='dumporders';
else if(formatindex==2)
	document.dumpform.act.value='dumpdetails';
else if(formatindex==3)
	document.dumpform.act.value='quickbooks';
else if(formatindex==4)
	document.dumpform.act.value='ouresolutionsxmldump';
}
// -->
</script>
      <table border="0" cellspacing="0" cellpadding="0" width="100%" bgcolor="">
        <tr>
          <td width="100%" align="center">
<%	themask = cStr(DateSerial(2003,12,11))
	themask = replace(themask,"2003","yyyy")
	themask = replace(themask,"12","mm")
	themask = replace(themask,"11","dd")
	if NOT success then response.write "<p><font color='#FF0000'>"&errmsg&"</font></p>" %>
			<span name="searchspan" id="searchspan" <% if request.cookies("powersearch")="1" then response.write "style=""display:block""" else response.write "style=""display:none"""%>>
            <table width="100%" border="0" cellspacing="1" cellpadding="1" bgcolor="">
			  <form method="post" action="adminorders.asp" name="psearchform">
			  <input type="hidden" name="powersearch" value="1" />
			  <tr bgcolor="#030133"><td colspan="4"><strong><font color="#E7EAEF">&nbsp;<%=yyPowSea%></font></strong></td></tr>
			  <tr bgcolor="#E7EAEF"> 
                <td align="right" width="25%"><strong><%=yyOrdFro%>:</strong></td>
				<td align="left" width="25%">&nbsp;<input type="text" size="14" name="fromdate" value="<%=fromdate%>" /> <input type=button onclick="popUpCalendar(this, document.forms.psearchform.fromdate, '<%=themask%>', 0)" value='DP' /></td>
				<td align="right" width="25%"><strong><%=yyOrdTil%>:</strong></td>
				<td align="left" width="25%">&nbsp;<input type="text" size="14" name="todate" value="<%=todate%>" /> <input type=button onclick="popUpCalendar(this, document.forms.psearchform.todate, '<%=themask%>', -205)" value='DP' /></td>
			  </tr>
			  <tr bgcolor="#EAECEB">
				<td align="right"><strong><%=yyOrdId%>:</strong></td>
				<td align="left">&nbsp;<input type="text" size="14" name="ordid" value="<%=ordid%>" /></td>
				<td align="right"><strong><%=yySeaTxt%>:</strong></td>
				<td align="left">&nbsp;<input type="text" size="24" name="searchtext" value="<%=origsearchtext%>" /></td>
			  </tr>
			  <tr bgcolor="#E7EAEF">
				<td align="right"><strong><%=yyOrdSta%>:</strong></td>
				<td align="left">&nbsp;<select name="ordstatus" size="5" multiple><option value="9999" <%if InStr(ordstatus,"9999")>0 then response.write "selected"%>><%=yyAllSta%></option><%
						if ordstatus<>"" then selstatus = Split(ordstatus, ",")
						for index=0 to UBOUND(allstatus,2)
							response.write "<option value=""" & allstatus(0,index) & """"
							if IsArray(selstatus) then
								for ii=0 to UBOUND(selstatus)
									if Int(selstatus(ii))=Int(allstatus(0,index)) then response.write " selected"
								next
							end if
							response.write ">" & allstatus(1,index) & "</option>"
						next %></select></td>
				<td colspan="2" align="center"><input type="checkbox" name="startwith" value="1" <% if request.cookies("powersearch")="1" then response.write "checked"%> /> <strong><%=yyStaPow%></strong><br /><br />
				  <input type="submit" value="<%=yySearch%>" /> <input type="button" value="Stats" onclick="document.forms.psearchform.action='adminstats.asp';document.forms.psearchform.submit();" /></td>
			  </tr>
			  <tr><td colspan="4">&nbsp;</td></tr>
			  </form>
			</table>
			</span>
			<table width="100%" border="0" cellspacing="0" cellpadding="2" bgcolor="">
			  <form method="post" action="adminorders.asp">
			  <tr> 
                <td align="center"><input type="button" value="<%=yyPowSea%>" onclick="displaysearch()" /></td>
				<td align="center"><strong><%=yyShoFrm%>:</strong> <select name="sd" size="1"><%
					gotmatch=false
					For rowcounter=0 to Day(thedate)-1
						response.write "<option value='"&thedate-rowcounter&"'"
						if thedate-rowcounter=sd then
							response.write " selected"
							gotmatch=true
						end if
						response.write ">"&thedate-rowcounter&"</option>"&vbCrLf
						smonth=thedate-rowcounter
					Next
					For rowcounter=1 to 12
						if NOT gotmatch AND DateAdd("m",0-rowcounter,smonth) < sd then
							response.write "<option value='"&sd&"' selected>" & sd & "</option>"
							gotmatch=true
						end if
						response.write "<option value='"&DateAdd("m",0-rowcounter,smonth)&"'"
						if DateAdd("m",0-rowcounter,smonth)=sd then
							response.write " selected"
							gotmatch=true
						end if
						response.write ">"&DateAdd("m",0-rowcounter,smonth)&"</option>"&vbCrLf
					Next
					if NOT gotmatch then response.write "<option value='"&sd&"' selected>" & sd & "</option>"
				%></select> <strong><%=yyTo%>:</strong> <select name="ed" size="1"><%
					gotmatch=false
					For rowcounter=0 to Day(thedate)-1
						response.write "<option value='"&thedate-rowcounter&"'"
						if thedate-rowcounter=ed then
							response.write " selected"
							gotmatch=true
						end if
						response.write ">"&thedate-rowcounter&"</option>"&vbCrLf
						smonth=thedate-rowcounter
					Next
					For rowcounter=1 to 12
						if NOT gotmatch AND DateAdd("m",0-rowcounter,smonth) < ed then
							response.write "<option value='"&ed&"' selected>" & ed & "</option>"
							gotmatch=true
						end if
						response.write "<option value='"&DateAdd("m",0-rowcounter,smonth)&"'"
						if DateAdd("m",0-rowcounter,smonth)=ed then
							response.write " selected"
							gotmatch=true
						end if
						response.write ">"&DateAdd("m",0-rowcounter,smonth)&"</option>"&vbCrLf
					Next
					if NOT gotmatch then response.write "<option value='"&ed&"' selected>" & ed & "</option>"
				%></select> <input type="submit" value="Go" /></td>
			  </tr>
			  <tr><td colspan="2">&nbsp;</td></tr>
			  </form>
			</table>
			<table width="100%" border="0" cellspacing="1" cellpadding="2" bgcolor="">
			  <tr bgcolor="#030133"> 
                <td align="center"><strong><font color="#E7EAEF"><%=yyOrdId%></font></strong></td>
				<td align="center"><strong><font color="#E7EAEF"><%=yyName%></font></strong></td>
				<td align="center"><strong><font color="#E7EAEF"><%=yyMethod%></font></strong></td>
				<td width="1%"><strong><font color="#E7EAEF">AVS</font></strong></td>
				<td width="1%"><strong><font color="#E7EAEF">CVV</font></strong></td>
				<td align="center"><strong><font color="#E7EAEF"><%=yyAutCod%></font></strong></td>
				<td align="center"><strong><font color="#E7EAEF"><%=yyDate%></font></strong></td>
				<td align="center"><strong><font color="#E7EAEF"><%=yyStatus%></font></strong></td>
			  </tr>
			  <form method="post" name="mainform" action="adminorders.asp">
			  <% if request.form("powersearch")="1" then %>
			  <input type="hidden" name="powersearch" value="1" />
			  <input type="hidden" name="fromdate" value="<%=Trim(request.form("fromdate"))%>" />
			  <input type="hidden" name="todate" value="<%=Trim(request.form("todate"))%>" />
			  <input type="hidden" name="ordid" value="<%=Trim(Replace(Replace(request.form("ordid"),"'",""),"""",""))%>" />
			  <input type="hidden" name="origsearchtext" value="<%=Trim(Replace(request.form("searchtext"),"""","&quot;"))%>" />
			  <input type="hidden" name="searchtext" value="<%=Trim(Replace(request.form("searchtext"),"""","&quot;"))%>" />
			  <input type="hidden" name="ordstatus" value="<%=Trim(request.form("ordstatus"))%>" />
			  <input type="hidden" name="startwith" value="<% if request.cookies("powersearch")="1" then response.write "1"%>" />
			  <% end if %>
			  <input type="hidden" name="act" value="xxx" />
			  <input type="hidden" name="id" value="xxx" />
			  <input type="hidden" name="authcode" value="xxx" />
			  <input type="hidden" name="ed" value="<%=DateValue(ed)%>" />
			  <input type="hidden" name="sd" value="<%=DateValue(sd)%>" />
<%	ordTot=0
	if IsArray(alldata) then
		for rowcounter=0 to UBOUND(alldata,2)
			if alldata(5,rowcounter)>=3 then ordTot=ordTot+alldata(6,rowcounter)
			if alldata(3,rowcounter)="" OR IsNull(alldata(3,rowcounter)) then
				startfont="<font color='#FF0000'>"
				endfont="</font>"
			else
				startfont=""
				endfont=""
			end if
			if bgcolor="#E7EAEF" then bgcolor="#EAECEB" else bgcolor="#E7EAEF"
%>
			  <tr bgcolor="<%=bgcolor%>"> 
                <td align="center"><a onclick="return(checkprinter(this,event));" href="adminorders.asp?id=<%=alldata(0,rowcounter)%>"><%="<strong>"&startfont&alldata(0,rowcounter)&endfont&"</strong>"%></a></td>
				<td align="center"><a onclick="return(checkprinter(this,event));" href="adminorders.asp?id=<%=alldata(0,rowcounter)%>"><%=startfont&alldata(1,rowcounter)&endfont%></a></td>
				<td align="center"><%=startfont&alldata(2,rowcounter)&IIfVr(alldata(2,rowcounter)="PayPal" AND trim(alldata(7,rowcounter)&"")<>""," CC","")&endfont%></td>
				<td align="center"><% if trim(alldata(8,rowcounter)&"")<>"" then response.write alldata(8,rowcounter) else response.write "&nbsp;" %></td>
				<td align="center"><% if trim(alldata(9,rowcounter)&"")<>"" then response.write alldata(9,rowcounter) else response.write "&nbsp;" %></td>
				<td align="center"><%
					if alldata(3,rowcounter)="" OR IsNull(alldata(3,rowcounter)) then
						isauthorized=false
						response.write "<input type='button' name='auth' value='"&yyAuthor&"' onclick=""authrec('"&alldata(0,rowcounter)&"')"" />"
					else
						isauthorized=true
						response.write "<a href=""#"" title="""&FormatEuroCurrency(alldata(6,rowcounter))&""" onclick=""authrec('"&alldata(0,rowcounter)&"');return(false);"">" & startfont & alldata(3,rowcounter) & endfont & "</a>"
					end if %></td>
				<td align="center"><font size="1"><%=startfont&Replace(alldata(4,rowcounter)&""," ","<br />",1,1)&endfont%></font></td>
				<td align="center"><input type="hidden" name="ordid<%=rowcounter%>" value="<%=alldata(0,rowcounter)%>" /><select name="ordstatus<%=rowcounter%>" size="1" onchange="checkcontrol(this,event)"<% if alldata(10,rowcounter)=20 then response.write " disabled" %>><%
						gotitem=false
						for index=0 to UBOUND(allstatus,2)
							if NOT isauthorized AND allstatus(0,index)>2 then exit for
							if NOT (alldata(5,rowcounter)<>2 AND allstatus(0,index)=2) then
								response.write "<option value=""" & allstatus(0,index) & """"
								if alldata(5,rowcounter)=allstatus(0,index) then
									response.write " selected"
									gotitem=true
								end if
								response.write ">" & allstatus(1,index) & "</option>"
							end if
						next
						if NOT gotitem then response.write "<option value="""" selected>"&yyUndef&"</option>" %></select></td>
			  </tr>
<%			if rowcounter>=250 then
				response.write "<tr><td colspan='8' align='center'><strong>Limit of "&rowcounter&" orders reached. Please refine your search.</strong></td></tr>"
				exit for
			end if
		next %>
			  <tr> 
				<td align="center"><%=FormatEuroCurrency(ordTot)%></td>
				<td align="center"><% if hasdeleted then %><input type="submit" value="<%=yyPurDel%>" onclick="document.mainform.act.value='purge';" /><% end if %></td>
				<td colspan="5"><select name="ctrlmod" size="1"><option value="0"><%=yyVieDet%></option><option value="1"><%=yyPPSlip%></option><option value="2"><%=yyEdOrd%></option></select>
				&nbsp;&nbsp;&nbsp;<%if orderstatusemail<>"" then %><input type="checkbox" name="emailstat" value="1" <% if request.form("emailstat")="1" OR alwaysemailstatus=true then response.write "checked"%> /> <%=yyEStat%><% end if %></td>
				<td align="center"><input type="hidden" name="maxitems" value="<%=rowcounter%>" /><input type="submit" value="<%=yyUpdate%>" onclick="document.mainform.act.value='status';" /> <input type="reset" value="<%=yyReset%>" /></td>
			  </tr>
			  </form>
			  <form method="post" action="dumporders.asp" name="dumpform">
<%		if request.form("powersearch")="1" then %>
			  <input type="hidden" name="powersearch" value="1" />
			  <input type="hidden" name="fromdate" value="<%=Trim(request.form("fromdate"))%>" />
			  <input type="hidden" name="todate" value="<%=Trim(request.form("todate"))%>" />
			  <input type="hidden" name="ordid" value="<%=Trim(Replace(Replace(request.form("ordid"),"'",""),"""",""))%>" />
			  <input type="hidden" name="origsearchtext" value="<%=Trim(Replace(request.form("searchtext"),"""","&quot;"))%>" />
			  <input type="hidden" name="searchtext" value="<%=Trim(Replace(request.form("searchtext"),"""","&quot;"))%>" />
			  <input type="hidden" name="ordstatus" value="<%=Trim(request.form("ordstatus"))%>" />
			  <input type="hidden" name="startwith" value="<% if request.cookies("powersearch")="1" then response.write "1"%>" />
<%		end if %>
			  <input type="hidden" name="sd" value="<%=DateValue(sd)%>" />
			  <input type="hidden" name="ed" value="<%=DateValue(ed)%>" />
			  <input type="hidden" name="act" value="" />
			  <tr> 
                <td colspan="8" align="center"><select name="filedump" size="1">
					<option value="1"><%=yyDmpOrd%></option>
					<option value="2"><%=yyDmpDet%></option>
<%		if false then %>
					<option value="3">Dump orders to Quickbooks format</option>
<%		end if
		if ouresolutionsxml<>"" then response.write "<option value=""4"">OurESolutions XML format</option>" %>
					</select> <input type="submit" value="<%=yySubmit%>" onclick="setdumpformat()" /></td>
			  </tr>
			  </form>
<%	else %>
			  <tr> 
                <td width="100%" colspan="8" align="center">
					<p><%
					if request.form("powersearch")="1" then
						response.write yyNoMat1
					elseif sd=ed then
						response.write yyNoMat2&" "&sd&"."
					else
						response.write yyNoMat3&" "&sd&" "&yyAnd&" "&ed&"."
					end if %></p>
				</td>
			  </tr>
			  <% if hasdeleted then %>
			  <tr>
				<td colspan="8"><input type="submit" value="<%=yyPurDel%>" onclick="document.mainform.act.value='purge';" /></td>
			  </tr>
			  <% end if %>
			  </form>
<%	end if %>
			  <tr> 
                <td width="100%" colspan="8" align="center"><p><br />
					<a href="adminorders.asp?sd=<%=DateAdd("m",-1,sd)%>&ed=<%=DateAdd("m",-1,ed)%>"><strong>- <%=yyMonth%></strong></a> | 
					<a href="adminorders.asp?sd=<%=DateValue(sd)-7%>&ed=<%=DateValue(ed)-7%>"><strong>- <%=yyWeek%></strong></a> | 
					<a href="adminorders.asp?sd=<%=DateValue(sd)-1%>&ed=<%=DateValue(ed)-1%>"><strong>- <%=yyDay%></strong></a> | 
					<a href="adminorders.asp"><strong><%=yyToday%></strong></a> | 
					<a href="adminorders.asp?sd=<%=DateValue(sd)+1%>&ed=<%=DateValue(ed)+1%>"><strong><%=yyDay%> +</strong></a> | 
					<a href="adminorders.asp?sd=<%=DateValue(sd)+7%>&ed=<%=DateValue(ed)+7%>"><strong><%=yyWeek%> +</strong></a> | 
					<a href="adminorders.asp?sd=<%=DateAdd("m",1,sd)%>&ed=<%=DateAdd("m",1,ed)%>"><strong><%=yyMonth%> +</strong></a>
				  </p></td>
			  </tr>
			</table>
		  </td>
		</tr>
      </table>
<%
end if
end if
cnn.Close
set rs = nothing
set cnn = nothing
%>
