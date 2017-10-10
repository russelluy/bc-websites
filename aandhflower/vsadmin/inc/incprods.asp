<%
if storesessionvalue="" then storesessionvalue="virtualstore"
if Session("loggedon") <> storesessionvalue OR disallowlogin=TRUE then response.end
Dim sSQL,rs,alldata,success,cnn,rowcounter,allsections,alloptions,errmsg,prodoptions,aFields(3),dorefresh,thecat
Sub writemenulevel(id,itlevel)
	Dim wmlindex
	if itlevel<10 then
		FOR wmlindex=0 TO ubound(alldata,2)
			if alldata(2,wmlindex)=id then
				response.write "<option value='"&alldata(0,wmlindex)&"'"
				if thecat=alldata(0,wmlindex) then response.write " selected>" else response.write ">"
				for index = 0 to itlevel-2
					response.write "&nbsp;&nbsp;&raquo;&nbsp;"
				next
				response.write alldata(1,wmlindex)&"</option>" & vbCrLf
				if alldata(3,wmlindex)=0 then call writemenulevel(alldata(0,wmlindex),itlevel+1)
			end if
		NEXT
	end if
end Sub
Function writepagebar(CurPage, iNumPages)
	Dim sLink, i, sStr, startPage, endPage
	sLink = "<a href=""adminprods.asp?rid="&request("rid")&"&stock="&request("stock")&"&scat="&request("scat")&"&stext="&server.urlencode(request("stext"))&"&stype="&request("stype")&"&sprice="&server.urlencode(maxprice)&IIfVr(minprice<>"","&sminprice="&minprice,"")&"&pg="
	startPage = vrmax(1,CInt(Int(CDbl(CurPage)/10.0)*10))
	endPage = vrmin(iNumPages,CInt(Int(CDbl(CurPage)/10.0)*10)+10)
	if CurPage > 1 then
		sStr = sLink & "1" & """><strong><font face=""Verdana"">&laquo;</font></strong></a> " & sLink & CurPage-1 & """>"&yyPrev&"</a> | "
	else
		sStr = "<strong><font face=""Verdana"">&laquo;</font></strong> "&yyPrev&" | "
	end if
	for i=startPage to endPage
		if i=CurPage then
			sStr = sStr & i & " | "
		else
			sStr = sStr & sLink & i & """>"
			if i=startPage AND i > 1 then sStr=sStr&"..."
			sStr = sStr & i
			if i=endPage AND i < iNumPages then sStr=sStr&"..."
			sStr = sStr & "</a> | "
		end if
	next
	if CurPage < iNumPages then
		writepagebar = sStr & sLink & CurPage+1 & """>"&yyNext&"</a> " & sLink & iNumPages & """><strong><font face=""Verdana"">&raquo;</font></strong></a>"
	else
		writepagebar = sStr & " "&yyNext&" <strong><font face=""Verdana"">&raquo;</font></strong>"
	end if
End function
'for each objItem in request.form
'	response.write objItem&":"&request.form(objItem)&"<br>"
'next
success=true
Set rs = Server.CreateObject("ADODB.RecordSet")
Set rs2 = Server.CreateObject("ADODB.RecordSet")
Set cnn=Server.CreateObject("ADODB.Connection")
cnn.open sDSN
alreadygotadmin = getadminsettings()
Session.LCID=1033
simpleOptions = (adminTweaks AND 2)=2
simpleSections = (adminTweaks AND 4)=4
if maxprodsects="" then maxprodsects=20
dorefresh=FALSE
if request.form("posted")="1" then
	pExemptions=0
	if Trim(request.form("pExemptions"))<>"" then
		pExemptArray=Split(request.form("pExemptions"), ",")
		for each pExemptObj in pExemptArray
			pExemptions = pExemptions + pExemptObj
		next
	end if
	if request.form("act")="delete" then
		sSQL = "DELETE FROM pricebreaks WHERE pbProdID='"&request.form("id")&"'"
		cnn.Execute(sSQL)
		sSQL = "DELETE FROM cpnassign WHERE cpaType=2 AND cpaAssignment='"&request.form("id")&"'"
		cnn.Execute(sSQL)
		sSQL = "DELETE FROM products WHERE pID='" & request.form("id")&"'"
		cnn.Execute(sSQL)
		sSQL = "DELETE FROM prodoptions WHERE poProdID='" & request.form("id")&"'"
		cnn.Execute(sSQL)
		sSQL = "DELETE FROM multisections WHERE pID='" & request.form("id")&"'"
		cnn.Execute(sSQL)
		sSQL = "DELETE FROM relatedprods WHERE rpProdID='" & request.form("id")&"' OR rpRelProdID='" & request.form("id")&"'"
		cnn.Execute(sSQL)
		dorefresh=TRUE
	elseif request.form("act")="updaterelations" then
		rid=trim(request.form("rid"))
		for each objItem in request.form
			if left(objItem,4)="updq" then
				theprodid=right(objItem,len(objItem)-4)
				sSQL = "DELETE FROM relatedprods WHERE rpProdID='"&replace(rid,"'","''")&"' AND rpRelProdID='"&replace(right(objItem,len(objItem)-4),"'","''")&"'"
				cnn.Execute(sSQL)
				if request.form("updr"&theprodid)="1" then
					sSQL = "INSERT INTO relatedprods (rpProdID,rpRelProdID) VALUES ('"&replace(rid,"'","''")&"','"&replace(right(objItem,len(objItem)-4),"'","''")&"')"
					cnn.Execute(sSQL)
				end if
			end if
		next
		dorefresh=TRUE
	elseif request.form("act")="domodify" then
		if Trim(Request.Form("newid")) <> Trim(Request.Form("id")) then
			sSQL = "SELECT * FROM products WHERE pID='"&Trim(request.form("newID"))&"'"
			rs.Open sSQL,cnn,0,1
			success = rs.EOF
			rs.Close
			if success then
				cnn.Execute("UPDATE pricebreaks SET pbProdID='"&request.form("newid")&"' WHERE pbProdID='"&request.form("id")&"'")
				cnn.Execute("UPDATE cpnassign SET cpaAssignment='"&request.form("newid")&"' WHERE cpaType=2 AND cpaAssignment='"&request.form("id")&"'")
			end if
		end if
		if success then
			pOrder = Trim(Request.Form("pOrder"))
			if NOT isnumeric(pOrder) then pOrder=0
			sSQL = "UPDATE products SET " & _
						"pID='"& Trim(Request.Form("newid")) &"', " & _
						"pName='"& Replace(Trim(Request.Form("pName")),"'","''") &"', " & _
						"pSection="& Trim(Request.Form("pSection")) &", " & _
						"pDropship="& Trim(Request.Form("pDropship")) &", " & _
						"pOrder="& pOrder &", " & _
						"pExemptions="& pExemptions &", " & _
						"pDescription='"& Replace(Trim(Request.Form("pDescription")),"'","''") &"', " & _
						"pImage='"& Replace(Trim(Request.Form("pImage")),"'","''") &"', " & _
						"pLongDescription='"& Replace(Trim(Request.Form("pLongDescription")),"'","''") &"', "
						for index=2 to adminlanguages+1
							if (adminlangsettings AND 1)=1 then sSQL = sSQL & "pName"&index&"='"& Replace(Trim(Request.Form("pName"&index)),"'","''") &"', "
							if (adminlangsettings AND 2)=2 then sSQL = sSQL & "pDescription"&index&"='"& Replace(Trim(Request.Form("pDescription"&index)),"'","''") &"', "
							if (adminlangsettings AND 4)=4 then sSQL = sSQL & "pLongDescription"&index&"='"& Replace(Trim(Request.Form("pLongDescription"&index)),"'","''") &"', "
						next
						sSQL = sSQL & "pLargeImage='"& Replace(Trim(Request.Form("pLargeImage")),"'","''") &"', "
						if Trim(Request.Form("pDisplay")) = "ON" then
							sSQL = sSQL & "pDisplay=1,"
						else
							sSQL = sSQL & "pDisplay=0,"
						end if
						if perproducttaxrate=true then
							sSQL = sSQL & "pTax=" & Trim(Request.Form("pTax")) & ","
						end if
						if stockManage<>0 AND IsNumeric(Trim(request.form("inStock"))) then
							sSQL = sSQL & "pInStock=" & Trim(request.form("inStock"))&","
						end if
						sSQL = sSQL & "pStockByOpts=" & IIfVr(Trim(Request.Form("pStockByOpts")) = "1", 1, 0) & ","
						sSQL = sSQL & "pStaticPage=" & IIfVr(Trim(Request.Form("pStaticPage")) = "1", 1, 0) & ","
						sSQL = sSQL & "pRecommend=" & IIfVr(Trim(Request.Form("pRecommend")) = "1", 1, 0) & ","
						sSQL = sSQL & "pSell=" & IIfVr(Trim(Request.Form("pSell")) = "ON", 1, 0) & ","
						if (adminUnits AND 12) > 0 then
							sSQL = sSQL & "pDims='" & Trim(request.form("plen")) & "x" & Trim(request.form("pwid")) & "x" & Trim(request.form("phei")) & "',"
						end if
						if digidownloads=true then
							sSQL = sSQL & "pDownload='" & trim(replace(request.form("pDownload"),"'","''")) & "',"
						end if
						if shipType=1 then
							if NOT IsNumeric(Trim(request.form("pShipping"))) then
								sSQL = sSQL & "pShipping=0,"
							else
								sSQL = sSQL & "pShipping="&Trim(request.form("pShipping"))&","
							end if
							if NOT IsNumeric(Trim(request.form("pShipping2"))) then
								sSQL = sSQL & "pShipping2=0,"
							else
								sSQL = sSQL & "pShipping2="&Trim(request.form("pShipping2"))&","
							end if
						elseif shipType=2 OR shipType=3 OR shipType=4 OR shipType=6 OR shipType=7 then
							if NOT IsNumeric(Trim(request.form("pShipping"))) then
								sSQL = sSQL & "pWeight=0,"
							else
								sSQL = sSQL & "pWeight="&Trim(request.form("pShipping"))&","
							end if
						end if
						if Trim(Request.Form("pWholesalePrice"))<>"" then
							sSQL = sSQL & "pWholesalePrice="& Trim(Request.Form("pWholesalePrice")) &","
						else
							sSQL = sSQL & "pWholesalePrice=0,"
						end if
						if Trim(Request.Form("pListPrice"))<>"" then
							sSQL = sSQL & "pListPrice="& Trim(Request.Form("pListPrice")) &","
						else
							sSQL = sSQL & "pListPrice=0,"
						end if
						sSQL = sSQL & "pPrice="& Trim(Request.Form("pPrice")) &" " & _
						"WHERE pID='"&Request.Form("id")&"'"
			on error resume next
			cnn.Execute(sSQL)
			sSQL = "DELETE FROM prodoptions WHERE poProdID='"&Request.Form("id")&"'"
			cnn.Execute(sSQL)
			for rowcounter=0 to maxprodopts-1
				if request.form("pOption"&rowcounter)<>"" AND request.form("pOption"&rowcounter)<>0 then
					sSQL = "INSERT INTO prodoptions (poProdID,poOptionGroup) VALUES ('"&Request.Form("newid")&"',"&request.form("pOption"&rowcounter)&")"
					cnn.Execute(sSQL)
				end if
			next
			sSQL = "DELETE FROM multisections WHERE pID='"&Request.Form("id")&"'"
			cnn.Execute(sSQL)
			for rowcounter=0 to maxprodsects-1
				if request.form("pSection"&rowcounter)<>"" AND request.form("pSection"&rowcounter)<>0 AND Request.Form("pSection")<>request.form("pSection"&rowcounter) then
					sSQL = "SELECT pID FROM multisections WHERE pID='" & Request.Form("newid") & "' AND pSection="&request.form("pSection"&rowcounter)
					rs.Open sSQL,cnn,0,1
					if rs.EOF then
						sSQL = "INSERT INTO multisections (pID,pSection) VALUES ('"&Request.Form("newid")&"',"&request.form("pSection"&rowcounter)&")"
						cnn.Execute(sSQL)
					end if
					rs.Close
				end if
			next
			if err.number<>0 then
				success=false
				errmsg = "There was an error writing to the database.<br />" & err.description
			else
				dorefresh=TRUE
			end if
			on error goto 0
		else
			errmsg = yyPrDup
		end if
	elseif request.form("act")="doaddnew" then
		sSQL = "SELECT * FROM products WHERE pID='"&Trim(request.form("newID"))&"'"
		rs.Open sSQL,cnn,0,1
		success = rs.EOF
		rs.Close
		if success then
			pOrder = Trim(Request.Form("pOrder"))
			if NOT isnumeric(pOrder) then pOrder=0
			sSQL = "INSERT INTO products (pID,pName,pSection,pDropship,pOrder,pExemptions,pDescription,pImage,pLongDescription,"
			for index=2 to adminlanguages+1
				if (adminlangsettings AND 1)=1 then sSQL = sSQL & "pName" & index & ","
				if (adminlangsettings AND 2)=2 then sSQL = sSQL & "pDescription" & index & ","
				if (adminlangsettings AND 4)=4 then sSQL = sSQL & "pLongDescription" & index & ","
			next
			sSQL = sSQL & "pLargeImage,pPrice,pWholesalePrice,pListPrice,"
			if shipType=1 then sSQL = sSQL & "pShipping,pShipping2,"
			sSQL = sSQL & "pDisplay,"
			if perproducttaxrate=true then sSQL = sSQL & "pTax,"
			if stockManage<>0 AND IsNumeric(Trim(request.form("inStock"))) then sSQL = sSQL & "pInStock,"
			if (adminUnits AND 12) > 0 then sSQL = sSQL & "pDims,"
			if digidownloads=true then sSQL = sSQL & "pDownload,"
			sSQL = sSQL & "pStockByOpts,pStaticPage,pRecommend,pSell,pWeight) VALUES (" & _
						"'"&Trim(request.form("newID"))&"'," & _
						"'"&replace(request.form("pName"),"'","''")&"'," & _
						request.form("pSection")&"," & _
						request.form("pDropship")&"," & _
						pOrder&"," & _
						pExemptions &"," & _
						"'"&replace(request.form("pDescription"),"'","''")&"'," & _
						"'"&replace(request.form("pImage"),"'","''")&"'," & _
						"'"&replace(request.form("pLongDescription"),"'","''")&"',"
						for index=2 to adminlanguages+1
							if (adminlangsettings AND 1)=1 then sSQL = sSQL & "'"& Replace(Trim(Request.Form("pName"&index)),"'","''") &"',"
							if (adminlangsettings AND 2)=2 then sSQL = sSQL & "'"& Replace(Trim(Request.Form("pDescription"&index)),"'","''") &"',"
							if (adminlangsettings AND 4)=4 then sSQL = sSQL & "'"& Replace(Trim(Request.Form("pLongDescription"&index)),"'","''") &"',"
						next
						sSQL = sSQL & "'"&replace(request.form("pLargeImage"),"'","''")&"'," & _
						Trim(request.form("pPrice"))&","
						if Trim(request.form("pWholesalePrice"))<>"" then
							sSQL = sSQL & Trim(request.form("pWholesalePrice")) & ","
						else
							sSQL = sSQL & "0,"
						end if
						if Trim(request.form("pListPrice"))<>"" then
							sSQL = sSQL & Trim(request.form("pListPrice")) & ","
						else
							sSQL = sSQL & "0,"
						end if
						if shipType=1 then
							if NOT IsNumeric(Trim(request.form("pShipping"))) then
								sSQL = sSQL & "0,"
							else
								sSQL = sSQL & Trim(request.form("pShipping"))&","
							end if
							if NOT IsNumeric(Trim(request.form("pShipping2"))) then
								sSQL = sSQL & "0,"
							else
								sSQL = sSQL & Trim(request.form("pShipping2"))&","
							end if
						end if
						if Trim(Request.Form("pDisplay")) = "ON" then
							sSQL = sSQL & "1,"
						else
							sSQL = sSQL & "0,"
						end if
						if perproducttaxrate=true then sSQL = sSQL & request.form("pTax") & ","
						if stockManage<>0 AND IsNumeric(Trim(request.form("inStock"))) then
							sSQL = sSQL & Trim(request.form("inStock"))&","
						end if
						if (adminUnits AND 12) > 0 then
							sSQL = sSQL & "'" & Trim(request.form("plen")) & "x" & Trim(request.form("pwid")) & "x" & Trim(request.form("phei")) & "',"
						end if
						if digidownloads=true then
							sSQL = sSQL & "'" & trim(replace(request.form("pDownload"),"'","''")) & "',"
						end if
						sSQL = sSQL & IIfVr(Trim(Request.Form("pStockByOpts")) = "1", 1, 0) & ","
						sSQL = sSQL & IIfVr(Trim(Request.Form("pStaticPage")) = "1", 1, 0) & ","
						sSQL = sSQL & IIfVr(Trim(Request.Form("pRecommend")) = "1", 1, 0) & ","
						sSQL = sSQL & IIfVr(Trim(Request.Form("pSell")) = "ON", 1, 0) & ","
						if shipType <= 1 OR NOT IsNumeric(Trim(request.form("pShipping"))) then
							sSQL = sSQL & "0"
						elseif shipType=2 OR shipType=3 OR shipType=4 OR shipType=6 OR shipType=7 then
							sSQL = sSQL & Trim(request.form("pShipping"))&""
						else
							sSQL = sSQL & Trim(request.form("pShipping"))&"."
							if Int(Trim(request.form("pShipping2"))) < 10 then sSQL = sSQL & "0"
							sSQL = sSQL & Trim(request.form("pShipping2"))
						end if
						sSQL = sSQL & ")"
			on error resume next
			cnn.Execute(sSQL)
			for rowcounter=0 to maxprodopts-1
				if request.form("pOption"&rowcounter)<>"" AND request.form("pOption"&rowcounter)<>0 then
					sSQL = "INSERT INTO prodoptions (poProdID,poOptionGroup) VALUES ('"&Request.Form("newid")&"',"&request.form("pOption"&rowcounter)&")"
					cnn.Execute(sSQL)
				end if
			next
			sSQL = "DELETE FROM multisections WHERE pID='"&Request.Form("newid")&"'"
			cnn.Execute(sSQL)
			for rowcounter=0 to maxprodsects-1
				if request.form("pSection"&rowcounter)<>"" AND request.form("pSection"&rowcounter)<>0 AND Request.Form("pSection")<>request.form("pSection"&rowcounter) then
					sSQL = "SELECT pID FROM multisections WHERE pID='" & Request.Form("newid") & "' AND pSection="&request.form("pSection"&rowcounter)
					rs.Open sSQL,cnn,0,1
					if rs.EOF then
						sSQL = "INSERT INTO multisections (pID,pSection) VALUES ('"&Request.Form("newid")&"',"&request.form("pSection"&rowcounter)&")"
						cnn.Execute(sSQL)
					end if
					rs.Close
				end if
			next
			if err.number<>0 then
				success=false
				errmsg = errmsg & err.description
			else
				dorefresh=TRUE
			end if
			on error goto 0
		else
			errmsg = yyPrDup
		end if
	elseif request.form("act")="dodiscounts" then
		sSQL = "INSERT INTO cpnassign (cpaCpnID,cpaType,cpaAssignment) VALUES ("&request.form("assdisc")&",2,'"&request.form("id")&"')"
		cnn.Execute(sSQL)
		dorefresh=TRUE
	elseif request.form("act")="deletedisc" then
		sSQL = "DELETE FROM cpnassign WHERE cpaID="&request.form("id")
		cnn.Execute(sSQL)
		dorefresh=TRUE
	end if
	if request.form("act")="modify" OR request.form("act")="clone" OR request.form("act")="addnew" then
		sSQL = "SELECT optGrpID, optGrpWorkingName FROM optiongroup ORDER BY optGrpWorkingName"
		rs.Open sSQL,cnn,0,1
		if NOT rs.EOF then alloptions=rs.getrows
		rs.Close
		if request.form("act")="modify" OR request.form("act")="clone" then
			sSQL = "SELECT poID, poOptionGroup FROM prodoptions WHERE poProdID='"&Trim(Request.Form("id"))&"'"
			rs.Open sSQL,cnn,0,1
			if NOT rs.EOF then prodoptions=rs.getrows
			rs.Close
			sSQL = "SELECT pSection FROM multisections WHERE pID='"&Trim(Request.Form("id"))&"'"
			rs.Open sSQL,cnn,0,1
			if NOT rs.EOF then prodsections=rs.getrows
			rs.Close
		end if
		sSQL = "SELECT sectionID,sectionWorkingName FROM sections WHERE rootSection=1 ORDER BY sectionWorkingName"
		rs.Open sSQL,cnn,0,1
		if NOT rs.EOF then allsections=rs.getrows
		rs.Close
		sSQL = "SELECT dsID,dsName FROM dropshipper ORDER BY dsName"
		rs.Open sSQL,cnn,0,1
		if NOT rs.EOF then alldropship=rs.getrows
		rs.Close
	end if
end if
if dorefresh then
	response.write "<meta http-equiv=""refresh"" content=""1; url=adminprods.asp"
	response.write "?rid="&request.form("rid")&"&stock=" & request.form("stock") & "&stext=" & server.urlencode(request.form("stext")) & "&sprice=" & server.urlencode(request.form("sprice")) & "&stype=" & request.form("stype") & "&scat=" & request.form("scat") & "&pg=" & request.form("pg")
	response.write """>"
end if
%>
<% if request.form("act")="addnew" OR request.form("act")="modify" OR request.form("act")="clone" then %>
<script language="javascript" type="text/javascript">
function checkastring(thestr,validchars){
  for (i=0; i < thestr.length; i++){
    ch = thestr.charAt(i);
    for (j = 0;  j < validchars.length;  j++)
      if (ch == validchars.charAt(j))
        break;
    if (j == validchars.length)
	  return(false);
  }
  return(true);
}
function formvalidator(theForm){
  if (theForm.newid.value == ""){
    alert("<%=yyPlsEntr%> \"<%=yyPrRef%>\".");
    theForm.newid.focus();
    return (false);
  }
  if (theForm.pSection.options[theForm.pSection.selectedIndex].value == ""){
    alert("<%=yyPlsSel%> \"<%=yySection%>\".");
    theForm.pSection.focus();
    return (false);
  }
  if (theForm.pName.value == ""){
    alert("<%=yyPlsEntr%> \"<%=yyPrNam%>\".");
    theForm.pName.focus();
    return (false);
  }
<%	for index=2 to adminlanguages+1
		if (adminlangsettings AND 1)=1 then %>
  if (theForm.pName<%=index%>.value == ""){
    alert("<%=yyPlsEntr%> \"<%=yyPrNam & " " & index%>\".");
    theForm.pName<%=index%>.focus();
    return (false);
  }
<%		end if
	next %>
  if (theForm.pPrice.value == ""){
    alert("<%=yyPlsEntr%> \"<%=yyPrPri%>\".");
    theForm.pPrice.focus();
    return (false);
  }
  var checkOK = "'\" ";
  var checkStr = theForm.newid.value;
  var allValid = true;
  for (i = 0;  i < checkStr.length;  i++){
    ch = checkStr.charAt(i);
    for (j = 0;  j < checkOK.length;  j++)
      if (ch == checkOK.charAt(j)){
	    allValid = false;
        break;
	  }
  }
  if (!allValid){
    alert("<%=yyQuoSpa%> \"<%=yyPrRef%>\".");
    theForm.newid.focus();
    return (false);
  }
  if (!checkastring(theForm.pPrice.value,"0123456789.")){
    alert("<%=yyOnlyDec%> \"<%=yyPrPri%>\".");
    theForm.pPrice.focus();
    return (false);
  }
  if (!checkastring(theForm.pWholesalePrice.value,"0123456789.")){
    alert("<%=yyOnlyDec%> \"<%=yyWhoPri%>\".");
    theForm.pWholesalePrice.focus();
    return (false);
  }
  if (!checkastring(theForm.pListPrice.value,"0123456789.")){
    alert("<%=yyOnlyDec%> \"<%=yyListPr%>\".");
    theForm.pListPrice.focus();
    return (false);
  }
<%	if (adminUnits AND 12) > 0 then %>
  var checkOK = "0123456789.";
  if (!checkastring(theForm.plen.value,checkOK)){
	alert("<%=yyOnlyDec%> \"<%=yyDims%>\".");
	theForm.plen.focus();
	return(false);
  }
  if (!checkastring(theForm.pwid.value,checkOK)){
	alert("<%=yyOnlyDec%> \"<%=yyDims%>\".");
	theForm.pwid.focus();
	return(false);
  }
  if (!checkastring(theForm.phei.value,checkOK)){
	alert("<%=yyOnlyDec%> \"<%=yyDims%>\".");
	theForm.phei.focus();
	return(false);
  }
<%	end if
	if shipType=1 OR shipType=2 OR shipType=3 OR shipType=4 OR shipType=6 OR shipType=7 then %>
  var checkOK = "0123456789.";
  if (!checkastring(theForm.pShipping.value,checkOK)){
<%   if shipType=1 then %>
    alert("<%=yyOnlyDec%> \"<%=yyShip & ": " & yyFirShi%>\".");
<%   else %>
    alert("<%=yyOnlyDec%> \"<%=yyPrWght%>\".");
<%   end if %>
    theForm.pShipping.focus();
    return (false);
  }
<%	end if
	if shipType=1 then %>
  if (!checkastring(theForm.pShipping2.value,"0123456789.")){
    alert("<%=yyOnlyDec%> \"<%=yyShip & ": " & yySubShi%>\".");
    theForm.pShipping2.focus();
    return (false);
  }
<%	end if
	if stockManage<>0 then %>
  if (!(theForm.pStockByOpts.selectedIndex==1) && theForm.inStock.value == ""){
    alert("<%=yyPlsEntr%> \"<%=yyInStk%>\".");
    theForm.inStock.focus();
    return (false);
  }
  if (!(theForm.pStockByOpts.selectedIndex==1) && !checkastring(theForm.inStock.value,"0123456789")){
    alert("<%=yyOnlyNum%> \"<%=yyInStk%>\".");
    theForm.inStock.focus();
    return (false);
  }
  if(theForm.pStockByOpts.selectedIndex==1 && theForm.pNumOptions.selectedIndex==0){
    alert("<%=yyStkWrn%>");
    theForm.pStockByOpts.focus();
    return (false);
  }
<%	end if
	if perproducttaxrate=true then %>
  if (theForm.pTax.value == ""){
	alert("<%=yyPlsEntr%> \"<%=yyTax%>\".");
	theForm.pTax.focus();
	return(false);
  }
  if (!checkastring(theForm.pTax.value,"0123456789.")){
    alert("<%=yyOnlyDec%> \"<%=yyTax%>\".");
    theForm.pTax.focus();
    return (false);
  }
<%	end if %>
  if (!checkastring(theForm.pOrder.value,"0123456789")){
    alert("<%=yyOnlyNum%> \"<%=yyProdOr%>\".");
    theForm.pOrder.focus();
    return (false);
  }
  return (true);
}
var prodOptGrpArr = new Array();
var prodSectGrpArr = new Array();
<%
rowcounter=0
if IsArray(prodoptions) then
	for rowcounter=0 to UBOUND(prodoptions,2)
		response.write "prodOptGrpArr["&rowcounter&"]="&prodoptions(1,rowcounter)&";"&vbCrLf
	next
end if
response.write "for(ii=" & rowcounter & ";ii<" & maxprodsects & ";ii++) prodOptGrpArr[ii]=0;" & vbCrLf
rowcounter=0
if IsArray(prodsections) then
	for rowcounter=0 to UBOUND(prodsections,2)
		response.write "prodSectGrpArr["&rowcounter&"]="&prodsections(0,rowcounter)&";"&vbCrLf
	next
end if
response.write "for(ii=" & rowcounter & ";ii<" & maxprodsects & ";ii++) prodSectGrpArr[ii]=0;" & vbCrLf
%>
function update_opts(index){
	var thisOption = document.getElementById('pOption'+index);
	prodOptGrpArr[index] = thisOption.options[thisOption.selectedIndex].value;
}
function update_sects(index){
	var thisSection = document.getElementById('pSection'+index);
	prodSectGrpArr[index] = thisSection.options[thisSection.selectedIndex].value;
}
function setprodoptions(){
	var noOpts = document.forms.mainform.pNumOptions.selectedIndex;
	var theElm;
	var theHTMLHead,theHTML="";
	var index=0;
	theElm = document.getElementById('prodoptions');
	theHTMLHead = '<table width="100%" border="0" cellspacing="0" cellpadding="3">';
	theHTML = theHTML + '<select size="1" id="pOptionGGREPLACEMExx" name="pOptionGGREPLACEMExx" onchange="update_opts(GGREPLACEMExx);"><option value="0"><%=yyNone%></option>';
	<%	if IsArray(alloptions) then
			for rowcounter=0 to UBOUND(alloptions,2)
				response.write "theHTML = theHTML +'<option value="""&alloptions(0,rowcounter)&""">"&replace(alloptions(1,rowcounter),"'","\'")&"</option>';" & vbCrLf
			next
		end if %>
	theHTML = theHTML + '</select>';
	for (index=0;index<noOpts;index++){
		if(index % 2 == 0) theHTMLHead = theHTMLHead + '<tr>';
		theHTMLHead = theHTMLHead + '<td width="25%" align="right"><%=yyPrdOpt%> '+(index+1)+':</td><td width="25%">'+theHTML.replace(/GGREPLACEMExx/g,index)+'</td>';
		if(index % 2 != 0) theHTMLHead = theHTMLHead + '</tr>';
	}
	if(index % 2 != 0) theHTMLHead = theHTMLHead + '<td width="50%" colspan="2">&nbsp;</td></tr>';
	theHTMLHead = theHTMLHead + '</table>';
	theElm.innerHTML=theHTMLHead;
	for (index=0;index<noOpts;index++){
		var thisOption = document.getElementById('pOption'+index);
		for (index2=0;index2<thisOption.length;index2++){
			if (thisOption[index2].value==prodOptGrpArr[index]){
				thisOption.selectedIndex=index2;
				thisOption.options[index2].selected = true;
			}
			else
				thisOption.options[index2].selected = false;
		}
	}
}
function setprodsections(){
	var noSects = document.forms.mainform.pNumSections.selectedIndex;
	var theHTMLHead,theHTML="";
	var index=0;
	var theElm = document.getElementById('prodsections');
	theHTMLHead = '<table width="100%" border="0" cellspacing="0" cellpadding="3">';
	theHTML = theHTML + '<select size="1" id="pSectionGGREPLACEMExx" name="pSectionGGREPLACEMExx" onchange="update_sects(GGREPLACEMExx);"><option value="0"><%=yyNone%></option>';
	<%	if IsArray(allsections) then
			for rowcounter=0 to UBOUND(allsections,2)
				response.write "theHTML = theHTML +'<option value="""&allsections(0,rowcounter)&""">"&replace(allsections(1,rowcounter)&"","'","\'")&"</option>';" & vbCrLf
			next
		end if %>
	theHTML = theHTML + '</select>';
	for (index=0;index<noSects;index++){
		if(index % 2 == 0) theHTMLHead = theHTMLHead + '<tr>';
		theHTMLHead = theHTMLHead + '<td width="25%" align="right">Prod. Section '+(index+1)+':</td><td width="25%">'+theHTML.replace(/GGREPLACEMExx/g,index)+'</td>';
		if(index % 2 != 0) theHTMLHead = theHTMLHead + '</tr>';
	}
	if(index % 2 != 0) theHTMLHead = theHTMLHead + '<td width="50%" colspan="2">&nbsp;</td></tr>';

	theHTMLHead = theHTMLHead + '</table>';
	theElm.innerHTML=theHTMLHead;
	for (index=0;index<noSects;index++){
		var thisSection = document.getElementById('pSection'+index);
		for (index2=0;index2<thisSection.length;index2++){
			if (thisSection[index2].value==prodSectGrpArr[index]){
				thisSection.selectedIndex=index2;
				thisSection.options[index2].selected = true;
			}
			else
				thisSection.options[index2].selected = false;
		}
	}
}
function setstocktype(){
var si = document.forms.mainform.pStockByOpts.selectedIndex;
document.forms.mainform.inStock.disabled=(si==1);
}
</script>
<%
end if
sub show_info()
%>
		<p><a name="info"></a><ul>
		  <li><font size="1"><%=yyPrEx1%></font></li>
		  <li><font size="1"><%=yyPrEx2%></font></li>
		</ul></p>
<%
end sub
  if request.form("posted")="1" AND (request.form("act")="modify" OR request.form("act")="clone" OR request.form("act")="addnew") then
		Dim pNames(10)
		if htmleditor="tinymce" then %>
<script language="javascript" type="text/javascript" src="tiny_mce.js"></script>
<script language="javascript" type="text/javascript">
	tinyMCE.init({
		theme : "simple",
		mode : "textareas",
		// save_callback : "customSave",
		valid_elements : "*[*]",
		extended_valid_elements : "a[class|href|target|name|onclick]," +
			"embed[quality|type|pluginspage|width|height|src|align]," +
			"hr[class|width|size|noshade]," + 
			"img[class|src|border|alt|title|hspace|vspace|width|height|align|onmouseover|onmouseout|name]," +
			"object[classid|codebase|width|height|align]," +
			"param[name|value]," +
			"input[checked|class|disabled|id|name|type|value|size|maxlength|src|width|height|readonly|tabindex|onfocus|onblur|onchange|onselect]",
		//plugins : "table",
		//theme_advanced_buttons3_add_before : "tablecontrols,separator",
		//invalid_elements : "a",
		//theme_advanced_styles : "Header 1=header1;Header 2=header2;Header 3=header3;Table Row=tableRow1", // Theme specific setting CSS classes
		//execcommand_callback : "myCustomExecCommandHandler",
		debug : false
	});
	tinyMCE.addToLang('',{
		plus_desc : 'Plus'
	});
</script>
<%		end if
		if request.form("act")="modify" OR request.form("act")="clone" then
			doaddnew = false
			sSQL = "SELECT pId,pName,pName2,pName3,pSection,pImage,pPrice,pWholesalePrice,pListPrice,pDisplay,pStaticPage,pRecommend,pStockByOpts,pSell,pShipping,pShipping2,pLargeImage,pWeight,pExemptions,pInStock,pDims,pTax,pDropship,pOrder,"
			if digidownloads=true then sSQL = sSQL & "pDownload,"
			sSQL = sSQL & "pDescription,pLongDescription FROM products WHERE pId='"&Request.Form("id")&"'"
			rs.Open sSQL,cnn,0,1
				pName = rs("pName")
				for index=2 to adminlanguages+1
					pNames(index)=rs("pName" & index)
				next
				pID = rs("pID")
				pSection = rs("pSection")
				pImage = rs("pImage")&""
				pPrice = rs("pPrice")
				pWholesalePrice = rs("pWholesalePrice")
				pListPrice = rs("pListPrice")
				pDisplay = rs("pDisplay")
				pStaticPage = rs("pStaticPage")
				pRecommend = rs("pRecommend")
				pStockByOpts = rs("pStockByOpts")
				pSell = rs("pSell")
				pShipping = rs("pShipping")
				pShipping2 = rs("pShipping2")
				pLargeImage = rs("pLargeImage")&""
				pWeight = rs("pWeight")
				pExemptions = rs("pExemptions")
				pInStock = rs("pInStock")
				pDims = rs("pDims")
				if digidownloads=true then pDownload = rs("pDownload")
				pTax = rs("pTax")
				pDropship = rs("pDropship")
				pOrder = rs("pOrder")
				pDescription = rs("pDescription")
				pLongDescription = rs("pLongDescription")
			rs.Close
		else
			doaddnew = true
			pID = ""
			if Trim(request.form("scat"))<>"" then pSection=Int(Trim(request.form("scat"))) else pSection = 0
			pImage = "prodimages/"
			pPrice = ""
			pWholesalePrice = ""
			pListPrice = 0
			pDisplay = 1
			pStaticPage = 0
			pRecommend = 0
			pStockByOpts = 0
			pSell = 1
			pShipping = ""
			pShipping2 = ""
			pLargeImage = "prodimages/"
			pWeight = ""
			pExemptions = 0
			pInStock = ""
			pDims = ""
			pDownload = ""
			pTax=0
			pDropship = 0
			pOrder=0
			pDescription = ""
			pLongDescription = ""
		end if
%>
	<form name="mainform" method="post" action="adminprods.asp" onsubmit="return formvalidator(this)">
      <table border="0" cellspacing="0" cellpadding="0" width="100%" bgcolor="" align="center">
		<tr>
		  <td width="100%">
			<input type="hidden" name="posted" value="1" />
			<%	if request.form("act")="modify" then %>
			<input type="hidden" name="act" value="domodify" />
			<input type="hidden" name="id" value="<%=pID%>" />
			<%	else %>
			<input type="hidden" name="act" value="doaddnew" />
			<%	end if
				call writehiddenvar("stock", request.form("stock"))
				call writehiddenvar("stext", request.form("stext"))
				call writehiddenvar("sprice", request.form("sprice"))
				call writehiddenvar("scat", request.form("scat"))
				call writehiddenvar("stype", request.form("stype"))
				call writehiddenvar("pg", request.form("pg")) %>
			<table width="100%" border="0" cellspacing="0" cellpadding="3" bgcolor="">
			  <tr> 
                <td width="100%" colspan="4" align="center"><strong><%
					if request.form("act")="modify" then
						response.write yyYouMod & " &quot;" & pName & "&quot;"
					elseif request.form("act")="addnew" then
						response.write yyPrUpd
					else
						response.write yyYouCln & " &quot;" & pName & "&quot;"
					end if
				%></strong><br />&nbsp;</td>
			  </tr>
			  <tr>
			    <td align="right"><font color="#FF0000">*</font><%=yyPrRef%>:</td><td><input type="text" name="newid" size="15" value="<%=pID%>" /></td>
			    <td align="right"><font color="#FF0000">*</font><%=yySection%>:</td><td><select size="1" name="pSection"><option value=""><%=yySelect%></option><%
					if IsArray(allsections) then
						for rowcounter=0 to UBOUND(allsections,2)
							response.write "<option value='"&allsections(0,rowcounter)&"'"
							if allsections(0,rowcounter)=pSection then response.write " selected"
							response.write ">"&allsections(1,rowcounter)&"</option>" &vbCrLf
						next
					end if %></select></td>
			  </tr>
			  <tr>
			    <td align="right"><font color="#FF0000">*</font><%=yyPrNam%>:</td><td><input type="text" name="pName" size="25" value="<%=replace(replace(pName,"&","&amp;"),"""","&quot;")%>" /></td>
			    <td align="right"><font color="#FF0000">*</font><%=yyPrPri%>:</td><td><input type="text" name="pPrice" size="15" value="<%=pPrice%>" /></td>
			  </tr>
<%				for index=2 to adminlanguages+1
					if (adminlangsettings AND 1)=1 then
			%><tr>
			    <td align="right"><font color="#FF0000">*</font><%=yyPrNam & " " & index%>:</td><td colspan="3"><input type="text" name="pName<%=index%>" size="25" value="<%=replace(replace(pNames(index)&"","&","&amp;"),"""","&quot;")%>" /></td>
			  </tr><%
					end if
				next %>
			  <tr>
				<% if useStockManagement then %>
				<td align="right">
				<input type="hidden" name="pSell" value="<% if int(pSell) <> 0 then response.write "ON" %>" />
				<select name="pStockByOpts" size="1" onchange="setstocktype();">
				<option value="0">&nbsp;&nbsp;&nbsp;<%=yyInStk%>:</option>
				<option value="1"<% if cint(pStockByOpts) <> 0 then response.write "selected" %>><%=yyByOpt%>:</option></select>
				</td><td><input type="text" name="inStock" size="10" value="<%=pInStock%>" /></td>
				<% else %>
				<input type="hidden" name="pStockByOpts" value="<% if cint(pStockByOpts)<>0 then response.write "1" %>" />
				<td align="right"><%=yySellBut%>:</td><td><input type="checkbox" name="pSell" value="ON" <% if int(pSell) <> 0 then response.write "checked" %> /></td>
				<% end if %>
			    <td align="right"><%=yyWhoPri%> <font size="1">(<a href="#info">info</a>)</font>:</td><td><input type="text" name="pWholesalePrice" size="15" value="<%=pWholesalePrice%>" /></td>
			  </tr>
			  <tr>
			    <td align="right"><%=yyDisPro%>:</td><td><input type="checkbox" name="pDisplay" value="ON" <% if cint(pDisplay)<>0 then response.write "checked" %> /></td>
				<td align="right"><%=yyListPr%> <font size="1">(<a href="#info">info</a>)</font>:</td><td><input type="text" name="pListPrice" size="15" value="<% if cDbl(pListPrice)<>0.0 then response.write pListPrice %>" /></td>
			  </tr>
			  <tr>
			    <td align="right"><%=yyImage%>:</td><td><input type="text" name="pImage" size="25" value="<%=replace(pImage,"""","&quot;")%>" /></td>
				<%	if (adminUnits AND 12) > 0 then
						proddims = split(pDims&"", "x") %>
				<td align="right"><%=yyDims%>:</td>
				<td><input type="text" name="plen" size="4" value="<%if UBOUND(proddims)>=0 then response.write proddims(0)%>" /> <strong>X</strong> 
				<input type="text" name="pwid" size="4" value="<%if UBOUND(proddims)>=1 then response.write proddims(1)%>" /> <strong>X</strong> 
				<input type="text" name="phei" size="4" value="<%if UBOUND(proddims)>=2 then response.write proddims(2)%>" /></td>
				<%	else %>
			    <td align="center" colspan="2"><strong><%
						if shipType=1 OR shipType=2 OR shipType=3 OR shipType=4 OR shipType=6 OR shipType=7 then
							response.write yyShpInf
						else
							response.write "&nbsp;"
						end if %></strong></td>
				<%	end if %>
			  </tr>
			  <tr>
                <td width="25%" align="right"><%=yyLgeImg%>:</td>
                <td width="25%" align="left"><input type="text" name="pLargeImage" size="25" value="<%=replace(pLargeImage,"""","&quot;")%>" /></td>
                <td width="25%" align="right"><%
				if shipType=1 then
					response.write yyShip & ":<br />" & yyFirShi
				elseif shipType=2 OR shipType=3 OR shipType=4 OR shipType=6 OR shipType=7 then
					response.write yyPrWght & ":"
				else
					response.write "&nbsp;"
				end if
				  %></td>
                <td width="25%" align="left"><%
				if shipType=1 then
					response.write "<input type=text name='pShipping' size='15' value='"&pShipping&"' />"
				elseif shipType=2 OR shipType=3 OR shipType=4 OR shipType=6 OR shipType=7 then
					response.write "<input type=text name='pShipping' size='9' value='"&pWeight&"' />"
					' response.write " <select name=""oversize""><option value=""0"">...</option><option value=""1"""&IIfVr(oversize=1," selected","")&">"&yyOversi&" 1</option><option value=""2"""&IIfVr(oversize=2," selected","")&">"&yyOversi&" 2</option><option value=""3"""&IIfVr(oversize=3," selected","")&">"&yyOversi&" 3</option></select>"
				else
					response.write "&nbsp;"
				end if
				%></td>
			  </tr>
			  <tr>
			<%	if simpleOptions then %>
				<td colspan="2">&nbsp;</td>
			<%	else %>
                <td align="right"><%=yyNumOpt%>:</td>
                <td>
				  <select size="1" name="pNumOptions" onchange="setprodoptions();">
					<option value='0'><%=yyNone%></option>
					<% for rowcounter=1 to maxprodopts
						   response.write "<option value='"&rowcounter&"'>"&rowcounter&"</option>"
					   next %>
				  </select></td>
			<%	end if %>
				<td width="25%" align="right"><%
				if shipType=1 then
					response.write yyShip & ":<br />" & yySubShi
				else
					response.write "&nbsp;"
				end if
				  %></td>
                <td width="25%" align="left"><%
				if shipType=1 then
					response.write "<input type=text name='pShipping2' size='15' value='"&pShipping2&"' />"
				else
					response.write "&nbsp;"
				end if
				  %></td>
			  </tr>
<%	if simpleOptions then
		for index=0 to maxprodopts-1
			if index MOD 2=0 then response.write "<tr>"
			response.write "<td align=""right"">" & yyPrdOpt & " " & index+1 & ":</td><td><select size=""1"" id=""pOption" & index & """ name=""pOption" & index & """><option value=""0"">"&yyNone&"</option>"
			if IsArray(alloptions) then
				for rowcounter=0 to UBOUND(alloptions,2)
					response.write "<option value="""&alloptions(0,rowcounter)&""""
					if IsArray(prodoptions) then
						if index <= UBOUND(prodoptions,2) then
							if prodoptions(1,index)=alloptions(0,rowcounter) then response.write " selected"
						end if
					end if
					response.write ">"&alloptions(1,rowcounter)&"</option>"
				next
			end if
			response.write "</select></td>"
			if index MOD 2<>0 then response.write "</tr>" & vbCrLf
		next
		if index MOD 2=0 then
			response.write "</tr>" & vbCrLf
		else
			response.write "<td colspan=""2"">&nbsp;</td></tr>" & vbCrLf
		end if
	else %>
			</table>
			<div name="prodoptions" id="prodoptions">
			</div>
			<table width="100%" border="0" cellspacing="0" cellpadding="3" bgcolor="">
<%	end if
	if digidownloads=true then %>
			  <tr>
                <td align="right"><%=yyDownl%>:</td>
                <td colspan="3" align="left"><input type="text" size="60" name="pDownload" value="<%=pDownload%>" /></td>
			  </tr>
<%	end if %>
			  <tr> 
                <td align="right"><%=yyDesc%>:</td>
                <td colspan="2"><textarea name="pDescription" cols="55" rows="8" wrap=virtual><%=replace(pDescription&"","&","&amp;")%></textarea></td>
				<td align="center"><%=yyDrSppr%>: <select name="pDropship" size="1">
				  <option value="0"><%=yyNone%></option>
				<%	if IsArray(alldropship) then
						for index=0 to UBOUND(alldropship, 2)
							response.write "<option value="""&alldropship(0,index)&""""
							if alldropship(0,index)=pDropship then response.write " selected"
							response.write ">"&alldropship(1,index)&"</option>"&vbCrLf
						next
					end if %>
				  </select>
				<br /><br />
				<%=yyExemp%> <font size="1">&lt;Ctrl>+Click</font><br />
					<select name="pExemptions" size="3" multiple>
					<option value="1" <%if (pExemptions AND 1)=1 then response.write "selected"%>><%=yyExStat%></option>
					<option value="2" <%if (pExemptions AND 2)=2 then response.write "selected"%>><%=yyExCoun%></option>
					<option value="4" <%if (pExemptions AND 4)=4 then response.write "selected"%>><%=yyExShip%></option>
					</select><br /><img src="images/clearpixel.gif" width="20" height="3" alt="" />
<%				if perproducttaxrate=TRUE then %>
					<br /><%=yyTax%>: <input type="text" style="text-align:right" size="6" name="pTax" value="<%=pTax%>" />%
<%				end if %>
				</td>
			  </tr>
<%				for index=2 to adminlanguages+1
					if (adminlangsettings AND 2)=2 then
						if NOT doaddnew then
							sSQL = "SELECT pDescription" & index & " FROM products WHERE pId='"&Request.Form("id")&"'"
							rs2.Open sSQL,cnn,0,1
							thedescription = rs2("pDescription" & index)
							rs2.Close
						end if
					%>
			  <tr>
				<td align="right"><%=yyDesc & " " & index%>:</td>
                <td colspan="3"><textarea name="pDescription<%=index%>" cols="55" rows="8" wrap=virtual><%=replace(thedescription&"","&","&amp;")%></textarea></td>
			  </tr>
<%					end if
				next %>
			  <tr>
                <td width="25%" align="right"><%=yyLnDesc%>:</td>
                <td colspan="3" align="left"><textarea name="pLongDescription" cols="65" rows="9" wrap=virtual><%=replace(pLongDescription&"","&","&amp;")%></textarea></td>
			  </tr>
<%				for index=2 to adminlanguages+1
					if (adminlangsettings AND 4)=4 then
						if NOT doaddnew then
							sSQL = "SELECT pLongDescription" & index & " FROM products WHERE pId='"&Request.Form("id")&"'"
							rs2.Open sSQL,cnn,0,1
							thedescription = rs2("pLongDescription" & index)
							rs2.Close
						end if %>
			  <tr>
				<td align="right"><%=yyLnDesc & " " & index%>:</td>
                <td colspan="3"><textarea name="pLongDescription<%=index%>" cols="65" rows="9" wrap=virtual><%=replace(thedescription&"","&","&amp;")%></textarea></td>
			  </tr>
<%					end if
				next %>
			  <tr>
				<td align="right"><%=yyStatPg%>:</td>
                <td><input type="checkbox" name="pStaticPage" value="1"<% if int(pStaticPage) <> 0 then response.write " checked" %>></td>
				<td align="right"><%=yyRecomd%>:</td>
                <td><input type="checkbox" name="pRecommend" value="1"<% if int(pRecommend) <> 0 then response.write " checked" %>></td>
			  </tr>
			  <tr>
				<td align="right"><%=yyProdOr%>:</td>
                <td colspan="3"><input type="text" name="pOrder" size="10" value="<%=pOrder%>"></td>
			  </tr>
			  <tr>
				<td width="25%" align="right"><strong><%=yyAddSec%>:</strong></td>
                <td colspan="4" align="left">
				<%	if NOT simpleSections then %>
				  <select size="1" name="pNumSections" onchange="setprodsections();">
					<option value='0'><%=yyNone%></option>
					<% for rowcounter=1 to maxprodsects
						   response.write "<option value='"&rowcounter&"'>"&rowcounter&"</option>"
					   next %>
				  </select>
				<%	end if %>&nbsp;</td>
			  </tr>
<%	if simpleSections then
		for index=0 to maxprodsects-1
			if index MOD 2=0 then response.write "<tr>"
			response.write "<td align=""right"">" & yyPrdSec & " " & index+1 & ":</td><td><select size=""1"" id=""pSection" & index & """ name=""pSection" & index & """><option value=""0"">"&yyNone&"</option>"
			if IsArray(allsections) then
				for rowcounter=0 to UBOUND(allsections,2)
					response.write "<option value="""&allsections(0,rowcounter)&""""
					if IsArray(prodsections) then
						if index <= UBOUND(prodsections,2) then
							if prodsections(0,index)=allsections(0,rowcounter) then response.write " selected"
						end if
					end if
					response.write ">"&allsections(1,rowcounter)&"</option>"
				next
			end if
			response.write "</select></td>"
			if index MOD 2<>0 then response.write "</tr>" & vbCrLf
		next
		if index MOD 2=0 then
			response.write "</tr>" & vbCrLf
		else
			response.write "<td colspan=""2"">&nbsp;</td></tr>" & vbCrLf
		end if
	else %>
			</table>
			<div name="prodsections" id="prodsections">
			</div>
			<table width="100%" border="0" cellspacing="0" cellpadding="3" bgcolor="">
<%	end if %>
			  <tr> 
                <td width="100%" colspan="4">
                  <p align="center"><input type="submit" value="<%=yySubmit%>" />&nbsp;&nbsp;<input type="reset" value="<%=yyReset%>" /></p>
<%	show_info() %>
                </td>
			  </tr>
            </table>
		  </td>
        </tr>
      </table>
	</form>
<% if NOT doaddnew then %>
<script language="javascript" type="text/javascript">
<!--
<%	if NOT simpleOptions then %>
document.forms.mainform.pNumOptions.selectedIndex=<% if IsArray(prodoptions) then response.write (UBOUND(prodoptions,2)+1) else response.write "0" %>;
document.forms.mainform.pNumOptions.options[<% if IsArray(prodoptions) then response.write (UBOUND(prodoptions,2)+1) else response.write "0" %>].selected = true;
setprodoptions();
<%	end if
	if NOT simpleSections then %>
document.forms.mainform.pNumSections.selectedIndex=<% if IsArray(prodsections) then response.write (UBOUND(prodsections,2)+1) else response.write "0" %>;
document.forms.mainform.pNumSections.options[<% if IsArray(prodsections) then response.write (UBOUND(prodsections,2)+1) else response.write "0" %>].selected = true;
setprodsections();
<%	end if
	if useStockManagement then %>
setstocktype();
<%	end if %>
//-->
</script>
<% end if
   elseif request.form("act")="discounts" then 
		sSQL = "SELECT pName FROM products WHERE pID='"&replace(request.form("id"),"'","''")&"'"
		rs.Open sSQL,cnn,0,1
		thisname=rs("pName")
		rs.Close
		alldata=""
		sSQL = "SELECT cpaID,cpaCpnID,cpnWorkingName,cpnSitewide,cpnEndDate,cpnType FROM cpnassign INNER JOIN coupons ON cpnassign.cpaCpnID=coupons.cpnID WHERE cpaType=2 AND cpaAssignment='" & request.form("id") & "'"
		rs.Open sSQL,cnn,0,1
		if NOT rs.EOF then alldata=rs.GetRows
		rs.Close
		alldata2=""
		tdt = Date()
		sSQL = "SELECT cpnID,cpnWorkingName,cpnSitewide FROM coupons WHERE cpnSitewide=0 AND cpnEndDate >="&datedelim&VSUSDate(tdt)&datedelim
		rs.Open sSQL,cnn,0,1
		if NOT rs.EOF then alldata2=rs.GetRows
		rs.Close
%>
<script language="javascript" type="text/javascript">
<!--
function drec(id){
cmsg = "<%=yyConAss%>\n"
if (confirm(cmsg)){
	document.mainform.id.value = id;
	document.mainform.act.value = "deletedisc";
	document.mainform.submit();
}
}
// -->
</script>
        <tr>
		<form name="mainform" method="post" action="adminprods.asp">
		  <td width="100%">
			<input type="hidden" name="posted" value="1" />
			<input type="hidden" name="act" value="dodiscounts" />
			<input type="hidden" name="id" value="<%=request.form("id")%>" />
<%				call writehiddenvar("stock", request.form("stock"))
				call writehiddenvar("stext", request.form("stext"))
				call writehiddenvar("sprice", request.form("sprice"))
				call writehiddenvar("scat", request.form("scat"))
				call writehiddenvar("stype", request.form("stype"))
				call writehiddenvar("pg", request.form("pg")) %>
            <table width="100%" border="0" cellspacing="0" cellpadding="3" bgcolor="">
			  <tr> 
                <td width="100%" colspan="4" align="center"><strong><%=yyAssPrd%> &quot;<%=thisname%>&quot;.</strong><br />&nbsp;</td>
			  </tr>
<%	gotone=false
	if IsArray(alldata2) then
		thestr = "<tr><td colspan='4' align='center'>"&yyAsDsCp&": <select name='assdisc' size='1'>"
		for index=0 to UBOUND(alldata2,2)
			alreadyassign=false
			if IsArray(alldata) then
				for index2=0 to UBOUND(alldata,2)
					if alldata2(0,index)=alldata(1,index2) then alreadyassign=true
				next
			end if
			if NOT alreadyassign then
				thestr = thestr & "<option value='"&alldata2(0,index)&"'>"&alldata2(1,index)&"</option>" & vbCrLf
				gotone=true
			end if
		next
		thestr = thestr & "</select> <input type='submit' value='Go' /></td></tr>"
	end if
	if gotone then
		response.write thestr
	else
%>
			  <tr> 
                <td width="100%" colspan="4" align="center"><br /><strong><%=yyNoDis%></td>
			  </tr>
<%
	end if
	if IsArray(alldata) then
%>
			  <tr> 
                <td width="100%" colspan="4" align="center"><br /><strong><%=yyCurDis%> &quot;<%=thisname%>&quot;.</strong><br />&nbsp;</td>
			  </tr>
			  <tr> 
                <td><strong><%=yyWrkNam%></strong></td>
				<td><strong><%=yyDisTyp%></strong></td>
				<td><strong><%=yyExpire%></strong></td>
				<td align="center"><strong><%=yyDelete%></strong></td>
			  </tr>
<%
		for index=0 to UBOUND(alldata,2)
			prefont = ""
			postfont = ""
			if alldata(3,index)=1 OR alldata(4,index)-Date() < 0 then
				prefont = "<font color=""#FF0000"">"
				postfont = "</font>"
			end if
%>
			  <tr> 
                <td><%=prefont & alldata(2,index) & postfont %></td>
				<td><%	if alldata(5,index)=0 then
							response.write prefont & yyFrSShp & postfont
						elseif alldata(5,index)=1 then
							response.write prefont & yyFlatDs & postfont
						elseif alldata(5,index)=2 then
							response.write prefont & yyPerDis & postfont
						end if %></td>
				<td><%	if alldata(4,index)=DateSerial(3000,1,1) then
							response.write yyNever
						elseif alldata(4,index)-Date() < 0 then
							response.write "<font color='#FF0000'>"&yyExpird&"</font>"
						else
							response.write prefont & alldata(4,index) & postfont
						end if %></td>
				<td align="center"><input type="button" value="Delete Assignment" onclick="drec('<%=alldata(0,index)%>')" /></td>
			  </tr>
<%
		next
	else
%>
			  <tr> 
                <td width="100%" colspan="4" align="center"><br /><strong><%=yyNoAss%></td>
			  </tr>
<%
	end if
%>
			  <tr>
                <td width="100%" colspan="4" align="center"><br />&nbsp;</td>
			  </tr>
			  <tr> 
                <td width="100%" colspan="4" align="center"><br />
                          <a href="admin.asp"><strong><%=yyAdmHom%></strong></a><br />
                          &nbsp;</td>
			  </tr>
            </table></td>
		  </form>
        </tr>
<% elseif request.form("posted")="1" AND success then %>
      <table border="0" cellspacing="0" cellpadding="0" width="100%" bgcolor="" align="center">
        <tr>
          <td width="100%">
			<table width="100%" border="0" cellspacing="0" cellpadding="3" bgcolor="">
			  <tr> 
                <td width="100%" colspan="2" align="center"><br /><strong><%=yyUpdSuc%></strong><br /><br /><%=yyNowFrd%><br /><br />
                        <%=yyNoAuto%> <a href="adminprods.asp<%
							response.write "?rid="&request.form("rid")&"&stock="&request.form("stock")&"&stext=" & server.urlencode(request.form("stext")) & "&sprice=" & server.urlencode(request.form("sprice")) & "&stype=" & request.form("stype") & "&scat=" & request.form("scat") & "&pg=" & request.form("pg")
						%>"><strong><%=yyClkHer%></strong></a>.<br />
                        <br />
				<img src="../images/clearpixel.gif" width="300" height="3" alt="" />
                </td>
			  </tr>
			</table></td>
        </tr>
      </table>
<% elseif request.form("posted")="1" then %>
      <table border="0" cellspacing="0" cellpadding="0" width="100%" bgcolor="" align="center">
        <tr>
          <td width="100%">
			<table width="100%" border="0" cellspacing="0" cellpadding="3" bgcolor="">
			  <tr> 
                <td width="100%" colspan="2" align="center"><br /><font color="#FF0000"><strong><%=yyOpFai%></strong></font><br /><br /><%=errmsg%><br /><br />
				<a href="javascript:history.go(-1)"><strong><%=yyClkBac%></strong></a></td>
			  </tr>
			</table></td>
        </tr>
      </table>
<% else %>
<script language="javascript" type="text/javascript">
<!--
function mrec(id,evt){
	document.mainform.action="adminprods.asp";
	document.mainform.id.value = id;
<% if (instr(Request.ServerVariables("HTTP_USER_AGENT"), "Gecko") > 0) then %>
	if(evt.ctrlKey || evt.altKey)
<% else %>
	theevnt=window.event;
	if(theevnt.ctrlKey)
<% end if %>
		document.mainform.act.value = "clone";
	else
		document.mainform.act.value = "modify";
	document.mainform.posted.value = "1";
	document.mainform.submit();
}
function rrec(id){
	document.mainform.action="adminprods.asp?related=go";
	document.mainform.rid.value = id;
	document.mainform.act.value = "search";
	document.mainform.posted.value = "";
	document.mainform.submit();
}
function updaterelations(){
	document.mainform.action="adminprods.asp";
	document.mainform.act.value = "updaterelations";
	document.mainform.posted.value = "1";
	document.mainform.submit();
}
function newrec(id){
	document.mainform.action="adminprods.asp";
	document.mainform.id.value = id;
	document.mainform.act.value = "addnew";
	document.mainform.posted.value = "1";
	document.mainform.submit();
}
function dscnts(id){
	document.mainform.action="adminprods.asp";
	document.mainform.id.value = id;
	document.mainform.act.value = "discounts";
	document.mainform.posted.value = "1";
	document.mainform.submit();
}
function startsearch(){
	document.mainform.action="adminprods.asp";
	document.mainform.act.value = "search";
	document.mainform.stock.value = "";
	document.mainform.posted.value = "";
	document.mainform.submit();
}
function searchoutstock(){
	document.mainform.action="adminprods.asp";
	document.mainform.act.value = "search";
	document.mainform.stock.value = "1";
	document.mainform.posted.value = "";
	document.mainform.submit();
}
function inventorymenu(){
	themenuitem=document.mainform.inventoryselect.options[document.mainform.inventoryselect.selectedIndex].value;
	if(themenuitem=="1") document.mainform.act.value = "stockinventory";
	if(themenuitem=="2") document.mainform.act.value = "fullinventory";
	if(themenuitem=="3") document.mainform.act.value = "dump2COinventory";
	document.mainform.action="dumporders.asp";
	document.mainform.submit();
}
function drec(id){
cmsg = "<%=yyConDel%>\n"
if (confirm(cmsg)){
	document.mainform.action="adminprods.asp";
	document.mainform.id.value = id;
	document.mainform.act.value = "delete";
	document.mainform.posted.value = "1";
	document.mainform.submit();
}
}
// -->
</script>
      <table border="0" cellspacing="0" cellpadding="0" width="100%" bgcolor="" align="center">
<%	rid = trim(request("rid"))
	ridarr = ""
	if rid<>"" then
		sSQL = "SELECT rpRelProdID FROM relatedprods WHERE rpProdID='" & replace(rid,"'","''") & "'"
		rs.Open sSQL,cnn,0,1
		if NOT rs.eof then ridarr = rs.getrows
		rs.Close
	end if
	if request.querystring("related")="go" then session("savesearch")= "stock="&request.form("stock")&"&stext=" & server.urlencode(request.form("stext")) & "&sprice=" & server.urlencode(request.form("sprice")) & "&stype=" & request.form("stype") & "&scat=" & request.form("scat") & "&pg=" & request.form("pg")
%>
        <tr>
		<form name="mainform" method="post" action="adminprods.asp">
		  <td width="100%">
			<input type="hidden" name="posted" value="1" />
			<input type="hidden" name="act" value="xxxxx" />
			<input type="hidden" name="stock" value="" />
			<input type="hidden" name="id" value="xxxxx" />
			<input type="hidden" name="rid" value="<%=rid%>" />
			<input type="hidden" name="pg" value="<%=IIfVr(request.form("act")="search", "1", request.querystring("pg"))%>" />
<%		thecat = request("scat")
		if thecat<>"" then thecat = int(thecat)
		sSQL = "SELECT sectionID,sectionWorkingName,topSection,rootSection FROM sections " & IIfVr(adminonlysubcats=true, "WHERE rootSection=1 ORDER BY sectionWorkingName", "ORDER BY sectionOrder")
		rs.Open sSQL,cnn,0,1
		if rs.eof then
			success=false
		else
			alldata=rs.getrows
			success=true
		end if
		rs.Close
		sSQL = "SELECT payProvEnabled,payProvData1 FROM payprovider WHERE payProvID=2"
		rs.Open sSQL,cnn,0,1
		if rs("payProvEnabled")=1 AND trim(rs("payProvData1")&"")<>"" then twocoinventory=TRUE else twocoinventory=FALSE
		rs.Close
%>				<table class="cobtbl" width="100%" border="0" bordercolor="#B1B1B1" cellspacing="1" cellpadding="3" bgcolor="#B1B1B1">
<%			if rid<>"" then %>
				  <tr><td class="cobhl" align="center" colspan="4" height="22"><strong> Products related to <%=rid %></strong> </td></tr>
<%			end if %>
				  <tr> 
	                <td class="cobhl" width="25%" align="right" bgcolor="#EBEBEB"><%=yySrchFr%>:</td>
					<td class="cobll" width="25%" bgcolor="#FFFFFF"><input type="text" name="stext" size="20" value="<%=request("stext")%>" /></td>
					<td class="cobhl" width="25%" align="right" bgcolor="#EBEBEB"><%=yySrchMx%>:</td>
					<td class="cobll" width="25%" bgcolor="#FFFFFF"><input type="text" name="sprice" size="10" value="<%=request("sprice")%>" /></td>
				  </tr>
				  <tr>
				    <td class="cobhl" width="25%" align="right" bgcolor="#EBEBEB"><%=yySrchTp%>:</td>
					<td class="cobll" width="25%" bgcolor="#FFFFFF"><select name="stype" size="1">
						<option value=""><%=yySrchAl%></option>
						<option value="any" <% if request("stype")="any" then response.write "selected"%>><%=yySrchAn%></option>
						<option value="exact" <% if request("stype")="exact" then response.write "selected"%>><%=yySrchEx%></option>
						</select>
					</td>
					<td class="cobhl" width="25%" align="right" bgcolor="#EBEBEB"><%=yySrchCt%>:</td>
					<td class="cobll" width="25%" bgcolor="#FFFFFF">
					  <select name="scat" size="1">
					  <option value=""><%=yySrchAC%></option>
						<%	if IsArray(alldata) then
								if adminonlysubcats=true then
									for rowcounter=0 to UBOUND(alldata,2)
										response.write "<option value='"&alldata(0,rowcounter)&"'"
										if alldata(0,rowcounter)=thecat then response.write " selected"
										response.write ">"&alldata(1,rowcounter)&"</option>" &vbCrLf
									next
								else
									call writemenulevel(0,1)
								end if
							end if %>
					  </select>
					</td>
	              </tr>
				  <tr>
				    <td class="cobhl" bgcolor="#EBEBEB">&nbsp;</td>
				    <td class="cobll" bgcolor="#FFFFFF" colspan="3"><table width="100%" cellspacing="0" cellpadding="0" border="0">
					    <tr>
						  <td class="cobll" bgcolor="#FFFFFF" align="center"><input type="button" value="<%=yyListPd%>" onclick="startsearch();" /> 
							<% if useStockManagement then response.write "<input type=""button"" value="""&yyOOStoc&""" onclick=""searchoutstock();"" />" %>
<%					if rid<>"" then %>
							<strong>&raquo;</strong> <input type="button" value="<%=yyBckLis%>" onclick="document.location='adminprods.asp?<%=session("savesearch")%>'">
<%					else %>
							<input type="button" value="<%=yyNewPr%>" onclick="newrec();" />
<%					end if %>
						  </td>
						  <td class="cobll" bgcolor="#FFFFFF" height="26" width="20%" align="right">
<%					if rid<>"" then %>
							<input type="button" value="<%=yyUpdRel%>" onclick="updaterelations()">
<%					else %>
						<select name="inventoryselect" size="1">
							<% if stockManage<>0 then response.write "<option value=""1"">"&yyStkInv&"</option>" %>
							<option value="2"><%=yyFulInv%></option>
							<% if twocoinventory then response.write "<option value=""3"">2Checkout Inventory</option>" %>
						</select>&nbsp;<input type="button" value="Go" onclick="javascript:inventorymenu();" />
<%					end if %></td>
						</tr>
					  </table></td>
				  </tr>
				</table>
            <table width="100%" border="0" cellspacing="0" cellpadding="1" bgcolor="">
<%	if request.form("act")="search" OR request.querystring("pg")<>"" then
		sub displayprodrow(xrs)
			%><tr bgcolor="<%=bgcolor%>"><td><%=xrs("pID")%></td><td><%
					if IsNull(xrs("rootSection")) then
						response.write "<font color='#FF0000'>*</font> "
						haveerrprods=true
					elseif cint(xrs("rootSection"))<>1 then
						response.write "<font color='#FF0000'>*</font> "
						haveerrprods=true
					end if
					stockbyoptions=false
					if stockManage<>0 then
						if cint(xrs("pStockByOpts"))<>0 then stockbyoptions=true
					end if
					hasstock = TRUE
					if cint(xrs("pDisplay")) = 0 OR ((stockManage<>0 AND xrs("pInStock") <= 0 AND NOT stockbyoptions) OR (stockManage=0 AND cint(xrs("pSell"))=0)) then hasstock=FALSE
					if NOT hasstock then response.write "<font color='#FF0000'>"
					if cint(xrs("pDisplay")) = 0 then response.write "<strike>"
					response.write xrs("pName")
					if cint(xrs("pDisplay")) = 0 then response.write "</strike>"
					if NOT hasstock then response.write "</font>"
					if stockManage>0 then
						if stockbyoptions then response.write " (-)" else response.write " (" & xrs("pInStock") & ")"
					end if %></td><td><input <%
				if IsArray(allcoupon) then
					for index=0 to UBOUND(allcoupon,2)
						if Trim(allcoupon(0,index))=xrs("pID") then
							response.write "style=""color: #FF0000"" "
							exit for
						end if
					next
				end if
			%>type="button" value="<%=yyAssign%>" onclick="dscnts('<%=replace(replace(xrs("pID"),"\","\\"),"'","\'")%>')" /></td><td><input type="button" value="<%=yyModify%>" onclick="mrec('<%=replace(replace(xrs("pID"),"\","\\"),"'","\'")%>',event)" /></td><%
				if rid<>"" then
			%><td align="center"><input type="hidden" name="updq<%=replace(xrs("pID"),"""","&quot;")%>" value="1"><input type="checkbox" name="updr<%=replace(xrs("pID"),"""","&quot;")%>" value="1" <%
					if rid=xrs("pID") then
						response.write "disabled "
					else
						if IsArray(ridarr) then
							for index=0 to UBOUND(ridarr,2)
								if ridarr(0,index)=xrs("pID") then response.write "checked " : exit for
							next
						end if
					end if %>/></td><%
				else
			%><td><input type="button" id="rrec<%=replace(xrs("pID"),"""","&quot;")%>" value="<%=yyRelate%>" onclick="rrec('<%=replace(replace(xrs("pID"),"\","\\"),"'","\'")%>')" /></td><%
				end if
			%><td><input type="button" value="<%=yyDelete%>" onclick="drec('<%=replace(replace(xrs("pID"),"\","\\"),"'","\'")%>')" /></td></tr><%
			response.write vbCrLf
		end sub
		sub displayheaderrow() %>
			<tr>
				<td><strong><%=yyPrId%></strong></td>
				<td><strong><%=yyPrName%></strong></td>
				<td width="5%" align="center"><font size="1"><strong><%=yyDiscnt%></strong></font></td>
				<td width="5%" align="center"><font size="1"><strong><%=yyModify%></strong></font></td>
				<td width="5%" align="center"><font size="1"><strong><%=yyRelate%></strong></font></td>
				<td width="5%" align="center"><font size="1"><strong><%=yyDelete%></strong></font></td>
			</tr>
<%		end sub
		allcoupon="" : pidlist=""
		sSQL = "SELECT DISTINCT cpaAssignment FROM cpnassign WHERE cpaType=2"
		rs.Open sSQL,cnn,0,1
		if NOT rs.EOF then allcoupon=rs.getrows
		rs.Close
		if request.querystring("related")="go" then
			if mysqlserver=TRUE then
				sSQL = "SELECT DISTINCT products.pID,pName,pDisplay,pSell,pInStock,rootSection,pStockByOpts FROM relatedprods INNER JOIN products ON products.pId=relatedprods.rpRelProdId LEFT OUTER JOIN sections ON products.pSection=sections.sectionID WHERE rpProdId='"&replace(rid,"'","''")&"'"
			else
				sSQL = "SELECT DISTINCT products.pID,pName,pDisplay,pSell,pInStock,rootSection,pStockByOpts FROM relatedprods INNER JOIN (products LEFT OUTER JOIN sections ON products.pSection=sections.sectionID) ON products.pId=relatedprods.rpRelProdId WHERE rpProdId='"&replace(rid,"'","''")&"'"
			end if
			rs.Open sSQL,cnn,0,1
			if NOT rs.EOF then
				displayheaderrow()
				do while NOT rs.EOF
					if bgcolor="#E7EAEF" then bgcolor="#FFFFFF" else bgcolor="#E7EAEF"
					displayprodrow(rs)
					rs.MoveNext
					Count = Count + 1
				loop
			else
				response.write "<tr><td width=""100%"" colspan=""6"" align=""center""><p>&nbsp;</p><p>"&yyPrNoRe&"</p><p>"&yyPrReSe&"</p><p>"&yyPrReLs&"</p>&nbsp;</td></tr>"
			end if
			rs.Close
		else
			whereand=" WHERE "
			if mysqlserver=true then
				sSQL = "SELECT DISTINCT products.pID,pName,pDisplay,pSell,pInStock,rootSection,pStockByOpts FROM multisections RIGHT JOIN products ON products.pId=multisections.pId LEFT OUTER JOIN sections ON products.pSection=sections.sectionID"
			else
				sSQL = "SELECT DISTINCT products.pID,pName,pDisplay,pSell,pInStock,rootSection,pStockByOpts FROM multisections RIGHT JOIN (products LEFT OUTER JOIN sections ON products.pSection=sections.sectionID) ON products.pId=multisections.pId"
			end if
			if thecat<>"" then
				sectionids = getsectionids(thecat, TRUE)
				if sectionids<>"" then sSQL = sSQL & whereand & "(products.pSection IN (" & sectionids & ") OR multisections.pSection IN (" & sectionids & "))"
				whereand=" AND "
			end if
			sprice = trim(request("sprice"))
			if sprice<>"" then
				if instr(sprice, "-") > 0 then
					pricearr=split(sprice, "-")
					if NOT isnumeric(pricearr(0)) then pricearr(0)=0
					if NOT isnumeric(pricearr(1)) then pricearr(1)=10000000
					sSQL = sSQL & whereand & "pPrice BETWEEN "&cDbl(replace(pricearr(0),"$",""))&" AND "&cDbl(replace(pricearr(1),"$",""))
					whereand=" AND "
				elseif IsNumeric(sprice) then
					sSQL = sSQL & whereand & "pPrice="&cDbl(replace(sprice,"$",""))&" "
					whereand=" AND "
				end if
			end if
			if Trim(request("stext"))<>"" then
				sText = replace(Trim(request("stext")),"'","''")
				aText = Split(sText)
				aFields(0)="products.pId"
				aFields(1)=getlangid("pName",1)
				aFields(2)=getlangid("pDescription",2)
				if request("stype")="exact" then
					sSQL=sSQL & whereand & "(products.pId LIKE '%"&sText&"%' OR "&getlangid("pName",1)&" LIKE '%"&sText&"%' OR "&getlangid("pDescription",2)&" LIKE '%"&sText&"%' OR "&getlangid("pLongDescription",2)&" LIKE '%"&sText&"%') "
					whereand=" AND "
				else
					sJoin="AND "
					if request("stype")="any" then sJoin="OR "
					sSQL=sSQL & whereand&"("
					whereand=" AND "
					for index=0 to 2
						sSQL=sSQL & "("
						for rowcounter=0 to UBOUND(aText)
							sSQL=sSQL & aFields(index) & " LIKE '%"&aText(rowcounter)&"%' "
							if rowcounter<UBOUND(aText) then sSQL=sSQL & sJoin
						next
						sSQL=sSQL & ") "
						if index < 2 then sSQL=sSQL & "OR "
					next
					sSQL=sSQL & ") "
				end if
			end if
			if request("stock")="1" then sSQL = sSQL & whereand & "(pInStock<=0 AND pStockByOpts=0)"
			if adminsortorder<>"" then
				if instr(lcase(adminsortorder), "pid") > 0 AND instr(lcase(adminsortorder), "products.pid") = 0 then adminsortorder = replace(lcase(adminsortorder), "pid", "products.pid")
				sSQL = sSQL & " ORDER BY " & adminsortorder
			else
				sSQL = sSQL & " ORDER BY pName"
			end if
			if adminproductsperpage="" then adminproductsperpage=200
			rs.CursorLocation = 3 ' adUseClient
			rs.CacheSize = adminproductsperpage
			rs.Open sSQL, cnn
			if rs.eof or rs.bof then
				success=false
				iNumOfPages=0
			else
				success=true
				rs.MoveFirst
				rs.PageSize = adminproductsperpage
				If Request.QueryString("pg") = "" Then
					CurPage = 1
				Else
					CurPage = Int(Request.QueryString("pg"))
				End If
				iNumOfPages = Int((rs.RecordCount + (adminproductsperpage-1)) / adminproductsperpage)
				rs.AbsolutePage = CurPage
			end if
			Count = 0
			haveerrprods=FALSE
			if NOT rs.EOF then
				If iNumOfPages > 1 Then Response.Write "<tr><td colspan=""6"" align=""center"">" & writepagebar(CurPage, iNumOfPages) & "</td></tr>"
				displayheaderrow()
				addcomma=""
				do while NOT rs.EOF And Count < rs.PageSize
					if bgcolor="#E7EAEF" then bgcolor="#FFFFFF" else bgcolor="#E7EAEF"
					displayprodrow(rs)
					pidlist=pidlist&addcomma&"'"&rs("pID")&"'"
					addcomma=","
					rs.MoveNext
					Count = Count + 1
				loop
				if haveerrprods then response.write "<tr><td width=""100%"" colspan=""6""><br /><strong><font color='#FF0000'>* </font></strong>"&yySeePr&"</td></tr>"
				If iNumOfPages > 1 Then Response.Write "<tr><td colspan=""6"" align=""center"">" & writepagebar(CurPage, iNumOfPages) & "</td></tr>"
			else
				response.write "<tr><td width=""100%"" colspan=""6"" align=""center""><br />"&yyPrNone&"<br />&nbsp;</td></tr>"
			end if
			rs.Close
		end if
		if pidlist<>"" AND rid="" then
			response.write vbCrLf & "<script language=""javascript"" type=""text/javascript"">function setcl(tid){document.getElementById('rrec'+tid).style.color='#FF0000';}" & vbCrLf
			rs.Open "SELECT DISTINCT rpProdId FROM relatedprods WHERE rpProdId IN ("&pidlist&")",cnn,0,1
			do while NOT rs.EOF
				response.write "setcl('"&rs("rpProdId")&"');" & vbCrLf
				rs.MoveNext
			loop
			rs.Close
			response.write "</script>"
		end if
	end if %>
			  <tr> 
                <td width="100%" colspan="6" align="center"><br />
                          <a href="admin.asp"><strong><%=yyAdmHom%></strong></a><br />
				<img src="../images/clearpixel.gif" width="300" height="3" alt="" /></td>
			  </tr>
            </table></td>
		  </form>
        </tr>
      </table>
<% end if
cnn.Close
set rs = nothing
set cnn = nothing
%>
