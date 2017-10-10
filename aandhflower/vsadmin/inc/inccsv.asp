<%
if storesessionvalue="" then storesessionvalue="virtualstore"
if Session("loggedon") <> storesessionvalue OR disallowlogin=TRUE then response.end
addsuccess = TRUE
success = TRUE
successlines = 0
faillines = 0
pidnotfoundlines = 0
stoppedonerror = FALSE
showaccount = TRUE
dorefresh = FALSE
isstockupdate=FALSE
CrLf = Chr(13) & Chr(10)
csvcurrpos = 1
csvlen = 0
Server.ScriptTimeout = 180
Set rs = Server.CreateObject("ADODB.RecordSet")
Set cnn=Server.CreateObject("ADODB.Connection")
cnn.open sDSN
sSQL = "SELECT adminEmail,emailObject,smtpserver,emailUser,emailPass,adminEmailConfirm,adminTweaks,adminProdsPerPage,adminStoreURL,adminHandling,adminPacking,adminDelUncompleted,adminDelCC,adminUSZones,adminStockManage,adminShipping,adminIntShipping,adminCanPostUser,adminZipCode,adminUnits,adminUSPSUser,adminUSPSpw,adminUPSUser,adminUPSpw,adminUPSAccess,FedexAccountNo,FedexMeter,adminlanguages,adminlangsettings,currRate1,currSymbol1,currRate2,currSymbol2,currRate3,currSymbol3,currConvUser,currConvPw,currLastUpdate,countryLCID,countryCurrency,countryName,countryCode,countryTax FROM admin INNER JOIN countries ON admin.adminCountry=countries.countryID WHERE adminID=1"
rs.Open sSQL,cnn,0,1
	splitUSZones = (Int(rs("adminUSZones"))=1)
rs.Close

	'**************************************
    ' Name: ANSI to Unicode
    ' Description:Converts from ANSI to Unic
    '     ode very fast. Inspired by code found in
    '     UltraFastAspUpload by Cakkie (on PSC). T
    '     his should work slightly faster then Cak
    '     kies due to how some of the code has bee
    '     n arranged.
    ' By: Lewis E. Moten III
    '
    ' This code is copyrighted and has
    ' limited warranties. Please see http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=7266&lngWId=4
	' for details.
    '**************************************
	function ANSIToUnicode(ByRef pbinBinaryData)
    	Dim lbinData	' Binary Data (ANSI)
    	Dim llngLength	' Length of binary data (byte count)
    	Dim lobjRs		' Recordset
    	Dim lstrData	' Unicode Data
    	' VarType Reference
    	'8 = Integer (this is expected var type)
    	'17 = Byte Subtype
    	' 8192 = Array
    	' 8209 = Byte Subtype + Array
    	Set lobjRs = Server.CreateObject("ADODB.Recordset")
    	if VarType(pbinBinaryData) = 8 Then
    		' Convert integers(4 bytes) To Byte Subtype Array (1 byte)
    		llngLength = LenB(pbinBinaryData)
    		if llngLength = 0 Then
    			lbinData = ChrB(0)
    		Else
    			Call lobjRs.Fields.Append("BinaryData", adLongVarBinary, llngLength)
    			Call lobjRs.Open()
    			Call lobjRs.AddNew()
    			Call lobjRs.Fields("BinaryData").AppendChunk(pbinBinaryData & ChrB(0)) ' + Null terminator
    			Call lobjRs.Update()
    			lbinData = lobjRs.Fields("BinaryData").GetChunk(llngLength)
    			Call lobjRs.Close()
    		End if
    	Else
    		lbinData = pbinBinaryData
    	End if
    	' Do REAL conversion now!	
    	llngLength = LenB(lbinData)
    	if llngLength = 0 Then
    		lstrData = ""
    	Else
    		Call lobjRs.Fields.Append("BinaryData", 201, llngLength)
    		Call lobjRs.Open()
    		Call lobjRs.AddNew()
    		Call lobjRs.Fields("BinaryData").AppendChunk(lbinData)
    		Call lobjRs.Update()
    		lstrData = lobjRs.Fields("BinaryData").Value
    		Call lobjRs.Close()
    	End if
    				
    	Set lobjRs = Nothing
    	ANSIToUnicode = lstrData
    End function
function getcsvline()
	getcsvline=""
	do while csvcurrpos <= csvlen
		tmpchar=mid(csvfile, csvcurrpos, 1)
		csvcurrpos=csvcurrpos+1
		if tmpchar=vbCr OR tmpchar=vbLf then exit do else getcsvline=getcsvline&tmpchar
	loop
	do while csvcurrpos <= csvlen
		tmpchar=mid(csvfile, csvcurrpos, 1)
		if tmpchar=vbCr OR tmpchar=vbLf then csvcurrpos=csvcurrpos+1 else exit do
	loop
end function
function GetFieldName(infoStr)
	sPos = InStr(infoStr, "name=")
	endPos = InStr(sPos + 6, infoStr, Chr(34) & ";")
	if endPos = 0 then
		endPos = inStr(sPos + 6, infoStr, Chr(34))
	end if
	GetFieldName = mid(infoStr, sPos + 6, endPos - (sPos + 6))
end function
'This function retreives a file field's filename
function GetFileName(infoStr)
	sPos = InStr(infoStr, "filename=")
	endPos = InStr(infoStr, Chr(34) & CrLf)
	GetFileName = mid(infoStr, sPos + 10, endPos - (sPos + 10))
end function
'This function retreives a file field's MIME type
function GetFileType(infoStr)
	sPos = InStr(infoStr, "Content-Type: ")
	GetFileType = mid(infoStr, sPos + 14)
end function
biData = Request.BinaryRead(Request.TotalBytes)
bidatalen = LenB(biData)
isposted = (bidatalen>0)
PostData = ""
if isposted then
	PostData = ANSIToUnicode(biData)
	ContentType = Request.ServerVariables("HTTP_CONTENT_TYPE")
	ctArray = Split(ContentType, ";")
	if Trim(ctArray(0)) = "multipart/form-data" then
		ErrMsg = ""
		' grab the form boundary...
		bArray = Split(Trim(ctArray(1)), "=")
		Boundary = Trim(bArray(1))
		'Now use that to split up all the variables!
		formData = Split(PostData, Boundary)
		'Extract the information for each variable and its data
		FileCount = 0
		for x = 0 to UBound(formData)
			'Two CrLfs mark the end of the information about this field; everything after that is the value
			Infoend = InStr(formData(x), CrLf & CrLf)
			if Infoend > 0 then
				'Get info for this field, minus stuff at the end
				varInfo = mid(formData(x), 3, Infoend - 3)
				'Get value for this field, being sure to skip CrLf pairs at the start and the CrLf at the end
				if (InStr(varInfo, "filename=") > 0) then ' Is this a file?
					if GetFieldName(varInfo)="csvfile" then
						csvfile = mid(formData(x), Infoend + 4, Len(formData(x)) - Infoend - 7) & vbCrLf ' add a "known elephant"
						csvlen = len(csvfile)
					end if
					' GetFileName(varInfo) : GetFileType(varInfo)
					FileCount = FileCount + 1
				else ' It's a regular field
					varValue = mid(formData(x), Infoend + 4, Len(formData(x)) - Infoend - 7)
					fieldname = GetFieldName(varInfo)
					select case fieldname
					case "show_errors"
						show_errors = (varValue="ON")
					case "stop_errors"
						stop_errors = (varValue="ON")
					case "theaction"
						isupdate = (varValue="update")
					end select
				end if
			end if
		next
	else
		ErrMsg = "Wrong encoding type!"
	end if
end if
progressevery=500
function csv_database_error()
	if show_errors then response.write "Line " & line_num & ", " & mysql_error & "<br />"
	csvsuccess=FALSE
	faillines=faillines+1
	successlines=successlines-1
end function
function execute_sql()
	if isstockupdate then
		' on error resume next
		if trim(valuesarray(4))<>"" then
			sSQL = "UPDATE options SET optStock=" & valuesarray(3) & " WHERE optID=" & valuesarray(4)
		else
			sSQL = "UPDATE products SET pInStock=" & valuesarray(3) & " WHERE pID='" & trim(valuesarray(0)) & "'"
		end if
		err.number = 0
		cnn.Execute(sSQL)
		on error goto 0
	elseif isupdate then
		if mysqlserver then rs.CursorLocation = 3
		rs.Open "SELECT * FROM products WHERE pID='" & replace(valuesarray(keycolumn), "'", "''") & "'",cnn,1,3,&H0001
		if rs.EOF then
			pidnotfoundlines=pidnotfoundlines+1
			successlines=successlines-1
		else
			for i=0 to columncount-1
				if i <> keycolumn then
					' sSQL = sSQL & addcomma & columnarray(i) & "='" & replace(valuesarray(i), "'", "''") & "'"
					if (rs.Fields(columnarray(i)).Type=3 OR rs.Fields(columnarray(i)).Type=5 OR rs.Fields(columnarray(i)).Type=11 OR rs.Fields(columnarray(i)).Type=17) AND trim(valuesarray(i)&"")="" then valuesarray(i)=0
					' response.write "Upd col: " & columnarray(i) & " - " & valuesarray(i) & " : " & rs.Fields(columnarray(i)).Type & "<br>" : response.flush
					on error resume next
					err.number = 0
					rs.Fields(columnarray(i)) = valuesarray(i)
					errnum=err.number
					on error goto 0
					if errnum<>0 then
						if show_errors then
							faillines=faillines+1
							successlines=successlines-1
							response.write "Data type mismatch adding " & valuesarray(i) & " to column " & columnarray(i) & "<br>"
						end if
						if stop_errors then
							csvcurrpos=csvlen+1
							stoppedonerror=TRUE
						end if
					end if
				end if
			next
			rs.Update
		end if
		rs.Close
		' sSQL = sSQL & " WHERE pID='" & replace(valuesarray(keycolumn), "'", "''") & "'"
		' response.write "<b>" & sSQL & "</b><br />"
	else
		on error resume next
		rs.Open "products",cnn,0,3
		rs.AddNew
		for i=0 to columncount-1
			' response.write "Add col: " & columnarray(i) & " - " & valuesarray(i) & "<br>"
			rs.Fields(columnarray(i)) = valuesarray(i)
		next
		err.number = 0
		rs.Update
		errnum=err.number
		errdesc=err.description
		if errnum<>0 then
			if errnum=-2147217887 OR errnum=-2147467259 then
				errdesc = "Error, duplicate ID column."
				pidnotfoundlines=pidnotfoundlines+1
			else
				faillines=faillines+1
			end if
			if show_errors then response.write "Adding pID: &quot;" & valuesarray(keycolumn) & "&quot; - " & errdesc & " (" &  errnum & ")<br>"
			if stop_errors then
				csvcurrpos=csvlen+1
				stoppedonerror=TRUE
			end if
			csvsuccess=FALSE
			successlines=successlines-1
		end if
		rs.Close
		on error goto 0
	end if
end function
if isposted then
	// response.write '<meta http-equiv="refresh" content="2; url=admincsv.asp">';
	time_start = timer()
	column_list = lcase(replace(getcsvline(),"""",""))
	if column_list="pid,pname,pprice,pinstock,optid,optiongroup,option" then
		isstockupdate=TRUE
	else
		on error resume next
		err.number = 0
		cnn.Execute("SELECT " & column_list & " FROM products WHERE pID='abcwxyz'")
		errnum=err.number
		errdesc=err.description
		on error goto 0
		if errnum<>0 then
			errmsg = errdesc
			success=FALSE
		end if
	end if
	columnarray = split(lcase(column_list), ",")
	valuesarray = columnarray
	columncount = UBOUND(columnarray)+1
	columnnum=0
	keycolumn=""
	for i=0 to columncount-1
		columnarray(i)=trim(columnarray(i))
		if columnarray(i)="pid" then keycolumn=i
	next
	if keycolumn="" then
		success=FALSE
		errmsg="There was no pID column specified."
	end if
	if success then
		if isupdate then
			response.write "&nbsp;Updating row: "
		else
			response.write "&nbsp;Adding row: "
		end if
		line_num = 1
		totallines=20
		do while csvcurrpos < csvlen
			thechar = mid(csvfile, csvcurrpos, 1)
			' response.write "&lt;"&thechar&">"
			if NOT needquote then
				if thiscol="" AND thechar = """" then
					needquote = TRUE
				elseif thechar <> "," AND thechar <> vbCr AND thechar <> vbLf then
					thiscol=thiscol&thechar
				else
					valuesarray(columnnum)=thiscol
					columnnum=columnnum+1
					' response.write "<b>Adding col:</b>" & columnnum & ": " & thiscol & "<br>"
					if columnnum=columncount then
						successlines=successlines+1
						columnnum=0
						execute_sql()
						if (line_num MOD progressevery) = 0 then
							response.write line_num & ", "
							response.flush()
						end if
						needquote=FALSE
						do while csvcurrpos<csvlen
							tmpchar = mid(csvfile, csvcurrpos+1, 1)
							if tmpchar = vbCr OR tmpchar = vbLf then csvcurrpos=csvcurrpos+1 else exit do
						loop
						line_num = line_num + 1
					end if
					thiscol=""
				end if
			elseif thechar = """" then
				if mid(csvfile, csvcurrpos+1, 1) = """" then
					thiscol = thiscol & """"
					csvcurrpos = csvcurrpos + 1
				else
					needquote=FALSE
				end if
			else
				pos = instr(csvcurrpos, csvfile, """")
				if pos=0 then
					thiscol=thiscol & mid(csvfile, csvcurrpos, (csvlen+1) - csvcurrpos)
					csvcurrpos = csvlen
				else
					thiscol=thiscol & mid(csvfile, csvcurrpos, pos - csvcurrpos)
					' response.write "<br>ADDING THIS CHUNK: " & mid(csvfile, csvcurrpos, pos - csvcurrpos) & "<br>"
					csvcurrpos=pos-1
				end if
			end if
			csvcurrpos=csvcurrpos+1
		loop
		response.write line_num-1 & "</p>"
	end if
	time_end = timer()
%>
      <table border="0" cellspacing="0" cellpadding="0" width="100%" bgcolor="" align="center">
        <tr>
          <td width="100%">
			<table width="100%" border="0" cellspacing="0" cellpadding="3" bgcolor="">
			  <tr> 
                <td width="100%" colspan="2" align="center"><%
					if NOT success then response.write "<p>ERROR: " & errmsg & "</p>"
					if isupdate then
						response.write "<p>Rows successfully updated " & successlines & "</p>"
						if faillines > 0 then response.write "<p>Error rows " & faillines & "</p>"
						if pidnotfoundlines > 0 then response.write "<p>Rows where pID not found " & pidnotfoundlines & "</p>"
					else
						response.write "<p>Rows successfully added " & successlines & "</p>"
						if faillines > 0 then response.write "<p>Error rows " & faillines & "</p>"
						if pidnotfoundlines > 0 then response.write "<p>Rows with duplicate product id (pID) " & pidnotfoundlines & "</p>"
					end if
					response.write "<p>This page took: " & round(time_end - time_start,4) & " seconds</p>"
					if successlines + faillines > 0 then response.write "<p>That is " & round((time_end - time_start) / (successlines + faillines), 4) & " seconds per row</p>"
                %></td>
			  </tr>
			  <tr> 
                <td width="100%" colspan="2" align="center"><br /><strong><% if stoppedonerror then response.write "<font color=""#FF0000"">" & yyOpFai & "</font>" else response.write yyUpdSuc%></strong><br /><br /><br /><br />
                        Please <a href="admin.asp"><strong><%=yyClkHer%></strong></a> for the admin home page or <a href="javascript:history.go(-1)"><strong><%=yyClkHer%></strong></a> to go back and try again.<br />
                        <br />
				<img src="../images/clearpixel.gif" width="300" height="3" alt="" />
                </td>
			  </tr>
			</table></td>
        </tr>
      </table>
<%
else
%>
<script language="javascript" type="text/javascript">
<!--
function modrec(id) {
	document.mainform.id.value = id;
	document.mainform.act.value = "modify";
	document.mainform.submit();
}
function newrec(id) {
	document.mainform.id.value = id;
	document.mainform.act.value = "addnew";
	document.mainform.submit();
}
function delrec(id) {
cmsg = "<%=yyConDel%>\n"
if (confirm(cmsg)) {
	document.mainform.id.value = id;
	document.mainform.act.value = "delete";
	document.mainform.submit();
}
}
// -->
</script>
		  <form name="mainform" method="post" action="admincsv.asp" enctype="multipart/form-data">
		  <input type="hidden" name="posted" value="1">
			<table width="100%" border="0" cellspacing="0" cellpadding="2" bgcolor="">
			  <tr>
				<td width="100%" align="center" colspan="2"><strong>CSV File Upload</strong><br />&nbsp;</td>
			  </tr>
			  <tr>
				<td align="right"><strong>CSV Filename:</strong></td>
				<td><input type="file" name="csvfile" /></td>
			  </tr>
			  <tr>
				<td align="right"><strong>Action:</strong></td>
				<td><select name="theaction" size="1">
					<option value="add">Add to database</option>
					<option value="update">Update database</option>
					</select></td>
			  </tr>
			  <tr>
				<td align="right"><strong>Show Errors:</strong></td>
				<td><input type="checkbox" name="show_errors" value="ON" checked /></td>
			  </tr>
			  <tr>
				<td align="right"><strong>Stop On Errors:</strong></td>
				<td><input type="checkbox" name="stop_errors" value="ON" checked /></td>
			  </tr>
			  <tr>
				<td width="100%" align="center" colspan="2">&nbsp;<br /><input type="submit" value="Submit"><br />&nbsp;</td>
			  </tr>
			  <tr> 
				<td width="100%" align="center" colspan="2"><br />
					  <a href="admin.asp"><strong><%=yyAdmHom%></strong></a><br />
			  <img src="../images/clearpixel.gif" width="300" height="3" alt="" /></td>
			  </tr>
			</table>
		  </form>
<%
end if
%>