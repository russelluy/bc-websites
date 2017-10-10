<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!--#include file="../../../Connections/catalogmanager.asp" -->
<%
' *** Edit Operations: declare variables

Dim MM_editAction
Dim MM_abortEdit
Dim MM_editQuery
Dim MM_editCmd

Dim MM_editConnection
Dim MM_editTable
Dim MM_editRedirectUrl
Dim MM_editColumn
Dim MM_recordId

Dim MM_fieldsStr
Dim MM_columnsStr
Dim MM_fields
Dim MM_columns
Dim MM_typeArray
Dim MM_formVal
Dim MM_delim
Dim MM_altVal
Dim MM_emptyVal
Dim MM_i

MM_editAction = CStr(Request.ServerVariables("SCRIPT_NAME"))
If (Request.QueryString <> "") Then
  MM_editAction = MM_editAction & "?" & Request.QueryString
End If

' boolean to abort record edit
MM_abortEdit = false

' query string to execute
MM_editQuery = ""
%>
<%
' *** Update Record: set variables

If (CStr(Request("MM_update")) = "UpdateDetail" And CStr(Request("MM_recordId")) <> "") Then

  MM_editConnection = MM_catalogmanager_STRING
  MM_editTable = "tblCatalogDetails"
  MM_editColumn = "DetailID"
  MM_recordId = "" + Request.Form("MM_recordId") + ""
  MM_editRedirectUrl = "../update.asp"
  MM_fieldsStr  = "DetailDate1|value|DetailNumber1|value|DetailDate2|value|DetailNumber2|value|DetailDate3|value|DetailNumber3|value|DetailDate4|value|DetailNumber4|value|DetailDate5|value|DetailNumber5|value|Detailtxt1|value|Detailtxt2|value|Detailtxt3|value|Detailtxt4|value|Detailtxt5|value|DetailImageFile1|value|DetailImageFile2|value|DetailImageFile3|value|DetailImageFile4|value|DetailImageFile5|value|DetailFlag|value|DetailFlag2|value|DetailFlag3|value|DetailFlag4|value|DetailFlag5|value|DetailMemo1|value|tDetailMemo2|value|DetailMemo3|value|DetailMemo4|value|DetailMemo5|value|DownloadFile1|value|DownloadFile2|value|DownloadFile3|value|DownloadFile4|value|DownloadFile5|value"
  MM_columnsStr = "DetailDate1|',none,NULL|DetailNumber1|none,none,NULL|DetailDate2|',none,NULL|DetailNumber2|none,none,NULL|DetailDate3|',none,NULL|DetailNumber3|none,none,NULL|DetailDate4|',none,NULL|DetailNumber4|none,none,NULL|DetailDate5|',none,NULL|DetailNumber5|none,none,NULL|Detailtxt1|',none,''|Detailtxt2|',none,''|Detailtxt3|',none,''|Detailtxt4|',none,''|Detailtxt5|',none,''|DetailImage1|',none,''|DetailImage2|',none,''|DetailImage3|',none,''|DetailImage4|',none,''|DetailImage5|',none,''|DetailFlag1|',none,''|DetailFlag2|',none,''|DetailFlag3|',none,''|DetailFlag4|',none,''|DetailFlag5|',none,''|DetailMemo1|',none,''|DetailMemo2|',none,''|DetailMemo3|',none,''|DetailMemo4|',none,''|DetailMemo5|',none,''|DetailDownloadFile1|',none,''|DetailDownloadFile2|',none,''|DetailDownloadFile3|',none,''|DetailDownloadFile4|',none,''|DetailDownloadFile5|',none,''"

  ' create the MM_fields and MM_columns arrays
  MM_fields = Split(MM_fieldsStr, "|")
  MM_columns = Split(MM_columnsStr, "|")
  
  ' set the form values
  For MM_i = LBound(MM_fields) To UBound(MM_fields) Step 2
    MM_fields(MM_i+1) = CStr(Request.Form(MM_fields(MM_i)))
  Next

  ' append the query string to the redirect URL
  If (MM_editRedirectUrl <> "" And Request.QueryString <> "") Then
    If (InStr(1, MM_editRedirectUrl, "?", vbTextCompare) = 0 And Request.QueryString <> "") Then
      MM_editRedirectUrl = MM_editRedirectUrl & "?" & Request.QueryString
    Else
      MM_editRedirectUrl = MM_editRedirectUrl & "&" & Request.QueryString
    End If
  End If

End If
%>
<%
' *** Update Record: construct a sql update statement and execute it

If (CStr(Request("MM_update")) <> "" And CStr(Request("MM_recordId")) <> "") Then

  ' create the sql update statement
  MM_editQuery = "update " & MM_editTable & " set "
  For MM_i = LBound(MM_fields) To UBound(MM_fields) Step 2
    MM_formVal = MM_fields(MM_i+1)
    MM_typeArray = Split(MM_columns(MM_i+1),",")
    MM_delim = MM_typeArray(0)
    If (MM_delim = "none") Then MM_delim = ""
    MM_altVal = MM_typeArray(1)
    If (MM_altVal = "none") Then MM_altVal = ""
    MM_emptyVal = MM_typeArray(2)
    If (MM_emptyVal = "none") Then MM_emptyVal = ""
    If (MM_formVal = "") Then
      MM_formVal = MM_emptyVal
    Else
      If (MM_altVal <> "") Then
        MM_formVal = MM_altVal
      ElseIf (MM_delim = "'") Then  ' escape quotes
        MM_formVal = "'" & Replace(MM_formVal,"'","''") & "'"
      Else
        MM_formVal = MM_delim + MM_formVal + MM_delim
      End If
    End If
    If (MM_i <> LBound(MM_fields)) Then
      MM_editQuery = MM_editQuery & ","
    End If
    MM_editQuery = MM_editQuery & MM_columns(MM_i) & " = " & MM_formVal
  Next
  MM_editQuery = MM_editQuery & " where " & MM_editColumn & " = " & MM_recordId

  If (Not MM_abortEdit) Then
    ' execute the update
    Set MM_editCmd = Server.CreateObject("ADODB.Command")
    MM_editCmd.ActiveConnection = MM_editConnection
    MM_editCmd.CommandText = MM_editQuery
    MM_editCmd.Execute
    MM_editCmd.ActiveConnection.Close

    If (MM_editRedirectUrl <> "") Then
      Response.Redirect(MM_editRedirectUrl)
    End If
  End If

End If
%>
<%
Dim item_detail__value1
item_detail__value1 = "0"
If (Request.QueryString("ItemID")  <> "") Then 
  item_detail__value1 = Request.QueryString("ItemID") 
End If
%>
<%
Dim item_detail
Dim item_detail_numRows

Set item_detail = Server.CreateObject("ADODB.Recordset")
item_detail.ActiveConnection = MM_catalogmanager_STRING
item_detail.Source = "SELECT tblCatalog.*, tblCatalogSubCategory.*, tblCatalogCategory.*, tblGPC.*, tblCatalogDetails.*, tblManufacturers.*  FROM ((((tblCatalogCategory RIGHT JOIN tblCatalogSubCategory ON tblCatalogCategory.CategoryID = tblCatalogSubCategory.CategoryIDkey) RIGHT JOIN tblCatalog ON tblCatalogSubCategory.SubCategoryID = tblCatalog.SubCategoryIDKey) LEFT JOIN tblGPC ON tblCatalogCategory.GPCIDkey = tblGPC.GPCID) LEFT JOIN tblCatalogDetails ON tblCatalog.ItemID = tblCatalogDetails.ItemIDKey) LEFT JOIN tblManufacturers ON tblCatalog.ManufacturerIDkey = tblManufacturers.ManufacturerID  WHERE ItemID = " + Replace(item_detail__value1, "'", "''") + ""
item_detail.CursorType = 0
item_detail.CursorLocation = 2
item_detail.LockType = 1
item_detail.Open()

item_detail_numRows = 0
%>
<html>
<head>
<title>Catalog Manager</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link href="../../styles.css" rel="stylesheet" type="text/css">
<script language="JavaScript" type="text/JavaScript">
<!--
function MM_openBrWindow(theURL,winName,features) { //v2.0
  window.open(theURL,winName,features);
}
//-->
</script>
</head>
<body>
<!--#include file="header.asp" -->
<form ACTION="<%=MM_editAction%>" METHOD="POST" name="UpdateDetail" id="UpdateDetail">
<table width="100%" align="center" class="tableborder">
  <tr>
    <td colspan="5" align="right" nowrap>Update Details: <strong><%=(item_detail.Fields.Item("ItemName").Value)%> </strong>-
      Item: <strong><%=(item_detail.Fields.Item("ItemID").Value)%></strong> </td>
  </tr>
  <tr>
    <td colspan="2" align="right" valign="top" nowrap>
      <table width="100%" border="0" cellpadding="3" cellspacing="1" class="tableborder">
        <tr>
          <td width="18%" class="tableheader">Extra Date 1:</td>
          <td width="31%"><input name="DetailDate1" type="text" id="DetailDate12" value="<%=(item_detail.Fields.Item("DetailDate1").Value)%>" size="20">
          </td>
          <td width="21%" class="tableheader">Extra Number 1:</td>
          <td width="30%"><input name="DetailNumber1" type="text" id="DetailNumber12" value="<%=(item_detail.Fields.Item("DetailNumber1").Value)%>" size="20">
          </td>
        </tr>
        <tr>
          <td class="tableheader">Extra Date 2:</td>
          <td><input name="DetailDate2" type="text" id="DetailDate22" value="<%=(item_detail.Fields.Item("DetailDate2").Value)%>" size="20">
          </td>
          <td class="tableheader">Extra Number 2:</td>
          <td><input name="DetailNumber2" type="text" id="DetailNumber22" value="<%=(item_detail.Fields.Item("DetailNumber2").Value)%>" size="20">
          </td>
        </tr>
        <tr>
          <td class="tableheader">Extra Date 3:</td>
          <td><input name="DetailDate3" type="text" id="DetailDate32" value="<%=(item_detail.Fields.Item("DetailDate3").Value)%>" size="20">
          </td>
          <td class="tableheader">Extra Number 3:</td>
          <td><input name="DetailNumber3" type="text" id="DetailNumber32" value="<%=(item_detail.Fields.Item("DetailNumber3").Value)%>" size="20">
          </td>
        </tr>
        <tr>
          <td class="tableheader">Extra Date 4:</td>
          <td><input name="DetailDate4" type="text" id="DetailDate42" value="<%=(item_detail.Fields.Item("DetailDate4").Value)%>" size="20">
          </td>
          <td class="tableheader">Extra Number 4:</td>
          <td><input name="DetailNumber4" type="text" id="DetailNumber42" value="<%=(item_detail.Fields.Item("DetailNumber4").Value)%>" size="20">
          </td>
        </tr>
        <tr>
          <td class="tableheader">Extra Date 5:</td>
          <td><input name="DetailDate5" type="text" id="DetailDate52" value="<%=(item_detail.Fields.Item("DetailDate5").Value)%>" size="20">
          </td>
          <td class="tableheader">Extra Number 5:</td>
          <td><input name="DetailNumber5" type="text" id="DetailNumber52" value="<%=(item_detail.Fields.Item("DetailNumber5").Value)%>" size="20">
          </td>
        </tr>
      </table>
      <br>
      <table width="100%" border="0" cellpadding="0" cellspacing="1" class="tableborder">
        <tr>
          <td width="18%" class="tableheader">Extra Text 1: </td>
          <td width="82%"><textarea name="Detailtxt1" cols="45" rows="2"><%=(item_detail.Fields.Item("Detailtxt1").Value)%></textarea>
          </td>
        </tr>
        <tr>
          <td class="tableheader">Extra Text 2: </td>
          <td><textarea name="Detailtxt2" cols="45" rows="2"><%=(item_detail.Fields.Item("Detailtxt2").Value)%></textarea>
          </td>
        </tr>
        <tr>
          <td class="tableheader">Extra Text 3: </td>
          <td><textarea name="Detailtxt3" cols="45" rows="2"><%=(item_detail.Fields.Item("Detailtxt3").Value)%></textarea>
          </td>
        </tr>
        <tr>
          <td class="tableheader">Extra Text 4: </td>
          <td><textarea name="Detailtxt4" cols="45" rows="2"><%=(item_detail.Fields.Item("Detailtxt4").Value)%></textarea>
          </td>
        </tr>
        <tr>
          <td class="tableheader">Extra Text 5: </td>
          <td><textarea name="Detailtxt5" cols="45" rows="2"><%=(item_detail.Fields.Item("Detailtxt5").Value)%></textarea>
          </td>
        </tr>
      </table>
      <br>
      <table width="100%" border="0" cellpadding="3" cellspacing="1" class="tableborder">
        <tr>
          <td width="19%" height="20" class="tableheader"><div align="left"> Extra Image 1</div>
          </td>
          <td width="47%"><div align="left">
              <input name="DetailImageFile1" type="text" id="DetailImageFile1" value="<%=(item_detail.Fields.Item("DetailImage1").Value)%>" size="15">
              | <a href="javascript:;" onClick="MM_openBrWindow('upload_extra_image1.asp?DetailID=<%=(item_detail.Fields.Item("DetailID").Value)%>','Image','width=300,height=150')">Add/Edit
              Image </a></div>
          </td>
          <td width="34%"><%		  						  
Dim objextraimage1
strImage = "../../../applications/CatalogManager/images/" & item_detail.Fields.Item("DetailImage1").Value
Set objextraimage1 = CreateObject("Scripting.FileSystemObject")
If objextraimage1.FileExists(Server.MapPath(strImage)) then
%>
              <% if item_detail.Fields.Item("DetailImage1").Value <> "" then %>
              <img src="../../../applications/CatalogManager/images/<%=(item_detail.Fields.Item("DetailImage1").Value)%>" width="50">
              <% end if ' image check %>
              <% end if %>
          </td>
        </tr>
        <tr>
          <td width="19%" height="36" class="tableheader">
            <div align="left">Extra Image 2</div>
          </td>
          <td>            <div align="left">
              <input name="DetailImageFile2" type="text" id="DetailImageFile2" value="<%=(item_detail.Fields.Item("DetailImage2").Value)%>" size="15">
              |  <a href="javascript:;" onClick="MM_openBrWindow('upload_extra_image2.asp?DetailID=<%=(item_detail.Fields.Item("DetailID").Value)%>','Image','width=300,height=150')">Add/Edit
              Image </a> </div>
          </td>
          <td><%		  						  
Dim objextraimage2
strImage = "../../../applications/CatalogManager/images/" & item_detail.Fields.Item("DetailImage2").Value
Set objextraimage2 = CreateObject("Scripting.FileSystemObject")
If objextraimage2.FileExists(Server.MapPath(strImage)) then
%>
              <% if item_detail.Fields.Item("DetailImage2").Value <> "" then %>
              <img src="../../../applications/CatalogManager/images/<%=(item_detail.Fields.Item("DetailImage2").Value)%>" width="50">
              <% end if ' image check %>
              <% end if %>
          </td>
        </tr>
        <tr>
          <td width="19%" height="36" class="tableheader">
            <div align="left">Extra Image 3</div>
          </td>
          <td>            <div align="left">
              <input name="DetailImageFile3" type="text" id="DetailImageFile3" value="<%=(item_detail.Fields.Item("DetailImage3").Value)%>" size="15">
              |  <a href="javascript:;" onClick="MM_openBrWindow('upload_extra_image3.asp?DetailID=<%=(item_detail.Fields.Item("DetailID").Value)%>','Image','width=300,height=150')">Add/Edit
              Image </a> </div>
          </td>
          <td>
            <%		  						  
Dim objextraimage3
strImage = "../../../applications/CatalogManager/images/" & item_detail.Fields.Item("DetailImage3").Value
Set objextraimage3 = CreateObject("Scripting.FileSystemObject")
If objextraimage3.FileExists(Server.MapPath(strImage)) then
%>
            <% if item_detail.Fields.Item("DetailImage3").Value <> "" then %>
            <img src="../../../applications/CatalogManager/images/<%=(item_detail.Fields.Item("DetailImage3").Value)%>" width="50">
            <% end if ' image check %>
            <% end if %>
          </td>
        </tr>
        <tr>
          <td height="36" class="tableheader">Extra Image 4</td>
          <td><input name="DetailImageFile4" type="text" id="DetailImageFile4" value="<%=(item_detail.Fields.Item("DetailImage4").Value)%>" size="15">
            |  <a href="javascript:;" onClick="MM_openBrWindow('upload_extra_image4.asp?DetailID=<%=(item_detail.Fields.Item("DetailID").Value)%>','Image','width=300,height=150')">Add/Edit
            Image </a></td>
          <td><%		  						  
Dim objextraimage4
strImage = "../../../applications/CatalogManager/images/" & item_detail.Fields.Item("DetailImage4").Value
Set objextraimage4 = CreateObject("Scripting.FileSystemObject")
If objextraimage4.FileExists(Server.MapPath(strImage)) then
%>
              <% if item_detail.Fields.Item("DetailImage4").Value <> "" then %>
              <img src="../../../applications/CatalogManager/images/<%=(item_detail.Fields.Item("DetailImage4").Value)%>" width="50">
              <% end if ' image check %>
              <% end if %>
          </td>
        </tr>
        <tr>
          <td height="36" class="tableheader">Extra Image 5</td>
          <td><input name="DetailImageFile5" type="text" id="DetailImageFile5" value="<%=(item_detail.Fields.Item("DetailImage5").Value)%>" size="15">
            |  <a href="javascript:;" onClick="MM_openBrWindow('upload_extra_image5.asp?DetailID=<%=(item_detail.Fields.Item("DetailID").Value)%>','Image','width=300,height=150')">Add/Edit
            Image </a></td>
          <td><%		  						  
Dim objextraimage5
strImage = "../../../applications/CatalogManager/images/" & item_detail.Fields.Item("DetailImage5").Value
Set objextraimage5 = CreateObject("Scripting.FileSystemObject")
If objextraimage5.FileExists(Server.MapPath(strImage)) then
%>
              <% if item_detail.Fields.Item("DetailImage5").Value <> "" then %>
              <img src="../../../applications/CatalogManager/images/<%=(item_detail.Fields.Item("DetailImage5").Value)%>" width="50">
              <% end if ' image check %>
              <% end if %>
          </td>
        </tr>
      </table>
      <br>
      <table width="100%" border="0" cellpadding="3" cellspacing="1" class="tableborder">
        <tr class="tableheader">
          <td width="20%" height="21"><div align="center">Extra Flag 1:
                  <input <%if item_detail.Fields.Item("DetailFlag1").Value <> "" then %><%If (CStr((item_detail.Fields.Item("DetailFlag1").Value)) = CStr("True")) Then Response.Write("checked") : Response.Write("")%> <%end if%> name="DetailFlag" type="checkbox" id="DetailFlag" value="True" >
            </div>
          </td>
          <td width="20%"><div align="center">Extra Flag 2:
                  <input <%if item_detail.Fields.Item("DetailFlag2").Value <> "" then %><%If (CStr((item_detail.Fields.Item("DetailFlag2").Value)) = CStr("True")) Then Response.Write("checked") : Response.Write("")%> <%end if%>  name="DetailFlag2" type="checkbox" id="DetailFlag2" value="True">
            </div>
          </td>
          <td width="20%"><div align="center">Extra Flag 3:
                  <input <%if item_detail.Fields.Item("DetailFlag3").Value <> "" then %><%If (CStr((item_detail.Fields.Item("DetailFlag3").Value)) = CStr("True")) Then Response.Write("checked") : Response.Write("")%> <%end if%>  name="DetailFlag3" type="checkbox" id="DetailFlag3" value="True">
            </div>
          </td>
          <td width="20%"><div align="center">Extra Flag 4:
                  <input <%if item_detail.Fields.Item("DetailFlag4").Value <> "" then %><%If (CStr((item_detail.Fields.Item("DetailFlag4").Value)) = CStr("True")) Then Response.Write("checked") : Response.Write("")%> <%end if%> name="DetailFlag4" type="checkbox" id="DetailFlag4" value="True">
            </div>
          </td>
          <td width="20%"><div align="center">Extra Flag 5:
                  <input <%if item_detail.Fields.Item("DetailFlag5").Value <> "" then %><%If (CStr((item_detail.Fields.Item("DetailFlag5").Value)) = CStr("True")) Then Response.Write("checked") : Response.Write("")%> <%end if%>  name="DetailFlag5" type="checkbox" id="DetailFlag5" value="True">
            </div>
          </td>
        </tr>
      </table>      <p>&nbsp;</p>
    </td>
    <td width="1%" valign="baseline">&nbsp;</td>
    <td width="51%" colspan="2" valign="top">
      <table width="100%" border="0" cellpadding="0" cellspacing="1" class="tableborder">
        <tr>
          <td width="18%" class="tableheader">Extra Memo 1: </td>
          <td width="82%"><textarea name="DetailMemo1" cols="45" rows="5"><%=(item_detail.Fields.Item("DetailMemo1").Value)%></textarea>
          </td>
        </tr>
        <tr>
          <td class="tableheader">Extra Memo 2: </td>
          <td><textarea name="tDetailMemo2" cols="45" rows="5"><%=(item_detail.Fields.Item("DetailMemo2").Value)%></textarea>
          </td>
        </tr>
        <tr>
          <td class="tableheader">Extra Memo 3: </td>
          <td><textarea name="DetailMemo3" cols="45" rows="5"><%=(item_detail.Fields.Item("DetailMemo3").Value)%></textarea>
          </td>
        </tr>
        <tr>
          <td class="tableheader">Extra Memo 4: </td>
          <td><textarea name="DetailMemo4" cols="45" rows="5"><%=(item_detail.Fields.Item("DetailMemo4").Value)%></textarea>
          </td>
        </tr>
        <tr>
          <td class="tableheader">Extra Memo 5: </td>
          <td><textarea name="DetailMemo5" cols="45" rows="5"><%=(item_detail.Fields.Item("DetailMemo5").Value)%></textarea>
          </td>
        </tr>
      </table>
      <br>
      <table width="100%" border="0" cellpadding="3" cellspacing="1" class="tableborder">
        <tr>
          <td width="26%" height="25" class="tableheader">Extra Download File
            1</td>
          <td width="74%">Enter URL address to external file:
            <input name="DownloadFile1" type="text" id="DownloadFile1" value="<%=(item_detail.Fields.Item("DetailDownloadFile1").Value)%>" size="32">              <br>
            OR <a href="javascript:;" onClick="MM_openBrWindow('upload_extra_file1.asp?DetailID=<%=(item_detail.Fields.Item("DetailID").Value)%>','File','width=300,height=150')"> Upload
            File To Server</a></td>
        </tr>
        <tr>
          <td height="25" class="tableheader">Extra Download File 2</td>
          <td> Enter URL address to external file:
            <input name="DownloadFile2" type="text" id="DownloadFile2" value="<%=(item_detail.Fields.Item("DetailDownloadFile2").Value)%>" size="32">
            <br>
            OR <a href="javascript:;" onClick="MM_openBrWindow('upload_extra_file2.asp?DetailID=<%=(item_detail.Fields.Item("DetailID").Value)%>','File','width=300,height=150')"> Upload
            File To Server</a>          </td>
        </tr>
        <tr>
          <td height="25" class="tableheader">Extra Download File 3</td>
          <td>Enter URL address to external file:
            <input name="DownloadFile3" type="text" id="DownloadFile3" value="<%=(item_detail.Fields.Item("DetailDownloadFile3").Value)%>" size="32">
            <br>
            OR <a href="javascript:;" onClick="MM_openBrWindow('upload_extra_file3.asp?DetailID=<%=(item_detail.Fields.Item("DetailID").Value)%>','File','width=300,height=150')"> Upload
            File To Server</a>          </td>
        </tr>
        <tr>
          <td height="25" class="tableheader">Extra Download File 4</td>
          <td>Enter URL address to external file:
            <input name="DownloadFile4" type="text" id="DownloadFile4" value="<%=(item_detail.Fields.Item("DetailDownloadFile4").Value)%>" size="32">
            <br>
            OR <a href="javascript:;" onClick="MM_openBrWindow('upload_extra_file4.asp?DetailID=<%=(item_detail.Fields.Item("DetailID").Value)%>','File','width=300,height=150')"> Upload
            File To Server</a>          </td>
        </tr>
        <tr>
          <td height="25" class="tableheader">Extra Download File 5</td>
          <td>Enter URL address to external file:
            <input name="DownloadFile5" type="text" id="DownloadFile5" value="<%=(item_detail.Fields.Item("DetailDownloadFile5").Value)%>" size="32">
            <br>
            OR <a href="javascript:;" onClick="MM_openBrWindow('upload_extra_file5.asp?DetailID=<%=(item_detail.Fields.Item("DetailID").Value)%>','File','width=300,height=150')"> Upload
            File To Server</a>          </td>
        </tr>
      </table>
      <br>
      <br>
      <div align="center"><br>
          <br>
          <input name="submit2" type="submit" value="Update Record">
      </div>
    </td>
  </tr>
</table>

<input type="hidden" name="MM_update" value="UpdateDetail">
<input type="hidden" name="MM_recordId" value="<%= item_detail.Fields.Item("DetailID").Value %>">
</form>
<p>&nbsp;</p>
</body>
</html>
<%
item_detail.Close()
Set item_detail = Nothing
%>
