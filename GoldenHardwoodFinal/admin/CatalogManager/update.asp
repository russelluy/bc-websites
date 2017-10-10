<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="../../Connections/catalogmanager.asp" -->
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

If (CStr(Request("MM_update")) = "form1" And CStr(Request("MM_recordId")) <> "") Then

  MM_editConnection = MM_catalogmanager_STRING
  MM_editTable = "tblCatalog"
  MM_editColumn = "ItemID"
  MM_recordId = "" + Request.Form("MM_recordId") + ""
  MM_editRedirectUrl = ""
  MM_fieldsStr  = "ItemName|value|ItemDesc|value|ItemDescShort|value|ItemPrice|value|UnitOfMeasure|value|ItemPrice2|value|ItemCost|value|ItemCost2|value|Feature1|value|Feature2|value|Feature3|value|Feature4|value|Feature5|value|Activated|value|OrderAvailabilityFlag|value|InStock|value|ManufacturerID|value|SubCategoryID|value|DownloadFile|value|DownloadFile2|value|OrderLink|value|ImageFile1|value|ImageFile2|value|ImageFileThumb|value|ImageFileThumb2|value"
  MM_columnsStr = "ItemName|',none,''|ItemDesc|',none,''|ItemDescShort|',none,''|ItemPrice|none,none,NULL|UnitOfMeasure|',none,''|ItemPrice2|none,none,NULL|ItemCost|none,none,NULL|ItemCost2|none,none,NULL|Feature1|',none,''|Feature2|',none,''|Feature3|',none,''|Feature4|',none,''|Feature5|',none,''|Activated|',none,''|OrderAvailabilityFlag|',none,''|InStock|',none,''|ManufacturerIDkey|none,none,NULL|SubCategoryIDKey|none,none,NULL|DownloadFile|',none,''|DownloadFile2|',none,''|OrderLink|',none,''|ImageFile|',none,''|ImageFile2|',none,''|ImageFileThumb|',none,''|ImageFileThumb2|',none,''"

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
Dim item_list__value1
item_list__value1 = "0"
If (Request.QueryString("ItemID")   <> "") Then 
  item_list__value1 = Request.QueryString("ItemID")  
End If
%>
<%
Dim item_list
Dim item_list_numRows

Set item_list = Server.CreateObject("ADODB.Recordset")
item_list.ActiveConnection = MM_catalogmanager_STRING
item_list.Source = "SELECT tblGPC.*, tblCatalogCategory.*, tblCatalogSubCategory.*, tblManufacturers.*, tblCatalog.*, tblCatalogDetails.*  FROM ((((tblCatalogDetails RIGHT JOIN tblCatalog ON tblCatalogDetails.ItemIDKey = tblCatalog.ItemID) LEFT JOIN tblManufacturers ON tblCatalog.ManufacturerIDkey = tblManufacturers.ManufacturerID) LEFT JOIN tblCatalogSubCategory ON tblCatalog.SubCategoryIDKey = tblCatalogSubCategory.SubCategoryID) LEFT JOIN tblCatalogCategory ON tblCatalogSubCategory.CategoryIDkey = tblCatalogCategory.CategoryID) LEFT JOIN tblGPC ON tblCatalogCategory.GPCIDkey = tblGPC.GPCID  WHERE ItemID = " + Replace(item_list__value1, "'", "''") + ""
item_list.CursorType = 2
item_list.CursorLocation = 2
item_list.LockType = 1
item_list.Open()

item_list_numRows = 0
%>
<%
set subcategorymenu = Server.CreateObject("ADODB.Recordset")
subcategorymenu.ActiveConnection = MM_catalogmanager_STRING
subcategorymenu.Source = "SELECT tblGPC.GPCID, tblGPC.GPCName, tblCatalogCategory.CategoryID, tblCatalogCategory.CategoryName, tblCatalogSubCategory.SubCategoryID, tblCatalogSubCategory.SubCategoryName  FROM (tblGPC INNER JOIN tblCatalogCategory ON tblGPC.GPCID = tblCatalogCategory.GPCIDkey) INNER JOIN tblCatalogSubCategory ON tblCatalogCategory.CategoryID = tblCatalogSubCategory.CategoryIDkey"
subcategorymenu.CursorType = 0
subcategorymenu.CursorLocation = 2
subcategorymenu.LockType = 3
subcategorymenu.Open()
subcategorymenu_numRows = 0
%>
<%
Dim manufacturer
Dim manufacturer_numRows

Set manufacturer = Server.CreateObject("ADODB.Recordset")
manufacturer.ActiveConnection = MM_catalogmanager_STRING
manufacturer.Source = "SELECT *  FROM tblManufacturers"
manufacturer.CursorType = 0
manufacturer.CursorLocation = 2
manufacturer.LockType = 1
manufacturer.Open()

manufacturer_numRows = 0
%>
<html>
<head>
<title>Catalog Manager</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link href="../styles.css" rel="stylesheet" type="text/css">
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
<form ACTION="<%=MM_editAction%>" METHOD="POST" name="form1">
  <table width="100%" align="center" class="tableborder">
    <tr>
      <td colspan="5" align="right" nowrap>Update Record</td>
    </tr>
    <tr>
      <td colspan="2" align="right" valign="top" nowrap>        <table width="100%" height="78" border="0" cellpadding="3" cellspacing="1" class="tableborder">
        <tr class="tableheader">
          <td height="20" colspan="2"><a href="javascript:;" onClick="MM_openBrWindow('CategoryManager/list.asp','Category','scrollbars=yes,width=600,height=400')">Add
              New Category</a> | <a href="javascript:;" onClick="MM_openBrWindow('ManufacturerManager/list.asp','Manufacturer','scrollbars=yes,width=600,height=400')">Add
              New Manufacturer</a></td>
        </tr>
        <tr>
          <td width="26%" height="27" class="tableheader">Category:</td>
          <td width="74%"><select name="SubCategoryID" id="SubCategoryID">
              <%
While (NOT subcategorymenu.EOF)
%>
              <option value="<%=(subcategorymenu.Fields.Item("SubCategoryID").Value)%>" <%If (Not isNull(item_list.Fields.Item("SubCategoryID").Value)) Then If (CStr(subcategorymenu.Fields.Item("SubCategoryID").Value) = CStr(item_list.Fields.Item("SubCategoryID").Value)) Then Response.Write("SELECTED") : Response.Write("")%> ><%=(subcategorymenu.Fields.Item("GPCName").Value)%>&nbsp;|&nbsp;<%=(subcategorymenu.Fields.Item("CategoryName").Value)%>&nbsp;|&nbsp;<%=(subcategorymenu.Fields.Item("SubCategoryName").Value)%></option>
              <%
  subcategorymenu.MoveNext()
Wend
If (subcategorymenu.CursorType > 0) Then
  subcategorymenu.MoveFirst
Else
  subcategorymenu.Requery
End If
%>
            </select>
          </td>
        </tr>
        <tr>
          <td height="27" class="tableheader">Manufacturer:</td>
          <td><select name="ManufacturerID" id="ManufacturerID">
              <%
While (NOT manufacturer.EOF)
%>
              <option value="<%=(manufacturer.Fields.Item("ManufacturerID").Value)%>" <%If (Not isNull((item_list.Fields.Item("ManufacturerID").Value))) Then If (CStr(manufacturer.Fields.Item("ManufacturerID").Value) = CStr((item_list.Fields.Item("ManufacturerID").Value))) Then Response.Write("SELECTED") : Response.Write("")%> ><%=(manufacturer.Fields.Item("Manufacturer").Value)%></option>
              <%
  manufacturer.MoveNext()
Wend
If (manufacturer.CursorType > 0) Then
  manufacturer.MoveFirst
Else
  manufacturer.Requery
End If
%>
            </select>
          </td>
        </tr>
      </table>        <br>
        <table width="100%" border="0" cellpadding="3" cellspacing="1" class="tableborder">
          <tr>
            <td class="tableheader">Item Name:</td>
            <td><textarea name="ItemName" cols="40" rows="2"><%=(item_list.Fields.Item("ItemName").Value)%></textarea></td>
          </tr>
          <tr>
            <td class="tableheader">Item Description:</td>
            <td>
              <div align="left">
                  <textarea name="ItemDesc" cols="40" rows="5"><%=(item_list.Fields.Item("ItemDesc").Value)%></textarea>
                </div></td>
          </tr>
          <tr>
            <td height="76" class="tableheader">Item Short Description:</td>
            <td>
              <div align="left">
                  <textarea name="ItemDescShort" cols="40" rows="3"><%=(item_list.Fields.Item("ItemDescShort").Value)%></textarea>
                </div></td>
          </tr>
        </table>
        <br>
        <table width="100%" border="0" cellpadding="0" cellspacing="1" class="tableborder">
          <tr>
            <td width="28%" class="tableheader">Feature 1: </td>
            <td width="72%"><textarea name="Feature1" cols="40" rows="2"><%=(item_list.Fields.Item("Feature1").Value)%></textarea>
            </td>
          </tr>
          <tr>
            <td class="tableheader">Feature 2:</td>
            <td><textarea name="Feature2" cols="40" rows="2"><%=(item_list.Fields.Item("Feature2").Value)%></textarea>
            </td>
          </tr>
          <tr>
            <td class="tableheader">Feature 3:</td>
            <td><textarea name="Feature3" cols="40" rows="2"><%=(item_list.Fields.Item("Feature3").Value)%></textarea>
            </td>
          </tr>
          <tr>
            <td class="tableheader">Feature 4:</td>
            <td><textarea name="Feature4" cols="40" rows="2"><%=(item_list.Fields.Item("Feature4").Value)%></textarea>
            </td>
          </tr>
          <tr>
            <td class="tableheader">Feature 5:</td>
            <td><textarea name="Feature5" cols="40" rows="2"><%=(item_list.Fields.Item("Feature5").Value)%></textarea>
            </td>
          </tr>
        </table>
      </td>
      <td width="2%" valign="baseline">&nbsp;</td>
      <td colspan="2" valign="top" width="100%"><table width="100%" border="0" cellpadding="3" cellspacing="1" class="tableborder">
        <tr>
          <td width="19%" class="tableheader">Item Price:</td>
          <td width="38%"><input type="text" name="ItemPrice" value="<%=(item_list.Fields.Item("ItemPrice").Value)%>" size="10">
              <strong>/</strong>
              <select name="UnitOfMeasure" id="select">
                <option value="Unit" selected <%If (Not isNull((item_list.Fields.Item("UnitOfMeasure").Value))) Then If ("Unit" = CStr((item_list.Fields.Item("UnitOfMeasure").Value))) Then Response.Write("SELECTED") : Response.Write("")%>>Unit</option>
                <option value="Hour" <%If (Not isNull((item_list.Fields.Item("UnitOfMeasure").Value))) Then If ("Hour" = CStr((item_list.Fields.Item("UnitOfMeasure").Value))) Then Response.Write("SELECTED") : Response.Write("")%>>Hour</option>
                <option value="Year" <%If (Not isNull((item_list.Fields.Item("UnitOfMeasure").Value))) Then If ("Year" = CStr((item_list.Fields.Item("UnitOfMeasure").Value))) Then Response.Write("SELECTED") : Response.Write("")%>>Year</option>
              </select>
          </td>
          <td width="18%" class="tableheader">Item Price 2:</td>
          <td width="25%"><input name="ItemPrice2" type="text" id="ItemPrice2" value="<%=(item_list.Fields.Item("ItemPrice2").Value)%>" size="10">
          </td>
        </tr>
        <tr>
          <td class="tableheader">Item Cost:</td>
          <td><input name="ItemCost" type="text" id="ItemCost3" value="<%=(item_list.Fields.Item("ItemCost").Value)%>" size="10">
          </td>
          <td class="tableheader">Item Cost 2:</td>
          <td><input name="ItemCost2" type="text" id="ItemCost22" value="<%=(item_list.Fields.Item("ItemCost2").Value)%>" size="10">
          </td>
        </tr>
      </table>
        <br>
        <table width="100%" border="0" cellpadding="3" cellspacing="1" class="tableborder">
          <tr class="tableheader">
            <td width="31%">Activated:
                <input <%If (CStr((item_list.Fields.Item("Activated").Value)) = CStr("True")) Then Response.Write("checked") : Response.Write("")%> name="Activated" type="checkbox" id="Activated" value="True">
            </td>
            <td width="39%"><div align="center">Available for Order?
                    <input <%If (CStr((item_list.Fields.Item("OrderAvailabilityFlag").Value)) = CStr("True")) Then Response.Write("checked") : Response.Write("")%> name="OrderAvailabilityFlag" type="checkbox" id="OrderAvailabilityFlag" value="True">
              </div>
            </td>
            <td width="30%"><div align="right">In Stock
                    <input <%If (CStr((item_list.Fields.Item("InStock").Value)) = CStr("True")) Then Response.Write("checked") : Response.Write("")%> name="InStock" type="checkbox" id="InStock2" value="True">
              </div>
            </td>
          </tr>
        </table>
        <br>
        <table width="100%" border="0" cellpadding="3" cellspacing="1" class="tableborder">
        <tr class="tableheader">
          <td width="50%" height="20"><div align="left"> Image 1: Large Size
              (400 x 400)</div>
          </td>
          <td><div align="left">Image 1: Thumbnail Size (150 x 150)</div>
          </td>
        </tr>
        <tr>
          <td width="50%" height="36">
            <div align="left">
              <input type="text" name="ImageFile1" value="<%=(item_list.Fields.Item("ImageFile").Value)%>" size="15">
        | <a href="javascript:;" onClick="MM_openBrWindow('upload_i1.asp?ItemID=<%=(item_list.Fields.Item("ItemID").Value)%>','Image','scrollbars=yes,width=300,height=150')">Add/Edit
        Image </a> </div>
          </td>
          <td>
            <div align="left">
              <input name="ImageFileThumb" type="text" id="ImageFileThumb" value="<%=(item_list.Fields.Item("ImageFileThumb").Value)%>" size="15">
        | <a href="javascript:;" onClick="MM_openBrWindow('upload_ti1.asp?ItemID=<%=(item_list.Fields.Item("ItemID").Value)%>','Image','width=300,height=150')">Add/Edit
        Image </a> </div>
          </td>
        </tr>
        <tr>
          <td width="50%" height="36">
            <div align="left">
              <%		  						  
Dim objimage
strImage = "../../applications/CatalogManager/images/" & item_list.Fields.Item("ImageFile").Value
Set objimage = CreateObject("Scripting.FileSystemObject")
If objimage.FileExists(Server.MapPath(strImage)) then
%>
              <% if item_list.Fields.Item("ImageFile").Value <> "" then %>
              <img src="../../applications/CatalogManager/images/<%=(item_list.Fields.Item("ImageFile").Value)%>" width="50">
              <% end if ' image check %>
              <% end if %>
            </div>
          </td>
          <td>
            <div align="left">
              <%		  						  
Dim objthumb
strImage = "../../applications/CatalogManager/images/" & item_list.Fields.Item("ImageFileThumb").Value
Set objthumb = CreateObject("Scripting.FileSystemObject")
If objthumb.FileExists(Server.MapPath(strImage)) then
%>
              <% if item_list.Fields.Item("ImageFileThumb").Value <> "" then %>
              <img src="../../applications/CatalogManager/images/<%=(item_list.Fields.Item("ImageFileThumb").Value)%>" width="50">
              <% end if ' image check %>
              <% end if %>
            </div>
          </td>
        </tr>
      </table>
        <br>
        <table width="100%" border="0" cellpadding="3" cellspacing="1" class="tableborder">
          <tr class="tableheader">
            <td width="50%" height="20"><div align="left"> Image 2: Large Size
                (400 x 400)</div>
            </td>
            <td><div align="left">Image 2: Thumbnail Size (150 x 150)</div>
            </td>
          </tr>
          <tr>
            <td width="50%" height="36">
              <div align="left">
                <input name="ImageFile2" type="text" id="ImageFile22" value="<%=(item_list.Fields.Item("ImageFile2").Value)%>" size="15">
        | <a href="javascript:;" onClick="MM_openBrWindow('upload_i2.asp?ItemID=<%=(item_list.Fields.Item("ItemID").Value)%>','Image','width=300,height=150')">Add/Edit
        Image </a> </div>
            </td>
            <td>
              <div align="left">
                <input name="ImageFileThumb2" type="text" id="ImageFileThumb22" value="<%=(item_list.Fields.Item("ImageFileThumb2").Value)%>" size="15">
        | <a href="javascript:;" onClick="MM_openBrWindow('upload_ti2.asp?ItemID=<%=(item_list.Fields.Item("ItemID").Value)%>','Image','width=300,height=150')">Add/Edit
        Image </a></div>
            </td>
          </tr>
          <tr>
            <td width="50%" height="36">
              <div align="left">
                <%		  						  
Dim objimage2
strImage = "../../applications/CatalogManager/images/" & item_list.Fields.Item("ImageFile2").Value
Set objimage2 = CreateObject("Scripting.FileSystemObject")
If objimage2.FileExists(Server.MapPath(strImage)) then
%>
                <% if item_list.Fields.Item("ImageFile2").Value <> "" then %>
                <img src="../../applications/CatalogManager/images/<%=(item_list.Fields.Item("ImageFile2").Value)%>" width="50">
                <% end if ' image check %>
                <% end if %>
              </div>
            </td>
            <td>
              <div align="left">
                <%		  						  
Dim objthumb2
strImage = "../../applications/CatalogManager/images/" & item_list.Fields.Item("ImageFileThumb2").Value
Set objthumb2 = CreateObject("Scripting.FileSystemObject")
If objthumb2.FileExists(Server.MapPath(strImage)) then
%>
                <% if item_list.Fields.Item("ImageFileThumb2").Value <> "" then %>
                <img src="../../applications/CatalogManager/images/<%=(item_list.Fields.Item("ImageFileThumb2").Value)%>" width="50">
                <% end if ' image check %>
                <% end if %>
              </div>
            </td>
          </tr>
        </table>  
        <br>
        <table width="100%" border="0" cellpadding="3" cellspacing="1" class="tableborder">
        <tr>
          <td width="26%" height="25" class="tableheader">Download File</td>
          <td width="74%">Enter URL address to external file:
            <input name="DownloadFile" type="text" id="DownloadFile3" value="<%=(item_list.Fields.Item("DownloadFile").Value)%>" size="32">
              <br>
              OR <a href="javascript:;" onClick="MM_openBrWindow('upload_file1.asp?ItemID=<%=(item_list.Fields.Item("ItemID").Value)%>','Image','scrollbars=yes,width=300,height=150')"> Upload
              File To Server</a></td>
        </tr>
        <tr>
          <td height="25" class="tableheader">Download File2</td>
          <td>Enter URL address to external file: 
            <input name="DownloadFile2" type="text" id="DownloadFile23" value="<%=(item_list.Fields.Item("DownloadFile2").Value)%>" size="32">
            <br>
            OR <a href="javascript:;" onClick="MM_openBrWindow('upload_file2.asp?ItemID=<%=(item_list.Fields.Item("ItemID").Value)%>','Image','scrollbars=yes,width=300,height=150')">
            Upload File To Server</a></td>
        </tr>
        <tr>
          <td height="25" class="tableheader">Order URL</td>
          <td><input name="OrderLink" type="text" id="OrderLink2" value="<%=(item_list.Fields.Item("OrderLink").Value)%>" size="32"> 
          (i.e. http://www.domain.com/order.htm)</td>
        </tr>
      </table>        
        <br>      
        <div align="center">
<% if item_list.Fields.Item("ItemIDkey").Value <> "" then %>
<a href="CatalogExtraDetails/update_extra_details.asp?ItemID=<%=(item_list.Fields.Item("ItemID").Value)%>">MODIFY ADDITIONAL DETAILS</a> 
<%else%>
<a href="CatalogExtraDetails/insert_extra_details.asp?ItemID=<%=(item_list.Fields.Item("ItemID").Value)%>">DEFINE ADDITIONAL DETAILS</a>
<% end if ' image check %></div>          
          <div align="center"><br>
              <br>
              <input name="submit" type="submit" value="Update Record">
          </div></td>
    </tr>
  </table>
     <input type="hidden" name="MM_update" value="form1">
  <input type="hidden" name="MM_recordId" value="<%= item_list.Fields.Item("ItemID").Value %>">

</form>
</body>
</html>
<%
item_list.Close()
Set item_list = Nothing
%>
<%
subcategorymenu.Close()
%>
<%
manufacturer.Close()
Set manufacturer = Nothing
%>
