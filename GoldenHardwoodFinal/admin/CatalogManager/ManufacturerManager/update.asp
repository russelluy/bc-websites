<%@LANGUAGE="VBSCRIPT"%>
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

If (CStr(Request("MM_update")) = "form1" And CStr(Request("MM_recordId")) <> "") Then

  MM_editConnection = MM_catalogmanager_STRING
  MM_editTable = "tblManufacturers"
  MM_editColumn = "ManufacturerID"
  MM_recordId = "" + Request.Form("MM_recordId") + ""
  MM_editRedirectUrl = "closewindow_redirect.asp"
  MM_fieldsStr  = "ItemName2|value|ManDesc|value|Web|value|ManufacturerImageFile2|value|Activated2|value"
  MM_columnsStr = "Manufacturer|',none,''|ManufacturerDesc|',none,''|ManufacturerWebsiteAddress|',none,''|ManufacturerImageFile|',none,''|ManufacturerActivated|',none,''"

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
Dim List_Manufacturers__value1
List_Manufacturers__value1 = "%"
If (Request.queryString("manid")   <> "") Then 
  List_Manufacturers__value1 = Request.queryString("manid")  
End If
%>
<%
set List_Manufacturers = Server.CreateObject("ADODB.Recordset")
List_Manufacturers.ActiveConnection = MM_catalogmanager_STRING
List_Manufacturers.Source = "SELECT *  FROM tblManufacturers  WHERE ManufacturerID LIKE '" + Replace(List_Manufacturers__value1, "'", "''") + "'  ORDER BY ManufacturerID DESC"
List_Manufacturers.CursorType = 0
List_Manufacturers.CursorLocation = 2
List_Manufacturers.LockType = 3
List_Manufacturers.Open()
List_Manufacturers_numRows = 0
%>
<html>
<head>
<title>Update</title>
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
      <form ACTION="<%=MM_editAction%>" METHOD="POST" name="form1">
        <table width="100%" align="center" class="tableborder">
          <tr align="right" valign="top">
            <td colspan="2" class="tableheader">Update</td>
          </tr>
          <tr align="right" valign="top">
            <td class="tableheader">Manufacturer Name:</td>
            <td align="left">
              <input name="ItemName2" type="text" id="ItemName2" value="<%=(List_Manufacturers.Fields.Item("Manufacturer").Value)%>" size="50">
              <img src="questionmark.gif" alt="Enter the name of the event" width="15" height="15"> </td>
          </tr>
          <tr align="right" valign="top">
            <td class="tableheader">Manufacturer Description:</td>
            <td align="left">
              <textarea name="ManDesc" cols="50" rows="2" id="ManDesc"><%=(List_Manufacturers.Fields.Item("ManufacturerDesc").Value)%></textarea>
              <img src="questionmark.gif" alt="Enter a description of the event" width="15" height="15"> </td>
          </tr>
          <tr align="right" valign="top">
            <td class="tableheader">Manufacturer Website Address:</td>
            <td align="left"> http://
<input name="Web" type="text" id="Web" value="<%=(List_Manufacturers.Fields.Item("ManufacturerWebsiteAddress").Value)%>" size="50">
              <img src="questionmark.gif" alt="Enter a short description of the event" width="15" height="15"></td>
          </tr>
          <tr align="right" valign="top">
            <td class="tableheader">Image:</td>
            <td align="left"><%		  						  
Dim objimage
strImage = "../../../applications/CatalogManager/images/" & List_Manufacturers.Fields.Item("ManufacturerImageFile").Value
Set objimage = CreateObject("Scripting.FileSystemObject")
If objimage.FileExists(Server.MapPath(strImage)) then
%>
              <% if List_Manufacturers.Fields.Item("ManufacturerImageFile").Value <> "" then %>
              <img src="../../../applications/CatalogManager/images/<%=(List_Manufacturers.Fields.Item("ManufacturerImageFile").Value)%>" width="50">
              <% end if ' image check %>
              <% end if %>
|
<input name="ManufacturerImageFile2" type="text" id="ManufacturerImageFile2" value="<%=(List_Manufacturers.Fields.Item("ManufacturerImageFile").Value)%>">
| <a href="javascript:;" onClick="MM_openBrWindow('upload_image.asp?manid=<%=(List_Manufacturers.Fields.Item("ManufacturerID").Value)%>','Image','scrollbars=yes,width=300,height=150')">Update
Image</a> <img src="questionmark.gif" alt="Upload image associated with the event" width="15" height="15"> </td>
          </tr>
          <tr align="right" valign="top">
            <td class="tableheader">Activated:</td>
            <td align="left">
              <input <%If (CStr((List_Manufacturers.Fields.Item("ManufacturerActivated").Value)) = CStr("True")) Then Response.Write("checked") : Response.Write("")%> type="checkbox" name="Activated2" value=True>
              <img src="questionmark.gif" alt="(Check if you want this link to be visible to the public)(Ucheck if you wish to hide)" width="15" height="15"> </td>
          </tr>
          <tr align="right" valign="top">
            <td class="tableheader">&nbsp;</td>
            <td align="left"><input name="submit2" type="submit" value="submit">
            </td>
          </tr>
        </table>
        
        

        <input type="hidden" name="MM_update" value="form1">
        <input type="hidden" name="MM_recordId" value="<%= List_Manufacturers.Fields.Item("ManufacturerID").Value %>">
      </form>
</body>
</html>
<%
List_Manufacturers.Close()
%>
