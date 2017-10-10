<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="../../Connections/photogallerymanager.asp" -->
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

  MM_editConnection = MM_photogallerymanager_STRING
  MM_editTable = "tblPhotoGallery"
  MM_editColumn = "ItemID"
  MM_recordId = "" + Request.Form("MM_recordId") + ""
  MM_editRedirectUrl = "admin.asp"
  MM_fieldsStr  = "CategoryID|value|ItemName|value|ItemDesc|value|ItemDescShort|value|DateAdded|value|ImageFileA|value|ImageThumbFileA|value|ImageFileB|value|ImageThumbFileB|value|Activated|value"
  MM_columnsStr = "CategoryID|none,none,NULL|ItemName|',none,''|ItemDesc|',none,''|ItemDescShort|',none,''|DateAdded|',none,NULL|ImageFileA|',none,''|ImageThumbFileA|',none,''|ImageFileB|',none,''|ImageThumbFileB|',none,''|Activated|',none,''"

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
Dim photogallery_list__value1
photogallery_list__value1 = "%"
If (Request.queryString("ItemID")  <> "") Then 
  photogallery_list__value1 = Request.queryString("ItemID") 
End If
%>
<%
set photogallery_list = Server.CreateObject("ADODB.Recordset")
photogallery_list.ActiveConnection = MM_photogallerymanager_STRING
photogallery_list.Source = "SELECT *  FROM tblPhotoGallery  WHERE ItemID LIKE '" + Replace(photogallery_list__value1, "'", "''") + "'"
photogallery_list.CursorType = 0
photogallery_list.CursorLocation = 2
photogallery_list.LockType = 3
photogallery_list.Open()
photogallery_list_numRows = 0
%>
<%
set Category = Server.CreateObject("ADODB.Recordset")
Category.ActiveConnection = MM_photogallerymanager_STRING
Category.Source = "SELECT *  FROM tblPhotoGalleryCategory  ORDER BY CategoryID"
Category.CursorType = 0
Category.CursorLocation = 2
Category.LockType = 3
Category.Open()
Category_numRows = 0
%>
<html>
<head>
<title>Update</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link href="../styles.css" rel="stylesheet" type="text/css">
<script language="JavaScript" type="text/JavaScript">
<!--
function MM_openBrWindow(theURL,winName,features) { //v2.0
  window.open(theURL,winName,features);
}
//-->
</script>
<style type="text/css">
<!--
.style11 {font-family: Arial, Helvetica, sans-serif}
.style12 {font-size: 12px}
.style14 {
	font-family: Arial, Helvetica, sans-serif;
	font-size: 12px;
	font-weight: bold;
}
.style15 {font-family: Arial, Helvetica, sans-serif; font-size: 12px; }
-->
</style>
</head>
<body>
<!--#include file="header.asp" -->
      <form ACTION="<%=MM_editAction%>" METHOD="POST" name="form1">
        <table width="100%" align="center" class="tableborder">
          <tr align="right" valign="top">
            <td width="31%" class="style12 style11 tableheader"><strong>Category:</strong></td>
            <td width="2%" align="left">&nbsp;</td>
            <td width="67%" align="left">
              <select name="CategoryID">
              <%
While (NOT Category.EOF)
%>
              <option value="<%=(Category.Fields.Item("CategoryID").Value)%>" <%If (Not isNull((photogallery_list.Fields.Item("CategoryID").Value))) Then If (CStr(Category.Fields.Item("CategoryID").Value) = CStr((photogallery_list.Fields.Item("CategoryID").Value))) Then Response.Write("SELECTED") : Response.Write("")%> ><%=(Category.Fields.Item("CategoryName").Value)%></option>
                <%
  Category.MoveNext()
Wend
If (Category.CursorType > 0) Then
  Category.MoveFirst
Else
  Category.Requery
End If
%>
              </select>
              <span class="style14">      | <a href="javascript:;" onClick="MM_openBrWindow('add_category.asp','Category','scrollbars=yes,width=400,height=300')">add/edit
      category</a> <img src="questionmark.gif" alt="Select a category that best describes the Image i.e. Sporting Image" width="15" height="15"></span></td>
          </tr>
          <tr align="right" valign="top">
            <td class="style12 style11 tableheader"><strong>Image Name:</strong></td>
            <td align="left">&nbsp;</td>
            <td align="left">
              <input name="ItemName" type="text" id="ItemName" value="<%=(photogallery_list.Fields.Item("ItemName").Value)%>" size="50">
              <img src="questionmark.gif" alt="Enter the name of the Image" width="15" height="15"> </td>
          </tr>
          <tr align="right" valign="top">
            <td class="style12 style11 tableheader"><strong>Image Description:</strong></td>
            <td align="left">&nbsp;</td>
            <td align="left">
              <textarea name="ItemDesc" cols="50" rows="5" id="ItemDesc"><%=(photogallery_list.Fields.Item("ItemDesc").Value)%></textarea>
              <img src="questionmark.gif" alt="Enter a description of the Image" width="15" height="15"> </td>
          </tr>
          <tr align="right" valign="top">
            <td class="style12 style11 tableheader"><strong>Image Short Description:</strong></td>
            <td align="left">&nbsp;</td>
            <td align="left">
            <textarea name="ItemDescShort" cols="50" rows="5" id="ItemDescShort"><%=(photogallery_list.Fields.Item("ItemDescShort").Value)%></textarea>
            <img src="questionmark.gif" alt="Enter a short description of the Image" width="15" height="15"></td>
          </tr>
          <tr align="right" valign="top">
            <td class="style12 style11 tableheader"><strong>Date Added:</strong></td>
            <td align="left">&nbsp;</td>
            <td align="left"><input name="DateAdded" type="text" id="DateAdded" value="<%=(photogallery_list.Fields.Item("DateAdded").Value)%>">
            <img src="questionmark.gif" alt="Enter the date the Image was added" width="15" height="15"></td>
          </tr>
          <tr align="right" valign="top">
            <td class="style12 style11 tableheader"><strong>Image A:<br>
            (Large Size)</strong></td>
            <td align="left">&nbsp;</td>
            <td align="left">              <% if photogallery_list.Fields.Item("ImageFileA").Value <> "" then %>
              <img src="../../applications/PhotoGalleryManager/images/<%=(photogallery_list.Fields.Item("ImageFileA").Value)%>" width="50">
              <% end if %> 
              | 
              <input name="ImageFileA" type="text" id="ImageFileA" value="<%=(photogallery_list.Fields.Item("ImageFileA").Value)%>"> 
              <span class="style15">| <a href="javascript:;" onClick="MM_openBrWindow('upload_imageA.asp?ItemID=<%=(photogallery_list.Fields.Item("ItemID").Value)%>','Image','scrollbars=yes,width=300,height=150')">Update
Image</a></span> <img src="questionmark.gif" alt="Upload image" width="15" height="15"></td>
          </tr>
          <tr align="right" valign="top">
            <td class="style12 style11 tableheader"><strong>Image A:<br>
            (Thumbnail Size)</strong></td>
            <td align="left">&nbsp;</td>
            <td align="left">              <% if photogallery_list.Fields.Item("ImageThumbFileA").Value <> "" then %>
                <img src="../../applications/PhotoGalleryManager/images/<%=(photogallery_list.Fields.Item("ImageThumbFileA").Value)%>" width="50">
                <% end if %>
  |              
  <input name="ImageThumbFileA" type="text" id="ImageThumbFileA" value="<%=(photogallery_list.Fields.Item("ImageThumbFileA").Value)%>">
  <span class="style15">| <a href="javascript:;" class="style15" onClick="MM_openBrWindow('upload_thumbimageA.asp?ItemID=<%=(photogallery_list.Fields.Item("ItemID").Value)%>','Image','scrollbars=yes,width=300,height=150')">Update
  Image</a></span> <img src="questionmark.gif" alt="Upload image" width="15" height="15"></td></tr>
		            <tr align="right" valign="top">
            <td class="style12 style11 tableheader"><strong>Image B:<br>
(Large Size)</strong></td>
            <td align="left">&nbsp;</td>
            <td align="left">              <% if photogallery_list.Fields.Item("ImageFileB").Value <> "" then %>
              <img src="../../applications/PhotoGalleryManager/images/<%=(photogallery_list.Fields.Item("ImageFileB").Value)%>" width="50">
              <% end if %> 
              | 
              <input name="ImageFileB" type="text" id="ImageFileB" value="<%=(photogallery_list.Fields.Item("ImageFileB").Value)%>"> 
              <span class="style15">| <a href="javascript:;" class="style15" onClick="MM_openBrWindow('upload_imageB.asp?ItemID=<%=(photogallery_list.Fields.Item("ItemID").Value)%>','Image','scrollbars=yes,width=300,height=150')">Update
Image</a></span> <img src="questionmark.gif" alt="Upload image" width="15" height="15"></td>
          </tr>
          <tr align="right" valign="top">
            <td class="style12 style11 tableheader"><strong>Image B:<br>
            (Thumbnail Size)</strong></td>
            <td align="left">&nbsp;</td>
            <td align="left">              <% if photogallery_list.Fields.Item("ImageThumbFileB").Value <> "" then %>
                <img src="../../applications/PhotoGalleryManager/images/<%=(photogallery_list.Fields.Item("ImageThumbFileB").Value)%>" width="50">
                <% end if %>
  |              
  <input name="ImageThumbFileB" type="text" id="ImageThumbFileB" value="<%=(photogallery_list.Fields.Item("ImageThumbFileB").Value)%>">
  <span class="style15">  | <a href="javascript:;" onClick="MM_openBrWindow('upload_thumbimageB.asp?ItemID=<%=(photogallery_list.Fields.Item("ItemID").Value)%>','Image','scrollbars=yes,width=300,height=150')">Update
  Image</a></span> <img src="questionmark.gif" alt="Upload image" width="15" height="15"></td></tr>
          <tr align="right" valign="top">
            <td class="style12 style11 tableheader"><strong>Activated:</strong></td>
            <td align="left">&nbsp;</td>
            <td align="left"><input name="Activated" type="checkbox" value="True" <%If (CStr((photogallery_list.Fields.Item("Activated").Value)) = CStr("True")) Then Response.Write("checked") : Response.Write("")%>>
              <img src="questionmark.gif" alt="(Check if you want this link to be visible to the public)(Ucheck if you wish to hide)" width="15" height="15"> </td>
          </tr>
          <tr align="right" valign="top">
            <td class="tableheader">&nbsp;</td>
            <td align="left">&nbsp;</td>
            <td align="left"><input name="submit" type="submit" value="Update">
</td>
          </tr>
        </table>
        

        <input type="hidden" name="MM_update" value="form1">
        <input type="hidden" name="MM_recordId" value="<%= photogallery_list.Fields.Item("ItemID").Value %>">
      </form>
</body>
</html>
<%
photogallery_list.Close()
%>
<%
Category.Close()
%>



