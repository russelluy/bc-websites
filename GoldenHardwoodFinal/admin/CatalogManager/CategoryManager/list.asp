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
' *** Insert Record: set variables

If (CStr(Request("MM_insert")) = "gpc") Then

  MM_editConnection = MM_catalogmanager_STRING
  MM_editTable = "tblGPC"
  MM_editRedirectUrl = ""
  MM_fieldsStr  = "GPCName|value"
  MM_columnsStr = "GPCName|',none,''"

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
' *** Insert Record: set variables

If (CStr(Request("MM_insert")) = "category") Then

  MM_editConnection = MM_catalogmanager_STRING
  MM_editTable = "tblCatalogCategory"
  MM_editRedirectUrl = ""
  MM_fieldsStr  = "GPCID|value|CategoryName|value"
  MM_columnsStr = "GPCIDkey|none,none,NULL|CategoryName|',none,''"

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
' *** Insert Record: set variables

If (CStr(Request("MM_insert")) = "subcategory") Then

  MM_editConnection = MM_catalogmanager_STRING
  MM_editTable = "tblCatalogSubCategory"
  MM_editRedirectUrl = ""
  MM_fieldsStr  = "CategoryID|value|SubCategoryName|value"
  MM_columnsStr = "CategoryIDkey|none,none,NULL|SubCategoryName|',none,''"

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
' *** Insert Record: construct a sql insert statement and execute it

Dim MM_tableValues
Dim MM_dbValues

If (CStr(Request("MM_insert")) <> "") Then

  ' create the sql insert statement
  MM_tableValues = ""
  MM_dbValues = ""
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
      MM_tableValues = MM_tableValues & ","
      MM_dbValues = MM_dbValues & ","
    End If
    MM_tableValues = MM_tableValues & MM_columns(MM_i)
    MM_dbValues = MM_dbValues & MM_formVal
  Next
  MM_editQuery = "insert into " & MM_editTable & " (" & MM_tableValues & ") values (" & MM_dbValues & ")"

  If (Not MM_abortEdit) Then
    ' execute the insert
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
set allcategories = Server.CreateObject("ADODB.Recordset")
allcategories.ActiveConnection = MM_catalogmanager_STRING
allcategories.Source = "SELECT tblGPC.*, tblCatalogCategory.*, tblCatalogSubCategory.*  FROM (tblGPC LEFT JOIN tblCatalogCategory ON tblGPC.GPCID = tblCatalogCategory.GPCIDkey) LEFT JOIN tblCatalogSubCategory ON tblCatalogCategory.CategoryID = tblCatalogSubCategory.CategoryIDkey  ORDER BY tblGPC.GPCID, tblCatalogCategory.CategoryID, tblCatalogSubCategory.SubCategoryID"
allcategories.CursorType = 0
allcategories.CursorLocation = 2
allcategories.LockType = 3
allcategories.Open()
allcategories_numRows = 0
%>
<%
set gpcmenu = Server.CreateObject("ADODB.Recordset")
gpcmenu.ActiveConnection = MM_catalogmanager_STRING
gpcmenu.Source = "SELECT *  FROM tblGPC"
gpcmenu.CursorType = 0
gpcmenu.CursorLocation = 2
gpcmenu.LockType = 3
gpcmenu.Open()
gpcmenu_numRows = 0
%>
<%
Dim categorymenu__value1
categorymenu__value1 = "%"
If (request.querystring("gpcid")   <> "") Then 
  categorymenu__value1 = request.querystring("gpcid")  
End If
%>
<%
set categorymenu = Server.CreateObject("ADODB.Recordset")
categorymenu.ActiveConnection = MM_catalogmanager_STRING
categorymenu.Source = "SELECT *  FROM tblCatalogCategory  WHERE GPCIDkey LIKE '" + Replace(categorymenu__value1, "'", "''") + "'"
categorymenu.CursorType = 0
categorymenu.CursorLocation = 2
categorymenu.LockType = 3
categorymenu.Open()
categorymenu_numRows = 0
%>
<%
Dim subcategorymenu__value1
subcategorymenu__value1 = "%"
If (request.querystring("cid")    <> "") Then 
  subcategorymenu__value1 = request.querystring("cid")   
End If
%>
<%
set subcategorymenu = Server.CreateObject("ADODB.Recordset")
subcategorymenu.ActiveConnection = MM_catalogmanager_STRING
subcategorymenu.Source = "SELECT *  FROM tblCatalogSubCategory  WHERE CategoryIDkey LIKE '" + Replace(subcategorymenu__value1, "'", "''") + "'"
subcategorymenu.CursorType = 0
subcategorymenu.CursorLocation = 2
subcategorymenu.LockType = 3
subcategorymenu.Open()
subcategorymenu_numRows = 0
%>
<%
Dim Repeat1__numRows
Dim Repeat1__index

Repeat1__numRows = -1
Repeat1__index = 0
allcategories_numRows = allcategories_numRows + Repeat1__numRows
%>
<%
' UltraDeviant - Row Number written by Owen Palmer (http://ultradeviant.co.uk)
Dim OP_RowNum
If MM_offset <> "" Then
	OP_RowNum = MM_offset + 1
Else
	OP_RowNum = 1
End If
%>
<html>
<head>
<title>Catalog Category Administrator</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link href="../../styles.css" rel="stylesheet" type="text/css">
<script language="JavaScript" type="text/JavaScript">
<!--
function MM_jumpMenu(targ,selObj,restore){ //v3.0
  eval(targ+".location='"+selObj.options[selObj.selectedIndex].value+"'");
  if (restore) selObj.selectedIndex=0;
}

function openPictureWindow_Fever(imageName,imageWidth,imageHeight,alt,posLeft,posTop) {
	newWindow = window.open("","newWindow","width="+imageWidth+",height="+imageHeight+",left="+posLeft+",top="+posTop);
	newWindow.document.open();
	newWindow.document.write('<html><title>'+alt+'</title><body bgcolor="#FFFFFF" leftmargin="0" topmargin="0" marginheight="0" marginwidth="0" onBlur="self.close()">'); 
	newWindow.document.write('<img src='+imageName+' width='+imageWidth+' height='+imageHeight+' alt='+alt+'>'); 
	newWindow.document.write('</body></html>');
	newWindow.document.close();
	newWindow.focus();
}

function MM_openBrWindow(theURL,winName,features) { //v2.0
  window.open(theURL,winName,features);
}
//-->
</script>
</head>
<!--#include file="header.asp" -->
<table width="100%" height="122" border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td><table width="100%" border="0" cellpadding="0" cellspacing="0" class="tableborder">
      <tr>
        <td>To create a Department Category Level, enter  the  name of the
          Department and press the create button.</td>
      </tr>
      <tr>
        <td><form action="<%=MM_editAction%>" method="POST" name="gpc" id="gpc">
              <strong>Create Department:</strong>              
              <input name="GPCName" type="text" id="GPCName" size="25">
          <input type="submit" value="Create" name="submit">
          <input type="hidden" name="MM_insert" value="gpc">
          <input name="CatalogTypeID" type="hidden" id="CatalogTypeID" value="1">
        </form></td>
      </tr>
    </table></td>
  </tr>
  <tr>
    <td><table width="100%" border="0" cellpadding="0" cellspacing="0" class="tableborder">
      <tr>
        <td>To create a Category: Select the Department, 
          enter the name of the Category you wish to create and press the
          create button. </td>
      </tr>
      <tr>
        <td>
		<form action="<%=MM_editAction%>" method="POST" name="category" id="category">
		    <strong>Create Category:</strong>            
		    <select name="GPCID" id="select">
		  <option selected value="">Department</option>
            <%
While (NOT gpcmenu.EOF)
%>
            <option value="<%=(gpcmenu.Fields.Item("GPCID").Value)%>"><%=(gpcmenu.Fields.Item("GPCName").Value)%></option>
            <%
  gpcmenu.MoveNext()
Wend
If (gpcmenu.CursorType > 0) Then
  gpcmenu.MoveFirst
Else
  gpcmenu.Requery
End If
%>
          </select>
          <input name="CategoryName" type="text" id="SubMenu2" size="25">
          <input type="submit" value="Create" name="submit2">
          <input type="hidden" name="MM_insert" value="category">
        </form></td>
      </tr>
    </table></td>
  </tr>
  <tr>
    <td>
      <table width="100%" border="0" cellpadding="0" cellspacing="0" class="tableborder">
        <tr>
          <td>To create a SubCategory: Select the Department, Select the Category,
            enter the name of the SubCategory you wish to create and press the
            create button. </td>
        </tr>
        <tr>
          <td><form action="<%=MM_editAction%>" method="POST" name="subcategory" id="subcategory">
            <strong>Create SubCategory:</strong>            
            <select name="GPCID" id="GPCID" onChange="MM_jumpMenu('parent',this,0)">
			<option selected value="">Department</option>
              <%
While (NOT gpcmenu.EOF)
%>
              <option value="?gpcid=<%=(gpcmenu.Fields.Item("GPCID").Value)%>" <%If (Not isNull(request.querystring("gpcid"))) Then If (CStr(gpcmenu.Fields.Item("GPCID").Value) = CStr(request.querystring("gpcid"))) Then Response.Write("SELECTED") : Response.Write("")%> ><%=(gpcmenu.Fields.Item("GPCName").Value)%></option>
              <%
  gpcmenu.MoveNext()
Wend
If (gpcmenu.CursorType > 0) Then
  gpcmenu.MoveFirst
Else
  gpcmenu.Requery
End If
%>
            </select>
            <select name="CategoryID" id="CategoryID">
			<option selected value="">Category</option>
              <%
While (NOT categorymenu.EOF)
%>
              <option value="<%=(categorymenu.Fields.Item("CategoryID").Value)%>" <%If (Not isNull(request.querystring("cid"))) Then If (CStr(categorymenu.Fields.Item("CategoryID").Value) = CStr(request.querystring("cid"))) Then Response.Write("SELECTED") : Response.Write("")%> ><%=(categorymenu.Fields.Item("CategoryName").Value)%></option>
              <%
  categorymenu.MoveNext()
Wend
If (categorymenu.CursorType > 0) Then
  categorymenu.MoveFirst
Else
  categorymenu.Requery
End If
%>
            </select>
            <input name="SubCategoryName" type="text" id="SubCategoryName" size="25">
            <input name="submit3" type="submit" id="submit3" value="Create">
            <input type="hidden" name="MM_insert" value="subcategory">
            </form>
          </td>
        </tr>
      </table></td>
  </tr>
</table>

<table width="100%" border="0" cellpadding="0" cellspacing="0" class="tableborder">
  <tr valign>
    <td height="81" valign="top">
<% 
While ((Repeat1__numRows <> 0) AND (NOT allcategories.EOF)) 
%>
	  <table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#CCCCCC">
        <tr>
          <td width="50%"><strong>				  			  
                  <% TFM_nest = allcategories.Fields.Item("GPCName").Value
If lastTFM_nest <> TFM_nest Then 
	lastTFM_nest = TFM_nest %>
&nbsp;&nbsp;&nbsp;&nbsp;                  <%=(allcategories.Fields.Item("GPCName").Value)%> 
</strong></td>
          <td width="5%"><div align="center"><a href="update_category_gpc.asp?gpcid=<%=(allcategories.Fields.Item("GPCID").Value)%>">Edit</a></div></td>
          <td width="30%">
		  <%		  						  
Dim objgpcimage
strImage = "../../../applications/CatalogManager/images/" & allcategories.Fields.Item("GPCImageFile").Value
Set objgpcimage = CreateObject("Scripting.FileSystemObject")
If objgpcimage.FileExists(Server.MapPath(strImage)) then
%>
          <% if allcategories.Fields.Item("GPCImageFile").Value <> "" then %>
          <a href="javascript:;"><img src="../../../applications/CatalogManager/images/<%=(allcategories.Fields.Item("GPCImageFile").Value)%>" width="25" border="0" onClick="openPictureWindow_Fever('../../../applications/CatalogManager/images/<%=(allcategories.Fields.Item("GPCImageFile").Value)%>','400','400','<%=(allcategories.Fields.Item("GPCName").Value)%>',')',')')"></a>
          <% end if ' image check %>
          <% else%>
          <a href="javascript:;" onClick="MM_openBrWindow('update_image_gpc.asp?gpcid=<%=(allcategories.Fields.Item("GPCID").Value)%>','image','width=300,height=150')">Add
          Image</a>
          <% end if ' image check %>
          </td>
          <td><strong><a href="delete_gpc.asp?gpcid=<%=(allcategories.Fields.Item("GPCID").Value)%>">Delete</a>
              <%End If 'End Basic-UltraDev Simulated Nested Repeat %>
          </strong></td>
        </tr>
      </table>


	  <table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#F8F8F8">
        <tr>
          <td width="50%"> 
                  <% TFM_nest2 = allcategories.Fields.Item("CategoryName").Value
If lastTFM_nest2 <> TFM_nest2 Then 
	lastTFM_nest2 = TFM_nest2 %>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&gt;&nbsp;&nbsp;&nbsp;&nbsp;<%=(allcategories.Fields.Item("CategoryName").Value)%></td>
          <td width="5%"><div align="center"><a href="update_category_category.asp?cid=<%=(allcategories.Fields.Item("CategoryID").Value)%>">Edit</a></div></td>
          <td width="30%"><%		  						  
Dim objcatimage
strImage = "../../../applications/CatalogManager/images/" & allcategories.Fields.Item("CategoryImageFile").Value
Set objcatimage = CreateObject("Scripting.FileSystemObject")
If objcatimage.FileExists(Server.MapPath(strImage)) then
%>
            <% if allcategories.Fields.Item("CategoryImageFile").Value <> "" then %>
            <a href="javascript:;"><img src="../../../applications/CatalogManager/images/<%=(allcategories.Fields.Item("CategoryImageFile").Value)%>" width="25" border="0" onClick="openPictureWindow_Fever('../../../applications/CatalogManager/images/<%=(allcategories.Fields.Item("CategoryImageFile").Value)%>','400','400','<%=(allcategories.Fields.Item("CategoryName").Value)%>',')',')')"></a>
             <% end if ' image check %>
           <% else%>
            <a href="javascript:;" onClick="MM_openBrWindow('update_image_category.asp?cid=<%=(allcategories.Fields.Item("CategoryID").Value)%>','image','width=300,height=150')">Add
            Image</a>
            <% end if ' image check %>
</td>
          <td><a href="delete_category.asp?cid=<%=(allcategories.Fields.Item("CategoryID").Value)%>">Delete</a>
            <%End If 'End Basic-UltraDev Simulated Nested Repeat %>
</td>
        </tr>
      </table>


      <table width="100%" border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td width="50%"><% if allcategories.Fields.Item("SubCategoryName").Value <>"" then %>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&gt;&nbsp;&nbsp;&nbsp;&nbsp;<%=(allcategories.Fields.Item("SubCategoryName").Value)%></td>
          <td width="5%"><div align="center"><a href="update_category_subcategory.asp?scid=<%=(allcategories.Fields.Item("SubCategoryID").Value)%>">Edit</a></div></td>
          <td width="30%"><%		  						  
Dim objsubcatimage
strImage = "../../../applications/CatalogManager/images/" & allcategories.Fields.Item("SubCategoryImageFile").Value
Set objsubcatimage = CreateObject("Scripting.FileSystemObject")
If objsubcatimage.FileExists(Server.MapPath(strImage)) then
%>
            <% if allcategories.Fields.Item("SubCategoryImageFile").Value <> "" then %>
            <a href="javascript:;"><img src="../../../applications/CatalogManager/images/<%=(allcategories.Fields.Item("SubCategoryImageFile").Value)%>" width="25" border="0" onClick="openPictureWindow_Fever('../../../applications/CatalogManager/images/<%=(allcategories.Fields.Item("SubCategoryImageFile").Value)%>','400','400','<%=(allcategories.Fields.Item("SubCategoryName").Value)%>',')',')')"></a>
            <% end if ' image check %>
            <% else%>
            <a href="javascript:;" onClick="MM_openBrWindow('update_image_subcategory.asp?scid=<%=(allcategories.Fields.Item("SubCategoryID").Value)%>','image','width=300,height=150')">Add
            Image</a>
            <% end if ' image check %>
</td>
          <td><a href="delete_subcategory.asp?scid=<%=(allcategories.Fields.Item("SubCategoryID").Value)%>">Delete </a>
            <%end if%>
</td>
        </tr>
      </table>
      <% 
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  allcategories.MoveNext()
Wend
%>
</td>
  </tr>
</table>

<h5>&nbsp;</h5>
<h3>&nbsp;</h3>

<p>&nbsp;</p>
</body>
</html>
<%
allcategories.Close()
Set allcategories = Nothing
%>
<%
gpcmenu.Close()
%>
<%
categorymenu.Close()
%>
<%
subcategorymenu.Close()
%>
