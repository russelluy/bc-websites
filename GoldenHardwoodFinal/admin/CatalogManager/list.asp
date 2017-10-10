<!--#include file="../../Connections/catalogmanager.asp" -->
<%
Dim category__value1
category__value1 = "%"
If (Request.QueryString("gpcid")  <> "") Then 
  category__value1 = Request.QueryString("gpcid") 
End If
%>
<%
set category = Server.CreateObject("ADODB.Recordset")
category.ActiveConnection = MM_catalogmanager_STRING
category.Source = "SELECT tblGPC.GPCID, tblCatalogCategory.CategoryID, tblCatalogCategory.CategoryName, tblCatalogCategory.GPCIDkey, tblCatalogCategory.CategoryDesc, tblCatalogCategory.CategoryImageFile  FROM ((tblGPC INNER JOIN tblCatalogCategory ON tblGPC.GPCID = tblCatalogCategory.GPCIDkey) INNER JOIN tblCatalogSubCategory ON tblCatalogCategory.CategoryID = tblCatalogSubCategory.CategoryIDkey) INNER JOIN tblCatalog ON tblCatalogSubCategory.SubCategoryID = tblCatalog.SubCategoryIDKey  WHERE (((tblCatalogCategory.GPCIDkey) LIKE '" + Replace(category__value1, "'", "''") + "'))  GROUP BY tblGPC.GPCID, tblCatalogCategory.CategoryID, tblCatalogCategory.CategoryName, tblCatalogCategory.GPCIDkey, tblCatalogCategory.CategoryDesc, tblCatalogCategory.CategoryImageFile"
category.CursorType = 0
category.CursorLocation = 2
category.LockType = 3
category.Open()
category_numRows = 0
%>
<%
set gpcmenu = Server.CreateObject("ADODB.Recordset")
gpcmenu.ActiveConnection = MM_catalogmanager_STRING
gpcmenu.Source = "SELECT tblGPC.GPCID, tblGPC.GPCName, tblGPC.GPCDesc, tblGPC.GPCImageFile  FROM ((tblGPC INNER JOIN tblCatalogCategory ON tblGPC.GPCID = tblCatalogCategory.GPCIDkey) INNER JOIN tblCatalogSubCategory ON tblCatalogCategory.CategoryID = tblCatalogSubCategory.CategoryIDkey) INNER JOIN tblCatalog ON tblCatalogSubCategory.SubCategoryID = tblCatalog.SubCategoryIDKey  GROUP BY tblGPC.GPCID, tblGPC.GPCName, tblGPC.GPCDesc, tblGPC.GPCImageFile"
gpcmenu.CursorType = 0
gpcmenu.CursorLocation = 2
gpcmenu.LockType = 3
gpcmenu.Open()
gpcmenu_numRows = 0
%>
<%
Dim subcategorymenu__value1
subcategorymenu__value1 = "2"
If (request.querystring("cid")     <> "") Then 
  subcategorymenu__value1 = request.querystring("cid")    
End If
%>
<%
set subcategorymenu = Server.CreateObject("ADODB.Recordset")
subcategorymenu.ActiveConnection = MM_catalogmanager_STRING
subcategorymenu.Source = "SELECT tblCatalogSubCategory.SubCategoryImageFile, tblCatalogSubCategory.SubCategoryID, tblCatalogSubCategory.SubCategoryName, tblCatalogSubCategory.SubCategoryDesc  FROM (((tblGPC RIGHT JOIN tblCatalogCategory ON tblGPC.GPCID = tblCatalogCategory.GPCIDkey) RIGHT JOIN tblCatalogSubCategory ON tblCatalogCategory.CategoryID = tblCatalogSubCategory.CategoryIDkey) RIGHT JOIN tblCatalog ON tblCatalogSubCategory.SubCategoryID = tblCatalog.SubCategoryIDKey) LEFT JOIN tblCatalogDetails ON tblCatalog.ItemID = tblCatalogDetails.ItemIDKey  WHERE (((tblCatalogSubCategory.CategoryIDkey) LIKE '" + Replace(subcategorymenu__value1, "'", "''") + "'))  GROUP BY tblCatalogSubCategory.SubCategoryImageFile, tblCatalogSubCategory.SubCategoryID, tblCatalogSubCategory.SubCategoryName, tblCatalogSubCategory.SubCategoryDesc"
subcategorymenu.CursorType = 0
subcategorymenu.CursorLocation = 2
subcategorymenu.LockType = 3
subcategorymenu.Open()
subcategorymenu_numRows = 0
%>
<%
Dim item_list__value2
item_list__value2 = "%"
If (Request.QueryString("cid")      <> "") Then 
  item_list__value2 = Request.QueryString("cid")     
End If
%>
<%
Dim item_list__value4
item_list__value4 = "%"
If (Request.QueryString("scid")      <> "") Then 
  item_list__value4 = Request.QueryString("scid")     
End If
%>
<%
Dim item_list__value5
item_list__value5 = "%"
If (Request.QueryString("gpcid")      <> "") Then 
  item_list__value5 = Request.QueryString("gpcid")     
End If
%>
<%
Dim item_list__MMColParam1
item_list__MMColParam1 = "%"
If (Request.Form("search") <> "") Then 
  item_list__MMColParam1 = Request.Form("search")
End If
%>
<%
Dim item_list
Dim item_list_numRows

Set item_list = Server.CreateObject("ADODB.Recordset")
item_list.ActiveConnection = MM_catalogmanager_STRING
item_list.Source = "SELECT tblCatalog.*, tblCatalogSubCategory.*, tblCatalogCategory.*, tblGPC.*, tblCatalogDetails.*, tblManufacturers.*  FROM ((((tblCatalogCategory RIGHT JOIN tblCatalogSubCategory ON tblCatalogCategory.CategoryID = tblCatalogSubCategory.CategoryIDkey) RIGHT JOIN tblCatalog ON tblCatalogSubCategory.SubCategoryID = tblCatalog.SubCategoryIDKey) LEFT JOIN tblGPC ON tblCatalogCategory.GPCIDkey = tblGPC.GPCID) LEFT JOIN tblCatalogDetails ON tblCatalog.ItemID = tblCatalogDetails.ItemIDKey) LEFT JOIN tblManufacturers ON tblCatalog.ManufacturerIDkey = tblManufacturers.ManufacturerID  WHERE CategoryID LIKE '" + Replace(item_list__value2, "'", "''") + "' AND SubCategoryID LIKE '" + Replace(item_list__value4, "'", "''") + "' AND GPCID LIKE '" + Replace(item_list__value5, "'", "''") + "'  AND  (ItemDesc Like '%" + Replace(item_list__MMColParam1, "'", "''") + "%' OR ItemName Like '%" + Replace(item_list__MMColParam1, "'", "''") + "%' OR Manufacturer Like '%" + Replace(item_list__MMColParam1, "'", "''") + "%')"
item_list.CursorType = 0
item_list.CursorLocation = 2
item_list.LockType = 1
item_list.Open()

item_list_numRows = 0
%>
<%
Dim Repeat1__numRows
Dim Repeat1__index

Repeat1__numRows = -1
Repeat1__index = 0
item_list_numRows = item_list_numRows + Repeat1__numRows
%>
<html>
<head>
<title>Catalog Manager</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link href="../styles.css" rel="stylesheet" type="text/css">
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
//-->
</script>

</head>
<Body>
<!--#include file="header.asp" -->
<table width="100%" height="42" border="0" cellpadding="0" cellspacing="0" class="tableborder">
  <tr> 
    <td width="70%" height="17"> 
      <form name="form1" method="post" action="">
        <div align="center">
          <% If Not gpcmenu.EOF Or Not gpcmenu.BOF Then %>
Search by Department
<select name="menu1" onChange="MM_jumpMenu('parent',this,0)">
  <option selected value="?gpcid=%25">Show All</option>
  <%
While (NOT gpcmenu.EOF)
%>
  <option value="?gpcid=<%=(gpcmenu.Fields.Item("GPCID").Value)%>" <%If (Not isNull(Request.QueryString("gpcid"))) Then If (CStr(gpcmenu.Fields.Item("GPCID").Value) = CStr(Request.QueryString("gpcid"))) Then Response.Write("SELECTED") : Response.Write("")%> ><%=(gpcmenu.Fields.Item("GPCName").Value)%></option>
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
<% End If ' end Not gpcmenu.EOF Or NOT gpcmenu.BOF %>
          <% If Not category.EOF Or Not category.BOF Then %>
Search by Category
<select name="menu2" onChange="MM_jumpMenu('parent',this,0)">
  <option selected value="?gpcid=%25">Show All</option>
  <%
While (NOT category.EOF)
%>
  <option value="?gpcid=<%=Request.QueryString("gpcid")%>&cid=<%=(category.Fields.Item("CategoryID").Value)%>" <%If (Not isNull(Request.QueryString("cid"))) Then If (CStr(category.Fields.Item("CategoryID").Value) = CStr(Request.QueryString("cid"))) Then Response.Write("SELECTED") : Response.Write("")%> ><%=(category.Fields.Item("CategoryName").Value)%></option>
  <%
  category.MoveNext()
Wend
If (category.CursorType > 0) Then
  category.MoveFirst
Else
  category.Requery
End If
%>
</select>
<% End If ' end Not category.EOF Or NOT category.BOF %>
          <% If Not subcategorymenu.EOF Or Not subcategorymenu.BOF Then %>
Search by Subcategory
<select name="menu3" onChange="MM_jumpMenu('parent',this,0)">
          <option selected value="?gpcid=%25" >Show All</option>
          <%
While (NOT subcategorymenu.EOF)
%>
          <option value="?gpcid=<%=Request.QueryString("gpcid")%>&cid=<%=Request.QueryString("cid")%>&amp;scid=<%=(subcategorymenu.Fields.Item("SubCategoryID").Value)%>" <%If (Not isNull(Request.QueryString("scid"))) Then If (CStr(subcategorymenu.Fields.Item("SubCategoryID").Value) = CStr(Request.QueryString("scid")))  Then Response.Write("SELECTED") : Response.Write("")%> ><%=(subcategorymenu.Fields.Item("SubCategoryName").Value)%></option>
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
<% End If ' end Not subcategorymenu.EOF Or NOT subcategorymenu.BOF %>
        </div>
      </form>
    </td>
    <td height="17" width="30%"> 
      <form name="form" method="post" action="">
        <div align="center">Search by Keyword 
          <input name="search" type="text" id="search">
          <input type="submit" value="Go" name="submit">
        </div>
      </form>
    </td>
  </tr>
</table>

<table width="100%" height="32" border="0" cellpadding="0" cellspacing="0" class="tableborder">
  <tr class="tableheader"> 
  <td colspan="2">Department</td>
    <td width="12%">Category</td>
    <td width="16%">Sub Category</td>
    <td width="16%">Name</td>
    <td width="10%">Image</td>
    <td width="27%"> <div align="center"><a href="insert.asp">Insert New Item</a></div></td>
  </tr>
  <% 
While ((Repeat1__numRows <> 0) AND (NOT item_list.EOF)) 
%>
  <tr class="<% 
RecordCounter = RecordCounter + 1
If RecordCounter Mod 2 = 1 Then
Response.Write "row1"
Else
Response.Write "row2"
End If
%>">
    <td width="2%" height="13">      <strong>
    <%Response.Write(RecordCounter)
RecordCounter = RecordCounter%>.      </strong>   </td>
    <td width="17%" height="13"><%=(item_list.Fields.Item("GPCName").Value)%></td>
    <td width="12%" height="13"><%=(item_list.Fields.Item("CategoryName").Value)%></td>
    <td width="16%" height="13"><%=(item_list.Fields.Item("SubCategoryName").Value)%></td>
    <td height="13"><%=(item_list.Fields.Item("ItemName").Value)%> </td>
    <td height="13">	                
	<%		  						  
Dim objthumb
strImage = "../../applications/CatalogManager/images/" & item_list.Fields.Item("ImageFileThumb").Value
Set objthumb = CreateObject("Scripting.FileSystemObject")
If objthumb.FileExists(Server.MapPath(strImage)) then
%>	              <% if item_list.Fields.Item("ImageFileThumb").Value <> "" then %>
                <a href="javascript:;"><img src="../../applications/CatalogManager/images/<%=(item_list.Fields.Item("ImageFileThumb").Value)%>" width="50" border="0" onClick="openPictureWindow_Fever('../../applications/CatalogManager/images/<%=(item_list.Fields.Item("ImageFile").Value)%>','400','400','<%=(item_list.Fields.Item("ItemName").Value)%>','','')"></a>
                <% end if ' image check %>
				<% end if ' image check %>
	</td>
    <td width="27%" height="13">
      <div align="center"><a href="update.asp?ItemID=<%=(item_list.Fields.Item("ItemID").Value)%>">Edit</a> | <a href="delete.asp?ItemID=<%=(item_list.Fields.Item("ItemID").Value)%>">Delete</a></div>
    </td>
  </tr>
  <% 
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  item_list.MoveNext()
Wend
%>

</table>
<br>
<% If item_list.EOF And item_list.BOF Then %>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td><div align="center">No Records Found...Please Try Again</div></td>
  </tr>
</table>
<% End If ' end item_list.EOF And item_list.BOF %>

</body>
</html>
<%
gpcmenu.Close()
%>
<%
subcategorymenu.Close()
%>
<%
item_list.Close()
Set item_list = Nothing
%>
<%
category.Close()
Set category = Nothing
%>
