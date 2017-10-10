<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="../../Connections/photogallerymanager.asp" -->
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
<%
Dim photogallery_list__MMColParam1
photogallery_list__MMColParam1 = "%"
If (Request.Form("search") <> "") Then 
  photogallery_list__MMColParam1 = Request.Form("search")        
End If
%>
<%
Dim photogallery_list__MMColParam2
photogallery_list__MMColParam2 = "%"
If (Request.QueryString("ItemID")  <> "") Then 
  photogallery_list__MMColParam2 = Request.QueryString("ItemID") 
End If
%>
<%
Dim photogallery_list__MMColParam3
photogallery_list__MMColParam3 = "%"
If (Request.Form("searchcat")  <> "") Then 
  photogallery_list__MMColParam3 = Request.Form("searchcat") 
End If
%>
<%
set photogallery_list = Server.CreateObject("ADODB.Recordset")
photogallery_list.ActiveConnection = MM_photogallerymanager_STRING
photogallery_list.Source = "SELECT tblPhotoGallery.*, tblPhotoGalleryCategory.CategoryDesc, tblPhotoGalleryCategory.CategoryName  FROM tblPhotoGalleryCategory INNER JOIN tblPhotoGallery ON tblPhotoGalleryCategory.CategoryID = tblPhotoGallery.CategoryID  WHERE tblPhotoGalleryCategory.CategoryName Like '" + Replace(photogallery_list__MMColParam3, "'", "''") + "'  AND tblPhotoGallery.ItemID Like '" + Replace(photogallery_list__MMColParam2, "'", "''") + "' AND (tblPhotoGallery.ItemDesc Like '%" + Replace(photogallery_list__MMColParam1, "'", "''") + "%' OR tblPhotoGallery.ItemName Like '%" + Replace(photogallery_list__MMColParam1, "'", "''") + "%' OR tblPhotoGallery.ItemDescShort Like '%" + Replace(photogallery_list__MMColParam1, "'", "''") + "%' )  ORDER BY tblPhotoGallery.CategoryID, DateAdded"
photogallery_list.CursorType = 0
photogallery_list.CursorLocation = 2
photogallery_list.LockType = 3
photogallery_list.Open()
photogallery_list_numRows = 0
%>
<%
Dim Repeat1__numRows
Dim Repeat1__index

Repeat1__numRows = -1
Repeat1__index = 0
photogallery_list_numRows = photogallery_list_numRows + Repeat1__numRows
%>
<html>
<head>
<title>Photo Gallery Controls</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link href="../styles.css" rel="stylesheet" type="text/css">
<script language="JavaScript" type="text/JavaScript">
<!--

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
<SCRIPT RUNAT=SERVER LANGUAGE=VBSCRIPT>	
function DoTrimProperly(str, nNamedFormat, properly, pointed, points)
  dim strRet
  strRet = Server.HTMLEncode(str)
  strRet = replace(strRet, vbcrlf,"")
  strRet = replace(strRet, vbtab,"")
  If (LEN(strRet) > nNamedFormat) Then
    strRet = LEFT(strRet, nNamedFormat)			
    If (properly = 1) Then					
      Dim TempArray								
      TempArray = split(strRet, " ")	
      Dim n
      strRet = ""
      for n = 0 to Ubound(TempArray) - 1
        strRet = strRet & " " & TempArray(n)
      next
    End If
    If (pointed = 1) Then
      strRet = strRet & points
    End If
  End If
  DoTrimProperly = strRet
End Function
</SCRIPT>
<style type="text/css">
<!--
.style1 {font-family: Arial, Helvetica, sans-serif}
.style3 {font-family: Arial, Helvetica, sans-serif; font-size: 12px; }
.style5 {
	font-family: Arial, Helvetica, sans-serif;
	font-size: 12px;
	font-weight: bold;
	color: #663300;
}
.style8 {
	font-family: Arial, Helvetica, sans-serif;
	font-size: 11px;
	font-weight: bold;
}
.style10 {font-family: Arial, Helvetica, sans-serif; font-size: 11px; }
.style12 {font-family: Arial, Helvetica, sans-serif; font-size: 12px; font-weight: bold; color: #666600; }
.style13 {color: #666600}
.style14 {color: #663300}
a:link {
	color: #663300;
}
a:visited {
	color: #663300;
}
a:hover {
	color: #666666;
}
-->
</style>
</head>
<Body>
<!--#include file="header.asp" -->
<table width="100%" border="0" cellspacing="0" cellpadding="0" height="24" class="tableborder">
  <tr>
    <td height="24" width="41%" valign="baseline">
      <form action="" method="post" name="form2" id="form2">
        <div align="left"><span class="style12">Search by Category
              </span>
          <select name="searchcat" id="searchcat" >
            <option selected value="%" <%If (Not isNull(Request.Form("searchcat"))) Then If ("%" = CStr(Request.Form("searchcat"))) Then Response.Write("SELECTED") : Response.Write("")%>>Show
                All</option>
            <%
While (NOT Category.EOF)
%>
            <option value="<%=(Category.Fields.Item("CategoryName").Value)%>" <%If (Not isNull(Request.Form("searchcat"))) Then If (CStr(Category.Fields.Item("CategoryName").Value) = CStr(Request.Form("searchcat"))) Then Response.Write("SELECTED") : Response.Write("")%> ><%=(Category.Fields.Item("CategoryName").Value)%></option>
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
            <input name="submit2" type="submit" value="Go">
        </div>
      </form>
    </td>
    <td height="24" width="59%" valign="baseline">
      <form name="form" method="post" action="">
        <div align="left"><span class="style5"><span class="style13">Search by Keyword</span>          
          <input name="search" type="text" id="search">
          </span>            <input type="submit" value="Go" name="submit">
        <br><br></div>
      </form>
    </td>
  </tr>
</table>
<table width="100%" height="32" border="0" cellpadding="0" cellspacing="0" class="tableborder">
  <tr class="tableheader"> 
  <td colspan="3"><div align="left"><span class="context style14"><strong>Category</strong></span></div></td>
    <td width="16%"><span class="context style14"><strong>Name</strong></span></td>
    <td><span class="context style14"><strong>Description</strong></span></td>
    <td width="21%"><div align="center"><span class="context style14"><strong>Image </strong></span></div></td>
    <td width="12%"><span class="context style14"><strong>Activated</strong></span></td>
    <td width="13%"> <div align="center" class="context"><a href="insert.asp">I<strong>nsert New Item</strong></a></div></td>
  </tr>
  <% 
While ((Repeat1__numRows <> 0) AND (NOT photogallery_list.EOF)) 
%>
  <tr class="<% 
RecordCounter = RecordCounter + 1
If RecordCounter Mod 2 = 1 Then
Response.Write "row1"
Else
Response.Write "row2"
End If
%>">
    <td width="2%">&nbsp;</td>
    <td width="2%" height="13">      <span class="style8">
    <%Response.Write(RecordCounter)
RecordCounter = RecordCounter%>
    .       </span></td>
    <td width="19%" height="13"><span class="style10"><%=(photogallery_list.Fields.Item("CategoryName").Value)%></span></td>
    <td height="13"><span class="style10"><%=(photogallery_list.Fields.Item("ItemName").Value)%> </span></td>
    <td width="15%"><span class="style10">
      <% =(DoTrimProperly((photogallery_list.Fields.Item("ItemDesc").Value), 50, 1, 1, "...")) %>
    </span></td>
    <td height="13">	                
	              <div align="center"><span class="style10">
                  <% if photogallery_list.Fields.Item("ImageThumbFileA").Value <> "" then %>
                  <a href="javascript:;"><img src="../../applications/PhotoGalleryManager/images/<%=(photogallery_list.Fields.Item("ImageFileA").Value)%>" width="50" border="0" onClick="openPictureWindow_Fever('../../applications/PhotoGalleryManager/images/<%=(photogallery_list.Fields.Item("ImageThumbFileA").Value)%>','400','400','<%=(photogallery_list.Fields.Item("ItemName").Value)%>','','')"></a>
                  <% end if ' image check %>  						  
                  <% if photogallery_list.Fields.Item("ImageThumbFileB").Value <> "" then %>
                  <a href="javascript:;"><img src="../../applications/PhotoGalleryManager/images/<%=(photogallery_list.Fields.Item("ImageThumbFileB").Value)%>" width="50" border="0" onClick="openPictureWindow_Fever('../../applications/PhotoGalleryManager/images/<%=(photogallery_list.Fields.Item("ImageFileB").Value)%>','400','400','<%=(photogallery_list.Fields.Item("ItemName").Value)%>',')','')"></a>
                  <% end if ' image check %>
                  <bg><br>
                  </span></div></td>
    <td>      <span class="style10">
      <% If photogallery_list.Fields.Item("Activated").Value = "True" Then %>
      Yes 
      <%else%> 
      No 
      <%end if%>
    </span></td>
    <td height="13">
      <div align="center" class="style10"><a href="update.asp?ItemID=<%=(photogallery_list.Fields.Item("ItemID").Value)%>">Edit</a> | <a href="delete.asp?ItemID=<%=(photogallery_list.Fields.Item("ItemID").Value)%>">Delete</a></div>
    </td>
  </tr>
  <tr class="<% 
RecordCounter = RecordCounter + 1
If RecordCounter Mod 2 = 1 Then
Response.Write "row1"
Else
Response.Write "row2"
End If
%>">
    <td height="13" colspan="8"><hr></td>
  </tr>
  <% 
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  photogallery_list.MoveNext()
Wend
%>

</table>
<br>
<% If photogallery_list.EOF And photogallery_list.BOF Then %>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td><div align="center" class="style5">No Records Found...Please Try Again</div></td>
  </tr>
</table>
<% End If ' end photogallery_list.EOF And photogallery_list.BOF %>

</body>
</html>
<%
Category.Close()
%>
<%
photogallery_list.Close()
Set photogallery_list = Nothing
%>
