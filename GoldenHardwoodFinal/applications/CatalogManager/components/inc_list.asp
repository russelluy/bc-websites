<!--#include file="../../../Connections/catalogmanager.asp" -->
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
Dim item_list__value6
item_list__value6 = "%"
If (Request.QueryString("manid")        <> "") Then 
  item_list__value6 = Request.QueryString("manid")       
End If
%>
<%
Dim item_list__value7
item_list__value7 = "%"
If (Request.Form("search") <> "") Then 
  item_list__value7 = Request.Form("search")
End If
%>
<%
Dim item_list__value8
item_list__value8 = "%"
If (Request.QueryString("ItemID")         <> "") Then 
  item_list__value8 = Request.QueryString("ItemID")        
End If
%>
<%
Dim item_list
Dim item_list_numRows

Set item_list = Server.CreateObject("ADODB.Recordset")
item_list.ActiveConnection = MM_catalogmanager_STRING
item_list.Source = "SELECT tblCatalog.*, tblCatalogSubCategory.*, tblCatalogCategory.*, tblGPC.*, tblCatalogDetails.*, tblManufacturers.*  FROM ((((tblCatalogCategory RIGHT JOIN tblCatalogSubCategory ON tblCatalogCategory.CategoryID = tblCatalogSubCategory.CategoryIDkey) RIGHT JOIN tblCatalog ON tblCatalogSubCategory.SubCategoryID = tblCatalog.SubCategoryIDKey) LEFT JOIN tblGPC ON tblCatalogCategory.GPCIDkey = tblGPC.GPCID) LEFT JOIN tblCatalogDetails ON tblCatalog.ItemID = tblCatalogDetails.ItemIDKey) LEFT JOIN tblManufacturers ON tblCatalog.ManufacturerIDkey = tblManufacturers.ManufacturerID  WHERE Activated = 'True' AND CategoryID LIKE '" + Replace(item_list__value2, "'", "''") + "' AND SubCategoryID LIKE '" + Replace(item_list__value4, "'", "''") + "' AND GPCID LIKE '" + Replace(item_list__value5, "'", "''") + "' AND ManufacturerID LIKE '" + Replace(item_list__value6, "'", "''") + "' AND (ItemName LIKE '%" + Replace(item_list__value7, "'", "''") + "%' OR ItemDesc LIKE '%" + Replace(item_list__value7, "'", "''") + "%' OR ItemDescShort LIKE '%" + Replace(item_list__value7, "'", "''") + "%') AND ItemID LIKE '" + Replace(item_list__value8, "'", "''") + "'"
item_list.CursorType = 0
item_list.CursorLocation = 2
item_list.LockType = 1
item_list.Open()

item_list_numRows = 0
%>
<%
Dim Repeat_item_list__numRows
Dim Repeat_item_list__index

Repeat_item_list__numRows = -1
Repeat_item_list__index = 0
item_list_numRows = item_list_numRows + Repeat_item_list__numRows
%>
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
<% If Not item_list.EOF Or Not item_list.BOF Then %>
<table width="100%" border="0" cellspacing="0" cellpadding="2">
          <tr>
          <td height="28" valign="top">
	        <font color="#744900" size="1" face="Arial, Helvetica, sans-serif">
	        <% if request.querystring("gpcid") <> "" then %>
		    <a href="<%=request.servervariables("URL")%><%If Request.QueryString("mid") <> "" Then %>?mid=<%=request.querystring("mid")%><%end if%><%If Request.QueryString ("mid2")<> "" Then %>&mid2=<%=request.querystring("mid2")%><%end if%><%If Request.QueryString ("mid3")<> "" Then %>&mid3=<%=request.querystring("mid3")%><%end if%><%If Request.QueryString ("incid")<> "" Then %>&incid=<%=request.querystring("incid")%><%end if%>">Show all</a> &gt; <a href="<%=request.servervariables("URL")%>?gpcid=<%=(item_list.Fields.Item("GPCID").Value)%><%If Request.QueryString("mid") <> "" Then %>&mid=<%=request.querystring("mid")%><%end if%><%If Request.QueryString ("mid2")<> "" Then %>&mid2=<%=request.querystring("mid2")%><%end if%><%If Request.QueryString ("mid3")<> "" Then %>&mid3=<%=request.querystring("mid3")%><%end if%><%If Request.QueryString ("incid")<> "" Then %>&incid=<%=request.querystring("incid")%><%end if%>"><%=(item_list.Fields.Item("GPCName").Value)%></a> &gt;
            <%end if%>
            <% if request.querystring("cid") <> "" then %>
            <a href="<%=request.servervariables("URL")%>?cid=<%=(item_list.Fields.Item("CategoryID").Value)%>&gpcid=<%=(item_list.Fields.Item("GPCIDkey").Value)%><%If Request.QueryString("mid") <> "" Then %>&mid=<%=request.querystring("mid")%><%end if%><%If Request.QueryString ("mid2")<> "" Then %>&mid2=<%=request.querystring("mid2")%><%end if%><%If Request.QueryString ("mid3")<> "" Then %>&mid3=<%=request.querystring("mid3")%><%end if%><%If Request.QueryString ("incid")<> "" Then %>&incid=<%=request.querystring("incid")%><%end if%>"><%=(item_list.Fields.Item("CategoryName").Value)%></a> &gt;
            <%end if%>
            <% if request.querystring("scid") <> "" then %>
            <a href="<%=request.servervariables("URL")%>?scid=<%=(item_list.Fields.Item("SubCategoryID").Value)%>&gpcid=<%=request.querystring("gpcid")%>&cid=<%=request.querystring("cid")%><%If Request.QueryString("mid") <> "" Then %>&mid=<%=request.querystring("mid")%><%end if%><%If Request.QueryString ("mid2")<> "" Then %>&mid2=<%=request.querystring("mid2")%><%end if%><%If Request.QueryString ("mid3")<> "" Then %>&mid3=<%=request.querystring("mid3")%><%end if%><%If Request.QueryString ("incid")<> "" Then %>&incid=<%=request.querystring("incid")%><%end if%>"><%=(item_list.Fields.Item("SubCategoryName").Value)%></a> &gt;
            <%end if%>
            <% if request.querystring("ItemID") <> "" then %>
            <%=(item_list.Fields.Item("ItemName").Value)%>
            <%end if%>
	        </font></td>
          </tr>
		  
</table>
<% if not request.querystring("ItemID") <> "" then %>
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="0">
  <% 
While ((Repeat_item_list__numRows <> 0) AND (NOT item_list.EOF)) 
%>
  <tr class="<% 
RecordCounter = RecordCounter + 1
If RecordCounter Mod 2 = 1 Then
Response.Write "row1"
Else
Response.Write "row2"
End If
%>">
    <td height="73" valign="top">
      <%		  						  
Dim objthumb
strImage = "/applications/CatalogManager/images/" & item_list.Fields.Item("ImageFileThumb").Value
Set objthumb = CreateObject("Scripting.FileSystemObject")
If objthumb.FileExists(Server.MapPath(strImage)) then
%>
      <% if item_list.Fields.Item("ImageFileThumb").Value <> "" then %>
     <a href="javascript:;"><img src="/applications/CatalogManager/images/<%=(item_list.Fields.Item("ImageFileThumb").Value)%>" alt="Click to Zoom" width="100" border="0" onClick="openPictureWindow_Fever('/applications/CatalogManager/images/<%=(item_list.Fields.Item("ImageFileThumb").Value)%>','400','400','<%=(item_list.Fields.Item("ImageFileThumb").Value)%>','','')"></a>
      <% end if ' image check %>
      <% end if %>
      <%		  						  
Dim objthumb2
strImage = "/applications/CatalogManager/images/" & item_list.Fields.Item("ImageFileThumb2").Value
Set objthumb2 = CreateObject("Scripting.FileSystemObject")
If objthumb2.FileExists(Server.MapPath(strImage)) then
%>
      <% if item_list.Fields.Item("ImageFileThumb2").Value <> "" then %>
           <a href="javascript:;"><img src="/applications/CatalogManager/images/<%=(item_list.Fields.Item("ImageFileThumb2").Value)%>" alt="Click to Zoom" width="100" border="0" onClick="openPictureWindow_Fever('/applications/CatalogManager/images/<%=(item_list.Fields.Item("ImageFileThumb2").Value)%>','400','400','<%=(item_list.Fields.Item("ImageFileThumb2").Value)%>','','')"></a>
      <% end if ' image check %>
      <% end if %>
    </td>
    <td width="90%" height="73" valign="top">
      <p><font size="2"><b><font color="#744900" face="Arial, Helvetica, sans-serif"><%=(item_list.Fields.Item("ItemName").Value)%></font></b> <font color="#744900" face="Arial, Helvetica, sans-serif">|</font><font color="#0066CC" face="Arial, Helvetica, sans-serif"> <a href="<%=request.servervariables("URL")%>?gpcid=<%=(item_list.Fields.Item("GPCID").Value)%>&cid=<%=(item_list.Fields.Item("CategoryID").Value)%>&scid=<%=(item_list.Fields.Item("SubCategoryID").Value)%>&ItemID=<%=(item_list.Fields.Item("ItemID").Value)%><%If Request.QueryString("mid") <> "" Then %>&mid=<%=request.querystring("mid")%><%end if%><%If Request.QueryString ("mid2")<> "" Then %>&mid2=<%=request.querystring("mid2")%><%end if%><%If Request.QueryString ("mid3")<> "" Then %>&mid3=<%=request.querystring("mid3")%><%end if%><%If Request.QueryString ("incid")<> "" Then %>&incid=<%=request.querystring("incid")%><%end if%>">More
                  Detail</a><br>
      <%=(item_list.Fields.Item("ItemDescShort").Value)%></font></font></p>      
      <font face="Arial, Helvetica, sans-serif">
      <% If item_list.Fields.Item("OrderLink").Value <> "" Then %>
      </font>
      <p align="right"><font size="1" face="Arial, Helvetica, sans-serif"> Order
          Now</font></p>      
      <font face="Arial, Helvetica, sans-serif">
      <% end if%>
      </font> </td>
  </tr>
  <% 
  Repeat_item_list__index=Repeat_item_list__index+1
  Repeat_item_list__numRows=Repeat_item_list__numRows-1
  item_list.MoveNext()
Wend
%>
</table>
<%end if%>
<% End If ' end Not item_list.EOF Or NOT item_list.BOF %>
<% If item_list.EOF And item_list.BOF Then %>  
 <table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td>  <div align="center"><font size="2" face="Arial, Helvetica, sans-serif">Sorry....No Records Found...Please <a href="javascript:history.go(-1);">try
          again</a>.</font></div></td>
  </tr>
</table>
<% End If ' end item_list.EOF And item_list.BOF %>

<%
item_list.Close()
Set item_list = Nothing
%>
