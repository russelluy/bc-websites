<%
set Category = Server.CreateObject("ADODB.Recordset")
Category.ActiveConnection = MM_catalogmanager_STRING
Category.Source = "SELECT *  FROM tblCatalogCategory  ORDER BY CategoryID"
Category.CursorType = 0
Category.CursorLocation = 2
Category.LockType = 3
Category.Open()
Category_numRows = 0
%>
<%
Dim RepeatCategoryList__numRows
Dim RepeatCategoryList__index

RepeatCategoryList__numRows = -1
RepeatCategoryList__index = 0
Category_numRows = Category_numRows + RepeatCategoryList__numRows
%>
<link href="../../../styles.css" rel="stylesheet" type="text/css">
<Body>
<table width="100%" border="0" cellpadding="0" cellspacing="0" class="tableborder">
  <tr>
    <td><strong>Category</strong></td>
  </tr>
  <tr>
    <td valign="top"><% 
While ((RepeatCategoryList__numRows <> 0) AND (NOT Category.EOF)) 
%>
<a href="../../../SoftwareLibrary/Demo/Catalog/Catalog/<%=request.servervariables("URL")%>?<%=request.servervariables("QUERY_STRING")%>&Search=<%=(Category.Fields.Item("CategoryName").Value)%>"><%=(Category.Fields.Item("CategoryName").Value)%></a><br>
    <% 
  RepeatCategoryList__index=RepeatCategoryList__index+1
  RepeatCategoryList__numRows=RepeatCategoryList__numRows-1
  Category.MoveNext()
Wend
%></td>
  </tr>
</table>
<%
Category.Close()
%>
