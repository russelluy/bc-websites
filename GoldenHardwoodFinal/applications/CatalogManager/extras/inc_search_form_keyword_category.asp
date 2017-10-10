<!--#include file="../../../SoftwareLibrary/Demo/Connections/catalogmanager.asp" -->
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
<link href="../../../styles.css" rel="stylesheet" type="text/css">
<Body>
<table width="100%" height="42" border="0" cellpadding="0" cellspacing="0">
  <tr> 
    <td height="17" width="50%"> 
      <form name="form1" method="post" action="../../../SoftwareLibrary/Demo/Catalog/Catalog/<%=request.servervariables("URL")%>?<%=request.servervariables("QUERY_STRING")%>">
        <div align="center">Search by Category
          <select name="Search" id="Search">
          <option value="%" <%If (Not isNull(Request.Form("Search"))) Then If ("%" = CStr(Request.Form("Search"))) Then Response.Write("SELECTED") : Response.Write("")%>>Show All</option>
          <%
While (NOT Category.EOF)
%>
          <option value="<%=(Category.Fields.Item("CategoryName").Value)%>" <%If (Not isNull(Request.Form("Search"))) Then If (CStr(Category.Fields.Item("CategoryName").Value) = CStr(Request.Form("Search"))) Then Response.Write("SELECTED") : Response.Write("")%> ><%=(Category.Fields.Item("CategoryName").Value)%></option>
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
          <input type="submit" value="Go" name="submit2">
        </div>
      </form>
    </td>
    <td height="17" width="50%"> 
      <form name="form" method="post" action="../../../SoftwareLibrary/Demo/Catalog/Catalog/<%=request.servervariables("URL")%>?<%=request.servervariables("QUERY_STRING")%>">
        <div align="center">Search by Keyword 
          <input type="text" name="Search">
          <input type="submit" value="Go" name="submit">
        </div>
      </form>
    </td>
  </tr>
</table>
<%
Category.Close()
%>
