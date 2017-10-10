<!--#include file="../../../Connections/catalogmanager.asp" -->
<%
Dim Category__value1
Category__value1 = "0"
If (Request.QueryString("gpcid")   <> "") Then 
  Category__value1 = Request.QueryString("gpcid")  
End If
%>
<%
set Category = Server.CreateObject("ADODB.Recordset")
Category.ActiveConnection = MM_catalogmanager_STRING
Category.Source = "SELECT tblCatalogCategory.CategoryID, tblGPC.GPCName, tblCatalogCategory.CategoryName, tblCatalogCategory.CategoryDesc, tblCatalogCategory.GPCIDkey, tblCatalogCategory.CategoryImageFile  FROM ((tblCatalogCategory RIGHT JOIN tblCatalogSubCategory ON tblCatalogCategory.CategoryID = tblCatalogSubCategory.CategoryIDkey) RIGHT JOIN tblCatalog ON tblCatalogSubCategory.SubCategoryID = tblCatalog.SubCategoryIDKey) LEFT JOIN tblGPC ON tblCatalogCategory.GPCIDkey = tblGPC.GPCID  WHERE (((tblCatalogCategory.GPCIDkey) Like '" + Replace(Category__value1, "'", "''") + "'))  GROUP BY tblCatalogCategory.CategoryID, tblGPC.GPCName, tblCatalogCategory.CategoryName, tblCatalogCategory.CategoryDesc, tblCatalogCategory.GPCIDkey, tblCatalogCategory.CategoryImageFile  ORDER BY tblCatalogCategory.CategoryID"
Category.CursorType = 0
Category.CursorLocation = 2
Category.LockType = 3
Category.Open()
Category_numRows = 0
%>
<%
Dim SubCategory__value1
SubCategory__value1 = "0"
If (request.querystring("cid")  <> "") Then 
  SubCategory__value1 = request.querystring("cid") 
End If
%>
<%
Dim SubCategory
Dim SubCategory_numRows

Set SubCategory = Server.CreateObject("ADODB.Recordset")
SubCategory.ActiveConnection = MM_catalogmanager_STRING
SubCategory.Source = "SELECT tblCatalogCategory.CategoryName, tblCatalogSubCategory.SubCategoryImageFile, tblCatalogSubCategory.SubCategoryID, tblCatalogSubCategory.SubCategoryName, tblCatalogSubCategory.SubCategoryDesc, tblCatalogSubCategory.CategoryIDkey  FROM ((tblCatalog LEFT JOIN tblCatalogSubCategory ON tblCatalog.SubCategoryIDKey = tblCatalogSubCategory.SubCategoryID) LEFT JOIN tblCatalogCategory ON tblCatalogSubCategory.CategoryIDkey = tblCatalogCategory.CategoryID) LEFT JOIN tblGPC ON tblCatalogCategory.GPCIDkey = tblGPC.GPCID  WHERE (((tblCatalogSubCategory.CategoryIDkey) Like '" + Replace(SubCategory__value1, "'", "''") + "'))  GROUP BY tblCatalogCategory.CategoryName, tblCatalogSubCategory.SubCategoryImageFile, tblCatalogSubCategory.SubCategoryID, tblCatalogSubCategory.SubCategoryName, tblCatalogSubCategory.SubCategoryDesc, tblCatalogSubCategory.CategoryIDkey"
SubCategory.CursorType = 0
SubCategory.CursorLocation = 2
SubCategory.LockType = 1
SubCategory.Open()

SubCategory_numRows = 0
%>
<%
Dim GPC
Dim GPC_numRows

Set GPC = Server.CreateObject("ADODB.Recordset")
GPC.ActiveConnection = MM_catalogmanager_STRING
GPC.Source = "SELECT tblGPC.GPCID, tblGPC.GPCName, tblGPC.GPCDesc, tblGPC.GPCImageFile  FROM ((tblGPC INNER JOIN tblCatalogCategory ON tblGPC.GPCID = tblCatalogCategory.GPCIDkey) INNER JOIN tblCatalogSubCategory ON tblCatalogCategory.CategoryID = tblCatalogSubCategory.CategoryIDkey) INNER JOIN tblCatalog ON tblCatalogSubCategory.SubCategoryID = tblCatalog.SubCategoryIDKey  GROUP BY tblGPC.GPCID, tblGPC.GPCName, tblGPC.GPCDesc, tblGPC.GPCImageFile"
GPC.CursorType = 0
GPC.CursorLocation = 2
GPC.LockType = 1
GPC.Open()

GPC_numRows = 0
%>
<script language="JavaScript" type="text/JavaScript">
<!--
function MM_jumpMenu(targ,selObj,restore){ //v3.0
  eval(targ+".location='"+selObj.options[selObj.selectedIndex].value+"'");
  if (restore) selObj.selectedIndex=0;
}
//-->
</script>
<table width="100%" height="37" border="0" cellpadding="0" cellspacing="0" class="tableborder">
  <tr> 
    <td width="75%" height="37"> 
      <div align="left">Search by Department
            <select name="menu1" onChange="MM_jumpMenu('parent',this,0)">
              <option selected value="?gpcid=%<%If Request.QueryString("mid") <> "" Then %>&mid=<%=request.querystring("mid")%><%end if%><%If Request.QueryString ("mid2")<> "" Then %>&mid2=<%=request.querystring("mid2")%><%end if%><%If Request.QueryString ("mid3")<> "" Then %>&mid3=<%=request.querystring("mid3")%><%end if%><%If Request.QueryString ("incid")<> "" Then %>?incid=<%=request.querystring("incid")%><%end if%>">Show All</option>
              <%
While (NOT GPC.EOF)
%>
              <option value="?gpcid=<%=(GPC.Fields.Item("GPCID").Value)%><%If Request.QueryString("mid") <> "" Then %>&mid=<%=request.querystring("mid")%><%end if%><%If Request.QueryString ("mid2")<> "" Then %>&mid2=<%=request.querystring("mid2")%><%end if%><%If Request.QueryString ("mid3")<> "" Then %>&mid3=<%=request.querystring("mid3")%><%end if%><%If Request.QueryString ("incid")<> "" Then %>&amp;incid=<%=request.querystring("incid")%><%end if%>" <%If (Not isNull(Request.QueryString("gpcid"))) Then If (CStr(GPC.Fields.Item("GPCID").Value) = CStr(Request.QueryString("gpcid"))) Then Response.Write("SELECTED") : Response.Write("")%> ><%=(GPC.Fields.Item("GPCName").Value)%></option>
              <%
  GPC.MoveNext()
Wend
If (GPC.CursorType > 0) Then
  GPC.MoveFirst
Else
  GPC.Requery
End If
%>
            </select>
          <% If Not Category.EOF Or Not Category.BOF Then %>
&nbsp;&nbsp;&nbsp;Search by Category
<select name="menu2" onChange="MM_jumpMenu('parent',this,0)">
    <option selected value="?gpcid=%<%If Request.QueryString("mid") <> "" Then %>&mid=<%=request.querystring("mid")%><%end if%><%If Request.QueryString ("mid2")<> "" Then %>&mid2=<%=request.querystring("mid2")%><%end if%><%If Request.QueryString ("mid3")<> "" Then %>&mid3=<%=request.querystring("mid3")%><%end if%><%If Request.QueryString ("incid")<> "" Then %>&incid=<%=request.querystring("incid")%><%end if%>">Show
    All</option>
    <%
While (NOT Category.EOF)
%>
    <option value="?gpcid=<%=Request.QueryString("gpcid")%>&amp;cid=<%=(Category.Fields.Item("CategoryID").Value)%><%If Request.QueryString("mid") <> "" Then %>&mid=<%=request.querystring("mid")%><%end if%><%If Request.QueryString ("mid2")<> "" Then %>&mid2=<%=request.querystring("mid2")%><%end if%><%If Request.QueryString ("mid3")<> "" Then %>&mid3=<%=request.querystring("mid3")%><%end if%><%If Request.QueryString ("incid")<> "" Then %>&amp;incid=<%=request.querystring("incid")%><%end if%>" <%If (Not isNull(Request.QueryString("cid"))) Then If (CStr(Category.Fields.Item("CategoryID").Value) = CStr(Request.QueryString("cid"))) Then Response.Write("SELECTED") : Response.Write("")%> ><%=(Category.Fields.Item("CategoryName").Value)%></option>
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
<% End If ' end Not Category.EOF Or NOT Category.BOF %>
          <% If Not SubCategory.EOF Or Not SubCategory.BOF Then %>
&nbsp;&nbsp;&nbsp;Search by Subcategory
<select name="menu3" onChange="MM_jumpMenu('parent',this,0)">
          <option selected value="?gpcid=%<%If Request.QueryString("mid") <> "" Then %>&mid=<%=request.querystring("mid")%><%end if%><%If Request.QueryString ("mid2")<> "" Then %>&mid2=<%=request.querystring("mid2")%><%end if%><%If Request.QueryString ("mid3")<> "" Then %>&mid3=<%=request.querystring("mid3")%><%end if%><%If Request.QueryString ("incid")<> "" Then %>&incid=<%=request.querystring("incid")%><%end if%>">Show All</option>
          <%
While (NOT SubCategory.EOF)
%>
          <option value="?gpcid=<%=Request.QueryString("gpcid")%>&amp;cid=<%=Request.QueryString("cid")%>&amp;scid=<%=(SubCategory.Fields.Item("SubCategoryID").Value)%><%If Request.QueryString("mid") <> "" Then %>&mid=<%=request.querystring("mid")%><%end if%><%If Request.QueryString ("mid2")<> "" Then %>&mid2=<%=request.querystring("mid2")%><%end if%><%If Request.QueryString ("mid3")<> "" Then %>&mid3=<%=request.querystring("mid3")%><%end if%><%If Request.QueryString ("incid")<> "" Then %>&amp;incid=<%=request.querystring("incid")%><%end if%>" <%If (Not isNull(Request.QueryString("scid"))) Then If (CStr(SubCategory.Fields.Item("SubCategoryID").Value) = CStr(Request.QueryString("scid")))  Then Response.Write("SELECTED") : Response.Write("")%> ><%=(SubCategory.Fields.Item("SubCategoryName").Value)%></option>
          <%
  SubCategory.MoveNext()
Wend
If (SubCategory.CursorType > 0) Then
  SubCategory.MoveFirst
Else
  SubCategory.Requery
End If
%>
        </select>
<% End If ' end Not SubCategory.EOF Or NOT SubCategory.BOF %>
        </form>
          </div></td>
    <td height="37" width="30%"> 
      <form name="form" method="post" action="">
        <div align="right">Search by Keyword 
          <input name="search" type="text" id="search" size="15">
          <input type="submit" value="Go" name="submit">
        </div>
      </form>
    </td>
  </tr>
</table>
<%
Category.Close()
Set Category = Nothing
%>
<%
SubCategory.Close()
Set SubCategory = Nothing
%>
<%
GPC.Close()
Set GPC = Nothing
%>
