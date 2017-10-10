<!--#include file="../../../Connections/catalogmanager.asp" -->
<%
Dim SubCategory__value1
SubCategory__value1 = "4"
If (request.querystring("cid")  <> "") Then 
  SubCategory__value1 = request.querystring("cid") 
End If
%>
<%
Dim SubCategory
Dim SubCategory_numRows

Set SubCategory = Server.CreateObject("ADODB.Recordset")
SubCategory.ActiveConnection = MM_catalogmanager_STRING
SubCategory.Source = "SELECT tblCatalogCategory.CategoryName, tblCatalogCategory.CategoryDesc, tblCatalogCategory.CategoryImageFile, tblCatalogSubCategory.SubCategoryImageFile, tblCatalogSubCategory.SubCategoryID, tblCatalogSubCategory.SubCategoryName, tblCatalogSubCategory.SubCategoryDesc, tblCatalogCategory.GPCIDkey, tblCatalogSubCategory.CategoryIDkey  FROM ((tblCatalog LEFT JOIN tblCatalogSubCategory ON tblCatalog.SubCategoryIDKey = tblCatalogSubCategory.SubCategoryID) LEFT JOIN tblCatalogCategory ON tblCatalogSubCategory.CategoryIDkey = tblCatalogCategory.CategoryID) LEFT JOIN tblGPC ON tblCatalogCategory.GPCIDkey = tblGPC.GPCID  WHERE (((tblCatalogSubCategory.CategoryIDkey) Like '" + Replace(SubCategory__value1, "'", "''") + "'))  GROUP BY tblCatalogCategory.CategoryName, tblCatalogCategory.CategoryDesc, tblCatalogCategory.CategoryImageFile, tblCatalogSubCategory.SubCategoryImageFile, tblCatalogSubCategory.SubCategoryID, tblCatalogSubCategory.SubCategoryName, tblCatalogSubCategory.SubCategoryDesc, tblCatalogCategory.GPCIDkey, tblCatalogSubCategory.CategoryIDkey"
SubCategory.CursorType = 0
SubCategory.CursorLocation = 2
SubCategory.LockType = 1
SubCategory.Open()

SubCategory_numRows = 0
%>
<script language="JavaScript" type="text/JavaScript">
<!--
function MM_jumpMenu(targ,selObj,restore){ //v3.0
  eval(targ+".location='"+selObj.options[selObj.selectedIndex].value+"'");
  if (restore) selObj.selectedIndex=0;
}
//-->
</script>
<% if not request.querystring("ItemID") <> "" then %>
<% If Not SubCategory.EOF Or Not SubCategory.BOF Then %>
<table width="100%" height="82" border="0" cellpadding="0" cellspacing="0" bgcolor="#E4E4E4" class="tableborder">
  <tr> 
    <td width="13%" height="82" valign="top" bgcolor="#FFFFFF"> 
      <div align="left">   
<%		  						  
Dim objsubcategoryimage
strImage = "images/" & SubCategory.Fields.Item("CategoryImageFile").Value
Set objsubcategoryimage = CreateObject("Scripting.FileSystemObject")
If objsubcategoryimage.FileExists(Server.MapPath(strImage)) then
%>
        <% if SubCategory.Fields.Item("CategoryImageFile").Value <> "" then %>
        <img src="images/<%=(SubCategory.Fields.Item("CategoryImageFile").Value)%>">
        <% end if ' image check %>
        <% end if %>
      </div>
    </td>
    <td height="82" width="87%">
      <h3><%=(SubCategory.Fields.Item("CategoryName").Value)%></h3>
      <p><%=(SubCategory.Fields.Item("CategoryDesc").Value)%></p>
      <% If Not SubCategory.EOF Or Not SubCategory.BOF Then %>
&nbsp;&nbsp;&nbsp;Search by Subcategory
      <select name="menu3" onChange="MM_jumpMenu('parent',this,0)">
          <option value="?gpcid=%&cid=<%=Request.QueryString("cid")%>&show=all<%If Request.QueryString("mid") <> "" Then %>&mid=<%=request.querystring("mid")%><%end if%><%If Request.QueryString ("mid2")<> "" Then %>&mid2=<%=request.querystring("mid2")%><%end if%><%If Request.QueryString ("mid3")<> "" Then %>&mid3=<%=request.querystring("mid3")%><%end if%><%If Request.QueryString ("incid")<> "" Then %>&incid=<%=request.querystring("incid")%><%end if%>">Show All</option>
          <%
While (NOT SubCategory.EOF)
%>
          <option value="?gpcid=<%=(SubCategory.Fields.Item("GPCIDkey").Value)%>&cid=<%=(SubCategory.Fields.Item("CategoryIDkey").Value)%>&scid=<%=(SubCategory.Fields.Item("SubCategoryID").Value)%><%If Request.QueryString("mid") <> "" Then %>&mid=<%=request.querystring("mid")%><%end if%><%If Request.QueryString ("mid2")<> "" Then %>&mid2=<%=request.querystring("mid2")%><%end if%><%If Request.QueryString ("mid3")<> "" Then %>&mid3=<%=request.querystring("mid3")%><%end if%><%If Request.QueryString ("incid")<> "" Then %>&incid=<%=request.querystring("incid")%><%end if%>" <%If (Not isNull(Request.QueryString("scid"))) Then If (CStr(SubCategory.Fields.Item("SubCategoryID").Value) = CStr(Request.QueryString("scid")))  Then Response.Write("SELECTED") : Response.Write("")%> ><%=(SubCategory.Fields.Item("SubCategoryName").Value)%></option>
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
    </td>
  </tr>
</table>
<% End If ' end Not SubCategory.EOF Or NOT SubCategory.BOF %>
<% If SubCategory.EOF And SubCategory.BOF Then %>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td>  <div align="center">Sorry....No Records Found...Please <a href="javascript:history.go(-1);">try
          again</a>.</div></td>
  </tr>
</table>
<% End If ' end SubCategory.EOF And SubCategory.BOF %>
<% end if%>
<%
SubCategory.Close()
Set SubCategory = Nothing
%>
