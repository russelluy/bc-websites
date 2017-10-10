<!--#include file="../../../Connections/catalogmanager.asp" -->
<%
Dim Category__value2
Category__value2 = "0"
If (Request.QueryString("gpcid")  <> "") Then 
  Category__value2 = Request.QueryString("gpcid") 
End If
%>
<%
set Category = Server.CreateObject("ADODB.Recordset")
Category.ActiveConnection = MM_catalogmanager_STRING
Category.Source = "SELECT tblCatalogCategory.CategoryID, tblGPC.GPCName, tblCatalogCategory.CategoryName, tblCatalogCategory.CategoryDesc, tblCatalogCategory.GPCIDkey, tblCatalogCategory.CategoryImageFile  FROM ((tblCatalogCategory RIGHT JOIN tblCatalogSubCategory ON tblCatalogCategory.CategoryID = tblCatalogSubCategory.CategoryIDkey) RIGHT JOIN tblCatalog ON tblCatalogSubCategory.SubCategoryID = tblCatalog.SubCategoryIDKey) LEFT JOIN tblGPC ON tblCatalogCategory.GPCIDkey = tblGPC.GPCID  WHERE (((tblCatalogCategory.GPCIDkey) Like '" + Replace(Category__value2, "'", "''") + "'))  GROUP BY tblCatalogCategory.CategoryID, tblGPC.GPCName, tblCatalogCategory.CategoryName, tblCatalogCategory.CategoryDesc, tblCatalogCategory.GPCIDkey, tblCatalogCategory.CategoryImageFile  ORDER BY tblCatalogCategory.CategoryID"
Category.CursorType = 0
Category.CursorLocation = 2
Category.LockType = 3
Category.Open()
Category_numRows = 0
%>
<%
Dim SubCategory__value2
SubCategory__value2 = "0"
If (request.querystring("cid")  <> "") Then 
  SubCategory__value2 = request.querystring("cid") 
End If
%>
<%
Dim SubCategory
Dim SubCategory_numRows

Set SubCategory = Server.CreateObject("ADODB.Recordset")
SubCategory.ActiveConnection = MM_catalogmanager_STRING
SubCategory.Source = "SELECT tblCatalogCategory.CategoryName, tblCatalogSubCategory.SubCategoryImageFile, tblCatalogSubCategory.SubCategoryID, tblCatalogSubCategory.SubCategoryName, tblCatalogSubCategory.SubCategoryDesc, tblCatalogSubCategory.CategoryIDkey  FROM ((tblCatalog LEFT JOIN tblCatalogSubCategory ON tblCatalog.SubCategoryIDKey = tblCatalogSubCategory.SubCategoryID) LEFT JOIN tblCatalogCategory ON tblCatalogSubCategory.CategoryIDkey = tblCatalogCategory.CategoryID) LEFT JOIN tblGPC ON tblCatalogCategory.GPCIDkey = tblGPC.GPCID  WHERE (((tblCatalogSubCategory.CategoryIDkey) Like '" + Replace(SubCategory__value2, "'", "''") + "'))  GROUP BY tblCatalogCategory.CategoryName, tblCatalogSubCategory.SubCategoryImageFile, tblCatalogSubCategory.SubCategoryID, tblCatalogSubCategory.SubCategoryName, tblCatalogSubCategory.SubCategoryDesc, tblCatalogSubCategory.CategoryIDkey"
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
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td align="center">
	 <% if Not Request.QueryString("ItemID") <> "" then %>
	 <% if Not Request.QueryString("cid") <> "" then %>
     <% if Not Request.QueryString("scid") <> "" then %>
	 <% if Not Request.QueryString("gpcid") <> "" then %>
     <table>
  <%
startrw = 0
endrw = CatalogHLooper3__index
numberColumns = 3
numrows = -1
while((numrows <> 0) AND (Not GPC.EOF))
	startrw = endrw + 1
	endrw = endrw + numberColumns
 %>
  <tr align="center" valign="top">
    <%
While ((startrw <= endrw) AND (Not GPC.EOF))
%>
    <td>
      <table width="200" border="0" cellspacing="0" cellpadding="0" class="tableborder">
        <tr>
          <td><div align="center"><a href="<%=request.servervariables("URL")%>?gpcid=<%=(GPC.Fields.Item("GPCID").Value)%><%If Request.QueryString("mid") <> "" Then %>&mid=<%=request.querystring("mid")%><%end if%><%If Request.QueryString ("mid2")<> "" Then %>&mid2=<%=request.querystring("mid2")%><%end if%><%If Request.QueryString ("mid3")<> "" Then %>&mid3=<%=request.querystring("mid3")%><%end if%><%If Request.QueryString ("incid")<> "" Then %>&incid=<%=request.querystring("incid")%><%end if%>"><%=(GPC.Fields.Item("GPCName").Value)%></a></div></td>
        </tr>
        <tr>
          <td><div align="center">
                <%		  						  
Dim objgpcimage
strImage = "images/" & GPC.Fields.Item("GPCImageFile").Value
Set objgpcimage = CreateObject("Scripting.FileSystemObject")
If objgpcimage.FileExists(Server.MapPath(strImage)) then
%>
                <% if GPC.Fields.Item("GPCImageFile").Value <> "" then %>
                <a href="<%=request.servervariables("URL")%>?gpcid=<%=(GPC.Fields.Item("GPCID").Value)%><%If Request.QueryString("mid") <> "" Then %>&mid=<%=request.querystring("mid")%><%end if%><%If Request.QueryString ("mid2")<> "" Then %>&mid2=<%=request.querystring("mid2")%><%end if%><%If Request.QueryString ("mid3")<> "" Then %>&mid3=<%=request.querystring("mid3")%><%end if%><%If Request.QueryString ("incid")<> "" Then %>&incid=<%=request.querystring("incid")%><%end if%>"><img src="images/<%=(GPC.Fields.Item("GPCImageFile").Value)%>" width="150" height="150" border="0"></a>
                <% end if ' image check %>
                <% end if ' image check %>
          </div></td>
        </tr>
      </table>
   </td>
    <%
	startrw = startrw + 1
	GPC.MoveNext()
	Wend
	%>
  </tr>
  <%
 numrows=numrows-1
 Wend
 %>
</table>
        <%end if%>
        <%end if%>
		<%end if%>
     <% if Not Request.QueryString("cid") <> "" then %>
     <% if Not Request.QueryString("scid") <> "" then %>
	  <table>
  <%
startrw = 0
endrw = CatalogHLooper1__index
numberColumns = 3
numrows = -1
while((numrows <> 0) AND (Not Category.EOF))
	startrw = endrw + 1
	endrw = endrw + numberColumns
 %>


  <tr align="center" valign="top">
    <%
While ((startrw <= endrw) AND (Not Category.EOF))
%>
    <td>
      <table width="200" height="64" border="0" cellpadding="0" cellspacing="0" class="tableborder">
              <tr>
                <td height="14"><div align="center"><a href="<%=request.servervariables("URL")%>?cid=<%=(Category.Fields.Item("CategoryID").Value)%><%If Request.QueryString("mid") <> "" Then %>&mid=<%=request.querystring("mid")%><%end if%><%If Request.QueryString ("mid2")<> "" Then %>&mid2=<%=request.querystring("mid2")%><%end if%><%If Request.QueryString ("mid3")<> "" Then %>&mid3=<%=request.querystring("mid3")%><%end if%><%If Request.QueryString ("incid")<> "" Then %>&incid=<%=request.querystring("incid")%><%end if%><%If Request.QueryString ("gpcid")<> "" Then %>&gpcid=<%=request.querystring("gpcid")%><%end if%>"> <%=(Category.Fields.Item("CategoryName").Value)%></a></div></td>
              </tr>
              <tr>
                <td valign="top"> <div align="center"><br>
                      <%		  						  
Dim objcatimage
strImage = "images/" & Category.Fields.Item("CategoryImageFile").Value
Set objcatimage = CreateObject("Scripting.FileSystemObject")
If objcatimage.FileExists(Server.MapPath(strImage)) then
%>
                      <% if Category.Fields.Item("CategoryImageFile").Value <> "" then %>
                      <a href="<%=request.servervariables("URL")%>?cid=<%=(Category.Fields.Item("CategoryID").Value)%><%If Request.QueryString("mid") <> "" Then %>&mid=<%=request.querystring("mid")%><%end if%><%If Request.QueryString ("mid2")<> "" Then %>&mid2=<%=request.querystring("mid2")%><%end if%><%If Request.QueryString ("mid3")<> "" Then %>&mid3=<%=request.querystring("mid3")%><%end if%><%If Request.QueryString ("incid")<> "" Then %>&incid=<%=request.querystring("incid")%><%end if%>"><img src="images/<%=(Category.Fields.Item("CategoryImageFile").Value)%>" width="150" height="150" border="0"></a>
                      <% end if ' image check %>
                      <% end if ' image check %>
                </div></td>
              </tr>
      </table>
    </td>
    <%
	startrw = startrw + 1
	Category.MoveNext()
	Wend
	%>
  </tr>
  <%
 numrows=numrows-1
 Wend
 %>
      </table>
	  <%end if%>
	  <%end if%>
<table>
  <%
startrw = 0
endrw = CatalogHLooper2__index
numberColumns = 3
numrows = -1
while((numrows <> 0) AND (Not SubCategory.EOF))
	startrw = endrw + 1
	endrw = endrw + numberColumns
 %>
  <tr align="center" valign="top">
    <%
While ((startrw <= endrw) AND (Not SubCategory.EOF))
%>
    <td>      <table width="200" height="64" border="0" cellpadding="0" cellspacing="0" class="tableborder">
              <tr>
                <td height="14"> <div align="center"><a href="<%=request.servervariables("URL")%>?scid=<%=(SubCategory.Fields.Item("SubCategoryID").Value)%><%If Request.QueryString("mid") <> "" Then %>&mid=<%=request.querystring("mid")%><%end if%><%If Request.QueryString ("mid2")<> "" Then %>&mid2=<%=request.querystring("mid2")%><%end if%><%If Request.QueryString ("mid3")<> "" Then %>&mid3=<%=request.querystring("mid3")%><%end if%><%If Request.QueryString ("incid")<> "" Then %>&incid=<%=request.querystring("incid")%><%end if%><%If Request.QueryString ("gpcid")<> "" Then %>&gpcid=<%=request.querystring("gpcid")%><%end if%><%If Request.QueryString ("cid")<> "" Then %>&cid=<%=request.querystring("cid")%><%end if%>"> <%=(SubCategory.Fields.Item("SubCategoryName").Value)%></a></div></td>
              </tr>
              <tr>
                <td valign="top"> <div align="center"><br>
                    <%		  						  
Dim objsubcatimage
strImage = "images/" & SubCategory.Fields.Item("SubCategoryImageFile").Value
Set objsubcatimage = CreateObject("Scripting.FileSystemObject")
If objsubcatimage.FileExists(Server.MapPath(strImage)) then
%>
                    <% if SubCategory.Fields.Item("SubCategoryImageFile").Value <> "" then %>
                      <a href="<%=request.servervariables("URL")%>?scid=<%=(SubCategory.Fields.Item("SubCategoryID").Value)%><%If Request.QueryString("mid") <> "" Then %>&mid=<%=request.querystring("mid")%><%end if%><%If Request.QueryString ("mid2")<> "" Then %>&mid2=<%=request.querystring("mid2")%><%end if%><%If Request.QueryString ("mid3")<> "" Then %>&mid3=<%=request.querystring("mid3")%><%end if%><%If Request.QueryString ("incid")<> "" Then %>&incid=<%=request.querystring("incid")%><%end if%>"><img src="images/<%=(SubCategory.Fields.Item("SubCategoryImageFile").Value)%>" width="150" height="150" border="0"></a>
                    <% end if ' image check %>
                    <% end if ' image check %>
                </div></td>
              </tr>
            </table>
    </td>
    <%
	startrw = startrw + 1
	SubCategory.MoveNext()
	Wend
	%>
  <%
 numrows=numrows-1
 Wend
 %>
</table>
<% end if%>
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
