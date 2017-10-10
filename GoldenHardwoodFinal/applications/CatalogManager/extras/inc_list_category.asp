<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="../../../SoftwareLibrary/Demo/Connections/catalogmanager.asp" -->
<%
Dim Category__value1
Category__value1 = "0"
If (Request.QueryString("gpcid")    <> "") Then 
  Category__value1 = Request.QueryString("gpcid")   
End If
%>
<%
set Category = Server.CreateObject("ADODB.Recordset")
Category.ActiveConnection = MM_catalogmanager_STRING
Category.Source = "SELECT *  FROM tblCatalogCategory  WHERE GPCIDkey LIKE '" + Replace(Category__value1, "'", "''") + "'  ORDER BY CategoryID"
Category.CursorType = 0
Category.CursorLocation = 2
Category.LockType = 3
Category.Open()
Category_numRows = 0
%>
<%
Dim SubCategory__value1
SubCategory__value1 = "0"
If (Request.QueryString("cid")    <> "") Then 
  SubCategory__value1 = Request.QueryString("cid")   
End If
%>
<%
set SubCategory = Server.CreateObject("ADODB.Recordset")
SubCategory.ActiveConnection = MM_catalogmanager_STRING
SubCategory.Source = "SELECT *  FROM tblCatalogSubCategory  WHERE CategoryIDkey LIKE '" + Replace(SubCategory__value1, "'", "''") + "'  ORDER BY SubCategoryID"
SubCategory.CursorType = 0
SubCategory.CursorLocation = 2
SubCategory.LockType = 3
SubCategory.Open()
SubCategory_numRows = 0
%>
<%
Dim GPC__value1
GPC__value1 = "2"
If (Request.QueryString("cattype")   <> "") Then 
  GPC__value1 = Request.QueryString("cattype")  
End If
%>
<%
set GPC = Server.CreateObject("ADODB.Recordset")
GPC.ActiveConnection = MM_catalogmanager_STRING
GPC.Source = "SELECT *  FROM tblGPC  WHERE CatalogTypeIDkey LIKE '" + Replace(GPC__value1, "'", "''") + "'  ORDER BY GPCID"
GPC.CursorType = 0
GPC.CursorLocation = 2
GPC.LockType = 3
GPC.Open()
GPC_numRows = 0
%>
<%
Dim HLooper1__numRows
HLooper1__numRows = -3
Dim HLooper1__index
HLooper1__index = 0
GPC_numRows = GPC_numRows + HLooper1__numRows
%>
<%
Dim HLooper2__numRows
HLooper2__numRows = -3
Dim HLooper2__index
HLooper2__index = 0
SubCategory_numRows = SubCategory_numRows + HLooper2__numRows
%>
<%
Dim HLooper3__numRows
HLooper3__numRows = -3
Dim HLooper3__index
HLooper3__index = 0
GPC_numRows = GPC_numRows + HLooper3__numRows
%>
<html>
<head>
<title>Catalog Manager</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link href="../../../styles.css" rel="stylesheet" type="text/css">
</head>
<body>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td align="center">
<% if Not Request.QueryString("cid") <> "" then %>
<% if Not Request.QueryString("gpcid") <> ""then %>
<% if Not Request.QueryString("scid") <> ""then %>
<% if Not Request.QueryString("ItemID") <> ""then %>	
<table>
  <%
startrw = 0
endrw = HLooper3__index
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
          <td><div align="center"><a href="../../../SoftwareLibrary/Demo/applications/CatalogManager/extras/<%=request.servervariables("URL")%>?gpcid=<%=(GPC.Fields.Item("GPCID").Value)%>"><%=(GPC.Fields.Item("GPCName").Value)%></a></div></td>
        </tr>
        <tr>
          <td><%		  						  
Dim objgpcimage
strImage = "../images/" & GPC.Fields.Item("GPCImageFile").Value
Set objgpcimage = CreateObject("Scripting.FileSystemObject")
If objgpcimage.FileExists(Server.MapPath(strImage)) then
%>
              <% if GPC.Fields.Item("GPCImageFile").Value <> "" then %>
              <a href="../../../SoftwareLibrary/Demo/applications/CatalogManager/extras/<%=request.servervariables("URL")%>?gpcid=<%=(GPC.Fields.Item("GPCID").Value)%>"><img src="../images/<%=(GPC.Fields.Item("GPCImageFile").Value)%>" width="150" border="0"></a>
              <% end if ' image check %>
              <% end if ' image check %>
          </td>
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
<%end if%>
<br>
      <br>
      <table>
  <%
startrw = 0
endrw = HLooper1__index
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
                <td height="14"><div align="center"><a href="../../../SoftwareLibrary/Demo/applications/CatalogManager/extras/<%=request.servervariables("URL")%>?cid=<%=(Category.Fields.Item("CategoryID").Value)%>"> <%=(Category.Fields.Item("CategoryName").Value)%></a></div></td>
              </tr>
              <tr>
                <td valign="top"> <br>
                    <%		  						  
Dim objcatimage
strImage = "../images/" & Category.Fields.Item("CategoryImageFile").Value
Set objcatimage = CreateObject("Scripting.FileSystemObject")
If objcatimage.FileExists(Server.MapPath(strImage)) then
%>
                    <% if Category.Fields.Item("CategoryImageFile").Value <> "" then %>
                    <a href="../../../SoftwareLibrary/Demo/applications/CatalogManager/extras/<%=request.servervariables("URL")%>?cid=<%=(Category.Fields.Item("CategoryID").Value)%>"><img src="../images/<%=(Category.Fields.Item("CategoryImageFile").Value)%>" width="150" border="0"></a>
                    <% end if ' image check %>
                    <% end if ' image check %>
                </td>
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
<br>
<table>
  <%
startrw = 0
endrw = HLooper2__index
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
                <td height="14"> <div align="center"><a href="../../../SoftwareLibrary/Demo/applications/CatalogManager/extras/<%=request.servervariables("URL")%>?scid=<%=(SubCategory.Fields.Item("SubCategoryID").Value)%>"> <%=(SubCategory.Fields.Item("SubCategoryName").Value)%></a></div></td>
              </tr>
              <tr>
                <td valign="top"> <br>
                  <%		  						  
Dim objsubcatimage
strImage = "../images/" & SubCategory.Fields.Item("SubCategoryImageFile").Value
Set objsubcatimage = CreateObject("Scripting.FileSystemObject")
If objsubcatimage.FileExists(Server.MapPath(strImage)) then
%>
                  <% if SubCategory.Fields.Item("SubCategoryImageFile").Value <> "" then %>
                  <a href="../../../SoftwareLibrary/Demo/applications/CatalogManager/extras/<%=request.servervariables("URL")%>?scid=<%=(SubCategory.Fields.Item("SubCategoryID").Value)%>"><img src="../images/<%=(SubCategory.Fields.Item("SubCategoryImageFile").Value)%>" width="150" border="0"></a>
                  <% end if ' image check %>
                  <% end if ' image check %>
                </td>
              </tr>
            </table>
    </td>
    <%
	startrw = startrw + 1
	SubCategory.MoveNext()
	Wend
	%>
  </tr>
  <%
 numrows=numrows-1
 Wend
 %>
</table></td>
  </tr>
</table>
</body>
</html>
<%
Category.Close()
%>
<%
SubCategory.Close()
%>
<%
GPC.Close()
%>
