
<!--#include file="../../../Connections/catalogmanager.asp" -->
<%
Dim Category__value1
Category__value1 = "0"
If (Request.QueryString("gpcid")  <> "") Then 
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
<%
Dim RepeatGPC__numRows
Dim RepeatGPC__index

RepeatGPC__numRows = -1
RepeatGPC__index = 0
GPC_numRows = GPC_numRows + RepeatGPC__numRows
%>
<%
Dim RepeatCategory__numRows
Dim RepeatCategory__index

RepeatCategory__numRows = -1
RepeatCategory__index = 0
Category_numRows = Category_numRows + RepeatCategory__numRows
%>
<%
Dim RepeatItemsAvail__numRows
Dim RepeatItemsAvail__index

RepeatItemsAvail__numRows = -1
RepeatItemsAvail__index = 0
SubCategory_numRows = SubCategory_numRows + RepeatItemsAvail__numRows
%>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td><table width="100%" border="0" cellpadding="5" cellspacing="0" class="tableborder">
        <tr>
          <td><font color="#744900" size="2" face="Arial, Helvetica, sans-serif"><strong>MORE PRODUCTS:</strong></font></td>
        </tr>
        <tr>
          <td valign="top">
            <div align="left">                <% 
While ((RepeatGPC__numRows <> 0) AND (NOT GPC.EOF)) 
%>
                  <font color="#0066CC" size="1" face="Arial, Helvetica, sans-serif"><font face="Arial, Helvetica, sans-serif"><a href="<%=request.servervariables("URL")%>?gpcid=<%=(GPC.Fields.Item("GPCID").Value)%><%If Request.QueryString("mid") <> "" Then %>&mid=<%=request.querystring("mid")%><%end if%><%If Request.QueryString ("mid2")<> "" Then %>&mid2=<%=request.querystring("mid2")%><%end if%><%If Request.QueryString ("mid3")<> "" Then %>&mid3=<%=request.querystring("mid3")%><%end if%><%If Request.QueryString ("incid")<> "" Then %>&incid=<%=request.querystring("incid")%><%end if%>">			      <%=(GPC.Fields.Item("GPCName").Value)%></a>	  
			      <br>
                  <% 
  RepeatGPC__index=RepeatGPC__index+1
  RepeatGPC__numRows=RepeatGPC__numRows-1
  GPC.MoveNext()
Wend
%>
                  </font></font></div>
          </td>
        </tr>
      </table>
 <% If Not Category.EOF Or Not Category.BOF Then %>
        <table width="100%" border="0" cellpadding="5" cellspacing="0" class="tableborder">
          <tr>
            <td><font color="#744900" size="1" face="Arial, Helvetica, sans-serif"><strong>Category 
            <% if Request.QueryString("gpcid") <> "" then %>  
            of 
		       
            <%=(Category.Fields.Item("GPCName").Value)%>
            <% end if%>
            </strong>	</font> </td>
          </tr>
          <tr>
            <td valign="top">
              <div align="left">
                <% 
While ((RepeatCategory__numRows <> 0) AND (NOT Category.EOF)) 
%>
                <font color="#0066CC" size="1" face="Arial, Helvetica, sans-serif"><font face="Arial, Helvetica, sans-serif"><a href="<%=request.servervariables("URL")%>?cid=<%=(Category.Fields.Item("CategoryID").Value)%>&gpcid=<%=(Category.Fields.Item("GPCIDkey").Value)%><%If Request.QueryString("mid") <> "" Then %>&mid=<%=request.querystring("mid")%><%end if%><%If Request.QueryString ("mid2")<> "" Then %>&mid2=<%=request.querystring("mid2")%><%end if%><%If Request.QueryString ("mid3")<> "" Then %>&mid3=<%=request.querystring("mid3")%><%end if%><%If Request.QueryString ("incid")<> "" Then %>&incid=<%=request.querystring("incid")%><%end if%>">	            <%=(Category.Fields.Item("CategoryName").Value)%></a>
                 <br>
                <% 
  RepeatCategory__index=RepeatCategory__index+1
  RepeatCategory__numRows=RepeatCategory__numRows-1
  Category.MoveNext()
Wend
%>
                </font></font></div>
            </td>
          </tr>
      </table>
        <% End If ' end Not Category.EOF Or NOT Category.BOF %>
		        <% If Not SubCategory.EOF Or Not SubCategory.BOF Then %><table width="100%" border="0" cellpadding="5" cellspacing="0" class="tableborder">
          <tr>
            <td><font color="#744900" size="1" face="Arial, Helvetica, sans-serif"><strong>Sub Category of <%=(SubCategory.Fields.Item("CategoryName").Value)%></strong></font></td>
          </tr>
          <tr>
            <td valign="top">
              <% 
While ((RepeatItemsAvail__numRows <> 0) AND (NOT SubCategory.EOF)) 
%>
                <font color="#0066CC" size="1" face="Arial, Helvetica, sans-serif"><font face="Arial, Helvetica, sans-serif"><a href="<%=request.servervariables("URL")%>?scid=<%=(SubCategory.Fields.Item("SubCategoryID").Value)%><% If Request.QueryString("gpcid") <> "" Then%>&gpcid=<%=request.querystring("gpcid")%><%end if%><% If Request.QueryString("cid") <> "" Then%>&cid=<%=request.querystring("cid")%><%end if%><%If Request.QueryString("mid") <> "" Then %>&mid=<%=request.querystring("mid")%><%end if%><%If Request.QueryString ("mid2")<> "" Then %>&mid2=<%=request.querystring("mid2")%><%end if%><%If Request.QueryString ("mid3")<> "" Then %>&mid3=<%=request.querystring("mid3")%><%end if%><%If Request.QueryString ("incid")<> "" Then %>&incid=<%=request.querystring("incid")%><%end if%>">			    <%=(SubCategory.Fields.Item("SubCategoryName").Value)%></a><br>
                <% 
  RepeatItemsAvail__index=RepeatItemsAvail__index+1
  RepeatItemsAvail__numRows=RepeatItemsAvail__numRows-1
  SubCategory.MoveNext()
Wend
%>
                </font></font></td>
          </tr>
        </table>
        <% End If ' end Not SubCategory.EOF Or NOT SubCategory.BOF %>	  
    </td>
  </tr>
</table>
<font color="#0066CC" size="1" face="Arial, Helvetica, sans-serif">
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
</font>