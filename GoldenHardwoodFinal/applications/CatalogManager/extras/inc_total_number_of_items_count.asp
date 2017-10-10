<!--#include file="../../../SoftwareLibrary/Demo/Connections/catalogmanager.asp" -->
<%
Dim count
Dim count_numRows

Set count = Server.CreateObject("ADODB.Recordset")
count.ActiveConnection = MM_catalogmanager_STRING
count.Source = "SELECT Count(tblCatalog.ItemID) AS CountOfItemID  FROM tblCatalog"
count.CursorType = 0
count.CursorLocation = 2
count.LockType = 1
count.Open()

count_numRows = 0
%>
  <font color="#FF0000"><strong><%=(count.Fields.Item("CountOfItemID").Value)%></strong></font>
      <%
count.Close()
Set count = Nothing
%>
