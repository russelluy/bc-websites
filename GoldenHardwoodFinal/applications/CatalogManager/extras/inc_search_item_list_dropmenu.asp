<!--#include file="../../../SoftwareLibrary/Demo/Connections/catalogmanager.asp" -->
<%
Dim item_listmenu
Dim item_listmenu_numRows

Set item_listmenu = Server.CreateObject("ADODB.Recordset")
item_listmenu.ActiveConnection = MM_catalogmanager_STRING
item_listmenu.Source = "SELECT tblCatalog.ItemName  FROM tblCatalog  GROUP BY tblCatalog.ItemName"
item_listmenu.CursorType = 0
item_listmenu.CursorLocation = 2
item_listmenu.LockType = 1
item_listmenu.Open()

item_listmenu_numRows = 0
%>
<link href="../../../styles.css" rel="stylesheet" type="text/css">
<script language="JavaScript" type="text/JavaScript">
<!--
function MM_jumpMenu(targ,selObj,restore){ //v3.0
  eval(targ+".location='"+selObj.options[selObj.selectedIndex].value+"'");
  if (restore) selObj.selectedIndex=0;
}
//-->
</script>
<form name="form1">
  <select name="menu1" onChange="MM_jumpMenu('parent',this,0)">
  <option selected>Show All</option>
  </select>
</form>
<%
item_listmenu.Close()
Set item_listmenu = Nothing
%>
