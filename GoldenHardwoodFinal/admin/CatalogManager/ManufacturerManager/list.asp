<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="../../../Connections/catalogmanager.asp" -->
<%
Dim List_Manufacturers__value2
List_Manufacturers__value2 = "%"
If (Request.Form("search")   <> "") Then 
  List_Manufacturers__value2 = Request.Form("search")  
End If
%>
<%
set List_Manufacturers = Server.CreateObject("ADODB.Recordset")
List_Manufacturers.ActiveConnection = MM_catalogmanager_STRING
List_Manufacturers.Source = "SELECT tblManufacturers.*  FROM tblManufacturers  WHERE Manufacturer LIKE '" + Replace(List_Manufacturers__value2, "'", "''") + "' OR ManufacturerDesc LIKE '%" + Replace(List_Manufacturers__value2, "'", "''") + "%'"
List_Manufacturers.CursorType = 0
List_Manufacturers.CursorLocation = 2
List_Manufacturers.LockType = 3
List_Manufacturers.Open()
List_Manufacturers_numRows = 0
%>
<%
Dim Repeat1__numRows
Repeat1__numRows = -1
Dim Repeat1__index
Repeat1__index = 0
List_Manufacturers_numRows = List_Manufacturers_numRows + Repeat1__numRows
%>
<html>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<head>
<title>Catalog Manager</title>
<script language="JavaScript" type="text/JavaScript">
<!--
function MM_jumpMenu(targ,selObj,restore){ //v3.0
  eval(targ+".location='"+selObj.options[selObj.selectedIndex].value+"'");
  if (restore) selObj.selectedIndex=0;
}
//-->
</script>
<link href="../../styles.css" rel="stylesheet" type="text/css">
</head>
<body>
<!--#include file="header.asp" -->

<table width="100%" border="0" cellspacing="0" cellpadding="0" height="24" class="tableborder">
  <tr> 
    <td height="24" valign="baseline">
      <form name="form" method="post" action="">
        <div align="center">Search by Keyword 
          <input name="search" type="text" id="search">
          <input type="submit" value="Go" name="submit">
        </div>
      </form>
    </td>
  </tr>
</table>

<% If Not List_Manufacturers.EOF Or Not List_Manufacturers.BOF Then %>
<table width="100%" height="32" border="0" cellpadding="0" cellspacing="0" class="tableborder">
  <tr class="tableheader"> 
    <td colspan="2"> Manufacturer Name</td>
    <td width="25%">Manufacturer Web site</td>
    <td width="18%"><div align="center">Activated</div></td>
    <td width="23%"> <a href="insert.asp">Insert New</a></td>
  </tr>
  <% 
While ((Repeat1__numRows <> 0) AND (NOT List_Manufacturers.EOF)) 
%>
  <tr class="<% 
RecordCounter = RecordCounter + 1
If RecordCounter Mod 2 = 1 Then
Response.Write "row1"
Else
Response.Write "row2"
End If
%>"> 
    <td width="3%" height="13">      <strong>
    <%Response.Write(RecordCounter)
RecordCounter = RecordCounter %>.      </strong>   </td>
    <td width="31%" height="13"><%=(List_Manufacturers.Fields.Item("Manufacturer").Value)%></td>
    <td height="13"><a href="http://<%=(List_Manufacturers.Fields.Item("ManufacturerWebsiteAddress").Value)%>" target="_blank"><%=(List_Manufacturers.Fields.Item("ManufacturerWebsiteAddress").Value)%></a> </td>
    <td width="18%" height="13"> 
      <div align="center">
        <input name="checkbox" type="checkbox" disabled="disabled" value="True" <%If (CStr((List_Manufacturers.Fields.Item("ManufacturerActivated").Value)) = CStr("True")) Then Response.Write("checked") : Response.Write("")%>>
      </div>
    </td>
    <td width="23%" height="13"><a href="update.asp?manid=<%=(List_Manufacturers.Fields.Item("ManufacturerID").Value)%>">Edit</a> | <a href="delete.asp?manid=<%=(List_Manufacturers.Fields.Item("ManufacturerID").Value)%>">Delete</a></td>
  </tr>
  <% 
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  List_Manufacturers.MoveNext()
Wend
%>
</table>
<% End If ' end Not List_Manufacturers.EOF Or NOT List_Manufacturers.BOF %>
<% If List_Manufacturers.EOF And List_Manufacturers.BOF Then %>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td><div align="center">No Records Found.....Please Try Again</div>
    </td>
  </tr>
</table>
<% End If ' end List_Manufacturers.EOF And List_Manufacturers.BOF %>
</body>
</html>
<%
List_Manufacturers.Close()
%>
