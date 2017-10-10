<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="../../Connections/catalogmanager.asp" -->
<%
' *** Edit Operations: declare variables
MM_editAction = CStr(Request("URL"))
If (Request.QueryString <> "") Then
  MM_editAction = MM_editAction & "?" & Request.QueryString
End If
' boolean to abort record edit
MM_abortEdit = false
' query string to execute
MM_editQuery = ""
%>
<%
' *** Delete Record: declare variables
if (CStr(Request("MM_delete")) <> "" And CStr(Request("MM_recordId")) <> "") Then
  MM_editConnection = MM_catalogmanager_STRING
  MM_editTable = "tblCatalogManufacturers"
  MM_editColumn = "ManufacturersID"
  MM_recordId = "" + Request.Form("MM_recordId") + ""
  MM_editRedirectUrl = "add_category.asp"
  ' append the query string to the redirect URL
  If (MM_editRedirectUrl <> "" And Request.QueryString <> "") Then
    If (InStr(1, MM_editRedirectUrl, "?", vbTextCompare) = 0 And Request.QueryString <> "") Then
      MM_editRedirectUrl = MM_editRedirectUrl & "?" & Request.QueryString
    Else
      MM_editRedirectUrl = MM_editRedirectUrl & "&" & Request.QueryString
    End If
  End If
End If
%>
<%
' *** Update Record: set variables
If (CStr(Request("MM_update")) = "form" And CStr(Request("MM_recordId")) <> "") Then
  MM_editConnection = MM_catalogmanager_STRING
  MM_editTable = "tblCatalogManufacturers"
  MM_editColumn = "ManufacturersID"
  MM_recordId = "" + Request.Form("MM_recordId") + ""
  MM_editRedirectUrl = ""
  MM_fieldsStr  = "ManufacturersName|value"
  MM_columnsStr = "ManufacturersName|',none,''"
  ' create the MM_fields and MM_columns arrays
  MM_fields = Split(MM_fieldsStr, "|")
  MM_columns = Split(MM_columnsStr, "|")
 ' set the form values
  For MM_i = LBound(MM_fields) To UBound(MM_fields) Step 2
    MM_fields(MM_i+1) = CStr(Request.Form(MM_fields(MM_i)))
  Next
  ' append the query string to the redirect URL
  If (MM_editRedirectUrl <> "" And Request.QueryString <> "") Then
    If (InStr(1, MM_editRedirectUrl, "?", vbTextCompare) = 0 And Request.QueryString <> "") Then
      MM_editRedirectUrl = MM_editRedirectUrl & "?" & Request.QueryString
    Else
      MM_editRedirectUrl = MM_editRedirectUrl & "&" & Request.QueryString
    End If
  End If
End If
%>
<%
' *** Delete Record: construct a sql delete statement and execute it
If (CStr(Request("MM_delete")) <> "" And CStr(Request("MM_recordId")) <> "") Then
  ' create the sql delete statement
  MM_editQuery = "delete from " & MM_editTable & " where " & MM_editColumn & " = " & MM_recordId
  If (Not MM_abortEdit) Then
    ' execute the delete
    Set MM_editCmd = Server.CreateObject("ADODB.Command")
    MM_editCmd.ActiveConnection = MM_editConnection
    MM_editCmd.CommandText = MM_editQuery
    MM_editCmd.Execute
    MM_editCmd.ActiveConnection.Close
    If (MM_editRedirectUrl <> "") Then
      Response.Redirect(MM_editRedirectUrl)
    End If
  End If
End If
%>
<%
' *** Update Record: construct a sql update statement and execute it
If (CStr(Request("MM_update")) <> "" And CStr(Request("MM_recordId")) <> "") Then
  ' create the sql update statement
  MM_editQuery = "update " & MM_editTable & " set "
  For MM_i = LBound(MM_fields) To UBound(MM_fields) Step 2
    MM_formVal = MM_fields(MM_i+1)
    MM_typeArray = Split(MM_columns(MM_i+1),",")
    MM_delim = MM_typeArray(0)
    If (MM_delim = "none") Then MM_delim = ""
    MM_altVal = MM_typeArray(1)
    If (MM_altVal = "none") Then MM_altVal = ""
    MM_emptyVal = MM_typeArray(2)
    If (MM_emptyVal = "none") Then MM_emptyVal = ""
    If (MM_formVal = "") Then
      MM_formVal = MM_emptyVal
    Else
      If (MM_altVal <> "") Then
        MM_formVal = MM_altVal
      ElseIf (MM_delim = "'") Then  ' escape quotes
        MM_formVal = "'" & Replace(MM_formVal,"'","''") & "'"
      Else
        MM_formVal = MM_delim + MM_formVal + MM_delim
      End If
    End If
    If (MM_i <> LBound(MM_fields)) Then
      MM_editQuery = MM_editQuery & ","
    End If
    MM_editQuery = MM_editQuery & MM_columns(MM_i) & " = " & MM_formVal
  Next
  MM_editQuery = MM_editQuery & " where " & MM_editColumn & " = " & MM_recordId

  If (Not MM_abortEdit) Then
    ' execute the update
    Set MM_editCmd = Server.CreateObject("ADODB.Command")
    MM_editCmd.ActiveConnection = MM_editConnection
    MM_editCmd.CommandText = MM_editQuery
    MM_editCmd.Execute
    MM_editCmd.ActiveConnection.Close

    If (MM_editRedirectUrl <> "") Then
      Response.Redirect(MM_editRedirectUrl)
    End If
  End If

End If
%>
<%
set Manufacturers = Server.CreateObject("ADODB.Recordset")
Manufacturers.ActiveConnection = MM_catalogmanager_STRING
Manufacturers.Source = "SELECT *  FROM tblManufacturers"
Manufacturers.CursorType = 0
Manufacturers.CursorLocation = 2
Manufacturers.LockType = 3
Manufacturers.Open()
Manufacturers_numRows = 0
%>
<%
Dim Repeat1__numRows
Repeat1__numRows = -1
Dim Repeat1__index
Repeat1__index = 0
Manufacturers_numRows = Manufacturers_numRows + Repeat1__numRows
%>
<%
' UltraDeviant - Row Number written by Owen Palmer (http://ultradeviant.co.uk)
Dim OP_RowNum
If MM_offset <> "" Then
	OP_RowNum = MM_offset + 1
Else
	OP_RowNum = 1
End If
%>
<html>
<head>
<title>Catalog Manager</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link href="../styles.css" rel="stylesheet" type="text/css">
</head>
<Body>
<!--#include file="header.asp" -->
<table width="100%"border="0" cellpadding="0" cellspacing="0" class="tableborder">
  <tr valign="middle"> 
    <td colspan="4" class="tableheader"><table width="100%" border="0" cellspacing="5" cellpadding="5">
        <tr>
          <td width="16%" valign="baseline"><strong>Add New Manufacturers </strong></td>
          <td width="84%" valign="baseline">
      <form name="form1">
              <input type="text" name="ManufacturersName" value="" size="32">
            <input type="submit" value="Insert New" name="submit">
      </form></td>
        </tr>
    </table>
    </td>
  </tr>
  <% 
While ((Repeat1__numRows <> 0) AND (NOT Manufacturers.EOF)) 
%>
  <tr class="<% 
RecordCounter = RecordCounter + 1
If RecordCounter Mod 2 = 1 Then
Response.Write "row1"
Else
Response.Write "row2"
End If
%>"> 
    <td width="2%" valign="baseline"><b> 
    <%Response.Write(RecordCounter)
RecordCounter = RecordCounter%>.</b>    </td>
    <td width="22%" valign="middle"><form name="form" method="POST" action="<%=MM_editAction%>">      
        <div align="left">
          <input name="ManufacturersName" type="text" id="ManufacturersName" value="<%=(Manufacturers.Fields.Item("ManufacturersName").Value)%>">
          <input type="hidden" name="MM_update" value="form">
          <input type="hidden" name="MM_recordId" value="<%= Manufacturers.Fields.Item("ManufacturersID").Value %>">
          <input type="submit" name="Submit" value="Update">
        </div>
    </form>
    </td>
    <td width="76%" valign="baseline"><form ACTION="<%=MM_editAction%>" METHOD="POST" name="delete">
        <div align="left">
          <input type="submit" name="Submit" value="Delete">
          <input type="hidden" name="MM_delete" value="delete">
<input type="hidden" name="MM_recordId" value="<%= Manufacturers.Fields.Item("ManufacturersID").Value %>">
      </div>
    </form></td>
  </tr>
  <% 
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  Manufacturers.MoveNext()
Wend
%>
</table>
</body>
</html>
<%
Manufacturers.Close()
%>