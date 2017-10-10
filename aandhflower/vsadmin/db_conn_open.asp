<%
DIM sDSN

sDSN = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("/access_db/vsproducts.mdb") ' Microsoft Access 2000 using mapped path
'sDSN = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=c:\inetpub\wwwroot\access_db\vsproducts.mdb;" ' Microsoft Access 2000

' Please note, for SQL Server you must have an SQL Server database available. Most people will want to use the Access database provided.
'sDSN = "driver={SQL Server};server=SERVERNAME;uid=USERNAME;pwd=PASSWORD;database=DATABASENAME" ' SQL Server
%>
