<%
function strip_tags(mistr)
Set toregexp = new RegExp
toregexp.pattern = "<[^>]+>"
toregexp.ignorecase = TRUE
toregexp.global = TRUE
mistr = toregexp.replace(mistr, "")
Set toregexp = Nothing
strip_tags = replace(mistr, """", "&quot;")
End Function
prodid=Trim(request.querystring("prod"))
catid=Trim(replace(request.querystring("cat"),"'",""))
sectionname=""
sectiondescription=""
productid=""
productname=""
productdescription=""
Set rs = Server.CreateObject("ADODB.RecordSet")
Set cnn=Server.CreateObject("ADODB.Connection")
cnn.open sDSN
if incfunctionsdefined=TRUE then
alreadygotadmin = getadminsettings()
sntxt = getlangid("sectionName",256)
sdtxt = getlangid("sectionDescription",512)
pntxt = getlangid("pName",1)
pdtxt = getlangid("pDescription",2)
else
sntxt = "sectionName"
sdtxt = "sectionDescription"
pntxt = "pName"
pdtxt = "pDescription"
end if
if prodid<>"" then
	rs.Open "SELECT pID,"&pntxt&","&sntxt&","&pdtxt&" FROM products INNER JOIN sections ON products.pSection=sections.sectionID WHERE pID='"&Replace(prodid,"'","''")&"'",cnn,0,1
	if NOT rs.EOF then
	productid=strip_tags(rs("pID")&"")
	productname=strip_tags(rs(pntxt)&"")
	productdescription=strip_tags(rs(pdtxt)&"")
	sectionname=strip_tags(rs(sntxt)&"")
	end if
	rs.Close
	if catid<>"" AND IsNumeric(catid) then
		rs.Open "SELECT "&sntxt&" FROM sections WHERE sectionID="&catid,cnn,0,1
		if NOT rs.EOF then sectionname=strip_tags(rs(sntxt)&"")
		rs.Close
	end if
elseif catid<>"" AND IsNumeric(catid) then
	topsection=0
	rs.Open "SELECT "&sntxt&","&sdtxt&",topSection FROM sections WHERE sectionID="&catid,cnn,0,1
	if NOT rs.EOF then
	sectionname=strip_tags(rs(sntxt)&"")
	sectiondescription=strip_tags(rs(sdtxt)&"")
	topsection=rs("topSection")&""
	end if
	rs.Close
	if topsection<>0 then
		rs.Open "SELECT sectionName FROM sections WHERE sectionID="&topsection,cnn,0,1
		if NOT rs.EOF then topsection=strip_tags(rs("sectionName")&"")
		rs.Close
	else
		topsection=""
	end if
end if
cnn.Close
set rs = nothing
set cnn = nothing
%>