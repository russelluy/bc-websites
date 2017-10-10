<%
sortBy = 1
pathtossl = ""
taxShipping=0
pagebarattop=0
productcolumns=2
useproductbodyformat=1
usesearchbodyformat=1
usedetailbodyformat=1
useemailfriend=true
nobuyorcheckout=false
noprice=false
expireaffiliate=30
sqlserver=false
usecategoryformat=1
allproductsimage=""
nogiftcertificate=false
showtaxinclusive=false
upspickuptype="03"
overridecurrency=false
	orcsymbol="AU$ "
	orcemailsymbol="AU$ "
	orcdecplaces=2
	orcpreamount=true
encryptmethod="aspencrypt"
commercialloc=true
showcategories=true
termsandconditions=false
showquantonproduct=false
showquantondetail=false
addshippinginsurance=0
noshipaddress=false
pricezeromessage=""
showproductid=false
currencyseparator=" "
noproductoptions=false
invoiceheader=""
invoiceaddress=""
invoicefooter=""
dumpccnumber=false
actionaftercart=1
dateadjust=0
emailorderstatus=3
htmlemails=false
categorycolumns=1
noshowdiscounts=false
catseparator="<br>&nbsp;"
willpickuptext=""
willpickupcost=0
extraorderfield1=""
extraorderfield1required=false
extraorderfield2=""
extraorderfield2required=false
enableclientlogin=true

' ===================================================================
' Please do not edit anything below this line
' ===================================================================

maintablebg=""
innertablebg=""
maintablewidth="98%"
innertablewidth="100%"
maintablespacing="0"
innertablespacing="0"
maintablepadding="1"
innertablepadding="6"
headeralign="left"

Session.LCID = 1033

const maxprodopts=15
const helpbaseurl="http://www.beancastle.com"

Function Max(a,b)
	if a > b then
		Max=a
	else
		Max=b
	end if
End function
Function Min(a,b)
	if a < b then
		Min=a
	else
		Min=b
	end if
End function
%>