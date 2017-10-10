<%
Response.Buffer = True
if Trim(request.querystring("id1"))<>"" AND Trim(request.querystring("id2"))<>"" then
	response.cookies("id1")=request.querystring("id1")
	response.cookies("id1").Expires = Date()+180
	response.cookies("id2")=request.querystring("id2")
	response.cookies("id2").Expires = Date()+180
elseif Trim(request.querystring("PARTNER"))<>"" then
	response.cookies("PARTNER")=Trim(request.querystring("PARTNER"))
	response.cookies("PARTNER").Expires = Date()+Int(request.querystring("EXPIRES"))
elseif Trim(request.querystring("WRITECKL")) <> "" then
	response.cookies("WRITECKL")=Trim(request.querystring("WRITECKL"))
	response.cookies("WRITECKL").Expires = Date()+365
	response.cookies("WRITECKP")=Trim(request.querystring("WRITECKP"))
	response.cookies("WRITECKP").Expires = Date()+365
elseif Trim(request.querystring("DELCK")) = "yes" then
	response.cookies("WRITECKL")=""
	response.cookies("WRITECKL").Expires = Date()-30
	response.cookies("WRITECKP")=""
	response.cookies("WRITECKP").Expires = Date()-30
elseif Trim(request.querystring("WRITECLL")) <> "" then
	response.cookies("WRITECLL")=Trim(request.querystring("WRITECLL"))
	if Trim(request.querystring("permanent")) = "Y" then response.cookies("WRITECLL").Expires = Date()+365
	response.cookies("WRITECLP")=Trim(request.querystring("WRITECLP"))
	if Trim(request.querystring("permanent")) = "Y" then response.cookies("WRITECLP").Expires = Date()+365
elseif Trim(request.querystring("DELCLL")) <> "" then
	response.cookies("WRITECLL")=""
	response.cookies("WRITECLL").Expires = Date()-30
	response.cookies("WRITECLP")=""
	response.cookies("WRITECLP").Expires = Date()-30
end if
response.flush
%>