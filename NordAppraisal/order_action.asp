<%
dim Body, Body2
dim comp_name, comp_addr, email, order_by, phone, fax, appraisal_type, property_type, appraisal_purpose,FHA_case, loan_no, st_addr, subj_city, country, cstate, zip, sales_price1,sales_price2, b_bath, b_story, b_year, build_area, lot_size, checkbox1,checkbox2,checkbox3,checkbox4,checkbox5, access, name_cont, access_cont, best_time, radiobutton, per_app, comment

comp_name=  Request.Form("comp_name")
comp_addr=  Request.Form("comp_addr")
email = Request.Form("email")
order_by = Request.Form("order_by")
phone =  Request.Form("phone")
fax = Request.Form("fax")
appraisal_type = Request.Form("appraisal_type")
property_type = Request.Form("property_type")

appraisal_purpose = Request.Form("appraisal_purpose")
FHA_case = Request.Form("FHA_case")
loan_no = Request.Form("loan_no")
st_addr = Request.Form("st_addr")
subj_city = Request.Form("subj_city")
country = Request.Form("country")
cstate = Request.Form("cstate")
zip = Request.Form("zip")

sales_price1 = Request.Form("sales_price1")
sales_price2 = Request.Form("sales_price2")
b_bath = Request.Form("b_bath")
b_story = Request.Form("b_story")
b_year = Request.Form("b_year")
build_area= Request.Form("build_area")
lot_size = Request.Form("lot_size")

checkbox1 = Request.Form("checkbox1")
checkbox2 = Request.Form("checkbox2")
checkbox3 = Request.Form("checkbox3")
checkbox4 = Request.Form("checkbox4")
checkbox5 = Request.Form("checkbox5")
access = Request.Form("access")
name_cont = Request.Form("name_cont")
access_cont = Request.Form("access_cont")
best_time = Request.Form("best_time")
radiobutton = Request.Form("radiobutton")
per_app = Request.Form("per_app")
comment = Request.Form("comment")

		body = "Nord Appraisal" & vbcrlf & vbcrlf & _
		"CONTACT IBO" & vbcrlf & vbcrlf & _
		"Autoresponder: DO NOT REPLY TO THIS EMAIL" & vbcrlf & vbcrlf & _
		"Thank you for Ordering for Nord Appraisal" & vbcrlf & vbcrlf & _
		"If you have any suggestions or comments,  let us know! " & vbcrlf & vbcrlf & _
		"Company name       : " & comp_name & vbcrlf &_
		"Company Address    : " & comp_addr & vbcrlf &_
		"Email Address      : " & email & vbcrlf &_
		"Order By           : " & order_by & vbcrlf &_
		
		"Phone              : " & phone& vbcrlf &_
		"Fax                : " & fax & vbcrlf &_
		"Appraisal Type     : " & appraisal_type & vbcrlf &_
		"Property Type      : " & property_type & vbcrlf &_
		"Appraisal Purpose  : " & appraisal_purpose & vbcrlf &_
		"FHA_case           : " & FHA_case & vbcrlf &_

		"Loan No            : " & loan_no & vbcrlf &_
		"Street Address     : " & st_addr & vbcrlf &_
		"Subject City       : " & subj_city & vbcrlf &_
		"Country            : " & country & vbcrlf &_
		"State              : " & cstate & vbcrlf &_
		"Zip                : " & zip & vbcrlf &_
		"Sales Price (1)    : " & sales_price1 & vbcrlf &_
		"Sales Price (2)    : " & sales_price2 & vbcrlf & vbcrlf &_
		
		"Bedrooms" & vbcrlf & _
		"Bath               : " & b_bath & vbcrlf &_
		"Story              : " & b_story & vbcrlf &_
		"Year Built         : " & b_year & vbcrlf & vbcrlf &_
	
		
		"Building Area      : " & build_area & "Lot Size: " & lot_size & vbcrlf &_
		"Amenities          : " & checkbox1 & vbcrlf &_
		"                     " & checkbox2 & vbcrlf &_
		"                     " & checkbox3 & vbcrlf &_
		"                     " & checkbox4 & vbcrlf &_
		"                     " & checkbox5 & vbcrlf &_
		
		"Contact:(Name)     : " & name_cont & vbcrlf &_
		"Access Contact     : " & access_cont & vbcrlf &_
		"Best Time of Call  : " & best_time & vbcrlf &_
		"Method of Appraisal Fee Payment: " & radiobutton & vbcrlf &_
		"Person or Company that pays appraisal fee: " & per_app & vbcrlf & vbcrlf &_
	
		"Comment            : " & comment & vbcrlf
		
		
		body1 = "Nord Appraisal" & vbcrlf & vbcrlf & _
		"CONTACT IBO" & vbcrlf & vbcrlf & _
		"Autoresponder: DO NOT REPLY TO THIS EMAIL" & vbcrlf & vbcrlf & _
		"Thank you for Ordering for Nord Appraisal" & vbcrlf & vbcrlf & _
		"If you have any suggestions or comments,  let us know! " & vbcrlf & vbcrlf & _
		"Company name       : " & comp_name & vbcrlf &_
		"Company Address    : " & comp_addr & vbcrlf &_
		"Email Address      : " & email & vbcrlf &_
		"Order By           : " & order_by & vbcrlf &_
		
		"Phone              : " & phone& vbcrlf &_
		"Fax                : " & fax & vbcrlf &_
		"Appraisal Type     : " & appraisal_type & vbcrlf &_
		"Property Type      : " & property_type & vbcrlf &_
		"Appraisal Purpose  : " & appraisal_purpose & vbcrlf &_
		"FHA_case           : " & FHA_case & vbcrlf &_

		"Loan No            : " & loan_no & vbcrlf &_
		"Street Address     : " & st_addr & vbcrlf &_
		"Subject City       : " & subj_city & vbcrlf &_
		"Country            : " & country & vbcrlf &_
		"State              : " & cstate & vbcrlf &_
		"Zip                : " & zip & vbcrlf &_
		"Sales Price (1)    : " & sales_price1 & vbcrlf &_
		"Sales Price (2)    : " & sales_price2 & vbcrlf & vbcrlf &_
		
		"Bedrooms" & vbcrlf & _
		"Bath               : " & b_bath & vbcrlf &_
		"Story              : " & b_story & vbcrlf &_
		"Year Built         : " & b_year & vbcrlf & vbcrlf &_
	
		
		"Building Area      : " & build_area & "Lot Size: " & lot_size & vbcrlf &_
		"Amenities          : " & checkbox1 & vbcrlf &_
		"                     " & checkbox2 & vbcrlf &_
		"                     " & checkbox3 & vbcrlf &_
		"                     " & checkbox4 & vbcrlf &_
		"                     " & checkbox5 & vbcrlf &_
		
		"Contact:(Name)     : " & name_cont & vbcrlf &_
		"Access Contact     : " & access_cont & vbcrlf &_
		"Best Time of Call  : " & best_time & vbcrlf &_
		"Method of Appraisal Fee Payment: " & radiobutton & vbcrlf &_
		"Person or Company that pays appraisal fee: " & per_app & vbcrlf & vbcrlf &_
	
		"Comment            : " & comment & vbcrlf
		
				
		Set objemail = Server.CreateObject("CDONTS.NewMail")
		objemail.To = email
		objemail.From = "nordmaurice@aol.com"
		objemail.Subject ="Online Order Form"
		objemail.Body = body
		objemail.MailFormat = 0
		objemail.Importance = 1 
		objemail.Send 
		Set objemail = Nothing
		
		Set objemail = Server.CreateObject("CDONTS.NewMail")
		objemail.To = "nordmaurice@aol.com"
		objemail.From = email
		objemail.Subject ="Online Order Form"
		objemail.Body = body1
		objemail.MailFormat = 0
		objemail.Importance = 1 
		objemail.Send 
		Set objemail = Nothing
        response.redirect("Thankyou.htm")
%>
