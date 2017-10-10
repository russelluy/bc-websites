<%@ Language=VBScript %>
<% 
'************************
'****************Header
'************************
Dim objSendMail
Dim body

'body = "Form Field Values" & vbCrLf

'************************
'***Collect the values 
'***in the form
'************************

For Each obj in request.form
body = body & obj & " : " & request.form(obj) & vbCrLf
Next
body = body & ""

'************************
'***Sends the Email
'************************


Set objSendMail = Server.CreateObject("CDONTS.NewMail")
objSendMail.From =request.form("email")
objSendMail.To ="nordmaurice@aol.com"
objSendMail.Subject = "Comments from Visitor"
objSendMail.Body = body
objSendMail.MailFormat = 0
objSendMail.Importance = 1 
objSendMail.Send 
Set objCDOMail = Nothing

response.redirect "thankyou02.htm"
%>
