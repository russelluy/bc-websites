<%
if menupoplimit="" then menupoplimit=9
Set rs = Server.CreateObject("ADODB.RecordSet")
Set cnn=Server.CreateObject("ADODB.Connection")
on error resume next
err.number = 0
cnn.open sDSN
alreadygotadmin = getadminsettings()
if Session("clientLoginLevel")<>"" then minloglevel=Session("clientLoginLevel") else minloglevel=0
rs.Open "SELECT sectionID,"&getlangid("sectionName",256)&",topSection,rootSection,sectionurl FROM sections WHERE sectionDisabled<="&minloglevel&" ORDER BY sectionOrder",cnn,0,1
if NOT rs.EOF then mAlldata = rs.getrows
rs.Close
cnn.Close
set rs = nothing
set cnn = nothing
Sub mwritemenulevel(id,itlevel)
	Dim mIndex
	if itlevel<=menupoplimit then
		if NOT (menucategoriesatroot=2 AND id=0) then
			for mIndex=0 TO ubound(mAlldata,2)
				if mAlldata(2,mIndex)=id then
					mTID = mAlldata(2,mIndex)
					if mTID = 0 then mTID = ""
					if menucategoriesatroot=1 then
						menuheadsec = "mymenu.addMenu("
					else
						menuheadsec = "mymenu.addSubMenu(""products"&mTID&""","
					end if
					if trim(mAlldata(4,mIndex)&"")<>"" then
						response.write menuheadsec&"""products"&mAlldata(0,mIndex)&""", """&menuprestr&replace(mAlldata(1,mIndex)&"","""","\""")&menupoststr&""", """&mAlldata(4,mIndex)&""");"&vbCrLf
					else
						if mAlldata(3,mIndex)=0 then
							response.write menuheadsec&"""products"&mAlldata(0,mIndex)&""", """&menuprestr&replace(mAlldata(1,mIndex)&"","""","\""")&menupoststr&""", ""categories.asp?cat="&mAlldata(0,mIndex)&""");"&vbCrLf
						else
							response.write menuheadsec&"""products"&mAlldata(0,mIndex)&""", """&menuprestr&Replace(mAlldata(1,mIndex)&"","""","\""")&menupoststr&""", ""products.asp?cat="&mAlldata(0,mIndex)&""");"&vbCrLf
						end if
					end if
				end if
			next
		end if
		FOR mIndex=0 TO ubound(mAlldata,2)
			if mAlldata(2,mIndex)=id AND mAlldata(3,mIndex)=0 AND menucategoriesatroot<>1 then call mwritemenulevel(mAlldata(0,mIndex),itlevel+1)
		NEXT
	end if
end Sub
sub writesubmenus()
	menucategoriesatroot=2
	call mwritemenulevel(0,2)
end sub
if err.number <> 0 then
	response.write "mymenu.addSubMenu(""products"", """", ""<strong>Please check your database connection</strong>"", ""http://www.beancastle.com"");"&vbCrLf
elseif IsArray(mAlldata) then
	call mwritemenulevel(0,1)
end if
on error goto 0
%>