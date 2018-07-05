<%

'	This page check browser type and accordingly
'	sets variables for sizes of form fields.
'	
'	Internet Explorer (IE) and Netscape display
'	form field sizes in a radically different way.

'Set bc = Server.CreateObject("MSWC.BrowserType")
Set bc = Request.ServerVariables("HTTP_USER_AGENT")

IE = InStr(bc,"MSIE")
Opera = InStrRev(bc,"Opera")
Lynx = InStrRev(bc,"Lynx")

if Opera > 0 then
	Session("browser") = "Opera"
elseif IE > 0 then
	Session("browser") = "IE"
elseif Lynx > 0 then
	Session("browser") = "Lynx"
else
	Session("browser") = "other"
end if


'Detect Macintosh browsers:
if InStr(bc,"Mac")>0 then
	Session("Mac") = true
end if

'Response.Write "bc=" & bc & "<br>"
'Response.Write "IE=" & IE & "<br>"
'Response.Write "Opera=" & Opera & "<br>"
'Response.Write "Lynx=" & Lynx & "<br>"
'Response.Write "Session(browser)=" & Session("browser") & "<br>"
'Response.Write "Session(Mac)=" & Session("Mac") & "<br>"

if Session("browser") = "IE" then 
	numcol = 65
	numcols = 40
	numcols2 = 35
	numcols3 = 30
	numcols4 = 20
	numcols5 = 17
	numcols6 = 8
	numcols7 = 5
	tdclass = "tiny"
else
	numcol = 55
	numcols = 40
	numcols2 = 30
	numcols3 = 18
	numcols4 = 17
	numcols5 = 12
	numcols6 = 5
	numcols7 = 3
	tdclass = "small"
end if
%>