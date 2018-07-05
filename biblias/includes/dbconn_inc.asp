<%
'----------------------------------
		
Function OpenDB	
	Set DB = Server.CreateObject("ADODB.Connection")
	Dim strConnectionString
	If Request.ServerVariables("SERVER_NAME") = "www.totacc.com" Then
		strConnectionString = "DRIVER={Microsoft Access Driver " &_
			"(*.mdb)};DBQ=" & Server.MapPath("biblias.mdb") & ";uid=;pw=;"	
	Else
		strConnectionString = "DRIVER={Microsoft Access Driver " &_
			"(*.mdb)};DBQ=" & Server.MapPath("\biblias\biblias.mdb") & ";uid=;pw=;"	
	End If
	'Response.Write strConnectionString
	DB.ConnectionString = strConnectionString
	'Response.Write "<p class=small>OpenDB</p>"
	'Response.Write "<p>" & DB.ConnectionString & "</p>"
	DB.Open
End Function
	
Function CloseDB
	'Response.Write "<p class=small>CloseDB</p>"
	DB.Close
	Set DB = Nothing	
End Function
%>
