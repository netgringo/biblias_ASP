<%
'Response.Write "<br>gen_request.asp<br><br>"

'REQUEST FIELD TO STRING VARIABLE GENERATOR

'This function takes all items submitted by a form
'and assigns them a non-Request variable name
'---------------------------------------------------

Function makeRequestVariables()
	if Len(Request.Form)>0 then
		For Each Item in Request.Form
		  For iCount = 1 to Request.Form(Item).Count
			fields = Item & "=Request.Form(""" & Item & """)"
			'dimfields = "Dim " & Item
			'Response.Write "<br>fields = " & fields
			'Execute the command contained in the string 
			'(Sets all VBScript variables)
			'Execute(dimfields)
			Execute(fields)
		  Next
		Next
	end if

	if Len(Request.QueryString)>0 then
		'Include QueryString
		For Each Item in Request.QueryString
		  For iCount = 1 to Request.QueryString(Item).Count
			fields = Item & "=Request.QueryString(""" & Item & """)"
			'dimfields = "Dim " & Item
			'Response.Write "<br>fields = " & fields
			'Execute the command contained in the string 
			'(Sets all VBScript variables)
			'Execute(dimfields)
			Execute(fields)
		  Next
		Next
	end if

End Function

'Use the following to invoke:
'<!--#INCLUDE FILE="gen_request.asp"-->
'<% makeRequestVariables

%>