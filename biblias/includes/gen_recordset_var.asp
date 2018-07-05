<%
'Response.Write "<br><-- gen_recordset_var.asp<br><br>"

'RECORDSET TO STRING VARIABLE GENERATOR

'This function takes all items in a recordset
'and assigns them a variable name
'---------------------------------------------------

Function makeRecordsetVariables()
	'Loop through recordset
	'-----------------------------
	if not objRec.EOF then
		'Response.Write "<br>iCount=" & objRec.Fields.Count & "<br>"
		For iCount = 0 to (objRec.Fields.Count-1)
			fields = objRec(iCount).Name & " = objRec(""" & objRec(iCount).Name & """)"
			'Response.Write fields & " = " & objRec(objRec(iCount).Name) & "<br>"
			'Response.Write "<br>" & fields
			'Execute the command contained in the string 
			'(Sets all VBScript variables)
			Execute(fields)
		next
		'objRec.MoveNext
	end if
End Function

'Use the following to invoke:
'<!--#INCLUDE FILE="gen_recordset_var.asp"-->
'<% makeRecordsetVariables

'Response.Write "<br>gen_recordset_var.asp --><br><br>"
%>