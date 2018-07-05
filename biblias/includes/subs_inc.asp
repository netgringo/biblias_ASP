<%
'---------------------------------------------------------------
' This file contains FUNCTIONS and SUBPROCEDURES
' used throughout the web application:
'---------------------------------------------------------------
' DebugCollection
' Cycles through Form, QueryString, Session, Application collections
' 
' FormCollection	
' Cycles through Form/QueryString collections
' 
' showSQL			
' Displays current value of variable strSQL
' 
' runQuery
' Sets up and runs new firehose recordset
' 
' runRecordCountQuery
' Sets up and runs a recordset than can return a recordcount
' 
'---------------------------------------------------------------
Sub DebugCollection
	fontcolor="gray" 
	%>
	<p>&nbsp;</p>
	<p>&nbsp;</p>
	<p>&nbsp;</p>
	<!--#INCLUDE VIRTUAL="/simple/debug/debug_collection_inc2.asp"-->
	<
End Sub

Sub FormCollection
	fontcolor="gray" 
	%>
	<!--#INCLUDE VIRTUAL="/simple/debug/form_collection_inc.asp"-->
	<%
End Sub

Sub showSQL
	Response.Write "<p>" & strSQL & "</p>" & VbCrLf
End Sub

Function runQuery
 	Set objRec = Server.CreateObject("ADODB.Recordset")
	objRec.Open strSQL, DB, 0, 1
End Function

Function runRecordCountQuery
 	Set objRec = Server.CreateObject("ADODB.Recordset")
	objRec.Open strSQL, DB, 1, 1
End Function

Sub buildDropMenu(table_name,id_field,display_name,id_value)
	' Loop through specified database table and build
	' dropmenu of all records
	'-------------------------------------------------
	strSQL = "SELECT " & id_field & ", " & display_name & " FROM " & table_name &_
			 " ORDER BY ID"
	'showSQL
	runQuery
	i = 0
	buildOptions = ""
	
	'Set ID field for this table
	Dim thisID
	
	While Not objRec.EOF
		i = i + 1
		makeRecordsetVariables
		
		thisID = Int(id_value)
		
		buildOptions = buildOptions & "<OPTION VALUE=""" & objRec(id_field) & """"
		
		' Pre-select access level based on user's access level 
		' passed to this subprocedure:
		if i = thisID then
			buildOptions = buildOptions & " SELECTED"
		end if
		
		buildOptions = buildOptions & ">" & objRec(display_name) & "</OPTION>" & VbCrLf
		
		objRec.MoveNext
	Wend

	buildOptions = "<SELECT NAME="""& id_field & """>" & VbCrLf &_
		"<OPTION VALUE="""">Please Select:</OPTION>" & VbCrLf &_
		buildOptions & VbCrLf &_
		"</SELECT>" & VbCrLf
		
	Response.Write buildOptions
End Sub
%>
