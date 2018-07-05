<%@ Language=VBScript %>
<% Option Explicit
' Include Dims
'---------------------------------------------------- %>
<!--#INCLUDE FILE="includes/variables_inc.asp"-->
<%
' Include some sub-routines and functions: 
'---------------------------------------------------- %>
<!--#INCLUDE FILE="includes/subs_inc.asp"-->
<%
' Include data connection (call with OpenDB/CloseDB): 
'---------------------------------------------------- %>
<!--#INCLUDE FILE="includes/dbconn_inc.asp"-->
<%
' Include browser detection script: 
'---------------------------------------------------- %>
<!--#INCLUDE FILE="includes/browser_check_inc.asp"-->
<%
' Include function makeRecordsetVariables
' to cycle through recordsets and make variables
' for instance, firstname = objRec("firstname"): 
'---------------------------------------------------- %>
<!--#INCLUDE FILE="includes/gen_recordset_var.asp"-->
<%
' Include function makeRequestVariables
' to cycle through collections and make variables
' for instance, firstname = Request("firstname"): 
'---------------------------------------------------- %>
<!--#INCLUDE FILE="includes/gen_request.asp"-->
<%
'------------------------------------------------------------
' If there are any items in the form collection or in the
' querystring, call function makeRequestVariables,
' which does this for each field name: 
'	fieldname = Request.Form("fieldname")
'------------------------------------------------------------
if Len(Request.Form)>0 OR Len(Request.QueryString)>0 then
	makeRequestVariables
end if
'------------------------------------------------------------
%>
<HTML>
<HEAD>
<TITLE>Comparación de cuatro versiones de la Biblia</TITLE>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<link REL="stylesheet" HREF="style.css" TYPE="text/css">
</HEAD>
<BODY BGCOLOR="#FFFF99" leftmargin="5" topmargin="5">
<% OpenDB
'--------------------------------------------
' This page allows users to search for 
' Scriptures in two Bible versions
'--------------------------------------------
%>
<FORM ACTION="default2.asp" METHOD="POST" NAME="myForm">

<TABLE BORDER="0" CELLPADDING="13" CELLSPACING="0">
<TR VALIGN="top"><TD ALIGN="right"><B CLASS="xlarge">Comparación de
<BR>cuatro versiones
<BR>de la Biblia</B></TD>
<TD>
<TABLE BORDER="1" CELLPADDING="3" CELLSPACING="0">
<!-- LIBRO BÍBLICO -->
<TR VALIGN="top"><TD ALIGN="right"><B>Libro bíblico</B></TD>
	<TD ALIGN="left">
	<%	'table_name,id_field,display_name,id_value
		Call buildDropMenu(_
			"LibrosBiblicos",_
			"ID",_
			"libro_nombre",_
			Request.Form("ID"))%>
	</TD></TR>
<!-- CAPÍTULO Y VERSÍCULO -->
<TR VALIGN="top"><TD ALIGN="right"><B>Capítulo y versículo</B></TD>
	<TD ALIGN="left">
			
		<INPUT TYPE="TEXT" NAME="cap" SIZE="<%=numcols7%>" MAXLENGTH="5" VALUE="<%=Request.Form("cap")%>">
		<INPUT TYPE="TEXT" NAME="vers_num" SIZE="<%=numcols7%>" MAXLENGTH="5" VALUE="<%=Request.Form("vers_num")%>">
		<INPUT TYPE="SUBMIT" NAME="Submit" VALUE="Search"></TD></TR>
</TABLE>
</TD>
<TD><BR><P CLASS="xsmall">
1. Choose the Bible book.<BR>
2. Type the chapter number.<BR>
3. Type the verse or verses. (For example: "17", or "17-20")</P></TD>
</TR></FORM></TABLE>
<%

'----------------------------------------
' DISPLAY SEARCH RESULTS
'----------------------------------------
if Len(Request.Form("ID")) then
	'---------------------------------------------------------
	' GET NAME OF BIBLE BOOK PASSED BY FORM
	'---------------------------------------------------------	
	strSQL = "SELECT libro_nombre FROM LibrosBiblicos" &_
				 " WHERE ID = " & Request.Form("ID")
	'showSQL
	runQuery
	if not objRec.EOF then
		makeRecordsetVariables
	end if
	
	Dim j, strNoItemsFound
	
	objRec.Close
	Set objRec = Nothing
	'---------------------------------------------------------
	' BUILD AND RUN QUERY based on values passed 
	' by the SEARCH FORM.
	'---------------------------------------------------------	
	Dim new_vers_num, new_ID, form_ID, new_cap
	for i = 1 to 4
		new_vers_num = Request.Form("vers_num")
		form_ID = Request.Form("ID")
		new_ID = form_ID
		new_cap = Request.Form("cap")
		
		Select Case i
			Case 1
				tablename = "ReinaValera"
				bibliaName = "Reina-Valera"
			Case 2
				tablename = "SagradasEscrituras"
				bibliaName = "Sagradas Escrituras"
			Case 3
				tablename = "BibliaJerusalen"
				bibliaName = "Biblia de Jerusalén"
				If form_ID >= 17 Then
					new_ID = form_ID + 2
					If form_ID >= 18 Then
						new_ID = new_ID + 2
					End If
					If form_ID >= 23 Then
						new_ID = new_ID + 2
					End If
					If form_ID >= 26 Then
						new_ID = new_ID + 1
						Select Case new_ID
							Case 38,64,70,71,72
								new_cap = 0
						End Select
					End If
				End If
			Case 4
				tablename = "ModernSpanish"
				bibliaName = "Modern Spanish"
		End Select
	
		strSQL = "SELECT ID, cap, vers_num, vers_txt" &_
				 " FROM " & tablename &_
				 " WHERE libro = " & new_ID
				 
		If Len(new_cap)>0 AND new_cap <> "0" Then
			strSQL = strSQL & " AND cap = " & new_cap
			if InStr(Request.Form("vers_num"),"-")>0 then
				new_vers_num = Replace(Request.Form("vers_num"),"-"," AND ")
				strSQL = strSQL & " AND vers_num BETWEEN " & new_vers_num
			else
				strSQL = strSQL & " AND vers_num = " & new_vers_num
			end if
		End If
		
		'showSQL
		runQuery
		' Begin main table for results list:
		j = 1
		if not objRec.EOF then
			While not objRec.EOF
				makeRecordsetVariables
				if i = 1 then
					if j = 1 then
						bibliaValera = bibliaValera & "<TR><TD class=""medium""><B>" & bibliaName & "</B></TD></TR>" & VbCrLf
						bibliaValera = bibliaValera & "<TR><TD CLASS=""small""><B>" & libro_nombre & " " & new_cap & ":" & Request("vers_num") & "</B></TD></TR>" & VbCrLf
					end if
					bibliaValera = bibliaValera & "<TR><TD CLASS=""small""><SUP>" & vers_num & "</SUP> " & vers_txt & "</TD></TR>" & VbCrLf
					j = j + 1
				elseif i = 2 then
					if j = 1 then
						bibliaOtra = bibliaOtra & "<TR><TD class=""medium""><B>" & bibliaName & "</B></TD></TR>" & VbCrLf
						bibliaOtra = bibliaOtra & "<TR><TD CLASS=""small""><B>" & libro_nombre & " " & new_cap & ":" & Request("vers_num") & "</B></TD></TR>" & VbCrLf
					end if
					bibliaOtra = bibliaOtra & "<TR><TD CLASS=""small""><SUP>" & vers_num & "</SUP> " & vers_txt & "</TD></TR>" & VbCrLf
					j = j + 1		
				elseif i = 3 then
					if j = 1 then
						bibliaJerusalen = bibliaJerusalen & "<TR><TD class=""medium""><B>" & bibliaName & "</B></TD></TR>" & VbCrLf
						bibliaJerusalen = bibliaJerusalen & "<TR><TD CLASS=""small""><B>" & libro_nombre & " " & new_cap & ":" & Request("vers_num") & "</B></TD></TR>" & VbCrLf
					end if
					bibliaJerusalen = bibliaJerusalen & "<TR><TD CLASS=""small""><SUP>" & vers_num & "</SUP> " & vers_txt & "</TD></TR>" & VbCrLf
					j = j + 1
				elseif i = 4 then
					if j = 1 then
						bibliaModerna = bibliaModerna & "<TR><TD class=""medium""><B>" & bibliaName & "</B></TD></TR>" & VbCrLf
						bibliaModerna = bibliaModerna & "<TR><TD><B CLASS=""small"">" & libro_nombre & " " & new_cap & ":" & Request("vers_num") & "</B></TD></TR>" & VbCrLf
					end if
					bibliaModerna = bibliaModerna & "<TR><TD CLASS=""small""><SUP>" & vers_num & "</SUP> " & vers_txt & "</TD></TR>" & VbCrLf
					j = j + 1	
				end if					
				objRec.MoveNext
			Wend
		else
			strNoItemsFound = "<P>Nothing found.</P>" & VbCrLf
			if i = 1 then
				if j = 1 then
					bibliaValera = bibliaValera & "<TR><TD class=""medium""><B>" & bibliaName & "</B></TD></TR>" & VbCrLf
					bibliaValera = bibliaValera & "<TR><TD CLASS=""small""><B>" & libro_nombre & " " & new_cap & ":" & Request("vers_num") & "</B></TD></TR>" & VbCrLf
				end if
				bibliaValera = bibliaValera & "<TR><TD CLASS=""small"">" & strNoItemsFound & "</TD></TR>" & VbCrLf
				j = j + 1
			elseif i = 2 then
				if j = 1 then
					bibliaOtra = bibliaOtra & "<TR><TD class=""medium""><B>" & bibliaName & "</B></TD></TR>" & VbCrLf
					bibliaOtra = bibliaOtra & "<TR><TD><B CLASS=""small"">" & libro_nombre & " " & new_cap & ":" & Request("vers_num") & "</B></TD></TR>" & VbCrLf
				end if
				bibliaOtra = bibliaOtra & "<TR><TD CLASS=""small"">" & strNoItemsFound & "</TD></TR>" & VbCrLf
				j = j + 1		
			elseif i = 3 then
				if j = 1 then
					bibliaJerusalen = bibliaJerusalen & "<TR><TD class=""medium""><B>" & bibliaName & "</B></TD></TR>" & VbCrLf
					bibliaJerusalen = bibliaJerusalen & "<TR><TD CLASS=""small""><B>" & libro_nombre & " " & new_cap & ":" & Request("vers_num") & "</B></TD></TR>" & VbCrLf
				end if
				bibliaJerusalen = bibliaJerusalen & "<TR><TD CLASS=""small"">" & strNoItemsFound & "</TD></TR>" & VbCrLf
				j = j + 1
			elseif i = 4 then
				if j = 1 then
					bibliaModerna = bibliaModerna & "<TR><TD class=""medium""><B>" & bibliaName & "</B></TD></TR>" & VbCrLf
					bibliaModerna = bibliaModerna & "<TR><TD CLASS=""small""><B>" & libro_nombre & " " & new_cap & ":" & Request("vers_num") & "</B></TD></TR>" & VbCrLf
				end if
				bibliaModerna = bibliaModerna & "<TR><TD CLASS=""small"">" & strNoItemsFound & "</TD></TR>" & VbCrLf
				j = j + 1
			end if	
		end if
	
		objRec.Close
		Set objRec = Nothing
		j = ""
	next

	Response.Write "<TABLE BORDER=""0"" CELLPADDING=""20"" CELLSPACING=""0"" WIDTH=""100%"">" & VbCrLf & "<TR><TD BGCOLOR=""pink"" VALIGN=""TOP"" WIDTH=""20%"">" & VbCrLf
	
	Response.Write "<TABLE BORDER=""0"" CELLPADDING=""3"" CELLSPACING=""0"" WIDTH=""100%"">" & VbCrLf
	Response.Write bibliaValera
	Response.Write "</TABLE>" & VbCrLf
	
	'Response.Write "</TD><TD WIDTH=""5%"">&nbsp;&nbsp;&nbsp;</TD><TD BGCOLOR=""lightgreen"" VALIGN=""TOP"" WIDTH=""25%"">"
	Response.Write "</TD><TD BGCOLOR=""lightgreen"" VALIGN=""TOP"" WIDTH=""20%"">" & VbCrLf

	Response.Write "<TABLE BORDER=""0"" CELLPADDING=""3"" CELLSPACING=""0"" WIDTH=""100%"">" & VbCrLf
	Response.Write bibliaOtra
	Response.Write "</TABLE>" & VbCrLf

	'Response.Write "</TD><TD WIDTH=""5%"">&nbsp;&nbsp;&nbsp;</TD><TD BGCOLOR=""lightblue"" VALIGN=""TOP"" WIDTH=""25%"">"
	Response.Write "</TD><TD BGCOLOR=""lightblue"" VALIGN=""TOP"" WIDTH=""20%"">" & VbCrLf

	Response.Write "<TABLE BORDER=""0"" CELLPADDING=""3"" CELLSPACING=""0"" WIDTH=""100%"">" & VbCrLf
	bibliaJerusalen = Replace(bibliaJerusalen,"|","<BR>")
	bibliaJerusalen = Replace(bibliaJerusalen,"#","")
	' Change equal sign in CLASS statement to protect it
	bibliaJerusalen = Replace(bibliaJerusalen," CLASS=""small"""," CLASS+""small""")
	' Kill the rest of the equal signs that may appear
	bibliaJerusalen = Replace(bibliaJerusalen,"=","")
	' Change equal sign back
	bibliaJerusalen = Replace(bibliaJerusalen," CLASS+""small"""," CLASS=""small""")
	Response.Write bibliaJerusalen
	Response.Write "</TABLE>" & VbCrLf

	'Response.Write "</TD><TD WIDTH=""5%"">&nbsp;&nbsp;&nbsp;</TD><TD BGCOLOR=""#c0c0c0"" VALIGN=""TOP"" WIDTH=""25%"">"
	Response.Write "</TD><TD BGCOLOR=""#FFCC99"" VALIGN=""TOP"" WIDTH=""20%"">" & VbCrLf

	Response.Write "<TABLE BORDER=""0"" CELLPADDING=""3"" CELLSPACING=""0"" WIDTH=""100%"">" & VbCrLf
	Response.Write bibliaModerna
	Response.Write "</TABLE>" & VbCrLf

	Response.Write "</TD></TR></TABLE>" & VbCrLf
end if
%>
<P>&nbsp;</P>
<% CloseDB %>

</BODY>
</HTML>
