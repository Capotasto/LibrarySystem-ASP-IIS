<%
Dim Con
Dim RS
Set Con = Server.CreateObject("ADODB.Connection")
Set RS = Server.CreateObject("ADODB.Recordset")
Con.Provider = "Microsoft.Jet.OLEDB.4.0"
'Con.ConnectionString = "C:\Users\norio.egi\Documents\My Web Sites\WebSite1\project\Library.mdb"
Con.ConnectionString = "\\Mac\Home\Documents\My Web Sites\WebSite1\project\Library.mdb"
Con.Open
Dim adopenDynamic
Dim adLockOptimistic
adopenDynamic = 2
adLockOptimistic = 3
'Rs.open "Select * from students", Con, adopenDynamic, adLockOptimistic
Rs.open "Select * f", Con, adopenDynamic, adLockOptimistic
'Response.write " So far the connection stuff is all right"
Rs.MoveFirst
'Reading data from the form

Dim varTitle, varAuthor, varSummary, varPublisher, varPublishedDate, varLanguage, varWeight, varGenre
varTitle = Request.Form ("title")
varAuthor = Request.Form ("author")
varSummary = Request.Form ("summary")
varPublisher = Request.Form("publisher")
varPublishedDate = Request.Form("date_pub")
varLanguage = Request.Form("language")
varWeight = Request.Form("weight")
varGenre = Request.Form("genre")

'Response.Write varTitle
'Response.Write "<BR>"
'Response.Write varAuthor
'Response.Write "<BR>"
'Response.Write varPublisher
'Response.Write "<BR>"
'Response.Write varPublishedDate
'Response.Write "<BR>"
'Response.Write varLanguage
'Response.Write "<BR>"
'Response.Write varWeight
'Response.Write "<BR>"
'Response.Write varGenre
'Response.Write "<BR>"

'Looking after ADO Empty Table bug
If Rs.eof = True And Rs.BOF = True Then
' The books table is Empty


End If
'Check the title is dupulicated
Dim titleDuplicated As Boolean
titleDuplicated = false
Do While Not RS.EOF
If RS.Fields("title") = varTitle Then
	titleDuplicated = true
End if
Rs.MoveNext
Loop

'Looking after ADO Empty Table bug
If Rs.eof = True And Rs.BOF = True Then
' The table is Empty
	Rs.AddNew  ' Creates a new empty row for me and sets the record pointer to the empty record
	Rs.Fields ("ID") = VarID
	Rs.Fields ("FirstName") = VarFirstName
	Rs.Fields ("SurName") = VarSurName
	Rs.Fields ("Age") = VarAge
	Rs.Fields ("Nationality") = VarNationality
	Rs.Update
End If
'look for duplicate record
Dim CriteriaString
CriteriaString = "ID = " + VarID
'Response.Write CriteriaString
Rs.Find CriteriaString
If Rs.Eof Then
	' Record Not found
	Rs.AddNew  ' Creates a new empty row for me and sets the record pointer to the empty record
	Rs.Fields ("ID") = VarID
	Rs.Fields ("FirstName") = VarFirstName
	Rs.Fields ("SurName") = VarSurName
	Rs.Fields ("Age") = VarAge
	Rs.Fields ("Nationality") = VarNationality
	Rs.Update
	Response.Write "Record Added Successfully, Please go back to add another record"
Else
	' Record found, It's Duplicate
	Response.Write "Duplicate Record"
	Response.Write "<br>"
	Response.Write "Please go back and modify your ID"
End if
%>
