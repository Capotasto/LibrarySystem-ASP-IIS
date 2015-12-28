<!--#include file="env.asp"-->
<%
Dim Con
Dim RS
Set Con = Server.CreateObject("ADODB.Connection")
Set RS = Server.CreateObject("ADODB.Recordset")
Con.Provider = "Microsoft.Jet.OLEDB.4.0"
Con.ConnectionString = DB_PATH
Con.Open
Dim adopenDynamic
Dim adLockOptimistic
adopenDynamic = 2
adLockOptimistic = 3

'Get values user has enterd to the form
Dim varID, varImage, varTitle, varAuthor, varSummary, varPublisher, varPublishedDate, varLanguage, varWeight, varGenre
varId = Request.Form ("bookId")
varImage = Request.Form ("image")
varTitle = Request.Form ("title")
varAuthor = Request.Form ("author")
varSummary = Request.Form ("summary")
varPublisher = Request.Form("publisher")
varPublishedDate = Request.Form("date_pub")
varLanguage = Request.Form("language")
varWeight = Request.Form("weight")
varGenre = Request.Form("genre")

'check the authors table
Dim authorId
Rs.open "SELECT * FROM authors WHERE name1 = '" + varAuthor + "'", Con, adopenDynamic, adLockOptimistic
'Looking after ADO Empty Table bug
If Rs.eof = True And Rs.BOF = True Then
	' The author table is Empty
	Rs.AddNew  ' Creates a new empty row for me and sets the record pointer to the empty record
	'newAuthorId = Rs("author_id").value
	Rs.Fields ("name1") = varAuthor
	Rs.Update
End If
authorId = Rs("author_id").value
Rs.close

'check the publishers table
Dim publisherId
Rs.open "Select * from publishers where name ='"+ varPublisher +"'", Con, adopenDynamic, adLockOptimistic
'Looking after ADO Empty Table bug
If Rs.eof = True And Rs.BOF = True Then
	Rs.AddNew
    Rs.Fields("name") = varPublisher
    Rs.Update
End If
publisherId = Rs("pub_id").value
Rs.Close

'check the books table
Dim booksId

Rs.open "Select * from books where book_id = "+ CStr(varId), Con, adopenDynamic, adLockOptimistic
'Looking after ADO Empty Table bug
Rs.MoveFirst
If Rs.eof = True And Rs.BOF = True Then
    Response.Cookies("screen") = "EditDB.asp"
    Response.Cookies("message") = "This book has already deleted!"
    Response.Redirect "./EditBook.asp"

Else
    Rs.Fields("title") = varTitle
    Rs.Fields("author_id") = authorId
    Rs.Fields("image") = varImage
    Rs.Fields("summary") = varSummary
    Rs.Fields("pub_id") = publisherId
    Rs.Fields("date_published") = CDate(varPublishedDate)
    Rs.Fields("lang_id") = varLanguage
    Rs.Fields("weight") = varWeight
    Rs.Fields("genre_id") = varGenre
    Rs.Update
    Response.Cookies("screen") = "EditDB.asp"
    Response.Cookies("message") = "Book editting is Succeeded!"
    Response.Cookies("bookId") = CStr(varId)
    Response.Redirect "./EditBook.asp"

End If
Rs.Close
%>
