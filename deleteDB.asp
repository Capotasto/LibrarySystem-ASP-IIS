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
Dim varID
varId = Request.Form ("remoteId")
Response.write varId

'check the books table

Rs.open "Select * from books where book_id = "+ CStr(varId), Con, adopenDynamic, adLockOptimistic
'Looking after ADO Empty Table bug

If Rs.eof = True And Rs.BOF = True Then
  Response.Cookies("screen") = "deleteDB.asp"
  Response.Cookies("alert") = "This book has already deleted!"
  Response.Redirect "./BookList.asp"

End If
Rs.MoveFirst
Rs.Delete
Rs.Update
Response.Cookies("screen") = "deleteDB.asp"
Response.Cookies("alert") = "Book deteting is Succeeded!"
Response.Redirect "./BookList.asp"
%>
