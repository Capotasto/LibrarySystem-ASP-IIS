<!DOCTYPE html>
<html>
<head>
  <meta charset="utf-8">
  <title>My Blog Page</title>
  <link rel="stylesheet" href="css/common.css" media="screen" title="no title" charset="utf-8">
</head>

<body>
  <div id="wrap">
    <!---header--------------------------------------------------->
    <div id="header">
      <p>
        <a href="./BookList.asp">
          <img src="img/logo.png" alt="" />
        </a>
      </p>
    </div>
    <!---mainmenu--------------------------------------------------->
    <div id="mainmenu">
      <ul>
        <li class="firstListItem"><a href="BookList.asp">Book List</a></li>
        <li><a href="EditBook.asp"><font color="black">Edit Book</font></a></li>
          <!--  <li><a href="gallery.html">Gallery Room</a></li>
        <li><a href="blog.html">Blog Page</a></li>
        <li><a href="contact.html">Contact Us</a></li> -->
      </ul>
    </div>
    <!---main_Content--------------------------------------------------->
    <div id="main_content">
      <div id="top_page_content">
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

        Dim varBookId
        varBookId = Request.Form("editId")

        If varBookId = "" Then
          varBookId = Request.Cookies("bookId")
        End If

        Rs.open "SELECT b.book_id, b.image, b.title, a.name1, b.summary, b.date_published, p.name, b.pages, l.lang_id, l.lang_name, b.weight, g.genre_id, g.name FROM ((((books  AS b INNER JOIN authors AS a ON b.author_id = a.author_id) INNER JOIN publishers AS p ON b.pub_id = p.pub_id) INNER JOIN  languages  AS l ON b.lang_id = l.lang_id) INNER JOIN  genres AS g ON b.genre_id = g.genre_id) WHERE b.book_id =" + CStr(varBookId), _
                   Con, adopenDynamic, adLockOptimistic
        Rs.MoveFirst
        'Response.Write "Edit this book infomations. And click the Confirm button"
        %>

        <table >
         <Form Method = "post" Action = "EditDB.asp">
        <input type="hidden" name="bookId" Id="bookId" value='<% Response.Write Rs.Fields("book_id") %>' >
        <Tr>
          <tD>Book Image :</TD>
          <TD><input type= "text" Name = "image" ID = "image" value = '<% Response.Write Rs.Fields("image") %>' ></TD>
        </Tr>
       	<Tr>
       		<tD>Book Title :</TD>
       		<TD><input type= "text" Name = "title" ID = "title" value = '<% Response.Write Rs.Fields("title") %>' ></TD>
       	</Tr>
       	<TR>
       		<TD>Author :</TD>
       		<TD><input type= "text" Name = "author" ID = "author" value = '<% Response.Write RS.Fields("name1") %>'></TD>
       	</Tr>
       	<TR>
       		<TD>Summary :</TD>
       		<TD><input type= "text" Name = "summary" ID = "summary" value = '<% Response.Write RS.Fields("summary") %>'></TD>
       	</Tr>
       	<TR>
       		<TD>Publisher :</TD>
       		<TD><input type= "text" Name = "publisher" ID = "publisher" value = '<% Response.Write RS.Fields("p.name") %>'></TD>
       	</Tr>
       	<TR>
       		<TD>Published Date :</TD>
       		<TD><input type= "date" Name = "date_pub" ID = "date_pub" value = '<% Response.Write Replace(RS.Fields("date_published"), "/","-") %>'></TD>
       	</Tr>
        <TR>
       		<TD>Language :</TD>
       		<TD>
            <select Name= "language">
              <%
              If 1 = CInt(RS.Fields("lang_id")) Then
                Response.Write "<option value='1' selected>English</option>"
                Response.Write "<option value='2'>Japanese</option>"
                Response.Write "<option value='3'>Portuguese</option>"
                Response.Write "<option value='4'>Spanish</option>"
              ElseIf 2 = CInt(RS.Fields("lang_id")) Then
                Response.Write "<option value='1'>English</option>"
                Response.Write "<option value='2' selected>Japanese</option>"
                Response.Write "<option value='3'>Portuguese</option>"
                Response.Write "<option value='4'>Spanish</option>"
              ElseIf 3 = CInt(RS.Fields("lang_id")) Then
                Response.Write "<option value='1'>English</option>"
                Response.Write "<option value='2'>Japanese</option>"
                Response.Write "<option value='3' selected>Portuguese</option>"
                Response.Write "<option value='4'>Spanish</option>"
              ElseIf 4 = CInt(RS.Fields("lang_id")) Then
                Response.Write "<option value='1'>English</option>"
                Response.Write "<option value='2'>Japanese</option>"
                Response.Write "<option value='3'>Portuguese</option>"
                Response.Write "<option value='4' selected>Spanish</option>"
              End If
              %>
            </select>
          </TD>
       	</Tr>
        <TR>
       		<TD>Weight :</TD>
       		<TD><input type= "text" Name = "weight" ID = "weight" value="<% Response.Write RS.Fields("weight") %>"></TD>
       	</Tr>
        <TR>
       		<TD>Genre :</TD>
       		<TD>
            <select Name= "genre">
              <%
              If 1 = CInt(RS.Fields("genre_id")) Then
                Response.Write "<option value='1' selected>Technorogy</option>"
                Response.Write "<option value='2'>Fantasy</option>"
                Response.Write "<option value='3'>Adventure</option>"
                Response.Write "<option value='4'>Education</option>"
                Response.Write "<option value='5'>Children</option>"
                Response.Write "<option value='6'>Other</option>"
              ElseIf 2 = CInt(RS.Fields("genre_id")) Then
                Response.Write "<option value='1'>Technorogy</option>"
                Response.Write "<option value='2' selected>Fantasy</option>"
                Response.Write "<option value='3'>Adventure</option>"
                Response.Write "<option value='4'>Education</option>"
                Response.Write "<option value='5'>Children</option>"
                Response.Write "<option value='6'>Other</option>"
              ElseIf 3 = CInt(RS.Fields("genre_id")) Then
                Response.Write "<option value='1'>Technorogy</option>"
                Response.Write "<option value='2'>Fantasy</option>"
                Response.Write "<option value='3' selected>Adventure</option>"
                Response.Write "<option value='4'>Education</option>"
                Response.Write "<option value='5'>Children</option>"
                Response.Write "<option value='6'>Other</option>"
              ElseIf 4 = CInt(RS.Fields("genre_id")) Then
                Response.Write "<option value='1'>Technorogy</option>"
                Response.Write "<option value='2'>Fantasy</option>"
                Response.Write "<option value='3'>Adventure</option>"
                Response.Write "<option value='4' selected>Education</option>"
                Response.Write "<option value='5'>Children</option>"
                Response.Write "<option value='6'>Other</option>"
              ElseIf 5 = CInt(RS.Fields("genre_id")) Then
                Response.Write "<option value='1'>Technorogy</option>"
                Response.Write "<option value='2'>Fantasy</option>"
                Response.Write "<option value='3'>Adventure</option>"
                Response.Write "<option value='4'>Education</option>"
                Response.Write "<option value='5' selected>Children</option>"
                Response.Write "<option value='6'>Other</option>"
              ElseIf 6 = CInt(RS.Fields("genre_id")) Then
                Response.Write "<option value='1'>Technorogy</option>"
                Response.Write "<option value='2'>Fantasy</option>"
                Response.Write "<option value='3'>Adventure</option>"
                Response.Write "<option value='4'>Education</option>"
                Response.Write "<option value='5'>Children</option>"
                Response.Write "<option value='6' selected>Other</option>"
              End If
              %>
          </TD>
       	</Tr>
       	<Tr>
       		<TD><input type ="Submit" Value = "Confrim"></TD>
            <%
                Dim screen
                Dim message
                screen = Request.Cookies("screen")
                message = Request.Cookies("message")
                Response.Write message
                Response.Cookies("screen") = "EditBook.asp"
                Response.Cookies("message") = "<br/>"
            %>
       	</Tr>
         </form>
         </table>
         <hr/>
        <table>
          <tr>
            <td>
              <a href="#">
                <img src="img/ads.jpg" alt="ads.jpg" />
              </a>
            </td>
            <td>
              <a href="#">
                <img src="img/ads2.jpg" alt="ads2.jpg" />
              </a>
            </td>
            <td>
              <a href="#">
                <img src="img/ads3.jpg" alt="ads3.jpg" />
              </a>
            </td>
          </tr>
        </table>
      </div>
    </div>
    <!---footer--------------------------------------------------->
    <div id="footer">
      <table>
        <tr>
          <th>About</th>
          <th>Background</th>
          <th>Recommendation</th>
          <th>Social Networking</th>
        </tr>
        <tr>
          <td><a href="#">Profile</a></td>
          <td><a href="#">Job Experience A</a></td>
          <td><a href="#">@MyFriendsA Tech Blog</a></td>
          <td><a href="#">Twitter</a></td>
        </tr>
        <tr>
          <td></td>
          <td><a href="#">Job Experience B</a></td>
          <td><a href="#">@MyFriendsB Tech Blog</a></td>
          <td><a href="#">Facebook</a></td>
        </tr>
        <tr>
          <td></td>
          <td><a href="#">Job Experience C</a></td>
          <td><a href="#">@MyFriendsC Tech Blog</a></td>
          <td><a href="#">LinkedIn</a></td>
        </tr>
      </table>
    </div>
    <!---copyright--------------------------------------------------->
    <div id="copyright">
      <p>
        Copyright &copy; 2015 <a href="./index.html">My Blog Page</a> All Rights Reserved.
      </p>
    </div>
  </div>
</body>

</html>
