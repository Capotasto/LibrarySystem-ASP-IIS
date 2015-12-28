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
        <li class="firstListItem"><a href="BookList.asp"><font color="black">Book List</font></a></li>
        <li><a href="AddBook.asp">Add Book</a></li>
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
        Response.Cookies("screen") = "BookList.asp"
        Response.Cookies("message") = "<br/>"
        Dim alertMsg
        alertMsg = Request.Cookies("alert")
        If alertMsg <> "" Then
          Response.Write("<script>alert('"+ alertMsg +"')</script>")
          Response.Cookies("alert") = ""
        End If
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
        Rs.open "SELECT b.book_id, b.image, b.title, a.name1, b.summary, b.date_published, p.name, b.pages, l.lang_name, b.weight, g.name FROM ((((books  AS b INNER JOIN authors AS a ON b.author_id = a.author_id) INNER JOIN publishers AS p ON b.pub_id = p.pub_id) INNER JOIN  languages  AS l ON b.lang_id = l.lang_id) INNER JOIN  genres AS g ON b.genre_id = g.genre_id)", _
                   Con, adopenDynamic, adLockOptimistic
        'Response.write " So far the connection stuff is all right"
        Rs.MoveFirst
        Do While Not RS.EOF
        Response.Write "<div class='list_item'>"
          Response.Write "<div class='left_side'>"
            Response.Write "<img src='./img/" + RS.Fields("image")+ "' alt='' />"
          Response.Write "</div>"
          Response.Write "<div class='right_side'>"
            Response.Write "<p class='title'>" + RS.Fields("title") + "</p>"
            Response.Write "<p class='author'>by " + RS.Fields("name1")+ "</p>"
            Response.Write "<div class='left_inside'>"
              Response.Write "<p class='genre'>Genre: " + RS.Fields("g.name")+ "</p>"
              Response.Write "<p class='publisher'>Publisher: "+ RS.Fields("p.name") +"</p>"
              Response.Write "<p class='language'>Language: "+ RS.Fields("lang_name") +"</p>"
            Response.Write "</div>"
            Response.Write "<div class='right_inside'>"
              Response.Write "<p class='pub_date'>Published Date: "+ CStr(RS.Fields("date_published"))+"</p>"
              Response.Write "<p class='pages'>Pages: "+ CStr(RS.Fields("pages")) +"</p>"
              Response.Write "<p class='weight'>Weight: "+ CStr(RS.Fields("weight")) +" g</p>"
            Response.Write "</div>"
            Response.Write "<div class='bottom_inside'>"
              Response.Write "<p class='summary'>"+ RS.Fields("summary") +"</p>"
            Response.Write "</div>"
            <!-- Edit button -->
            Response.Write "<FORM class='form_edit' ACTION='EditBook.asp' METHOD='POST'>"
            Response.Write "<INPUT type='hidden' NAME='editId' VALUE= '" + CStr(RS.Fields("book_id")) + "'>"
            Response.Write "<INPUT TYPE='submit' VALUE='Edit'>"
            Response.Write "</FORM>"
            <!-- Remove button -->
            Response.Write "<FORM class='form_remove' ACTION='deleteDB.asp' METHOD='POST'>"
            Response.Write "<INPUT type='hidden' NAME='remoteId' VALUE= '" + CStr(RS.Fields("book_id")) + "'>"
            Response.Write "<INPUT  TYPE='submit' VALUE='Remove'>"
            Response.Write "</FORM>"
          Response.Write "</div>"
        Response.Write "</div>"
        Response.Write "<hr/>"
        Rs.MoveNext
        Loop
        %>
        <!---add--------------------------------------------------->
        <table id="add_content">
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
