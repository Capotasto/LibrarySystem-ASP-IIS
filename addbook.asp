﻿<!DOCTYPE html>
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
        <li><a href="AddBook.asp"><font color="black">Add Book</font></a></li>
          <!--  <li><a href="gallery.html">Gallery Room</a></li>
        <li><a href="blog.html">Blog Page</a></li>
        <li><a href="contact.html">Contact Us</a></li> -->
      </ul>
    </div>
    <!---main_Content--------------------------------------------------->
    <div id="main_content">
      <div id="top_page_content">

        <table >
         <Form Method = "post" Action = "AddtoDb.asp">
       	<Tr>
       		<tD>Book Title :</TD>
       		<TD><input type= "text" Name = "title" ID = "title"></TD>
       	</Tr>
       	<TR>
       		<TD>Author :</TD>
       		<TD><input type= "text" Name = "author" ID = "author"></TD>
       	</Tr>
       	<TR>
       		<TD>Summary :</TD>
       		<TD><input type= "text" Name = "summary" ID = "summary"></TD>
       	</Tr>
       	<TR>
       		<TD>Publisher :</TD>
       		<TD><input type= "text" Name = "publisher" ID = "publisher"></TD>
       	</Tr>
       	<TR>
       		<TD>Published Date :</TD>
       		<TD><input type= "date" Name = "date_pub" ID = "date_pub"></TD>
       	</Tr>
        <TR>
       		<TD>Language :</TD>
       		<TD>
            <select Name= "language">
              <option value="1">English</option>
              <option value="2">Japanese</option>
              <option value="3">Portuguese</option>
              <option value="4">Spanish</option>
            </select>
          </TD>
       	</Tr>
        <TR>
       		<TD>Weight :</TD>
       		<TD><input type= "text" Name = "weight" ID = "weight"></TD>
       	</Tr>
        <TR>
       		<TD>Genre :</TD>
       		<TD>
            <select Name= "genre">
              <option value="1">Technorogy</option>
              <option value="2">Fantasy</option>
              <option value="3">Adventure</option>
              <option value="4">Education</option>
              <option value="5">Children</option>
              <option value="6">Other</option>
            </select>
          </TD>
       	</Tr>
       	<Tr>
       		<TD><input type ="Submit" Value = "Add info to table"></TD>
            <%
                Dim screen
                Dim message
                screen = Request.Cookies("screen")
                message = Request.Cookies("message")
                Response.Write message
                Response.Cookies("screen") = "addbook.asp"
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
