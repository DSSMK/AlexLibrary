<body background="images/bg.jpg" bgproperties="fixed" text=white>
<%  if request.form("subid")="" then %>
        <b><h1> Please Type in your MemberShip Id To Activate Your Membership Period</h1></b>
       <center>
<input type="submit" value="Back To Alex Lib Home Page " onclick="window.history.back(1)">
</center>
<% response.end %>
<% end if %>
<%
a=request.form("subid")
Response.Buffer = true
dim cnn,rst
set cnn = Server.CreateObject("ADODB.Connection")
set rst = Server.CreateObject("ADODB.RecordSet")
Dim CONN_STRING
CONN_STRING = "DBQ=" & Server.MapPath("db/alex.mdb") & ";"
CONN_STRING = CONN_STRING & "Driver={Microsoft Access Driver (*.mdb)};"
cnn.Open CONN_STRING
sqltext = "SELECT * FROM members where member='"&a &"'"
rst.Open sqltext,cnn,3,3
%>
<% if rst.eof   then %>
<center>
<br><h1>Sorry, You are not a Member</h1></b><br>
</center>
<% else %>
<br><h1> Your Membership Info is</h1></b><br>
<center>
<b><h3>
Start Date : <% =rst("subsdate")%><br>
End Date  : <% =rst("subedate") %><br></center>
</h3></b>
<% end if %>
<% 
rst.close
cnn.close
set rst=nothing 
set cnn=nothing
%><br>
<center>
<input type="submit" value="Back To Alex Lib Home Page " onclick="window.history.back(1)">
</center>
</body>