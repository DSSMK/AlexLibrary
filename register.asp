<body background="images/bg.jpg" bgproperties="fixed" text=white>
<%  if request.form("usr")="" then %>
        <b><h1> Please Type in your Login Name</h1></b>
         <center>
<input type="submit" value="Back To Alex Lib Home Page " onclick="window.history.back(1)">
</center>
<% response.end %>
<% end if %>
<%  if request.form("pass")="" then %>
        <b><h1> Please Type in your Password</h1></b>
         <center>
<input type="submit" value="Back To Alex Lib Home Page " onclick="window.history.back(1)">
</center>
<% response.end %>
<% end if %>
<%
Response.Buffer = true
dim cnn,rst
set cnn = Server.CreateObject("ADODB.Connection")
set rst = Server.CreateObject("ADODB.RecordSet")
Dim CONN_STRING
CONN_STRING = "DBQ=" & Server.MapPath("db/alex.mdb") & ";"
CONN_STRING = CONN_STRING & "Driver={Microsoft Access Driver (*.mdb)};"
cnn.Open CONN_STRING
sqltext = "SELECT * FROM members"
rst.Open sqltext,cnn,3,3
rst.addnew
rst("usr")=request.form("usr")
rst("pass")=request.form("pass")
rst("tel")=request.form("tel")
rst("email")=request.form("email")
rst("member")=request.form("name")
rst("address")=request.form("address")
rst.update
%>
<br><h1>CongratulationsYour Acount Info is</h1></b><br>
<center>
<b><h1>
user Name : <% =rst("usr")%><br>
Password  : <% =rst("pass") %><br></center>
</h1></b>
<% 
rst.close
cnn.close
set rst=nothing 
set cnn=nothing
%><br>
<center>
<input type="submit" value="Back To Alex Lib Home Page " onclick="window.history.back(1)">
</center>


<%
set cnn = Server.CreateObject("ADODB.Connection")
set rst = Server.CreateObject("ADODB.RecordSet")
CONN_STRING = "DBQ=" & Server.MapPath("db/alex.mdb") & ";"
CONN_STRING = CONN_STRING & "Driver={Microsoft Access Driver (*.mdb)};"
cnn.Open CONN_STRING
sqltext = "SELECT * FROM members where usr='"+usr+"';"
rst.Open sqltext,cnn,3,3 
if rst.eof <>true then 
k=rst("mid")
end if
rst.close
cnn.close
%>


</body>