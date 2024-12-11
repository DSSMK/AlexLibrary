<head>

</head>
<body topmargin="0" leftmargin="0" background="images/bg.jpg" bgproperties="fixed" text=white>
<%
Response.Buffer = true
dim cnn,rst
%>
<% 
   dim srch,tpe
   srch=Request.Form("srchtxt")
   tpe=Request.Form("type") 
   if srch = "" then
   error = "<b><h2>Please Type in  The Words to be searched.</h2></b><br>"
   Response.Write error
   end if
%>

<%

set cnn = Server.CreateObject("ADODB.Connection")
set rst = Server.CreateObject("ADODB.RecordSet")
Dim CONN_STRING
CONN_STRING = "DBQ=" & Server.MapPath("db/alex.mdb") & ";"
CONN_STRING = CONN_STRING & "Driver={Microsoft Access Driver (*.mdb)};"
cnn.Open CONN_STRING

dim ysrch

%>




<% select case tpe %>

<% case "author" %>

<%
sqltext = "SELECT * FROM books where author='"&srch&"' "
rst.Open sqltext,cnn,3,3 
IF NOT rst.EOF THEN

    book=rst("book")
   publisher=rst("publisher")
   edate=rst("edate")
   field=rst("field")
   pagecount=rst("pagecount")
   copies=rst("copies")
   found=true
  %>
<TABLE bgColor=#ffffff >
  <TBODY>
  <TR bgColor=#99ccff>
    <TH bgColor=#99ccff><FONT 
      color=#0000cc>гАъйгх</FONT></TH>
    <TH bgColor=#99ccff><FONT 
      color=#0000cc>гАДгтя</FONT></TH>
    <TH bgColor=#99ccff><FONT 
      color=#0000cc>зоо гАущмгй</FONT></TH>
    <TH bgColor=#99ccff><FONT 
      color=#0000cc>зоо гАДсн</FONT></TH>
    <TH bgColor=#99ccff><FONT 
      color=#0000cc>гАйнуу</FONT></TH>
  </TR>

<tr bgColor=#99c00f><td><%=book%></td><td><%=publisher%></td><td><center><%=pagecount%></td><td><%=copies%></center></td><td><%=field%></td>


</TBODY></TABLE>
<form method="post">
<center>
<input type="submit" value="Back" onclick="window.history.back(2)">
</center>
</form>

<% else %>
Аг йФло Дйгфл ААхмк

<%

end if

%>






<% case "book" %> 
<%
sqltext = "SELECT * FROM books where book='"&srch&"' "
rst.Open sqltext,cnn,3,3 
IF NOT rst.EOF THEN

   author=rst("author")
   publisher=rst("publisher")
   edate=rst("edate")
   field=rst("field")
   pagecount=rst("pagecount")
   copies=rst("copies")
                                        
 %>
<TABLE bgColor=#ffffff >
  <TBODY>
  <TR bgColor=#99ccff>
    <TH bgColor=#99ccff><FONT 
      color=#0000cc>гАЦдАщ</FONT></TH>
    <TH bgColor=#99ccff><FONT 
      color=#0000cc>гАДгтя</FONT></TH>
    <TH bgColor=#99ccff><FONT 
      color=#0000cc>зоо гАущмгй</FONT></TH>
    <TH bgColor=#99ccff><FONT 
      color=#0000cc>зоо гАДсн</FONT></TH>
    <TH bgColor=#99ccff><FONT 
      color=#0000cc>гАйнуу</FONT></TH>
  </TR>

<tr bgColor=#99c00f><td><%=author%></td><td><%=publisher%></td><td><center><%=pagecount%></td><td><%=copies%></center></td><td><%=field%></td>


</TBODY></TABLE>

<form method="post">
<center>
<input type="submit" value="Back" onclick="window.history.back(2)">
</center>
</form>
<% else %>
Аг йФло Дйгфл ААхмк

<%

end if

%>

<% end select %>
