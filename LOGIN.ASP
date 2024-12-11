<body topmargin="0" leftmargin="0" background="images/bg.jpg" bgproperties="fixed" text=white>

<% Response.Buffer = true %>
<% 
dim cnn,rs
dim pass,usr
dim found 
set cnn = Server.CreateObject("ADODB.Connection")
set rs = Server.CreateObject("ADODB.RecordSet")
Dim CONN_STRING
CONN_STRING = "DBQ=" & Server.MapPath("db/alex.mdb") & ";"
CONN_STRING = CONN_STRING & "Driver={Microsoft Access Driver (*.mdb)};"
cnn.Open CONN_STRING
sqltext = "SELECT * FROM members"
rs.Open sqltext,cnn
do While not rs.eof 
    k=instr(rs("usr"),request.form("usr")) 
     if k>0 then  
      found=true
  pass=rs("pass")
  usr=rs("usr")
       if request.form("pass") <> pass then
          msg="<b><h2>Sorry Dear User.Please,Type Your Acount     </h2></b>" 
          msg=msg+"<center><input type=submit value=Back onclick=window.history.back(1)></center>"
           response.write(msg) 
        else
                         
              response.redirect "index.html"
     
          
           
       end if 

 end if
 rs.movenext 
 loop 

if not found then 
msg="<b><h2>Sorry Dear User.Please,Type Your assword </h2></b>"
msg=msg+"<center><input type=submit value=Back onclick=window.history.back(1)></center>"
response.write(msg) 

end if 

rs.close
cnn.close


%>


































