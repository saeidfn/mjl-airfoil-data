<%@ LANGUAGE="VBSCRIPT" %>

<html>
<head>
<title>USERS DATA</title>
<link rel="STYLESHEET" href="../mystyle.css" type="text/css">
</head>

<body BGPROPERTIES="FIXED"
face="ARIAL" size="3" text="#000000" bgcolor="#cccccc">

<br><br>
<br>
<p align="center"><font size="+1"><b>Data on the users of the
Wind Turbine Airfoil Catalogue website</b></font></p><br>
<p align="center"><b><font
color="#ff0000">!!!Restricted page!!!</font></b></p>

<%

 const ForReading=1

 dim FSys,usrfile
 dim pathusrf,usrnb,usr_name,usr_email,usr_company,usr_id
 dim usr_nb,nb

 usr_nb = 0
     nb = 0

 pathusrf = Left(Request.ServerVariables("PATH_TRANSLATED"),InStrRev(Request.ServerVariables("PATH_TRANSLATED"),"\"))&"..\..\..\_private\usrdata.txt"

 Set FSys = Server.CreateObject("Scripting.FileSystemObject")

 Set usrfile = FSys.OpenTextFile(pathusrf,ForReading,FALSE)

  usrnb       = usrfile.ReadLine
  usr_nb = Cint(usrnb)

%>

<br>
<p>There are <% =usr_nb %> users.</p>

<TABLE BORDER="3" CELLSPACING="3" CELLPADDING="3" WIDTH ="100%">
<CAPTION> <b>USERS DATA</b> </CAPTION>

<TR><TH> User Name </TH>
    <TH> User e-mail </TH>
    <TH> User Company </TH>
    <TH> IP address </TH></TR>

<%
 FOR nb=1 TO usr_nb

  usr_name    = usrfile.ReadLine
  usr_email   = usrfile.ReadLine
  usr_company = usrfile.ReadLine
  usr_id      = usrfile.ReadLine
%>


<TR><TD> <% =usr_name %>    </TD>
    <TD> <% =usr_email %>   </TD>
    <TD> <% =usr_company %> </TD>
    <TD> <% =usr_id %>      </TD></TR>

<%
 NEXT

 usrfile.Close
 Set usrfile = nothing
%>

</TABLE>

<br><br>

<p>Continue to
the <a href="http://www.risoe.dk/vea/profcat/WWW/HTML/index.htm">Wind
Turbine Airfoil Catalogue home page</a>.
</p>

<br>

<A HREF="http://www.risoe.dk/vea/Research/research.htm"
><B>Back to Wind Energy and Atmospheric Physics - Research
Activities</B></A>

<br><br><br>

<p>
<font size="+2"><b><font color="#ff0000">Information</font> on
"usrdata.txt" file</b></font>
</p>

<p><b>Path to file "<em>usrdata.txt</em>":</b><br>
<% = pathusrf %></p><br>

<p><b>Replica of file "<em>usrdata.txt</em>":</b></p>

<font size="-1" color="#FF00FF">

<%
 Set usrfile = FSys.OpenTextFile(pathusrf,ForReading,FALSE)

  usrnb       = usrfile.ReadLine
  usr_nb = Cint(usrnb)
%>
<% =usr_nb %>      <br>
<%
 FOR nb=1 TO usr_nb

  usr_name    = usrfile.ReadLine
  usr_email   = usrfile.ReadLine
  usr_company = usrfile.ReadLine
  usr_id      = usrfile.ReadLine
%>
<% =usr_name %>    <br>
<% =usr_email %>   <br>
<% =usr_company %> <br>
<% =usr_id %>      <br>

<%
 NEXT

 usrfile.Close
 Set usrfile = nothing

 Set FSys = nothing
%>

</font>

<br><br><br>

<p>
<font size="+2"><b><font color="#ff0000">Renew</font> "usrdata.txt" file</b></font>
</p>

<p>The "usrdata.txt" file will be replaced by the "usrlist.txt" file
present in the folder "_private". <A href="./renew.htm"><b>
Proceed here...</b></A></p>

<br><br>

</body>
</html>
