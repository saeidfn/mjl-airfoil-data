<%@ LANGUAGE="VBSCRIPT" %>

<%

IF Request.Form("renew_passw")="Renew" THEN

 const ForReading=1,ForWriting=2,ForAppending=8,lgnImportance=1

 dim FSys,listfile,usrfile
 dim pathlist,pathusrf,usrnb
 dim usr_nb,nb

 usr_nb = 0
     nb = 0

 Set FSys = Server.CreateObject("Scripting.FileSystemObject") 

 ' **************
 ' **************
 ' READ USER LIST

 pathlist = Left(Request.ServerVariables("PATH_TRANSLATED"),InStrRev(Request.ServerVariables("PATH_TRANSLATED"),"\"))&"..\..\..\_private\usrlist.txt"

 Set listfile = FSys.OpenTextFile(pathlist,ForReading,FALSE)

   usrnb           = listfile.ReadLine
   usr_nb = Cint(usrnb)

  dim   usr_name()      ,usr_email()      ,usr_company()      ,usr_id()      
  redim usr_name(usr_nb),usr_email(usr_nb),usr_company(usr_nb),usr_id(usr_nb)

  FOR nb=1 TO usr_nb

   usr_name(nb)    = listfile.ReadLine
   usr_email(nb)   = listfile.ReadLine
   usr_company(nb) = listfile.ReadLine
   usr_id(nb)      = listfile.ReadLine

  NEXT

 listfile.Close
 Set listfile = nothing

 ' **************
 ' **************
 ' READ USER LIST


  pathusrf = Left(Request.ServerVariables("PATH_TRANSLATED"),InStrRev(Request.ServerVariables("PATH_TRANSLATED"),"\"))&"..\..\..\_private\usrdata.txt"

 Set usrfile = FSys.OpenTextFile(pathusrf,ForWriting,FALSE)

   usrfile.WriteLine usr_nb

  FOR nb=1 TO usr_nb

   usrfile.WriteLine usr_name(nb)
   usrfile.WriteLine usr_email(nb)
   usrfile.WriteLine usr_company(nb)
   usrfile.WriteLine usr_id(nb)

  NEXT

 usrfile.Close
 Set usrfile = nothing

 Set FSys = nothing

 ' **************
 ' **************
 ' END OF RENEWAL

 Response.Redirect "http://www.risoe.dk/vea/profcat/WWW/HTML/USR/frba.asp"
 Response.end

ELSE

 Response.Redirect "http://www.risoe.dk/vea/profcat/WWW/HTML/USR/no.htm"
 Response.end

END IF

%>
