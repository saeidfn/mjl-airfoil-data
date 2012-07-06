<%@ LANGUAGE="VBSCRIPT" %>

<%

 ' **************
 ' **************
 ' INITIALISATION

 const ForReading=1,ForWriting=2,ForAppending=8,lgnImportance=1

 dim usrfile,admfile,FSys,ObjAdmMail,ObjUsrMail
 dim pathusrf,pathadmi,MyEmail,EmailSubj,EmailText
 dim user_name,user_email,user_company,user_id,usrnb
 dim usr_dim,usr_nb,nb

  usr_dim = 0
  usr_nb  = 0
      nb  = 0

  Session("SES_username") = Request.Form("usrname")
  Session("SES_permission") = "NOT"

  pathadmi = Left(Request.ServerVariables("PATH_TRANSLATED"),InStrRev(Request.ServerVariables("PATH_TRANSLATED"),"\"))&"admin_addr.txt"

  pathusrf = Left(Request.ServerVariables("PATH_TRANSLATED"),InStrRev(Request.ServerVariables("PATH_TRANSLATED"),"\"))&"..\..\..\_private\usrdata.txt"


 '********************
 '********************
 ' READ USER DATA FILE

 Set FSys = Server.CreateObject("Scripting.FileSystemObject")

 ' READ ADMINISTRATOR ADDRESS

 Set admfile = FSys.OpenTextFile(pathadmi,ForReading,FALSE)

     MyEmail = admfile.ReadLine

 admfile.Close
 Set admfile = nothing

 ' READ USER DATA FILE

 Set usrfile = FSys.OpenTextFile(pathusrf,ForReading,FALSE)

     usrnb           = usrfile.ReadLine

  usr_nb  = Cint(usrnb)
  usr_dim = usr_nb + 1

 dim   usr_name()       ,usr_email()       ,usr_company()       ,usr_id()       
 redim usr_name(usr_dim),usr_email(usr_dim),usr_company(usr_dim),usr_id(usr_dim)

 FOR nb=1 TO usr_nb

     usr_name(nb)    = usrfile.ReadLine
     usr_email(nb)   = usrfile.ReadLine
     usr_company(nb) = usrfile.ReadLine
     usr_id(nb)      = usrfile.ReadLine

 NEXT

 usrfile.Close
 Set usrfile = nothing

 Set FSys = nothing


 ' ***********************
 ' ***********************
 ' FIRST TIME REGISTRATION

IF Request.Form("action")="register" THEN

  user_name    = Request.Form("usrname")
  user_email   = Request.Form("email")
  user_company = Request.Form("company")
  user_id      = Request.ServerVariables("REMOTE_ADDR")

   ' ***************
   ' Check user data
    IF user_name<>"" AND user_email<>"" AND user_company<>"" THEN

   '*****************************
   ' Check if available user name

  nb = 0

 DO WHILE (nb < usr_nb) AND (Session("SES_permission")<>"NOA")

  nb = nb + 1

  IF user_name=usr_name(nb) THEN
   Session("SES_permission") = "NOA"
  END IF

 LOOP

   '************************
   ' Write user data in file
     IF Session("SES_permission")<>"NOA" THEN

  usr_name(usr_dim)    = user_name
  usr_email(usr_dim)   = user_email
  usr_company(usr_dim) = user_company
  usr_id(usr_dim)      = user_id

 Set FSys = Server.CreateObject("Scripting.FileSystemObject") 

 Set usrfile = FSys.OpenTextFile(pathusrf,ForWriting,FALSE)

     usrfile.WriteLine usr_dim

 FOR nb=1 TO usr_dim

     usrfile.WriteLine usr_name(nb)
     usrfile.WriteLine usr_email(nb)
     usrfile.WriteLine usr_company(nb)
     usrfile.WriteLine usr_id(nb)

 NEXT

 usrfile.Close
 Set usrfile = nothing
 Set FSys = nothing

 Session("SES_permission") = "YES"

   '**********************
   ' Send new user warning

   ' To ADMINISTRATOR

 Set ObjAdmMail = Server.CreateObject("CDONTS.NewMail")

 EmailSubj="New User of Wind Turbine Airfoil Catalogue"
 EmailText="NAME=" & user_name & " + EMAIL=" & user_email & " + COMPANY=" & user_company & vbCrLf
 EmailText=EmailText & "IP address=" & user_id

 ObjAdmMail.Send user_email,MyEmail,EmailSubj,EmailText,lgnImportance

 Set ObjAdmMail = nothing

   ' To USER

 Set ObjUsrMail = Server.CreateObject("CDONTS.NewMail")

 EmailSubj="Wind Turbine Airfoil Catalogue Registration"
 EmailText=            "You are now a registered user of the"                    & vbCrLf
 EmailText=EmailText & "Wind Turbine Airfoil Catalogue website."                 & vbCrLf & vbCrLf
 EmailText=EmailText & "Your login name is: " & user_name                        & vbCrLf & vbCrLf
 EmailText=EmailText & "You are welcome to contact us at the following address:" & vbCrLf
 EmailText=EmailText & MyEmail                                                   & vbCrLf & vbCrLf & vbCrLf
 EmailText=EmailText & "Sincerely,"                                              & vbCrLf & vbCrLf
 EmailText=EmailText & "The Wind Turbine Airfoil Catalogue authors"              & vbCrLf & vbCrLf
 EmailText=EmailText & "http://www.risoe.dk/vea/profcat"

 ObjUsrMail.Send MyEmail,user_email,EmailSubj,EmailText,lgnImportance

 Set ObjUsrMail = nothing

     END IF


   ' ***************
   ' Data Missing
    ELSE

 Session("SES_permission")="MIS"

    END IF


 ' ***********************
 ' ***********************
 ' USER ALREADY REGISTERED

ELSE

 user_name    = Request.form("usrname")
 user_company = Request.form("company")

 Session("SES_permission") = "UNK"

  nb = 0

 DO WHILE (nb < usr_nb) AND (Session("SES_permission")<>"YES")

  nb = nb + 1

  IF user_name=usr_name(nb) THEN
   Session("SES_permission") = "YES"
   Session("SES_username")   = usr_name(nb)
  END IF

 LOOP

END IF


 ' ***********************
 ' ***********************
 ' REDIRECT TO PROPER PAGE

perm = Session("SES_permission")
SELECT CASE perm
CASE "YES"
 IF (user_name="frba") THEN
 '*************
 ' test company
   IF user_company=usr_company(1) THEN
 Response.Redirect "http://www.risoe.dk/vea/profcat/WWW/HTML/USR/frba.asp"
 Response.end
   ELSE
 Response.Redirect "http://www.risoe.dk/vea/profcat/WWW/HTML/USR/no.htm"
 Response.end
   END IF
 ELSE
 Response.Redirect "http://www.risoe.dk/vea/profcat/WWW/HTML/index.htm"
 Response.end
 END IF
CASE "MIS"
 Response.Redirect "http://www.risoe.dk/vea/profcat/reg_form.asp"
 Response.end
CASE "UNK"
 Response.Redirect "http://www.risoe.dk/vea/profcat/reg_form.asp"
 Response.end
CASE "NOA"
 Response.Redirect "http://www.risoe.dk/vea/profcat/reg_form.asp"
 Response.end
CASE ELSE
 Response.Redirect "http://www.risoe.dk/vea/profcat"
 Response.end
END SELECT

%>
