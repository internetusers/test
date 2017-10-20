




        <%@ language="vbscript"%>
        <%

Dim messagebody
Dim txtSubject
Dim promise
Dim email_address
Dim message
Dim semail_address
Dim squantity
Dim sflname
dim phone_number
Dim flname
dim sphone_number
dim zipcode

dim precinct      
dim phonecalls
dim sign
dim hostevent
dim donation

dim sdonation
dim sprecinct      
dim sphonecalls
dim ssign
dim shostevent
dim szipcode
dim zip

squantity = request.form("message")
semail_address = request.form("email_address")
sflname = request.form("flname")
sphone_number = request.form("phone_number")
szipcode = request.form("zipcode")

sdonation = request.form("donation")
sprecinct = request.form("precinct")
sphonecalls = request.form("phonecalls")
ssign = request.form("sign")
shostevent = request.form("hostevent")


%>
<!--METADATA TYPE="typelib" 
UUID="CD000000-8B95-11D1-82DB-00C04FB1625D" 
NAME="CDO for Windows 2000 Library" -->

<!--METADATA TYPE="typelib" 
UUID="00000205-0000-0010-8000-00AA006D2EA4" 
NAME="ADODB Type Library" -->
<%

Dim objMail 
Set objMail = Server.CreateObject("CDO.Message") 
Set objConfig = Server.CreateObject("CDO.Configuration") 

'Configuration: 
objConfig.Fields(cdoSendUsingMethod) = cdoSendUsingPort

objConfig.Fields(cdoSMTPServer)="smtp.1and1.com" 
objConfig.Fields(cdoSMTPServerPort)=25 
objConfig.Fields(cdoSMTPAuthenticate)=cdoBasic 
objConfig.Fields(cdoSendUserName) = "m39758306-1"
objConfig.Fields(cdoSendPassword) = "K3v1nC@rr"

'Update configuration 
objConfig.Fields.Update 
Set objMail.Configuration = objConfig 



messagebody = vbCrLf &  "Email Address: " & semail_address 



     
objMail.From ="kevin@internetusers.com" 
objMail.To =  "kevin@internetusers.com"
objMail.bcc ="kevin@internetusers.com" 
objMail.Subject ="Kevin Carr For Council " & semail_address
objMail.TextBody= messagebody
objMail.Send 


Response.Redirect "http://www.kevincarrforcouncil.com"




If Err.Number = 0 Then
  Response.Write("error")
Else
  Response.Write("Error sending mail. Code: " & Err.Number)
  Err.Clear
End If
Set objMail=Nothing 
Set objConfig=Nothing 


%>


<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>

<title></title>

<META http-equiv="Content-Type" content="text/html; charset=iso-8859-1"/>
<META name="ROBOTS" content="noINDEX, noFOLLOW">



<style type="text/css">
<!--
body {
	background-image: url(Crosshatch_grey_21x16.gif);
}
.font {
	font-family: Arial, Helvetica, sans-serif;
}
-->
</style>


</head>

<body>

Thank You 
        
        
         
</body>
</html>
