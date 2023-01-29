<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="../Connections/NationStar.asp" -->

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<HEAD></HEAD>
<BODY>
<TABLE border=0 width=500 >
<TR><TD>
<%
'ON THE FIRST TIME THAT THIS PAGE LOADS, SESSION("ERRORS") HAS NO VALUE AND SO 
'EQUALS 0.  ON SUBSEQUENT VISITS, SESSION("ERRORS WILL HAVE A VALUE OF 1 OR 
'MORE IF ERRORS WERE MADE  

if  Session("Errors")=0 then 
  response.write "Please fill out the form"
else
  'ERRORS WERE MADE SO LIST THEM IN THE REST OF THE TABLE
  'reset our error counter
  Session("Errors")="0"
  response.write "<BR>There are errors in your data.  " & _
          "Please make  the following  changes before " & _
          "clicking the submit button:<br>"
  response.write"<TABLE border=0 width=' 400' align='center'>"

  'THESE SESSION VARIABLES ARE SET IN THE SECOND PAGE
  ' IF THERE WERE ERRORS 'VALUE IS "F"  
  If Session("badFirstName") = "T" then 
     Response.write "<TR><TD><font color='red'>The First " & _
                    "Name field must be completed.</font>" 
     Response.write "</TD></TR>"
     Session("badFirstName")="F"
  End If 


  If Session("badLastName") = "T" then
    Response.write "<TR><TD><font color='red'>The Last Name " & _
                   "field must be completed. </font>" 
    Response.write "</TD></TR>"
    Session("badLastName")="F"
  End If 

  If Session("badDate") = "T" then
    Response.write "<TR><TD><font color='red'>Date must be " & _
                   "in the format mm/dd/yyyy. </font>" 
    Response.write "</TD></TR>"
    Session("badDate")="F"
  End If 

  'END THE ERRORS TABLE
  response.write "</TD></TR></TABLE>"
End If
%>	

<!--START THE FORM -->
<FORM ACTION="test.asp" NAME="frmUser" METHOD="POST">

<!--FORM IS IN A TABLE TO ALIGN THE TEXTBOXES -->
<TABLE ALIGN="" BORDER=0 ALIGN="">
<TR>
<TD align=right><B>First Name</B></TD>
<TD>
  <INPUT TYPE="text" Name="FirstName" VALUE="<%=Session("FName")%>">
</TD>
</TR>

<TD align=right><B>Last Name</B></TD>
<TD>
  <INPUT TYPE="text" Name="LastName" VALUE="<%=Session("LName")%>">
</TD>
</TR>

<TR><TD align="right">
  <B>Date </B></td><TD><INPUT TYPE="Text" NAME="SendDate" 
             VALUE="<%=Session("CompletionDate") %>"> 
             <font color="blue">(mm/dd/yyyy)</font>
</TD></TR>
</TABLE>
<!-- end table for data entry form -->

<INPUT TYPE="Submit" VALUE="Submit">  <INPUT TYPE="RESET">
</FORM>

</TD></TR></TABLE>
<!-- end table used for page alignment -->
</BODY>
</HTML>


<%

Session("FName")=Request("FirstName")
Session("LName")=Request("LastName")
Session("CompletionDate")=Request("SendDate")

'HERE IS ONE WAY OF CHECKING FOR AN EMPTY TEXT BOX
if not len(Request("FirstName")) > 0 then
  Session("badFirstName")="T"
  Session("Errors")=Session("Errors") + 1
end if

'AND HERE IS ANOTHER, BOTH SHOULD WORK
if Request("LastName")= "" then
  Session("badLastName")="T"
  Session("Errors")=Session("Errors") + 1
end if


if not IsDate(Request("SendDate")) then
  Session("badDate")="T"
  Session("Errors")=Session("Errors") + 1
end if


if Session("Errors") > 0 then
  ‘there were errors, so send back to form
  'response.redirect "form.asp"
  response.write "fails"
else
  'there were no errors, so do the update to the database and redirect to a thank you page
  'response.redirect "thanks.asp"
  response.write "Thanks"
end if
%>

<%
CWCount.Close()
Set CWCount = Nothing
%>