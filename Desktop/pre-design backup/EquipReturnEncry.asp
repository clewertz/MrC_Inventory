<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="Connections/NationStar.asp" -->

<%
Dim Groups__DInsite
Groups__DInsite = "%"
If (Request.Form("DInsite") <> "") Then 
  Groups__DInsite = Request.Form("DInsite")
End If
%>


<%
Dim DInEncryp
Dim DInEncryp_cmd
Dim DInEncryp_numRows

Set DInEncryp_cmd = Server.CreateObject ("ADODB.Command")
DInEncryp_cmd.ActiveConnection = MM_NationStar_STRING
DInEncryp_cmd.CommandText = "SELECT SerialNum, ComputerName, DInAssignedUser, DInModel, DInStatus, EnStatus, DInSite FROM dbo.vDInInvBHead ORDER BY DInSite, EnStatus" 
DInEncryp_cmd.Prepared = true

Set DInEncryp = DInEncryp_cmd.Execute
DInEncryp_numRows = 0
%>

<%
Dim DInEncryptest
Dim DInEncryptest_cmd
Dim DInEncryptest_numRows

Set DInEncryptest_cmd = Server.CreateObject ("ADODB.Command")
DInEncryptest_cmd.ActiveConnection = MM_NationStar_STRING
DInEncryptest_cmd.CommandText = "SELECT SerialNum FROM vDInInvBHead" 
DInEncryptest_cmd.Prepared = true

Set DInEncryptest = DInEncryptest_cmd.Execute
DInEncryptest_numRows = 0
%>

<%
Dim Repeat2__numRows
Dim Repeat2__index

Repeat2__numRows = -1
Repeat2__index = 0
DInEncryptest_numRows = DInEncryptest_numRows + Repeat2__numRows
%>




<%
Dim Repeat1__numRows
Dim Repeat1__index

Repeat1__numRows = -1
Repeat1__index = 0
DInEncryp_numRows = DInEncryp_numRows + Repeat1__numRows
%>
<%
Dim MM_paramName 
%>
<%
' *** Go To Record and Move To Record: create strings for maintaining URL and Form parameters

Dim MM_keepNone
Dim MM_keepURL
Dim MM_keepForm
Dim MM_keepBoth

Dim MM_removeList
Dim MM_item
Dim MM_nextItem

' create the list of parameters which should not be maintained
MM_removeList = "&index="
If (MM_paramName <> "") Then
  MM_removeList = MM_removeList & "&" & MM_paramName & "="
End If

MM_keepURL=""
MM_keepForm=""
MM_keepBoth=""
MM_keepNone=""

' add the URL parameters to the MM_keepURL string
For Each MM_item In Request.QueryString
  MM_nextItem = "&" & MM_item & "="
  If (InStr(1,MM_removeList,MM_nextItem,1) = 0) Then
    MM_keepURL = MM_keepURL & MM_nextItem & Server.URLencode(Request.QueryString(MM_item))
  End If
Next

' add the Form variables to the MM_keepForm string
For Each MM_item In Request.Form
  MM_nextItem = "&" & MM_item & "="
  If (InStr(1,MM_removeList,MM_nextItem,1) = 0) Then
    MM_keepForm = MM_keepForm & MM_nextItem & Server.URLencode(Request.Form(MM_item))
  End If
Next

' create the Form + URL string and remove the intial '&' from each of the strings
MM_keepBoth = MM_keepURL & MM_keepForm
If (MM_keepBoth <> "") Then 
  MM_keepBoth = Right(MM_keepBoth, Len(MM_keepBoth) - 1)
End If
If (MM_keepURL <> "")  Then
  MM_keepURL  = Right(MM_keepURL, Len(MM_keepURL) - 1)
End If
If (MM_keepForm <> "") Then
  MM_keepForm = Right(MM_keepForm, Len(MM_keepForm) - 1)
End If

' a utility function used for adding additional parameters to these strings
Function MM_joinChar(firstItem)
  If (firstItem <> "") Then
    MM_joinChar = "&"
  Else
    MM_joinChar = ""
  End If
End Function
%>






<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<link rel="stylesheet" type="text/css" href="desktop.css">
<script src="SpryAssets/SpryMenuBar.js" type="text/javascript"></script>
<link href="SpryAssets/SpryMenuBarVertical.css" rel="stylesheet" type="text/css" />
<title>Equip Return</title>
</head>

<body>

<div id="header" align = 'center'>Laptop Encryption Status</div>



<table border="0">
  <tr>
    <td valign = 'top'>
<!-- #include virtual= "nationstar/desktop/Library/sidebar.lbi" -->
	</td>
	<td>
		<table width="1100" border="1">

			<tr>
	

			
	

    <td width="100"><b>PC Name</b></td>
	<td width="100"><b>User</b></td>
	<td width="100"><b>Serial</b></td>
    <td width="150"><b>Model</b></td>
    <td width="25"><b>PC Status</b></td>
	<td width="25"><b>Encryp Status</b></td>
	<td width="100"><b>Location</b></td>

    </tr>
  
<% 
While ((Repeat1__numRows <> 0) AND (NOT DInEncryp.EOF)) 
%>

    <tr>
      	
		
		<td width="200"><%=(DInEncryp.Fields.Item("ComputerName").Value)%></td>	
		<td width="200"><%=(DInEncryp.Fields.Item("DInAssignedUser").Value)%></td>
		<td width="200"><%=(DInEncryp.Fields.Item("SerialNum").Value)%></td>
      	<td width="250"><%=(DInEncryp.Fields.Item("DInModel").Value)%></td>
		<td width="250"><%=(DInEncryp.Fields.Item("DInStatus").Value)%></td>
      	<td width="50"><%=(DInEncryp.Fields.Item("EnStatus").Value)%></td>
		<td width="150"><%=(DInEncryp.Fields.Item("DInSite").Value)%></td>


    </tr>

  <% 
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  DInEncryp.MoveNext()
Wend

%>

			</table>
    
		</td>
	</tr>
</table>




</body>
</html>
<%
DInEncryp.Close()
Set DInEncryp = Nothing
%>

