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
Dim Groups__DInimagever
Groups__DInimagever = "%"
If (Request.Form("DInimagever") <> "") Then 
  Groups__DInimagever = Request.Form("DInimagever")
End If
%>
<%
Dim Groups__DInmodel
Groups__DInmodel = "%"
If (Request.Form("DInmodel") <> "") Then 
  Groups__DInmodel = Request.Form("DInmodel")
End If
%>

<%
Dim Groups__DInstatus
Groups__DInstatus = "%"
If (Request.Form("DInstatus") <> "") Then 
  Groups__DInstatus = Request.Form("DInstatus")
End If
%>


<%
Dim Groups
Dim Groups_cmd
Dim Groups_numRows

Set Groups_cmd = Server.CreateObject ("ADODB.Command")
Groups_cmd.ActiveConnection = MM_NationStar_STRING
Groups_cmd.CommandText = "SELECT DInSerialNum, DInComputerName, DInUserID,  DInStatus, DInSite, DInImageVer, DInModel FROM dbo.DesktopInventory WHERE DInSite LIKE ? AND DInModel LIKE ? AND DInStatus LIKE ? AND DInImageVer LIKE ? ORDER BY DInModel ASC" 
Groups_cmd.Prepared = true
Groups_cmd.Parameters.Append Groups_cmd.CreateParameter("param3", 200, 1, 255, Groups__DInsite) ' adVarChar
Groups_cmd.Parameters.Append Groups_cmd.CreateParameter("param3", 200, 1, 255, Groups__DInmodel) ' adVarChar
Groups_cmd.Parameters.Append Groups_cmd.CreateParameter("param3", 200, 1, 255, Groups__DInstatus) ' adVarChar
Groups_cmd.Parameters.Append Groups_cmd.CreateParameter("param3", 200, 1, 255, Groups__DInimagever) ' adVarChar


Set Groups = Groups_cmd.Execute
Groups_numRows = 0%>

<%
Dim DIndesktop
Dim DIndesktop_cmd
Dim DIndesktop_numRows

Set DIndesktop_cmd = Server.CreateObject ("ADODB.Command")
DIndesktop_cmd.ActiveConnection = MM_NationStar_STRING
DIndesktop_cmd.CommandText = "SELECT * FROM dbo.DesktopInventory" 
DIndesktop_cmd.Prepared = true

Set DIndesktop = DIndesktop_cmd.Execute
DIndesktop_numRows = 0
%>

<%
Dim Repeat1__numRows
Dim Repeat1__index

Repeat1__numRows = -1
Repeat1__index = 0
Groups_numRows = Groups_numRows + Repeat1__numRows
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

<div id="header" align = 'center'>Equipment Return by Type</div>



<table border="0">
  <tr>
    <td valign = 'top'>
<!-- #include virtual= "nationstar/desktop/Library/sidebar.lbi" -->
	</td>
	<td>
		<table width="1200" border="1">

			<tr>
	
	<td width="30"><b>Edit Link</b></td>
			
	

    <td width="200"><b>PC Name</b></td>
	<td width="100"><b>User ID</b></td>
	<td width="250"><b>Model</b></td>
    <td width="50"><b>Serial</b></td>
    <td width="50"><b>Status</b></td>
	<td width="100"><b>Location</b></td>
	<td width="50"><b>Image</b></td>

    </tr>
  
<% 
While ((Repeat1__numRows <> 0) AND (NOT Groups.EOF)) 
%>

    <tr>
      	
		<td width="30"><a href="equip_edit.asp?<%= Server.HTMLEncode(MM_keepURL) & MM_joinChar(MM_keepURL) & "DInSerialNum=" & Groups.Fields.Item("DInSerialNum").Value %>">Edit</a></td>
			
		<td width="200"><%=(Groups.Fields.Item("DInComputerName").Value)%></td>
		<td width="100"><%=(Groups.Fields.Item("DInUserID").Value)%></td>
      	<td width="250"><%=(Groups.Fields.Item("DInModel").Value)%></td>
      	<td width="50"><a href="equip_detail.asp?<%= Server.HTMLEncode(MM_keepURL) & MM_joinChar(MM_keepURL) & "DInSerialNum=" & Groups.Fields.Item("DInSerialNum").Value %>"><%=(Groups.Fields.Item("DInSerialNum").Value)%></a></td>
      	<td width="50"><%=(Groups.Fields.Item("DInStatus").Value)%></td>
		<td width="100"><%=(Groups.Fields.Item("DInSite").Value)%></td>
		<td width="50"><%=(Groups.Fields.Item("DInImageVer").Value)%></td>

    </tr>

  <% 
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  Groups.MoveNext()
Wend

%>

			</table>
    
		</td>
	</tr>
</table>




</body>
</html>
<%
Groups.Close()
Set Groups = Nothing
%>

