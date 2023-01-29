<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="Connections/NationStar.asp" -->
<%
Dim Groups__bfsite
Groups__equiptype = ""
If (Request.Form("bfsite") <> "") Then 
  Groups__bfsite = Request.Form("bfsite")
End If
%>

<%
Dim Groups__image_vers
Groups__image_vers = ""
If (Request.Form("image_vers") <> "") Then 
  Groups__image_vers = Request.Form("image_vers")
End If
%>
<%
Dim Groups__bfmodel
Groups__bfmodel = ""
If (Request.Form("bfmodel") <> "") Then 
  Groups__bfmodel = Request.Form("bfmodel")
End If
%>

<%
Dim Groups__bfstatus
Groups__bfstatus = ""
If (Request.Form("bfstatus") <> "") Then 
  Groups__bfstatus = Request.Form("bfstatus")
End If
%>

<%
Dim Groups
Dim Groups_cmd
Dim Groups_numRows

Set Groups_cmd = Server.CreateObject ("ADODB.Command")
Groups_cmd.ActiveConnection = MM_NationStar_STRING
Groups_cmd.CommandText = "SELECT BF_Type, BF_Serial, BF_Status, BF_Notes, BF_Site, Image_Ver, BF_Model FROM dbo.BF_Desktop WHERE BF_Site =? AND Image_Ver = ? AND BF_Model = ? AND BF_Status = ? ORDER BY BF_Model ASC" 
Groups_cmd.Prepared = true
Groups_cmd.Parameters.Append Groups_cmd.CreateParameter("param1", 200, 1, 255, Groups__bfsite) ' adVarChar
Groups_cmd.Parameters.Append Groups_cmd.CreateParameter("param4", 200, 1, 255, Groups__image_vers) ' adVarChar
Groups_cmd.Parameters.Append Groups_cmd.CreateParameter("param2", 200, 1, 255, Groups__bfmodel) ' adVarChar
Groups_cmd.Parameters.Append Groups_cmd.CreateParameter("param2", 200, 1, 255, Groups__bfstatus) ' adVarChar



Set Groups = Groups_cmd.Execute
Groups_numRows = 0
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
Dim MM_keepForm


Dim MM_removeList
Dim MM_item
Dim MM_nextItem

' create the list of parameters which should not be maintained
MM_removeList = "&index="
If (MM_paramName <> "") Then
  MM_removeList = MM_removeList & "&" & MM_paramName & "="
End If

MM_keepForm=""
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
<link rel="stylesheet" type="text/css" href="citrix_migration_prep.css">
<script src="SpryAssets/SpryMenuBar.js" type="text/javascript"></script>
<link href="SpryAssets/SpryMenuBarVertical.css" rel="stylesheet" type="text/css" />
<title>Group Location Return</title>
</head>

<body>

<div id="header">Equipment Return by Type</div>

<div id="mainbody">


  <table width="1000" border="1">

<% 
While ((Repeat1__numRows <> 0) AND (NOT Groups.EOF)) 
%>

    <tr>
      	
		<td width="75"><a href="equip_edit.asp?<%= Server.HTMLEncode(MM_keepURL) & MM_joinChar(MM_keepURL) & "BF_Serial=" & Users.Fields.Item("BF_Serial").Value %>">Edit</a>
			<a href="equip_delete.asp?<%= Server.HTMLEncode(MM_keepURL) & MM_joinChar(MM_keepURL) & "BF_Serial=" & Users.Fields.Item("BF_Serial").Value %>">Delete</a></td>
		
		<td width="70"><%=(Groups.Fields.Item("BF_Type").Value)%></td>
      	<td width="140"><%=(Groups.Fields.Item("BF_Model").Value)%></td>
      	<td width="100"><%=(Groups.Fields.Item("BF_Serial").Value)%></td>
      	<td width="70"><%=(Groups.Fields.Item("BF_Status").Value)%></td>
		<td width="30"><%=(Groups.Fields.Item("BF_Site").Value)%></td>
		<td width="70"><%=(Groups.Fields.Item("Image_Ver").Value)%></td>
		<td width="400"><%=(Groups.Fields.Item("BF_Notes").Value)%></td>
    </tr>

  <% 
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  Groups.MoveNext()
Wend

%>

   </table>
    



</div>
<p>&nbsp;</p>
</body>
</html>
<%
Groups.Close()
Set Groups = Nothing
%>

