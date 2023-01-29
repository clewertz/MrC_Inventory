<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="Connections/NationStar.asp" -->
<%
Dim SerialNum
Dim AssetNum
Dim ComputerName
Dim UserID
Dim CostCenterNum

SerialNum = "%"
AssetNum = "%"
ComputerName = "%"
UserID = "%"
CostCenterNum = "%"

If (Request.Form("DInSearchSerial") <> "") Then 
  SerialNum = Request.Form("DInSearchSerial")
End If

If (Request.Form("DInAssetNum") <> "") Then 
  AssetNum = Request.Form("DInAssetNum")
End If

If (Request.Form("DInSearchPCName") <> "") Then 
  ComputerName = Request.Form("DInSearchPCName")
End If

If (Request.Form("DInSearchUserName") <> "") Then 
  UserID = Request.Form("DInSearchUserName")+"%"
End If

If (Request.Form("DInCostCenter") <> "") Then 
  CostCenterNum = Request.Form("DInCostCenter")
End If
%>

<%
Dim EquipSearch
Dim EquipSearch_cmd
Dim EquipSearch_numRows

Set EquipSearch_cmd = Server.CreateObject ("ADODB.Command")
EquipSearch_cmd.ActiveConnection = MM_NationStar_STRING
	EquipSearch_cmd.CommandText = "SELECT DInType, DInModel, DInSerialNum, DInStatus, DInSite, DInImageVer, DInUserID, DInComputerName  FROM dbo.DesktopInventory WHERE DInSerialNum LIKE ? AND DInComputerName LIKE ? AND DInUserID LIKE ? AND DInAssetTagNum LIKE ? AND DInCostCenter LIKE ? ORDER BY DInSite ASC, DInModel ASC, DInStatus ASC"
	EquipSearch_cmd.Prepared = true
	EquipSearch_cmd.Parameters.Append EquipSearch_cmd.CreateParameter("param1", 200, 1, 20, SerialNum) ' adVarChar
	EquipSearch_cmd.Parameters.Append EquipSearch_cmd.CreateParameter("param2", 200, 1, 20, ComputerName) ' adVarChar
	EquipSearch_cmd.Parameters.Append EquipSearch_cmd.CreateParameter("param3", 200, 1, 20, UserID) ' adVarChar
	EquipSearch_cmd.Parameters.Append EquipSearch_cmd.CreateParameter("param4", 200, 1, 20, AssetNum) ' adVarChar
	EquipSearch_cmd.Parameters.Append EquipSearch_cmd.CreateParameter("param5", 200, 1, 20, CostCenterNum) ' adVarChar
Set EquipSearch = EquipSearch_cmd.Execute
EquipSearch_numRows = 0
%>

<%
Dim Repeat1__numRows
Dim Repeat1__index

Repeat1__numRows = -1
Repeat1__index = 0
EquipSearch_numRows = EquipSearch_numRows + Repeat1__numRows
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
<title>Serial Return</title>
</head>

<body>

<div id="header">Search Results</div>


<table border = "0" width = "1400" >
	<tr>
		<td valign = "top"><!-- #include virtual= "nationstar/desktop/Library/sidebar.lbi" --></td>
		<td>
			<table width="1200" border="1">
<tr>
	
	<td width="75"><b>Edit Link</b></td>
			
	
	<td width="150"><b>Computer Name</b></td>
    <td width="150"><b>Serial</b></td>
    <td width="50"><b>User ID</b></td>
    <td width="150"><b>Model</b></td>
	<td width="50"><b>Status</b></td>
	<td width="100"><b>Location</b></td>

    </tr>
<% 
While ((Repeat1__numRows <> 0) AND (NOT EquipSearch.EOF)) 
%>

    <tr>
	
		<td width="75"><a href="equip_edit.asp?<%= Server.HTMLEncode(MM_keepURL) & MM_joinChar(MM_keepURL) & "DInSerialNum=" & EquipSearch.Fields.Item("DInSerialNum").Value %>">Edit</a></td>
		<td width="50"><%=(EquipSearch.Fields.Item("DInComputerName").Value)%></td>
		<td width="20"><%=(EquipSearch.Fields.Item("DInSerialNum").Value)%></td>
		<td width="20"><%=(EquipSearch.Fields.Item("DInUserID").Value)%></td>
		<td width="80"><%=(EquipSearch.Fields.Item("DInModel").Value)%></td>
		<td width="20"><%=(EquipSearch.Fields.Item("DInStatus").Value)%></td>
		<td width="50"><%=(EquipSearch.Fields.Item("DInSite").Value)%></td>
		
		

    </tr>

  <% 
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  EquipSearch.MoveNext()
Wend

%>

   </table>
		
		</td> 
	</tr>
</table>

  
    




<p>&nbsp;</p>
</body>
</html>
<%
EquipSearch.Close()
Set EquipSearch = Nothing
%>

