<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="Connections/NationStar.asp" -->

<%
Dim bgCo
Dim SerialNum
Dim AssetNum
Dim ComputerName
Dim UserID
Dim CostCenterNum
Dim MacAddress

SerialNum = "%"
AssetNum = "%"
ComputerName = "%"
UserID = "%"
CostCenterNum = "%"
MacAddress = "%"

If (Request.Form("DInSearchSelect") = 1) Then 
  SerialNum = Request.Form("txtSearch")

ElseIf(Request.Form("DInSearchSelect") = 2) Then 
  AssetNum = Request.Form("txtSearch")

ElseIf (Request.Form("DInSearchSelect") = 4) Then 
  ComputerName = "%"+Request.Form("txtSearch")+"%"

ElseIf (Request.Form("DInSearchSelect") = 3) Then 
  UserID = "%"+Request.Form("txtSearch")+"%"
  
ElseIf (Request.Form("DInSearchSelect") = 5) Then 
  CostCenterNum = Request.Form("txtSearch")
  
ElseIf (Request.Form("DInSearchSelect") = 6) Then 
  MacAddress = Request.Form("txtSearch")
  
End If
%>

<%
Dim EquipSearch
Dim EquipSearch_cmd
Dim EquipSearch_numRows

Set EquipSearch_cmd = Server.CreateObject ("ADODB.Command")
EquipSearch_cmd.ActiveConnection = MM_NationStar_STRING
	EquipSearch_cmd.CommandText = "SELECT DInType, DInModel, DInSerialNum, DInStatus, DInSite, DInImageVer, DInUserID, DInComputerName  FROM dbo.DesktopInventory WHERE DInSerialNum LIKE ? AND DInComputerName LIKE ? AND DInUserID LIKE ? AND DInAssetTagNum LIKE ? AND DInCostCenter LIKE ? AND DInMacAddress LIKE ? ORDER BY DInSite ASC, DInModel ASC, DInStatus ASC"
	EquipSearch_cmd.Prepared = true
	EquipSearch_cmd.Parameters.Append EquipSearch_cmd.CreateParameter("param1", 200, 1, 35, SerialNum) ' adVarChar
	EquipSearch_cmd.Parameters.Append EquipSearch_cmd.CreateParameter("param2", 200, 1, 20, ComputerName) ' adVarChar
	EquipSearch_cmd.Parameters.Append EquipSearch_cmd.CreateParameter("param3", 200, 1, 20, UserID) ' adVarChar
	EquipSearch_cmd.Parameters.Append EquipSearch_cmd.CreateParameter("param4", 200, 1, 20, AssetNum) ' adVarChar
	EquipSearch_cmd.Parameters.Append EquipSearch_cmd.CreateParameter("param5", 200, 1, 20, CostCenterNum) ' adVarChar
	EquipSearch_cmd.Parameters.Append EquipSearch_cmd.CreateParameter("param6", 200, 1, 20, MacAddress) ' adVarChar
Set EquipSearch = EquipSearch_cmd.Execute
EquipSearch_numRows = 0
%>

<%
Dim HEquipSearch
Dim HEquipSearch_cmd
Dim HEquipSearch_numRows

Set HEquipSearch_cmd = Server.CreateObject ("ADODB.Command")
HEquipSearch_cmd.ActiveConnection = MM_NationStar_STRING
	HEquipSearch_cmd.CommandText = "SELECT DInType, DInModel, DInSerialNum, DInStatus, DInSite, DInImageVer, DInUserID, DInComputerName  FROM dbo.HistoricDesktopInventory WHERE DInSerialNum LIKE ? AND DInComputerName LIKE ? AND DInUserID LIKE ? AND DInAssetTagNum LIKE ? AND DInCostCenter LIKE ? AND DInMacAddress LIKE ? ORDER BY DInLastChange DESC"
	HEquipSearch_cmd.Prepared = true
	HEquipSearch_cmd.Parameters.Append HEquipSearch_cmd.CreateParameter("param1", 200, 1, 35, SerialNum) ' adVarChar
	HEquipSearch_cmd.Parameters.Append HEquipSearch_cmd.CreateParameter("param2", 200, 1, 20, ComputerName) ' adVarChar
	HEquipSearch_cmd.Parameters.Append HEquipSearch_cmd.CreateParameter("param3", 200, 1, 20, UserID) ' adVarChar
	HEquipSearch_cmd.Parameters.Append HEquipSearch_cmd.CreateParameter("param4", 200, 1, 20, AssetNum) ' adVarChar
	HEquipSearch_cmd.Parameters.Append HEquipSearch_cmd.CreateParameter("param5", 200, 1, 20, CostCenterNum) ' adVarChar
	HEquipSearch_cmd.Parameters.Append HEquipSearch_cmd.CreateParameter("param6", 200, 1, 20, MacAddress) ' adVarChar
Set HEquipSearch = HEquipSearch_cmd.Execute
HEquipSearch_numRows = 0
%>

<%
Dim Repeat1__numRows
Dim Repeat1__index

Repeat1__numRows = -1
Repeat1__index = 0
EquipSearch_numRows = EquipSearch_numRows + Repeat1__numRows
%>

<%
Dim Repeat2__numRows
Dim Repeat2__index

Repeat2__numRows = -1
Repeat2__index = 0
HEquipSearch_numRows = HEquipSearch_numRows + Repeat2__numRows
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

<title>Nationstar Desktop Equipment Search Results</title>
<!--[if IE]><script src="http://html5shiv.googlecode.com/svn/trunk/html5.js"></script><![endif]-->
<link rel="stylesheet" type="text/css" href="dash.css" />

</head>
<body>
    <div id="wrapper">
        <div id="headerwrap">
        <div id="header">
            <img src="graphics/NSM-logo2.png" style="padding: 8px;">
			<h1>Desktop Equipment Search Results</h1>
        </div>
        </div>
        <div id="navigationwrap">
        <div id="navigation">
			<ul id="menu-bar">
				<li><a href="/desktop/">Dashboard</a></li>
				<li class="active"><a href="equip_search.asp">Equipment</a></li>
				<li><a href="#">Administrative</a>
					<ul>
					<li><a href="https://wd5.myworkday.com/nationstar/login.flex">Workday</a></li>
					<li><a href="https://login.daptiv.com/">Daptiv</a></li>
					<li><a href="http://teams/sites/desktopservices/">Sharepoint</a></li>
					<li><a href="http://teams/sites/oncall/Lists/OnCall/by%20Support%20Area.aspx">On Call</a></li>
					<li><a href="https://console.beachheadsolutions.net/Administration">BeachHead</a></li>
					<li><a href="http://vrsccm01/smsremote/">SCCM Tools</a></li>
					</ul>
				</li>
			</ul>
        </div>
        </div>
        <div id="contentwrap">
        <div id="content" >
		
<table class="eqTable">
	<tr><td>&nbsp;</td>
		<td><h3>Computer Name</h3></td>
		<td><h3>Serial</h3></td>
		<td><h3>User ID</h3></td>
		<td><h3>Model</h3></td>
		<td><h3>Status</h3></td>
		<td><h3>Location</h3></td></tr>
<% 
While ((Repeat1__numRows <> 0) AND (NOT EquipSearch.EOF)) 
if Repeat1__numRows MOD 2 = 0 then bgCo="cccccc" else bgCo="eaeaea"
%>

	<tr style="background-color:#<%=bgCo%>; height: 22px; vertical-align: middle;">
		<td><a href="equip_edit.asp?<%= Server.HTMLEncode(MM_keepURL) & MM_joinChar(MM_keepURL) & "DInSerialNum=" & EquipSearch.Fields.Item("DInSerialNum").Value %>">Edit</a></td>
		<td><%=(EquipSearch.Fields.Item("DInComputerName").Value)%></td>
		<td><%=(EquipSearch.Fields.Item("DInSerialNum").Value)%></td>
		<td><%=(EquipSearch.Fields.Item("DInUserID").Value)%></td>
		<td><%=(EquipSearch.Fields.Item("DInModel").Value)%></td>
		<td><%=(EquipSearch.Fields.Item("DInStatus").Value)%></td>
		<td><%=(EquipSearch.Fields.Item("DInSite").Value)%></td></tr>

<% 
Repeat1__index=Repeat1__index+1
Repeat1__numRows=Repeat1__numRows-1
EquipSearch.MoveNext()
Wend

%>

</table>	

<br>

<table class="eqTable">
	<tr><td><h2>Historical Data</h2><td></tr>
	<tr>
		<td><h3>Computer Name</h3></td>
		<td><h3>Serial</h3></td>
		<td><h3>User ID</h3></td>
		<td><h3>Model</h3></td>
		<td><h3>Status</h3></td>
		<td><h3>Location</h3></td></tr>
<% 
While ((Repeat2__numRows <> 0) AND (NOT HEquipSearch.EOF)) 
if Repeat2__numRows MOD 2 = 0 then bgCo="cccccc" else bgCo="eaeaea"
%>

	<tr style="background-color:#<%=bgCo%>; height: 22px; vertical-align: middle;">
		<td><%=(HEquipSearch.Fields.Item("DInComputerName").Value)%></td>
		<td><%=(HEquipSearch.Fields.Item("DInSerialNum").Value)%></td>
		<td><%=(HEquipSearch.Fields.Item("DInUserID").Value)%></td>
		<td><%=(HEquipSearch.Fields.Item("DInModel").Value)%></td>
		<td><%=(HEquipSearch.Fields.Item("DInStatus").Value)%></td>
		<td><%=(HEquipSearch.Fields.Item("DInSite").Value)%></td></tr>

<% 
Repeat2__index=Repeat2__index+1
Repeat2__numRows=Repeat2__numRows-1
HEquipSearch.MoveNext()
Wend

%>

</table>	


		
			
			
		
        </div>
        </div>
        <div id="footerwrap">
        <div id="footer">
            <p>&nbsp;</p>
        </div>
        </div>
    </div>
</body>
</html>


<%
EquipSearch.Close()
Set EquipSearch = Nothing
%>

