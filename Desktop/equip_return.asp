
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
Dim bgCo
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

<title>Nationstar Desktop Dashboard</title>
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
					<li><a href="https://vrassetcore01:1610/console/jws/jnlp/vrassetcore01_1610/console.jnlp">AssetCore</a></li>
					</ul>
				</li>
			</ul>
        </div>
        </div>
        <div id="contentwrap">
        <div id="content">
		
			<table class="eqTable">
				<tr><td>&nbsp;</td>
					<td><h3>PC Name</b></td>
					<td><h3>User ID</b></td>
					<td><h3>Model</b></td>
					<td><h3>Serial</b></td>
					<td><h3>Status</b></td>
					<td><h3>Location</b></td>
					<td><h3>Image</b></td>
				</tr>

			<% 
			While ((Repeat1__numRows <> 0) AND (NOT Groups.EOF)) 
			if Repeat1__numRows MOD 2 = 0 then bgCo="cccccc" else bgCo="eaeaea"
			%>
			<tr style="background-color:#<%=bgCo%>; height: 22px; vertical-align: middle;">
				<td><a href="equip_edit.asp?<%= Server.HTMLEncode(MM_keepURL) & MM_joinChar(MM_keepURL) & "DInSerialNum=" & Groups.Fields.Item("DInSerialNum").Value %>">Edit</a></td>
				<td><%=(Groups.Fields.Item("DInComputerName").Value)%></td>
				<td><%=(Groups.Fields.Item("DInUserID").Value)%></td>
				<td><%=(Groups.Fields.Item("DInModel").Value)%></td>
				<td><a href="equip_detail.asp?<%= Server.HTMLEncode(MM_keepURL) & MM_joinChar(MM_keepURL) & "DInSerialNum=" & Groups.Fields.Item("DInSerialNum").Value %>"><%=(Groups.Fields.Item("DInSerialNum").Value)%></a></td>
				<td><%=(Groups.Fields.Item("DInStatus").Value)%></td>
				<td><%=(Groups.Fields.Item("DInSite").Value)%></td>
				<td><%=(Groups.Fields.Item("DInImageVer").Value)%></td>
			</tr>

			<% 
			Repeat1__index=Repeat1__index+1
			Repeat1__numRows=Repeat1__numRows-1
			Groups.MoveNext()
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
Groups.Close()
Set Groups = Nothing
%>

