<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="Connections/NationStar.asp" -->

<%
Dim CWCount
Dim CWCount_cmd
Dim CWCount_numRows

Set CWCount_cmd = Server.CreateObject ("ADODB.Command")
CWCount_cmd.ActiveConnection = MM_NationStar_STRING
CWCount_cmd.CommandText = "SELECT DISTINCT PCTypeS, PCCount, CWStock FROM(SELECT  DISTINCT DInType AS PCTypeC,  COUNT(DInType) AS PCCount FROM DesktopInventory WHERE DInSite = 'Cypress Waters' GROUP BY ALL DInType) AS PCCount INNER JOIN (SELECT  DISTINCT DInType AS PCTypeS,  COUNT(DInType)AS CWStock FROM DesktopInventory WHERE DInSite = 'Cypress Waters' AND DInStatus = 'Ready' GROUP BY ALL DInType) AS CWStock ON PCTypeC = PCTypeS" 
CWCount_cmd.Prepared = true
Set CWCount = CWCount_cmd.Execute
CWCount_numRows = 0
%>

<%
Dim HWCount
Dim HWCount_cmd
Dim HWCount_numRows

Set HWCount_cmd = Server.CreateObject ("ADODB.Command")
HWCount_cmd.ActiveConnection = MM_NationStar_STRING
HWCount_cmd.CommandText = "SELECT DISTINCT PCTypeS, PCCount, HWStock FROM(SELECT  DISTINCT DInType AS PCTypeC,  COUNT(DInType) AS PCCount FROM DesktopInventory WHERE DInSite = 'Horizon Way' GROUP BY ALL DInType) AS PCCount INNER JOIN (SELECT  DISTINCT DInType AS PCTypeS,  COUNT(DInType)AS HWStock FROM DesktopInventory WHERE DInSite = 'Horizon Way' AND DInStatus = 'Ready' GROUP BY ALL DInType) AS HWStock ON PCTypeC = PCTypeS"   
HWCount_cmd.Prepared = true
Set HWCount = HWCount_cmd.Execute
HWCount_numRows = 0
%>

<%
Dim CVCount
Dim CVCount_cmd
Dim CVCount_numRows

Set CVCount_cmd = Server.CreateObject ("ADODB.Command")
CVCount_cmd.ActiveConnection = MM_NationStar_STRING
CVCount_cmd.CommandText = "SELECT DISTINCT PCTypeS, PCCount, CVStock FROM(SELECT  DISTINCT DInType AS PCTypeC,  COUNT(DInType) AS PCCount FROM DesktopInventory WHERE DInSite = 'Convergence' GROUP BY ALL DInType) AS PCCount INNER JOIN (SELECT  DISTINCT DInType AS PCTypeS,  COUNT(DInType)AS CVStock FROM DesktopInventory WHERE DInSite = 'Convergence' AND DInStatus = 'Ready' GROUP BY ALL DInType) AS CVStock ON PCTypeC = PCTypeS" 
CVCount_cmd.Prepared = true
Set CVCount = CVCount_cmd.Execute
CVCount_numRows = 0
%>

<%
Dim CHCount
Dim CHCount_cmd
Dim CHCount_numRows

Set CHCount_cmd = Server.CreateObject ("ADODB.Command")
CHCount_cmd.ActiveConnection = MM_NationStar_STRING
CHCount_cmd.CommandText = "SELECT DISTINCT PCTypeS, PCCount, CHStock FROM(SELECT  DISTINCT DInType AS PCTypeC,  COUNT(DInType) AS PCCount FROM DesktopInventory WHERE DInSite = 'Chandler' GROUP BY ALL DInType) AS PCCount INNER JOIN (SELECT  DISTINCT DInType AS PCTypeS,  COUNT(DInType)AS CHStock FROM DesktopInventory WHERE DInSite = 'Chandler' AND DInStatus = 'Ready' GROUP BY ALL DInType) AS CHStock ON PCTypeC = PCTypeS"
CHCount_cmd.Prepared = true
Set CHCount = CHCount_cmd.Execute
CHCount_numRows = 0
%>

<%
Dim Repeat1__numRows
Dim Repeat1__index

Repeat1__numRows = 10
Repeat1__index = 0
CWCount_numRows = CWCount_numRows + Repeat1__numRows
%>



<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">


<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<link rel="stylesheet" type="text/css" href="desktop.css">
<script src="SpryAssets/SpryMenuBar.js" type="text/javascript"></script>
<link href="SpryAssets/SpryMenuBarVertical.css" rel="stylesheet" type="text/css" />
<title>Dash Board</title>



</head>

<body>

<div id="header" align="center">Desktop Dash Board</div>
<div id="menu" align="center"></div>
<br>
<table width="850" border="0">
	<tr>
		<td width="142"><!-- #include virtual= "nationstar/desktop/Library/sidebar.lbi" --></td>
		<td>
			<table width="550" border="1" align="center">
				<tr>
					<td colspan="3" align="center"><h2>Cypress Waters</h2></td>
				</tr>
				<tr>
					<td><b>PC</b></td>
					<td><b>Count</b></td>
					<td><b>In Stock</b></td>
				</tr>

				<%While ((Repeat1__numRows <> 0) AND (NOT CWCount.EOF))%>
				<tr>
					<td><%=(CWCount.Fields.Item("PCTypeS").Value)%></td>
					<td><%=(CWCount.Fields.Item("PCCount").Value)%></td>
					<td><%=(CWCount.Fields.Item("CWStock").Value)%></td>
				</tr>
				<% 
				Repeat1__index=Repeat1__index+1
				Repeat1__numRows=Repeat1__numRows-1
				CWCount.MoveNext()
				Wend%>	
				
			</table>
		</td>
	</tr>
	<tr>
		<td></td>
		<td><img src="graphics/bar_break.png" width="800" height="10" /></td>
	</tr>
	<tr>
		<td></td>
		<td>
			<table width="550" border="1" align="center">
				<tr>
					<td colspan="3" align="center"><h2>Horizon Way</h2></td>
				</tr>
				<tr>
					<td><b>PC</b></td>
					<td><b>Count</b></td>
					<td><b>In Stock</b></td>
				</tr>
				<%While ((Repeat1__numRows <> 0) AND (NOT HWCount.EOF))%>
				<tr>
					<td><%=(HWCount.Fields.Item("PCTypeS").Value)%></td>
					<td><%=(HWCount.Fields.Item("PCCount").Value)%></td>
					<td><%=(HWCount.Fields.Item("HWStock").Value)%></td>
				</tr>
				<% 
				Repeat1__index=Repeat1__index+1
				Repeat1__numRows=Repeat1__numRows-1
				HWCount.MoveNext()
				Wend%>	
			</table>
		</td>
	</tr>    
	<tr>
		<td></td>
		<td><img src="graphics/bar_break.png" width="800" height="10" /></td>
	</tr>	
	<tr>
		<td></td>
		<td>
			<table width="550" border="1" align="center">
				<tr>
					<td colspan="3" align="center"><h2>Convergence</h2></td>
				</tr>
				<tr>
					<td><b>PC</b></td>
					<td><b>Count</b></td>
					<td><b>In Stock</b></td>
				</tr>
				<%While ((Repeat1__numRows <> 0) AND (NOT CVCount.EOF))%>
				<tr>
					<td><%=(CVCount.Fields.Item("PCTypeS").Value)%></td>
					<td><%=(CVCount.Fields.Item("PCCount").Value)%></td>
					<td><%=(CVCount.Fields.Item("CVStock").Value)%></td>
				</tr>
				<% 
				Repeat1__index=Repeat1__index+1
				Repeat1__numRows=Repeat1__numRows-1
				CVCount.MoveNext()
				Wend%>
			</table>
		</td>
	</tr>
	<tr>
		<td></td>
		<td><img src="graphics/bar_break.png" width="800" height="10" /></td>
	</tr>
	<tr>
		<td></td>
		<td>
			<table width="550" border="1" align="center">
				<tr>
					<td colspan="3" align="center"><h2>Chandler</h2></td>
				</tr>
				<tr>
					<td><b>PC</b></td>
					<td><b>Count</b></td>
					<td><b>In Stock</b></td>
				</tr>
				<%While ((Repeat1__numRows <> 0) AND (NOT CHCount.EOF))%>
				<tr>
					<td><%=(CHCount.Fields.Item("PCTypeS").Value)%></td>
					<td><%=(CHCount.Fields.Item("PCCount").Value)%></td>
					<td><%=(CHCount.Fields.Item("CHStock").Value)%></td>
				</tr>
				<% 
				Repeat1__index=Repeat1__index+1
				Repeat1__numRows=Repeat1__numRows-1
				CHCount.MoveNext()
				Wend%>
			</table>
		</td>
	</tr>
	<tr>
		<td></td>
		<td><img src="graphics/bar_break.png" width="800" height="10" /></td>
	</tr>
</table>




</body>
</html>

<%
CWCount.Close()
Set CWCount = Nothing
%>



