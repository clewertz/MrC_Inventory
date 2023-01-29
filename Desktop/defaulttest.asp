<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="Connections/NationStar.asp" -->

<%
Dim CWCount
Dim CWCount_cmd
Dim CWCount_numRows

Set CWCount_cmd = Server.CreateObject ("ADODB.Command")
CWCount_cmd.ActiveConnection = MM_NationStar_STRING
CWCount_cmd.CommandText = "SELECT DISTINCT PCTypeS, PCCount, CWStock FROM(SELECT  DISTINCT DInType AS PCTypeC,  COUNT(DInType) AS PCCount FROM DesktopInventory WHERE DInSite = 'Cypress Waters' AND DInStatus = 'Deployed' GROUP BY ALL DInType) AS PCCount INNER JOIN (SELECT  DISTINCT DInType AS PCTypeS,  COUNT(DInType)AS CWStock FROM DesktopInventory WHERE DInSite = 'Cypress Waters' AND DInStatus = 'Ready' GROUP BY ALL DInType) AS CWStock ON PCTypeC = PCTypeS" 
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
HWCount_cmd.CommandText = "SELECT DISTINCT PCTypeS, PCCount, HWStock FROM(SELECT  DISTINCT DInType AS PCTypeC,  COUNT(DInType) AS PCCount FROM DesktopInventory WHERE DInSite = 'Horizon Way' AND DInStatus = 'Deployed' GROUP BY ALL DInType) AS PCCount INNER JOIN (SELECT  DISTINCT DInType AS PCTypeS,  COUNT(DInType)AS HWStock FROM DesktopInventory WHERE DInSite = 'Horizon Way' AND DInStatus = 'Ready' GROUP BY ALL DInType) AS HWStock ON PCTypeC = PCTypeS"   
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
CVCount_cmd.CommandText = "SELECT DISTINCT PCTypeS, PCCount, CVStock FROM(SELECT  DISTINCT DInType AS PCTypeC,  COUNT(DInType) AS PCCount FROM DesktopInventory WHERE DInSite = 'Convergence' AND DInStatus = 'Deployed' GROUP BY ALL DInType) AS PCCount INNER JOIN (SELECT  DISTINCT DInType AS PCTypeS,  COUNT(DInType)AS CVStock FROM DesktopInventory WHERE DInSite = 'Convergence' AND DInStatus = 'Ready' GROUP BY ALL DInType) AS CVStock ON PCTypeC = PCTypeS" 
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
CHCount_cmd.CommandText = "SELECT DISTINCT PCTypeS, PCCount, CHStock FROM(SELECT  DISTINCT DInType AS PCTypeC,  COUNT(DInType) AS PCCount FROM DesktopInventory WHERE DInSite = 'Chandler' AND DInStatus = 'Deployed' GROUP BY ALL DInType) AS PCCount INNER JOIN (SELECT  DISTINCT DInType AS PCTypeS,  COUNT(DInType)AS CHStock FROM DesktopInventory WHERE DInSite = 'Chandler' AND DInStatus = 'Ready' GROUP BY ALL DInType) AS CHStock ON PCTypeC = PCTypeS"
CHCount_cmd.Prepared = true
Set CHCount = CHCount_cmd.Execute
CHCount_numRows = 0
%>

<%
Dim TotCount
Dim TotCount_cmd
Dim TotCount_numRows

Set TotCount_cmd = Server.CreateObject ("ADODB.Command")
TotCount_cmd.ActiveConnection = MM_NationStar_STRING
TotCount_cmd.CommandText = "SELECT DISTINCT PCTypeS, Deployed, Stock FROM (SELECT  DISTINCT DInType AS PCTypeC,  COUNT(DInType) AS Deployed FROM DesktopInventory WHERE DInStatus = 'Deployed' GROUP BY ALL DInType) AS Deployed INNER JOIN (SELECT  DISTINCT DInType AS PCTypeS,  COUNT(DInType)AS Stock FROM DesktopInventory	WHERE DInStatus = 'Ready' GROUP BY ALL DInType) AS Stock ON PCTypeC = PCTypeS"
TotCount_cmd.Prepared = true
Set TotCount = TotCount_cmd.Execute
TotCount_numRows = 0
%>

<%
Dim AppleCount
Dim AppleCount_cmd
Dim AppleCount_numRows

Set AppleCount_cmd = Server.CreateObject ("ADODB.Command")
AppleCount_cmd.ActiveConnection = MM_NationStar_STRING
AppleCount_cmd.CommandText = "SELECT t1.AppleStock, t2.AppleDeployed FROM (SELECT DISTINCT COUNT(DInImageVer) AS AppleDeployed FROM DesktopInventory WHERE DInStatus = 'Deployed' AND DInImageVer = 'OSX')t2, (SELECT DISTINCT COUNT(DInImageVer) AS AppleStock FROM DesktopInventory WHERE DInStatus = 'Ready' AND DInImageVer = 'OSX')t1"
AppleCount_cmd.Prepared = true
Set AppleCount = AppleCount_cmd.Execute
AppleCount_numRows = 0
%>

<%
Dim ChenCount
Dim ChenCount_cmd
Dim ChenCount_numRows

Set ChenCount_cmd = Server.CreateObject ("ADODB.Command")
ChenCount_cmd.ActiveConnection = MM_NationStar_STRING
ChenCount_cmd.CommandText = "SELECT DISTINCT PCTypeS, PCCount, ChenStock FROM(SELECT  DISTINCT DInType AS PCTypeC,  COUNT(DInType) AS PCCount FROM DesktopInventory WHERE DInSite = 'Chennai' AND DInStatus = 'Deployed' GROUP BY ALL DInType) AS PCCount INNER JOIN (SELECT  DISTINCT DInType AS PCTypeS,  COUNT(DInType)AS ChenStock FROM DesktopInventory WHERE DInSite = 'Chennai' AND DInStatus = 'Ready' GROUP BY ALL DInType) AS ChenStock ON PCTypeC = PCTypeS"   
ChenCount_cmd.Prepared = true
Set ChenCount = ChenCount_cmd.Execute
ChenCount_numRows = 0
%>

<%
Dim SenecaCount
Dim SenecaCount_cmd
Dim SenecaCount_numRows

Set SenecaCount_cmd = Server.CreateObject ("ADODB.Command")
SenecaCount_cmd.ActiveConnection = MM_NationStar_STRING
SenecaCount_cmd.CommandText = "SELECT DISTINCT PCTypeS, PCCount, SenecaStock FROM(SELECT  DISTINCT DInType AS PCTypeC,  COUNT(DInType) AS PCCount FROM DesktopInventory WHERE DInSite = 'Seneca' AND DInStatus = 'Deployed' GROUP BY ALL DInType) AS PCCount INNER JOIN (SELECT  DISTINCT DInType AS PCTypeS,  COUNT(DInType)AS SenecaStock FROM DesktopInventory WHERE DInSite = 'Seneca' AND DInStatus = 'Ready' GROUP BY ALL DInType) AS SenecaStock ON PCTypeC = PCTypeS"   
SenecaCount_cmd.Prepared = true
Set SenecaCount = SenecaCount_cmd.Execute
SenecaCount_numRows = 0
%>
<!--
<%
Dim LongviewCount
Dim LongviewCount_cmd
Dim LongviewCount_numRows

Set LongviewCount_cmd = Server.CreateObject ("ADODB.Command")
LongviewCount_cmd.ActiveConnection = MM_NationStar_STRING
LongviewCount_cmd.CommandText = "SELECT DISTINCT PCTypeS, PCCount, LongviewStock FROM(SELECT  DISTINCT DInType AS PCTypeC,  COUNT(DInType) AS PCCount FROM DesktopInventory WHERE DInSite = 'Longview' AND DInStatus = 'Deployed' GROUP BY ALL DInType) AS PCCount INNER JOIN (SELECT  DISTINCT DInType AS PCTypeS,  COUNT(DInType)AS LongviewStock FROM DesktopInventory WHERE DInSite = 'Longview' AND DInStatus = 'Ready' GROUP BY ALL DInType) AS LongviewStock ON PCTypeC = PCTypeS"   
LongviewCount_cmd.Prepared = true
Set LongviewCount = LongviewCount_cmd.Execute
LongviewCount_numRows = 0
%>
-->

<%
Dim Repeat1__numRows
Dim Repeat1__index

Repeat1__numRows = 14
Repeat1__index = 0
CWCount_numRows = CWCount_numRows + Repeat1__numRows
%>



<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">

<head>
<title>Nationstar Desktop Dashboard</title>
<!--[if IE]><script src="http://html5shiv.googlecode.com/svn/trunk/html5.js"></script><![endif]-->
<link rel="stylesheet" type="text/css" href="dash.css" />

<!--
<link rel="icon" type="image/png" href="favicon.png" />
-->

</head>
<body>
    <div id="wrapper">
        <div id="headerwrap">
        <div id="header">
            <img src="graphics/NSM-logo2.png" style="padding: 8px;">
			<h1>Desktop Dashboard</h1>
        </div>
        </div>
        <div id="navigationwrap">
        <div id="navigation">
			<ul id="menu-bar">
				<li class="active"><a href="/desktop/">Dashboard</a></li>
				<li><a href="equip_search.asp">Equipment</a></li>
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
	<tr><td>
			<table class="eqTable">
				<tr><td nowrap><h2>Cypress Waters</h2></td>
					<td colspan="3" width="85%">&nbsp;</td></tr>
				<tr><td rowspan="3">&nbsp;</td>
					<td><h3>Chassis</h3></td>
					<td><h3>Count</h3></td>
					<td><h3>In Stock</h3></td>
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
	
	<tr><td>
			<table class="eqTable">
				<tr><td nowrap><h2>Horizon Way</h2></td>
					<td colspan="3" width="85%">&nbsp;</td></tr>
				<tr><td rowspan="3">&nbsp;</td>
					<td><h3>Chassis</h3></td>
					<td><h3>Count</h3></td>
					<td><h3>In Stock</h3></td>
				</tr>
				<%While ((Repeat1__numRows <> 0) AND (NOT HWCount.EOF))%>
				<tr><td><%=(HWCount.Fields.Item("PCTypeS").Value)%></td>
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
	
	<tr><td>
			<table class="eqTable">
				<tr><td nowrap><h2>Convergence</h2></td>
					<td colspan="3" width="85%">&nbsp;</td></tr>
				<tr><td rowspan="3">&nbsp;</td>
					<td><h3>Chassis</h3></td>
					<td><h3>Count</h3></td>
					<td><h3>In Stock</h3></td>
				</tr>
				<%While ((Repeat1__numRows <> 0) AND (NOT CVCount.EOF))%>
				<tr><td><%=(CVCount.Fields.Item("PCTypeS").Value)%></td>
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

	<tr><td>
			<table class="eqTable">
				<tr><td nowrap><h2>Chandler</h2></td>
					<td colspan="3" width="85%">&nbsp;</td></tr>
				<tr><td rowspan="3">&nbsp;</td>
					<td><h3>Chassis</h3></td>
					<td><h3>Count</h3></td>
					<td><h3>In Stock</h3></td>
				</tr>
				<%While ((Repeat1__numRows <> 0) AND (NOT CHCount.EOF))%>
				<tr><td><%=(CHCount.Fields.Item("PCTypeS").Value)%></td>
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
	
	<tr><td>
			<table class="eqTable">
				<tr><td nowrap><h2>Chennai</h2></td>
					<td colspan="3" width="85%">&nbsp;</td></tr>
				<tr><td rowspan="3">&nbsp;</td>
					<td><h3>Chassis</h3></td>
					<td><h3>Count</h3></td>
					<td><h3>In Stock</h3></td>
				</tr>
				<%While ((Repeat1__numRows <> 0) AND (NOT ChenCount.EOF))%>
				<tr><td><%=(ChenCount.Fields.Item("PCTypeS").Value)%></td>
					<td><%=(ChenCount.Fields.Item("PCCount").Value)%></td>
					<td><%=(ChenCount.Fields.Item("ChenStock").Value)%></td>
				</tr>
				<% 
				Repeat1__index=Repeat1__index+1
				Repeat1__numRows=Repeat1__numRows-1
				ChenCount.MoveNext()
				Wend%>
			</table>
		</td>
	</tr>
	
		<tr><td>
			<table class="eqTable">
				<tr><td nowrap><h2>Seneca</h2></td>
					<td colspan="3" width="85%">&nbsp;</td></tr>
				<tr><td rowspan="3">&nbsp;</td>
					<td><h3>Chassis</h3></td>
					<td><h3>Count</h3></td>
					<td><h3>In Stock</h3></td>
				</tr>
				<%While ((Repeat1__numRows <> 0) AND (NOT SenecaCount.EOF))%>
				<tr><td><%=(SenecaCount.Fields.Item("PCTypeS").Value)%></td>
					<td><%=(SenecaCount.Fields.Item("PCCount").Value)%></td>
					<td><%=(SenecaCount.Fields.Item("SenecaStock").Value)%></td>
				</tr>
				<% 
				Repeat1__index=Repeat1__index+1
				Repeat1__numRows=Repeat1__numRows-1
				SenecaCount.MoveNext()
				Wend%>
			</table>
		</td>
	</tr>
	
			<tr><td>
			<table class="eqTable">
				<tr><td nowrap><h2>Longview</h2></td>
					<td colspan="3" width="85%">&nbsp;</td></tr>
				<tr><td rowspan="3">&nbsp;</td>
					<td><h3>Chassis</h3></td>
					<td><h3>Count</h3></td>
					<td><h3>In Stock</h3></td>
				</tr>
				<%While ((Repeat1__numRows <> 0) AND (NOT LongviewCount.EOF))%>
				<tr><td><%=(LongviewCount.Fields.Item("PCTypeS").Value)%></td>
					<td><%=(LongviewCount.Fields.Item("PCCount").Value)%></td>
					<td><%=(LongviewCount.Fields.Item("LongviewStock").Value)%></td>
				</tr>
				<% 
				Repeat1__index=Repeat1__index+1
				Repeat1__numRows=Repeat1__numRows-1
				LongviewCount.MoveNext()
				Wend%>
			</table>
		</td>
	</tr>

	
		<tr><td>
			<table class="eqTable">
				<tr><td nowrap><h2>Apple Products</h2></td>
					<td colspan="3" width="85%">&nbsp;</td></tr>
				<tr><td rowspan="3">&nbsp;</td>
					<td><h3>Chassis</h3></td>
					<td><h3>Count</h3></td>
					<td><h3>In Stock</h3></td>
				</tr>
				<tr><td>Apple</td>
					<td><%=(AppleCount.Fields.Item("AppleDeployed").Value)%></td>
					<td><%=(AppleCount.Fields.Item("AppleStock").Value)%></td>
				</tr>
			</table>
		</td>
	</tr> 
	
	
	<tr><td>
			<table class="eqTable">
				<tr><td nowrap><h2>Company Total</h2></td>
					<td colspan="3" width="85%">&nbsp;</td></tr>
				<tr><td rowspan="3">&nbsp;</td>
					<td><h3>Chassis</h3></td>
					<td><h3>Count</h3></td>
					<td><h3>In Stock</h3></td>
				</tr>
				<%While ((Repeat1__numRows <> 0) AND (NOT TotCount.EOF))%>
				<tr><td><%=(TotCount.Fields.Item("PCTypeS").Value)%></td>
					<td><%=(TotCount.Fields.Item("Deployed").Value)%></td>
					<td><%=(TotCount.Fields.Item("Stock").Value)%></td>
				</tr>
				<% 
				Repeat1__index=Repeat1__index+1
				Repeat1__numRows=Repeat1__numRows-1
				TotCount.MoveNext()
				Wend%>
			</table>
		</td>		
	</tr>

	
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
CWCount.Close()
Set CWCount = Nothing
%>