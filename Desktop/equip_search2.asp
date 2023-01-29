<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="Connections/NationStar.asp" -->

<%
Dim bgCo
Dim Repeat1__numRows
Dim Repeat1__index

Repeat1__numRows = -1
Repeat1__index = 0
users_numRows = users_numRows + Repeat1__numRows
%>
<%
Dim Repeat2__numRows
Dim Repeat2__index

Repeat2__numRows = -1
Repeat2__index = 0
users_software_numRows = users_software_numRows + Repeat2__numRows
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


<%
Dim Repeat__numRows
Dim Repeat__index

Repeat__numRows = 10
Repeat__index = 0
ycount_numRows = ycount_numRows + Repeat__numRows
%>

<head>
<title>Nationstar Desktop Dashboard : Equipment Search</title>
<!--[if IE]><script src="http://html5shiv.googlecode.com/svn/trunk/html5.js"></script><![endif]-->


<link rel="stylesheet" type="text/css" href="dash.css" />
</head>
<body>
    <div id="wrapper">
        <div id="headerwrap">
        <div id="header">
            <img src="graphics/NSM-logo2.png" style="padding: 8px;">
			<h1>Desktop Equipment</h1>
        </div>
        </div>
        <div id="navigationwrap">
        <div id="navigation">
			<ul id="menu-bar">
				<li><a href="/desktop/"">Dashboard</a></li>
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
					<li><a href="http://crdskweb01/desktop/equip_search2.asp">Batch Search</a></li>
					</ul>
				</li>
			</ul>
        </div>
        </div>

        <div id="content">
		
		<form id="DInSearchForm" name="DInSearchForm" method="post" action="equip_subset2.asp">
		<div style="float: left; margin-right: 18px; margin-bottom: 10px;">
			<select onChange = "UnhideOption()" name="DInSearchSelect" id="DInSearchSelect">
				<option value = "1">&nbsp;&nbsp;PC Name</option>

			</select>
				
		  <!--<td width="200" align="center"><label for="DInserial"></label><input name="DInserial" type="text" id="DInserial" maxlength="15" /><label for="DInserial"></label></td>-->
				<textarea id ="txtSearch" name="txtSearch"></textarea>
				<!--<input class="txtBox" name="txtSearch" type="text" id="txtSearch" maxlength="25" placeholder="Search"/> -->
	
				
			<input class="btn" type="submit" name="Submit" id="Submit" value="     Search     " />
		 </div></form>

		

        </div>
        <div id="footerwrap">
        <div id="footer">
            <p>&nbsp;</p>
        </div>
        </div>
    </div>
</body>
</html>




