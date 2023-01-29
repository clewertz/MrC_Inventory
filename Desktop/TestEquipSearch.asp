<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="Connections/NationStar.asp" -->


<%
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
Dim DInmodel
Dim DInmodel_cmd
Dim DInmodel_numRows

Set DInmodel_cmd = Server.CreateObject ("ADODB.Command")
DInmodel_cmd.ActiveConnection = MM_NationStar_STRING
DInmodel_cmd.CommandText = "SELECT equipment FROM dbo.BF_Resource WHERE equipment_num = '2'  ORDER BY equipment ASC" 
DInmodel_cmd.Prepared = true

Set DInmodel = DInmodel_cmd.Execute
DInmodel_numRows = 0
%>

<%
Dim DInimagever
Dim DInimagever_cmd
Dim DInimagever_numRows

Set DInimagever_cmd = Server.CreateObject ("ADODB.Command")
DInimagever_cmd.ActiveConnection = MM_NationStar_STRING
DInimagever_cmd.CommandText = "SELECT equipment FROM dbo.BF_Resource WHERE equipment_num = '3'  ORDER BY equipment ASC" 
DInimagever_cmd.Prepared = true

Set DInimagever = DInimagever_cmd.Execute
DInimagever_numRows = 0
%>

<%
Dim DInstatus
Dim DInstatus_cmd
Dim DInstatus_numRows

Set DInstatus_cmd = Server.CreateObject ("ADODB.Command")
DInstatus_cmd.ActiveConnection = MM_NationStar_STRING
DInstatus_cmd.CommandText = "SELECT equipment FROM dbo.BF_Resource WHERE equipment_num = '5'  ORDER BY equipment ASC" 
DInstatus_cmd.Prepared = true

Set DInstatus = DInstatus_cmd.Execute
DInstatus_numRows = 0
%>

<%
Dim DInsite
Dim DInsite_cmd
Dim DInsite_numRows

Set DInsite_cmd = Server.CreateObject ("ADODB.Command")
DInsite_cmd.ActiveConnection = MM_NationStar_STRING
DInsite_cmd.CommandText = "SELECT equipment FROM dbo.BF_Resource WHERE equipment_num = '4'  ORDER BY equipment ASC" 
DInsite_cmd.Prepared = true

Set DInsite = DInsite_cmd.Execute
DInsite_numRows = 0
%>

<%
Dim RecentImaged
Dim RecentImaged_cmd
Dim RecentImaged_numRows

Set RecentImaged_cmd = Server.CreateObject ("ADODB.Command")
RecentImaged_cmd.ActiveConnection = MM_NationStar_STRING
RecentImaged_cmd.CommandText = "SELECT DInSerialNum, DInComputerName, DInModel, DInSite FROM dbo.DesktopInventory WHERE DInLastChange > DATEADD(d, -5, GETDATE()) ORDER BY DInLastChange DESC"
RecentImaged_cmd.Prepared = true

Set RecentImaged = RecentImaged_cmd.Execute
Icount_numRows = 0
%>

<%
Dim Repeat__numRows
Dim Repeat__index

Repeat__numRows = 10
Repeat__index = 0
ycount_numRows = ycount_numRows + Repeat__numRows
%>



<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<link rel="stylesheet" type="text/css" href="desktop.css">
<script src="SpryAssets/SpryMenuBar.js" type="text/javascript"></script>
<link href="SpryAssets/SpryMenuBarVertical.css" rel="stylesheet" type="text/css" />
<title>PC Equip Search</title>



<script type="text/javascript">
function UnhideOption() {
	var s = document.getElementById('DInSearchSelect');
	var SelSer = s.options[s.selectedIndex].value;
	
	if (SelSer == 1) {
        document.getElementById('DInSearchSerial').style.display = 'block';
        document.getElementById('DInSearchUserName').style.display = 'none';
		document.getElementById('DInSearchUserName').value = '';
        document.getElementById('DInSearchPCName').style.display = 'none';
		document.getElementById('DInSearchPCName').value = '';
		document.getElementById('DInAssetNum').style.display = 'none';
		document.getElementById('DInAssetNum').value = '';
		document.getElementById('DInCostCenter').style.display = 'none';
		document.getElementById('DInCostCenter').value = '';
	}
	else if (SelSer == 2) {
		document.getElementById('DInSearchSerial').style.display = 'none';
		document.getElementById('DInSearchSerial').value = '';
        document.getElementById('DInSearchUserName').style.display = 'none';
		document.getElementById('DInSearchUserName').value = '';
        document.getElementById('DInSearchPCName').style.display = 'none';
		document.getElementById('DInSearchPCName').value = '';
		document.getElementById('DInAssetNum').style.display = 'block';
		document.getElementById('DInCostCenter').style.display = 'none';
		document.getElementById('DInCostCenter').value = '';
	}	
	else if (SelSer == 3) {
        document.getElementById('DInSearchSerial').style.display = 'none';
		document.getElementById('DInSearchSerial').value = '';
        document.getElementById('DInSearchUserName').style.display = 'block';
        document.getElementById('DInSearchPCName').style.display = 'none';
		document.getElementById('DInSearchPCName').value = '';
		document.getElementById('DInAssetNum').style.display = 'none';
		document.getElementById('DInAssetNum').value = '';
		document.getElementById('DInCostCenter').style.display = 'none';
		document.getElementById('DInCostCenter').value = '';
	}
	else if (SelSer == 4) {
        document.getElementById('DInSearchSerial').style.display = 'none';
		document.getElementById('DInSearchSerial').value = '';
        document.getElementById('DInSearchUserName').style.display = 'none';
		document.getElementById('DInSearchUserName').value = '';
        document.getElementById('DInSearchPCName').style.display = 'block';
		document.getElementById('DInAssetNum').style.display = 'none';
		document.getElementById('DInAssetNum').value = '';
		document.getElementById('DInCostCenter').style.display = 'none';
		document.getElementById('DInCostCenter').value = '';
	} 

	else {
		document.getElementById('DInSearchSerial').style.display = 'none';
		document.getElementById('DInSearchSerial').value = '';
        document.getElementById('DInSearchUserName').style.display = 'none';
		document.getElementById('DInSearchUserName').value = '';
        document.getElementById('DInSearchPCName').style.display = 'none';
		document.getElementById('DInSearchPCName').value = '';
		document.getElementById('DInAssetNum').style.display = 'none';
		document.getElementById('DInAssetNum').value = '';
		document.getElementById('DInCostCenter').style.display = 'block';
	}

}
</script>

</head>









<body>

<div id="header" align="center">Desktop</div>
<div id="menu" align="center"></div>
<br>
<table width="850" border="0">
  <tr>
    <td width="142">
<!-- #include virtual= "nationstar/desktop/Library/sidebar.lbi" -->
</td>
    <td width="193"><form id="equip_search" name="equip_search" method="post" action="SearchReturn.asp">

<table width="214" border="0" align="center">
      <tr>
        <td>Model: </td>
        <td>Image: </td>
        <td>Status:</td>
        <td>Location:  </td>
      </tr>
      <tr>
        <td align="center"><select name="DInmodel" id="DInmodel">
          <option value=""></option>
          <%
While (NOT DInmodel.EOF)
%>
          <option value="<%=(DInmodel.Fields.Item("equipment").Value)%>"><%=(DInmodel.Fields.Item("equipment").Value)%></option>
          <%
  DInmodel.MoveNext()
Wend
If (DInmodel.CursorType > 0) Then
  DInmodel.MoveFirst
Else
  DInmodel.Requery
End If
%>
        </select></td>
        <td align="center"><select name="DInimagever" id="DInimagever">
          <option value=""></option>
          <%
While (NOT DInimagever.EOF)
%>
          <option value="<%=(DInimagever.Fields.Item("equipment").Value)%>"><%=(DInimagever.Fields.Item("equipment").Value)%></option>
          <%
  DInimagever.MoveNext()
Wend
If (DInimagever.CursorType > 0) Then
  DInimagever.MoveFirst
Else
  DInimagever.Requery
End If
%>
        </select></td>
		
        <td align="center"><select name="DInstatus" id="DInstatus">
          <option value=""></option>
          <%
While (NOT DInstatus.EOF)
%>
          <option value="<%=(DInstatus.Fields.Item("equipment").Value)%>"><%=(DInstatus.Fields.Item("equipment").Value)%></option>
          <%
  DInstatus.MoveNext()
Wend
If (DInstatus.CursorType > 0) Then
  DInstatus.MoveFirst
Else
  DInstatus.Requery
End If
%>
        </select></td>
        <td align="center"><select name="DInsite" id="DInsite">
          <option value=""></option>
          <%
While (NOT DInsite.EOF)
%>
          <option value="<%=(DInsite.Fields.Item("equipment").Value)%>"><%=(DInsite.Fields.Item("equipment").Value)%></option>
          <%
  DInsite.MoveNext()
Wend
If (DInsite.CursorType > 0) Then
  DInsite.MoveFirst
Else
  DInsite.Requery
End If
%>
        </select></td>
      </tr>
	  <tr height = "20">
	  </tr>
            <tr>
        <td colspan="4"  align="center"><input type="submit" name="Search" id="Search" value="Search" /></td>
        </tr>
	</form>
 </table>
 


    <td width="194"><form id="DInSearchForm" name="DInSearchForm" method="post" action="SearchReturn.asp">
	
<table width="200" border="0" align="center">
	<tr>
		<td align = "left">
				Search:	<select onChange = "UnhideOption()" name="DInSearchSelect" id="DInSearchSelect">
				<option value = "1">Serial</option>
				<option value = "2">Asset Number</option>
				<option value = "3">User Name</option>
				<option value = "4">PC Name</option>
				<option value = "5">Cost Center</option>
				
				</select>
				
				
				
		</td>
	</tr>	
	<tr>
      <!--<td width="200" align="center"><label for="DInserial"></label><input name="DInserial" type="text" id="DInserial" maxlength="15" /><label for="DInserial"></label></td>-->
		<td>
			<input value = "Serial" style = "display:block" name="DInSearchSerial" type="text" id="DInSearchSerial" size="20" maxlength="25"/>
			<input value = "User" style = "display:none" name="DInSearchUserName" type="text" id="DInSearchUserName" size="20" maxlength="25"/>
			<input value = "PCName" style = "display:none" name="DInSearchPCName" type="text" id="DInSearchPCName" size="20" maxlength="25"/>
			<input value = "Asset" style = "display:none" name="DInAssetNum" type="text" id="DInAssetNum" size="20" maxlength="25"/>
			<input value = "CC" style = "display:none" name="DInCostCenter" type="text" id="DInCostCenter" size="20" maxlength="25"/>		
		</td>
  </tr>
  <tr align="left">
    <td colspan="2"><br><input type="submit" name="Submit" id="Submit" value="Search" />
      </td>
	</tr>
</table>

	<tr>
		<td>
		</td>
		<td colspan = "2">
			<img src="graphics/bar_break.png" width="800" height="10" />
		</td>
	</tr>
</form>
		<tr>
			<td height = "20"></td>
			<td>Equipment that has checked in (past 5 days)</td>
		</tr>
		<tr>
		<td></td>
		<td colspan = "2">
			<table border = "1">
				<tr>
					<td align = "center">Serial #</td>
					<td align = "center">Computer Name</td>
					<td align = "center">Model</td>
					<td align = "center">Site</td>
				</tr>
				<%While ((Repeat__numRows <> 0) AND (NOT RecentImaged.EOF))%>
					<tr>
						<td width="150"><a href="equip_edit.asp?<%= Server.HTMLEncode(MM_keepURL) & MM_joinChar(MM_keepURL) & "DInSerialNum=" & RecentImaged.Fields.Item("DInSerialNum").Value %>"><%=RecentImaged.Fields.Item("DInSerialNum").Value %></a></td>
						<td align="left" width = "220"><%=(RecentImaged.Fields.Item("DInComputerName").Value)%></td>
						<td align="left" width = "220"><%=(RecentImaged.Fields.Item("DInModel").Value)%></td>
						<td align="left" width = "175"><%=(RecentImaged.Fields.Item("DInSite").Value)%></td>
					</tr>
				<% 
					Repeat__index=Repeat__index+1
					Repeat__numRows=Repeat__numRows-1
					RecentImaged.MoveNext()
					Wend
				%>
				
			</table>
			
		</td>
	</tr>
</table>



</body>
</html>


<%
DInmodel.Close()
Set DInmodel = Nothing
%>
<%
DInimagever.Close()
Set DInimagever = Nothing
%>
<%
DInsite.Close()
Set DInsite = Nothing
%>
<%
DInstatus.Close()
Set DInstatus = Nothing
%>
<%
RecentImaged.Close()
Set RecentImaged = Nothing
%>
