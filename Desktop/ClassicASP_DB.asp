<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="Connections/NationStar.asp" -->


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


<body>
    <form id="equip_search" name="equip_search" method="post" action="equip_return.asp">

		</select>
		
		<select name="DInimagever" id="DInimagever">
			<option value="">&nbsp;&nbsp;Image ---</option>
		<%
		While (NOT DInimagever.EOF)%>
		
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
		</select>
		
	</form>
		

</body>
</html>

<%
DInimagever.Close()
Set DInimagever = Nothing
%>
