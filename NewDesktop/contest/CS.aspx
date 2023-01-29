<%@ Page Language="C#" AutoEventWireup="true" CodeFile="CS.aspx.cs" Inherits="CS" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
    <style type="text/css">
        body
        {
			background-image:url('pacman1.gif');
			background-repeat: no-repeat;
			background-attachment: fixed;
			background-position: center;
            font-family: Arial;
            font-size: 10pt;
        }
        input
        {
            width: 200px;
        }
        table
        {
            border: 1px solid #ccc;
        }
        table th
        {
            background-color: #F7F7F7;
            color: #333;
            font-weight: bold;
        }
        table th, table td
        {
			background-color: #fff;
            padding: 5px;
            border-color: #ccc;
        }
		
		#tableContainer-1 {
			height: 100%;
			width: 100%;
			display: table;
		}
		
		#tableContainer-2 {
			width: 400px;
			height: 250px;

			position:absolute; /*it can be fixed too*/
			left:0; right:0;
			top:0; bottom:0;
			margin:auto;

			/*this to solve "the content will not be cut when the window is smaller than the content": */
			max-width:100%;
			max-height:100%;
			overflow:auto;
		}
		#myTable {
			margin: 0 auto;
			padding: 1;
			border-spacing: 0;
			border: 1px solid black;
			width: 100%

		}
    </style>
</head>
<body>
    <form id="form1" runat="server">
		<div id="tableContainer-1">
			<div id="tableContainer-2">
				<table id="myTable" border="0">
					<tr>
						<th colspan="3">
							Pac-Man Challenge
						</th>
					</tr>
					<tr>
						<td>
							First Name
						</td>
						<td>
							<asp:TextBox ID="txtFirstName" runat="server" />
						</td>
						<td>
							<asp:RequiredFieldValidator ErrorMessage="Required" ForeColor="Red" ControlToValidate="txtFirstName"
								runat="server" />
						</td>
					</tr>
					<tr>
						<td>
							Last Name
						</td>
						<td>
							<asp:TextBox ID="txtLastName" runat="server"/>
						</td>
						<td>
							<asp:RequiredFieldValidator ErrorMessage="Required" ForeColor="Red" ControlToValidate="txtLastName"
								runat="server" />
						</td>
					</tr>
					<tr>
						<td>
							Email
						</td>
						<td>
							<asp:TextBox ID="txtEmail" runat="server" />
						</td>
						<td>
							<asp:RequiredFieldValidator ErrorMessage="Required" Display="Dynamic" ForeColor="Red"
								ControlToValidate="txtEmail" runat="server" />
							<asp:RegularExpressionValidator runat="server" Display="Dynamic" ValidationExpression="\w+([-+.']\w+)*@\w+([-.]\w+)*\.\w+([-.]\w+)*"
								ControlToValidate="txtEmail" ForeColor="Red" ErrorMessage="Invalid email address." />
						</td>
					</tr>
					<tr>
						<td>
						</td>
						<td>
							<asp:Button Text="Submit" runat="server" OnClick="RegisterUser" />
						</td>
						<td>
						</td>
					</tr>
				</table>
			</div>
		</div
    </form>
</body>
</html>
