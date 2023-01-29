Imports System.Data.SqlClient
Imports System.Data

Partial Class VB
    Inherits System.Web.UI.Page
    Protected Sub Page_Load(sender As Object, e As EventArgs) Handles Me.Load
        If Not IsPostBack Then
            gvCustomers.DataSource = GetData("select top 10 * from Customers")
            gvCustomers.DataBind()
        End If
    End Sub

    Private Shared Function GetData(query As String) As DataTable
        Dim strConnString As String = ConfigurationManager.ConnectionStrings("constr").ConnectionString
        Using con As New SqlConnection(strConnString)
            Using cmd As New SqlCommand()
                cmd.CommandText = query
                Using sda As New SqlDataAdapter()
                    cmd.Connection = con
                    sda.SelectCommand = cmd
                    Using ds As New DataSet()
                        Dim dt As New DataTable()
                        sda.Fill(dt)
                        Return dt
                    End Using
                End Using
            End Using
        End Using
    End Function

    Protected Sub OnRowDataBound(sender As Object, e As GridViewRowEventArgs)
        If e.Row.RowType = DataControlRowType.DataRow Then
            Dim customerId As String = gvCustomers.DataKeys(e.Row.RowIndex).Value.ToString()
            Dim gvOrders As GridView = TryCast(e.Row.FindControl("gvOrders"), GridView)
            gvOrders.DataSource = GetData(String.Format("select top 3 * from Orders where CustomerId='{0}'", customerId))
            gvOrders.DataBind()
        End If
    End Sub

End Class

