protected void Page_Load(object sender, EventArgs e)
{
    if (!IsPostBack)
    {
        gvCustomers.DataSource = GetData("select top 10 * from Customers");
        gvCustomers.DataBind();
    }
}
 
private static DataTable GetData(string query)
{
    string strConnString = ConfigurationManager.ConnectionStrings["constr"].ConnectionString;
    using (SqlConnection con = new SqlConnection(strConnString))
    {
        using (SqlCommand cmd = new SqlCommand())
        {
            cmd.CommandText = query;
            using (SqlDataAdapter sda = new SqlDataAdapter())
            {
                cmd.Connection = con;
                sda.SelectCommand = cmd;
                using (DataSet ds = new DataSet())
                {
                    DataTable dt = new DataTable();
                    sda.Fill(dt);
                    return dt;
                }
            }
        }
    }
}