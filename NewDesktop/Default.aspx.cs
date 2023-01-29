using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

public partial class _Default : System.Web.UI.Page
{
    private const string DESKTOP_INVENTORY_CONNECTION_STRING = @"Persist Security Info=False;Database=NationStar;Server=VRSQLNSM01;uid=ctxmigration;pwd=kI3g2ZQsKkCQiiEOjBgu;";
    private const string sql_getAll = @" SET TRANSACTION ISOLATION LEVEL SERIALIZABLE; 
    BEGIN TRANSACTION;
    SELECT * FROM DesktopInventory";
    private const string sql_getSites = @" SET TRANSACTION ISOLATION LEVEL SERIALIZABLE; 
    BEGIN TRANSACTION;
    SELECT equipment FROM BF_Resource WHERE equipment_num = 4 ORDER BY equipment";

    protected void Page_Load(object sender, EventArgs e)
    {
       
    }


    private void loadSites()
    {
        string sitesQueryCount = @"SELECT COUNT(equipment) FROM BF_Resource WHERE equipment_num = 4 GROUP BY equipment_num";
        object[] countArray = sqlQueryDesktopInventory(sitesQueryCount, 1);
        int arrayCount = Convert.ToInt32(countArray[0]);

        string sitesQuery = @"SELECT equipment FROM BF_Resource WHERE equipment_num = 4 ORDER BY equipment";
        object[] sites = sqlQueryDesktopInventory(sitesQuery, arrayCount);

        for (int i = 0; i < sites.Length; i++)
        {
            string siteStock = sites[i] + "Stock";
            string sitesInvQuery = @"SELECT DISTINCT PCTypeS, PCCount, "+ siteStock+ @" FROM(SELECT  DISTINCT DInType AS PCTypeC,  COUNT(DInType) AS PCCount FROM DesktopInventory WHERE DInSite = '"+sites[i]+"' AND DInStatus = 'Deployed' GROUP BY ALL DInType) AS PCCount INNER JOIN (SELECT  DISTINCT DInType AS PCTypeS,  COUNT(DInType)AS "+siteStock+" FROM DesktopInventory WHERE DInSite = '"+sites[i]+"' AND DInStatus = 'Ready' GROUP BY ALL DInType) AS "+siteStock+" ON PCTypeC = PCTypeS";
            //object[] siteStock = object[] sites = sqlQueryDesktopInventory(sitesInvQuery, arrayCount);
        }

    }

    private object[] sqlQueryDesktopInventory(string query, int arrayCount)
    {
        object[] results = new object[arrayCount];

        try
        {
            using (SqlConnection conn = new SqlConnection(DESKTOP_INVENTORY_CONNECTION_STRING))
            {
                conn.Open();

                string script = String.Format(query);

                using (SqlCommand cmd = new SqlCommand(script, conn))
                {
                    SqlCommand showresult = new SqlCommand(script, conn);

                }
                conn.Close();

            }
        }
        catch (System.Exception)
        {
            //do nothing
        }

        return results;
    }

    private void TableCreation()
    {
        Table siteTable = new Table();
        siteTable.CellSpacing = 10;

        int numberOfColumns = 6;
        for (int x = 0; x < numberOfColumns; x++)
        {
           
        }

    }
}