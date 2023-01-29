using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

public partial class _Default : Page
{
    private const string DESKTOP_INVENTORY_CONNECTION_STRING = @"Persist Security Info=False;Database=NationStar;Server=VRSQLNSM01;uid=ctxmigration;pwd=kI3g2ZQsKkCQiiEOjBgu;";
    private const string sql_getComp = @" SET TRANSACTION ISOLATION LEVEL SERIALIZABLE; 
		BEGIN TRANSACTION;
		SELECT DInComputerName FROM DesktopInventory WHERE DInComputerName <> '0' ORDER BY DinComputerName";
    private const string sql_getPW = @" SET TRANSACTION ISOLATION LEVEL SERIALIZABLE; 
		BEGIN TRANSACTION;
		SELECT DInAdmPwd FROM DesktopTest WHERE DInComputerName = '0'";

    private const string sql_getPCList = @" SET TRANSACATION ISOLATION LEVEL SERIALIZABLE;
        BEGIN TRANSACTION;
        SELECT DInComputerName, DInSerialNum, DInUserID,  DInModel, DInStatus, DInSite, DInImageVer
        FROM DesktopInventory
        WHERE DinComputerName IN ('0')";

    protected void Page_Load(object sender, EventArgs e)
    {

    }

    protected void btnSearch_Click(object sender, EventArgs e)
    {
        string[] lst = txtPCList.Text.Split(new Char[] { '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries);
        lblReturn.Text = lst[1];

        for (int i = 0; i < lst.Length; i++)
        {
            
        }
    }
}