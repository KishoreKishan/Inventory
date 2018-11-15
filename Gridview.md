using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace WebApplication3
{
    public partial class Stocks : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {

            if (!IsPostBack)
            {
                this.BindGrid();
            }

        }
        private void BindGrid()
        {

            SqlConnection con = new SqlConnection(@"Data Source=(localdb)\ProjectsV13;Initial Catalog=kishore;Integrated Security=True;Connect Timeout=30;Encrypt=False;TrustServerCertificate=False;ApplicationIntent=ReadWrite;MultiSubnetFailover=False");
            con.Open();
            SqlCommand cmd = new SqlCommand("SELECT [Inventory Bucket ID],MPN,[IB Acq Invoice Num],PONumber,Qtyy,Ship_date FROM Calculate WHERE Ship_date<= dateadd(month, -24, getdate()) and Qtyy>0 order by Ship_date DESC", con);
            //SqlDataReader rdr = cmd.ExecuteReader();
            //SqlCommand cmd = new SqlCommand("Select Older_Stocks from Sheets", con);
            DataTable dt = new DataTable();
            SqlDataAdapter sda = new SqlDataAdapter();
            sda.SelectCommand = cmd;
            //GridView1.DataSource = rdr;
            //DataTable dt = new DataTable();

            sda.Fill(dt);
            GridView3.DataSource = dt;
            GridView3.DataBind();
            //Export.Visible = true;

        }

        protected void GridView3_Sorting(object sender, GridViewSortEventArgs e)
        {
            SqlConnection con = new SqlConnection(@"Data Source=(localdb)\ProjectsV13;Initial Catalog=kishore;Integrated Security=True;Connect Timeout=30;Encrypt=False;TrustServerCertificate=False;ApplicationIntent=ReadWrite;MultiSubnetFailover=False");
            con.Open();
            //SqlCommand cmd = new SqlCommand("SELECT Reporteddd.MPN, Kumar.CalculatedSUM, sum(Reporteddd.Qty) AS ReportedSUM, (Kumar.CalculatedSUM - sum(Reporteddd.Qty)) Delta  from Reporteddd inner join (select MPN, sum(Qtyy) AS CalculatedSUM from Calculateddd group by MPN) AS Kumar on Reporteddd.MPN = Kumar.MPN group by Reporteddd.MPN, Kumar.CalculatedSUM having Kumar.CalculatedSUM<> sum(Reporteddd.Qty)order by Reporteddd.MPN", con);
            //SqlDataReader rdr = cmd.ExecuteReader();
            SqlCommand cmd = new SqlCommand("SELECT [Inventory Bucket ID],MPN,[IB Acq Invoice Num],PONumber,Qtyy,Ship_date FROM Calculate WHERE Ship_date<= dateadd(month, -24, getdate()) and Qtyy>0 order by Ship_date DESC", con);
            DataTable dt = new DataTable();
            SqlDataAdapter sda = new SqlDataAdapter();
            sda.SelectCommand = cmd;
            sda.Fill(dt);
            DataView view = dt.DefaultView;
            view.Sort = String.Format("{0} {1}", e.SortExpression, GetSortingDirection());
           
            GridView3.DataSource = view;
            GridView3.DataBind();
        }


        //protected void STLOGO_Click(object sender, ImageClickEventArgs e)
        //{
        //    Response.Redirect("Webform2.aspx");
        //}

        protected string GetSortingDirection()
        {
            //if (ViewState["SortDirection"] == null)
            //        //ViewState["SortDirection"] = "ASC";

            //else if (ViewState["SortDirection"] == "ASC")
            // (ViewState["SortDirection"] == "ASC")
            //    ViewState["SortDirection"] = "DESC";
            //else
            //    ViewState["SortDirection"] = "ASC";

            //return ViewState["SortDirection"].ToString();

            if (ViewState["SortDirection"] == null)
                ViewState["SortDirection"] = "ASC";

                if (ViewState["SortDirection"] == "ASC")
            {
                ViewState["SortDirection"] = "DESC";
            
                // return ViewState["SortDirection"].ToString();
            }
            else if (ViewState["SortDirection"] == "DESC")
            {
                ViewState["SortDirection"] = "ASC";
                
            }
        
            //else
            //    ViewState["SortDirection"] = "DESC";
            return ViewState["SortDirection"].ToString();

    }

        protected void GridView3_PageIndexChanging(object sender, GridViewPageEventArgs e)
        {
            GridView3.PageIndex = e.NewPageIndex;
            //GridView1.AllowPaging = true;
            //GridView1.DataSource = Class1.GetAll();
            GridView3.DataSource = Class2.GetAll();

            //GridView3.DataSource = BindGrid();
            GridView3.DataBind();

        }
        public override void VerifyRenderingInServerForm(Control control)
        {

        }

        protected void STLOGO_Click(object sender, ImageClickEventArgs e)
        {
            Response.Redirect("Webform2.aspx");
        }

        protected void Excel_Click(object sender, ImageClickEventArgs e)
        {
            ExportGridToExcel();
        }
        private void ExportGridToExcel()
        {
            Response.Clear();
            Response.Buffer = true;
            Response.ClearContent();
            Response.ClearHeaders();
            //Response.Charset = "";
            string FileName = "Stock_Inventory" + DateTime.Now + ".xls";
            StringWriter strwritter = new StringWriter();
            HtmlTextWriter htmltextwrtter = new HtmlTextWriter(strwritter);

            GridView3.AllowPaging = false;
            GridView3.AllowSorting = false;
            this.BindGrid();


            // Response.Cache.SetCacheability(HttpCacheability.NoCache);
            Response.ContentType = "application/vnd.ms-excel";
            Response.AddHeader("Content-Disposition", "attachment;filename=" + FileName);

            GridView3.HeaderRow.BackColor = Color.White;
            foreach (TableCell cell in GridView3.HeaderRow.Cells)
            {
                cell.BackColor = GridView3.HeaderStyle.BackColor;
            }
            foreach (GridViewRow row in GridView3.Rows)
            {
                row.BackColor = Color.White;
                foreach (TableCell cell in row.Cells)
                {
                    if (row.RowIndex % 2 == 0)
                    {
                        cell.BackColor = GridView3.AlternatingRowStyle.BackColor;
                    }
                    else
                    {
                        cell.BackColor = GridView3.RowStyle.BackColor;
                    }
                    cell.CssClass = "textmode";
                }
            }
            // GridView1.GridLines = GridLines.Both;
            //GridView1.HeaderStyle.Font.Bold = true;
            GridView3.RenderControl(htmltextwrtter);
            Response.Write(strwritter.ToString());
            Response.End();

        }

        protected void Search_Click(object sender, EventArgs e)
        {
            Excel.Visible = true;
            SqlConnection con = new SqlConnection(@"Data Source=(localdb)\ProjectsV13;Initial Catalog=kishore;Integrated Security=True;Connect Timeout=30;Encrypt=False;TrustServerCertificate=False;ApplicationIntent=ReadWrite;MultiSubnetFailover=False");
            con.Open();
            SqlCommand cmd = new SqlCommand("SELECT [Inventory Bucket ID],MPN,[IB Acq Invoice Num],PONumber,Qtyy,Ship_date FROM Calculate WHERE Ship_date<= dateadd(month, -24, getdate()) and Qtyy>0 and  MPN = '" + TextBox1.Text + "'order by Ship_date DESC", con);
            //SqlCommand cmd = new SqlCommand("SELECT Reporteddd.MPN, Kumar.CalculatedSUM, sum(Reporteddd.Qty) AS ReportedSUM, coalesce(sum(Kumar.CalculatedSUM), 0) - coalesce(sum(Reporteddd.Qty), 0) as Delta  from  Reporteddd inner join (select MPN,sum(Qty) AS CalculatedSUM from Calculateddd group by MPN) AS Kumar on Reporteddd.MPN=Kumar.MPN where Reporteddd.MPN= '" + TextBox1.Text +"' group by ReportedddMPN, Kumar.CalculatedSUM having Kumar.CalculatedSUM<>sum(Reporteddd.Qty)order by Reporteddd.MPN" , con);
            //SqlDataReader rdr = cmd.ExecuteReader();

            DataTable dt = new DataTable();
            SqlDataAdapter sda = new SqlDataAdapter();
            sda.SelectCommand = cmd;
            //GridView1.DataSource = rdr;
            //DataTable dt = new DataTable();

            sda.Fill(dt);
            GridView3.DataSource = dt;
            GridView3.DataBind();
        }
    }
}
