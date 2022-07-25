using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data;
using Excel = Microsoft.Office.Interop.Excel;


using System.Diagnostics;
using System.IO;
using PdfSharp;
using PdfSharp.Drawing;
using PdfSharp.Drawing.Layout;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;
using System.Collections;


namespace AUTO_COMM
{
    public partial class OMSSD_BULK : Form
    {
        public string nitdate1, EMDdt1, EAdt1, cost2dt, refnd;
       public  string emdprty, emddep, emddis;
        public string note, noofprevioustender, noofpresenttender, prevnitdate, preveadate, currentnitdt, currenteadt, nitdate_mnth, nitdate_year, finyr, qty, rprtrks, rprtot, rprurs, rpronly, rpremail, dep, dis, stock, qtyprpsd, totqtyput, totresut, pubdate;
        //public DateTime nitdate, pubdate, EMDdate, EAdate;
        public int flag = 0;
        public double refund, dep2, remn;
        public OMSSD_BULK()
        {
            InitializeComponent();
        }

        private void OMSSD_BULK_Load(object sender, EventArgs e)
        {
           /* getCon();
            DataSet ds = new DataSet();
       
            string qry = "Select NITdate from Publication order by NITDate";
            ds = select_data(qry);
            comboBox1.DataSource = ds.Tables[0];
            comboBox1.ValueMember = "NITdate";
            comboBox1.DisplayMember = "NITdate";
            comboBox1.Text = "-Select-";*/
        }


        # region  Database Connection

        OleDbConnection connectionString;
      /*  public static DataSet ImportExcelXLS(string FileName)
        {
            string strConn;
            if (FileName.Substring(FileName.LastIndexOf('.')).ToLower() == ".xlsx")
                strConn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + FileName + ";Extended Properties=Excel 12.0;";
            else
                strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + FileName + ";Extended Properties=Excel 8.0; ";

            DataSet output = new DataSet();

            using (OleDbConnection conn = new OleDbConnection(strConn))
            {
                conn.Open();

                DataTable schemaTable = conn.GetOleDbSchemaTable(
                    OleDbSchemaGuid.Tables, new object[] { null, null, null, "TABLE" });

                foreach (DataRow schemaRow in schemaTable.Rows)
                {
                    string sheet = schemaRow["TABLE_NAME"].ToString();

                    if (!sheet.EndsWith("_"))
                    {
                        try
                        {
                            OleDbCommand cmd = new OleDbCommand("SELECT Total FROM [" + sheet + "]", conn);
                            cmd.CommandType = CommandType.Text;

                            DataTable outputTable = new DataTable(sheet);
                            output.Tables.Add(outputTable);
                            new OleDbDataAdapter(cmd).Fill(outputTable);
                        }
                        catch (Exception ex)
                        {
                            throw new Exception(ex.Message + string.Format("Sheet:{0}.File:F{1}", sheet, FileName), ex);
                        }
                    }
                }
            }
            return output;
        }*/
        private DataSet getConexcel(string filename)
        {
            DataSet ds = new DataSet();
            try
            {
                //string filename_right = filename.Split('\\').Last();
                string strExcelConn;
               
                if (filename.Substring(filename.LastIndexOf('.')).ToLower() == ".xlsx")
                    strExcelConn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + filename + ";Extended Properties='Excel 12.0;HDR=Yes'";
                else
                    strExcelConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + filename + ";Extended Properties='Excel 8.0;HDR=Yes' ";

                OleDbConnection connExcel = new OleDbConnection(strExcelConn);
                OleDbCommand cmdExcel = new OleDbCommand();
                cmdExcel.Connection = connExcel;
                connExcel.Open();
                DataTable dtExcelSchema;
                dtExcelSchema = connExcel.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                connExcel.Close();
                
                string SheetName = dtExcelSchema.Rows[0]["TABLE_NAME"].ToString();//for first sheet         

               // MessageBox.Show(SheetName);
               //cmdExcel.CommandText = "SELECT * From [" + SheetName + "$A1:D1]";
                cmdExcel.CommandText = "SELECT District,Depot,Total From [" + SheetName + "]";//wheat,wheaturs
                //cmdExcel.CommandText = "SELECT FCIDistric,Terminal,WHEATRake  From [" + SheetName + "]";//rakes
                OleDbDataAdapter da = new OleDbDataAdapter();
                da.SelectCommand = cmdExcel;
                da.Fill(ds);
                            
            }
            catch (Exception excp)
            {

                MessageBox.Show(excp.ToString(), "Not a valid report selected");
            }
            return ds;
        }
        private DataSet getConexcel1W(string filename)
        {
            DataSet ds = new DataSet();
            try
            {
                //string filename_right = filename.Split('\\').Last();
                string strExcelConn;

                if (filename.Substring(filename.LastIndexOf('.')).ToLower() == ".xlsx")
                    strExcelConn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + filename + ";Extended Properties='Excel 12.0;HDR=Yes'";
                else
                    strExcelConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + filename + ";Extended Properties='Excel 8.0;HDR=Yes' ";

                OleDbConnection connExcel = new OleDbConnection(strExcelConn);
                OleDbCommand cmdExcel = new OleDbCommand();
                cmdExcel.Connection = connExcel;
                connExcel.Open();
                DataTable dtExcelSchema;
                dtExcelSchema = connExcel.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                connExcel.Close();

                string SheetName = dtExcelSchema.Rows[0]["TABLE_NAME"].ToString();//for first sheet         

                // MessageBox.Show(SheetName);
                //cmdExcel.CommandText = "SELECT * From [" + SheetName + "$A1:D1]";
                // cmdExcel.CommandText = "SELECT District,Depot,Total From [" + SheetName + "]";//wheat,wheaturs
               // cmdExcel.CommandText = "SELECT FCIDistric,Terminal,WHEATRake  From [" + SheetName + "]";//rakes
             cmdExcel.CommandText = "SELECT Depot,TotalWheat_,Date From [" + SheetName + "]";//total wheat
              //cmdExcel.CommandText = "SELECT * From [" + SheetName + "]";//total wheat

                OleDbDataAdapter da = new OleDbDataAdapter();
                da.SelectCommand = cmdExcel;
                da.Fill(ds);

            }
            catch (Exception excp)
            {

                MessageBox.Show(excp.ToString(), "Not a valid report selected");
            }
            return ds;
        }
        private DataSet getConexcel1R(string filename)
        {
            DataSet ds = new DataSet();
            try
            {
                //string filename_right = filename.Split('\\').Last();
                string strExcelConn;

                if (filename.Substring(filename.LastIndexOf('.')).ToLower() == ".xlsx")
                    strExcelConn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + filename + ";Extended Properties='Excel 12.0;HDR=Yes'";
                else
                    strExcelConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + filename + ";Extended Properties='Excel 8.0;HDR=Yes' ";

                OleDbConnection connExcel = new OleDbConnection(strExcelConn);
                OleDbCommand cmdExcel = new OleDbCommand();
                cmdExcel.Connection = connExcel;
                connExcel.Open();
                DataTable dtExcelSchema;
                dtExcelSchema = connExcel.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                connExcel.Close();

                string SheetName = dtExcelSchema.Rows[0]["TABLE_NAME"].ToString();//for first sheet         

                // MessageBox.Show(SheetName);
                //cmdExcel.CommandText = "SELECT * From [" + SheetName + "$A1:D1]";
                // cmdExcel.CommandText = "SELECT District,Depot,Total From [" + SheetName + "]";//wheat,wheaturs
                // cmdExcel.CommandText = "SELECT FCIDistric,Terminal,WHEATRake  From [" + SheetName + "]";//rakes
               cmdExcel.CommandText = "SELECT Depot,RiceRaw_,Date From [" + SheetName + "]";//total rice
               //cmdExcel.CommandText = "SELECT * From [" + SheetName + "]";//total rice
                OleDbDataAdapter da = new OleDbDataAdapter();
                da.SelectCommand = cmdExcel;
                da.Fill(ds);

            }
            catch (Exception excp)
            {

                MessageBox.Show(excp.ToString(), "Not a valid report selected");
            }
            return ds;
        }
        
        
        private DataSet getConexcel2(string filename)
        {
            DataSet ds = new DataSet();
            try
            {
                //string filename_right = filename.Split('\\').Last();
                string strExcelConn;

                if (filename.Substring(filename.LastIndexOf('.')).ToLower() == ".xlsx")
                    strExcelConn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + filename + ";Extended Properties='Excel 12.0;HDR=Yes'";
                else
                    strExcelConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + filename + ";Extended Properties='Excel 8.0;HDR=Yes' ";

                OleDbConnection connExcel = new OleDbConnection(strExcelConn);
                OleDbCommand cmdExcel = new OleDbCommand();
                cmdExcel.Connection = connExcel;
                connExcel.Open();
                DataTable dtExcelSchema;
                dtExcelSchema = connExcel.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                connExcel.Close();

                string SheetName = dtExcelSchema.Rows[0]["TABLE_NAME"].ToString();//for first sheet         

                // MessageBox.Show(SheetName);
                //cmdExcel.CommandText = "SELECT * From [" + SheetName + "$A1:D1]";
               // cmdExcel.CommandText = "SELECT District,Depot,Total From [" + SheetName + "]";//wheat,wheaturs
                cmdExcel.CommandText = "SELECT FCIDistrict,Terminal,WHEATRake  From [" + SheetName + "]";//rakes
                OleDbDataAdapter da = new OleDbDataAdapter();
                da.SelectCommand = cmdExcel;
                da.Fill(ds);

            }
            catch (Exception excp)
            {

                MessageBox.Show(excp.ToString(),"Not a valid report selected");
            }
            return ds;
        }
        private DataSet getConexcel3(string filename)
        {
            DataSet ds = new DataSet();
            try
            {
                //string filename_right = filename.Split('\\').Last();
                string strExcelConn;

                if (filename.Substring(filename.LastIndexOf('.')).ToLower() == ".xlsx")
                    strExcelConn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + filename + ";Extended Properties='Excel 12.0;HDR=Yes'";
                else
                    strExcelConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + filename + ";Extended Properties='Excel 8.0;HDR=Yes' ";

                OleDbConnection connExcel = new OleDbConnection(strExcelConn);
                OleDbCommand cmdExcel = new OleDbCommand();
                cmdExcel.Connection = connExcel;
                connExcel.Open();
                DataTable dtExcelSchema;
                dtExcelSchema = connExcel.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                connExcel.Close();

                string SheetName = dtExcelSchema.Rows[0]["TABLE_NAME"].ToString();//for first sheet         

                // MessageBox.Show(SheetName);
                //cmdExcel.CommandText = "SELECT * From [" + SheetName + "$A1:D1]";
                // cmdExcel.CommandText = "SELECT District,Depot,Total From [" + SheetName + "]";//wheat,wheaturs
               // cmdExcel.CommandText = "SELECT FCIDistrict,Terminal,WHEATRake  From [" + SheetName + "]";//rakes
                cmdExcel.CommandText = "SELECT DEPOT,RATES From [" + SheetName + "]";//rates
                OleDbDataAdapter da = new OleDbDataAdapter();
                da.SelectCommand = cmdExcel;
                da.Fill(ds);

            }
            catch (Exception excp)
            {

                MessageBox.Show(excp.ToString(), "Not a valid report selected");
            }
            return ds;
        }

        private DataSet getConexcel_result(string filename)
        {
            DataSet ds = new DataSet();
            try
            {
                //string filename_right = filename.Split('\\').Last();
                string strExcelConn;

                if (filename.Substring(filename.LastIndexOf('.')).ToLower() == ".xlsx")
                    strExcelConn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + filename + ";Extended Properties='Excel 12.0;HDR=Yes'";
                else
                    strExcelConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + filename + ";Extended Properties='Excel 8.0;HDR=Yes' ";

                OleDbConnection connExcel = new OleDbConnection(strExcelConn);
                OleDbCommand cmdExcel = new OleDbCommand();
                cmdExcel.Connection = connExcel;
                connExcel.Open();
                DataTable dtExcelSchema;
                dtExcelSchema = connExcel.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                connExcel.Close();

                string SheetName = dtExcelSchema.Rows[0]["TABLE_NAME"].ToString();//for first sheet         

                // MessageBox.Show(SheetName);
                //cmdExcel.CommandText = "SELECT * From [" + SheetName + "$A1:D1]";
                cmdExcel.CommandText = "SELECT * From [" + SheetName + "]";//wheat,wheaturs
                //cmdExcel.CommandText = "SELECT FCIDistric,Terminal,WHEATRake  From [" + SheetName + "]";//rakes
                OleDbDataAdapter da = new OleDbDataAdapter();
                da.SelectCommand = cmdExcel;
                da.Fill(ds);

            }
            catch (Exception excp)
            {

                MessageBox.Show(excp.ToString(), "Not a valid report selected");
            }
            return ds;
        }
        private DataSet getConexcelemail(string filename)
        {
            DataSet ds = new DataSet();
            try
            {
                //string filename_right = filename.Split('\\').Last();
                string strExcelConn;

                if (filename.Substring(filename.LastIndexOf('.')).ToLower() == ".xlsx")
                    strExcelConn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + filename + ";Extended Properties='Excel 12.0;HDR=Yes'";
                else
                    strExcelConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + filename + ";Extended Properties='Excel 8.0;HDR=Yes' ";

                OleDbConnection connExcel = new OleDbConnection(strExcelConn);
                OleDbCommand cmdExcel = new OleDbCommand();
                cmdExcel.Connection = connExcel;
                connExcel.Open();
                DataTable dtExcelSchema;
                dtExcelSchema = connExcel.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                connExcel.Close();

                string SheetName = dtExcelSchema.Rows[0]["TABLE_NAME"].ToString();//for first sheet         

                // MessageBox.Show(SheetName);
                //cmdExcel.CommandText = "SELECT * From [" + SheetName + "$A1:D1]";
                cmdExcel.CommandText = "SELECT * From [" + SheetName + "]";//wheat,wheaturs
                //cmdExcel.CommandText = "SELECT FCIDistric,Terminal,WHEATRake  From [" + SheetName + "]";//rakes
                OleDbDataAdapter da = new OleDbDataAdapter();
                da.SelectCommand = cmdExcel;
                da.Fill(ds);

            }
            catch (Exception excp)
            {

                MessageBox.Show(excp.ToString(), "Not a valid report selected");
            }
            return ds;
        }
        
        private OleDbConnection getCon()
        {
            try
            {
                connectionString = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=..\\OMSSD_BULK.accdb;Persist Security Info=True;Jet OLEDB:Database Password=12345");
                connectionString.Open();
                
            }
            catch (Exception excp) {

                MessageBox.Show(excp.ToString());
            }
            return connectionString;
        }
        private DataSet select_data(string str)
        {
            DataSet ds = new DataSet();
            try
            {
                getCon();
                OleDbDataAdapter adp = new OleDbDataAdapter(str, connectionString);                
                adp.Fill(ds);
                connectionString.Dispose();
                
            }
            catch (Exception excp)
            {

                MessageBox.Show(excp.ToString());
            }
            return ds;
        }
        private void insert_update_deleted(string str)
        {
            try
           {

                getCon();
                OleDbCommand cmd = new OleDbCommand(str, connectionString);
                cmd.ExecuteNonQuery();
                connectionString.Dispose();
                MessageBox.Show("Succesful");
                

            }
            

            catch (Exception exp) {

                MessageBox.Show("In insertion function" + " " + exp.ToString());
            }
        }
        #endregion

        #region publication

        # region calculation of dates


        /*  public void calnitdate1()
        {

            string dayofweek;
            int dtofdt, monthofdt, month, prevmonth;
            try 
            {
                if (cmbmonth.SelectedIndex !=-1)
                {
                    month = (cmbmonth.SelectedIndex) + 1; //month selectedfrom combo
                    prevmonth = month - 1;// previous month 1 less than current month


                    dayofweek = dttmPUB.Value.DayOfWeek.ToString(); //day of the week extracted from pub date
                    monthofdt = dttmPUB.Value.Month; //month extracted from pub date
                    dtofdt = dttmPUB.Value.Date.Day;//date extracted from pub date 

                    
                    if (dayofweek == "Monday" && monthofdt == month && dtofdt >= 1 && dtofdt <= 8)
                    {

                        nitdate1 = dttmPUB.Value.ToShortDateString();

                    }
                    else if (dayofweek == "Friday" && monthofdt == prevmonth && dtofdt >= 22 && dtofdt <= 31)
                    {

                        nitdate1 = dttmPUB.Value.AddDays(3).ToShortDateString();
                    }
                    else
                    {
                        lblManualNIT.Visible = true;
                        dttm1stNIT.Visible = true;
                        
                    }
                }
                      
            }
            catch (Exception e) 
            {
                MessageBox.Show("Error in conversions");
            }
        }*/
          
           public void caldts()
           {

            
          try{
            nitdate1 = dttm1stNIT.Value.ToShortDateString();          
            EMDdt1 = (Convert.ToDateTime(nitdate1)).AddDays(2).ToShortDateString();
            EAdt1 = (Convert.ToDateTime(nitdate1)).AddDays(3).ToShortDateString();
           
          }           
           

            catch (Exception e) 
            {
                MessageBox.Show("Error in conversions");
            }


        }
      # endregion

        # region adding dynamic controls to tablelayout panel  
     /*  public void adddynamiccontols()
        {

           for (int i = 1; i <=4; i++)
            {                
                    Label lblSl = new Label();
                    lblSl.Name = "lblSl" + i.ToString();
                    lblSl.Text = i.ToString();
                    lblSl.ForeColor = Color.White;
                    tableLayoutPanelDates.Controls.Add(lblSl, 0, i);                   
                
            }
            Label lblSl1=new Label();
            tableLayoutPanelDates.Controls.Add(lblSl1, 0, 1);
            lblSl1.Text = "1";
            lblSl1.ForeColor = Color.White;
            Label lblSl2 = new Label();
            tableLayoutPanelDates.Controls.Add(lblSl2, 0, 2);
            lblSl2.Text = "2";
            lblSl2.ForeColor = Color.White;
            Label lblSl3 = new Label();
            tableLayoutPanelDates.Controls.Add(lblSl3, 0, 3);
            lblSl3.Text = "3";
            lblSl3.ForeColor = Color.White;
            Label lblSl4 = new Label();
            tableLayoutPanelDates.Controls.Add(lblSl4, 0, 4);
            lblSl4.Text = "4";
            lblSl4.ForeColor = Color.White;
           Label lblNIT1 = new Label();
            tableLayoutPanelDates.Controls.Add(lblNIT1, 1, 1);
            lblNIT1.Text = nitdate1;
            lblNIT1.ForeColor = Color.White;

        }*/
    #endregion

            

        private void cmbmonth_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cmbmonth.SelectedIndex != -1 || cmbmonth.SelectedIndex != 0)
                     {
                         lbldt.Visible = true;
                         dttmPUB.Visible = true;
                         lblpbdt.Visible = true;
                     }
            else
            {
                lbldt.Visible = false;
                dttmPUB.Visible = false;
                lblpbdt.Visible = false;
            }
                              
                // dttmPUB.Enabled = true;                        
            
        }
  
        private void dttmPUB_VisibleChanged(object sender, EventArgs e)
        {
            if (dttmPUB.Visible == true)
            {
                lblManualNIT.Visible = true;
                dttm1stNIT.Visible = true;
                lblnit.Visible = true;
            }
        }
        private void dttm1stNIT_ValueChanged(object sender, EventArgs e)
        {
           // btnCalDt.Visible = true;
            //btnShowDt.Visible = true;

        }

        private void dttm1stNIT_VisibleChanged(object sender, EventArgs e)
        {
            btnCalDt.Visible = true;
            btnShowDt.Visible = true;
            lblalldtsmnth.Visible = true;

        }
        #region insertdates
        private void btnDates_Click(object sender, EventArgs e)
        {
            // tableLayoutPanelDates.Visible = true;                    
            caldts();
           try {
                
               string pubdate = dttmPUB.Value.ToShortDateString();

               if (cmbmonth.SelectedIndex != -1 && cmbpubfinyr.SelectedIndex != -1)
                {
                    if (cmbmonth.SelectedIndex != 0)
                    {
                        // to prevent duplicate entry in primary key field NITDate
                        string nitdate = dttm1stNIT.Value.ToLongDateString();
                        //nitdate = Convert.ToDateTime(nitdate1);               
                        DataSet ds_selectdts = new DataSet();
                        string qry = "Select * from Publication where  NITDate= #" + nitdate + "#";
                        ds_selectdts = select_data(qry);
                        if (ds_selectdts.Tables[0].Rows.Count == 0)// no duplicate entry correponding to the NIT date
                        {
                            // publication date and nit date not matching with month, data not inserted
                            if ((cmbmonth.SelectedIndex == dttmPUB.Value.Month) && (cmbmonth.SelectedIndex == dttm1stNIT.Value.Month))
                            {

                                if ((Convert.ToDateTime(nitdate1).DayOfWeek).ToString() == "Monday")// insert when NIT DATE is MONDAY
                                {
                                    string str = "insert into Publication (Monthname,Finyr,PubDate,NITDate,EMDDate,EAuctionDate) values('" + cmbmonth.SelectedItem + "','"+cmbpubfinyr.SelectedItem.ToString()+"', '" + pubdate + "','" + nitdate1 + "','" + EMDdt1 + "','" + EAdt1 + "')";
                                    insert_update_deleted(str);
                                }
                                else
                                {
                                    DialogResult result=MessageBox.Show("NITdate not Monday-Do you want to continue?", "Check!", MessageBoxButtons.YesNo,MessageBoxIcon.Question);
                                    if(result==System.Windows.Forms.DialogResult.Yes)
                                    {
                                        string str = "insert into Publication (Monthname,Finyr,PubDate,NITDate,EMDDate,EAuctionDate) values('" + cmbmonth.SelectedItem + "', '" + cmbpubfinyr.SelectedItem.ToString() + "','" + pubdate + "','" + nitdate1 + "','" + EMDdt1 + "','" + EAdt1 + "')";
                                        insert_update_deleted(str);
                                    }
                                    
                                    

                                }
                            }
                            else
                            {
                                MessageBox.Show(" Either " + dttmPUB.Value.ToShortDateString() + " or " + dttm1stNIT.Value.ToShortDateString() + " or " + " both "+"  doesnot fall in  " + cmbmonth.SelectedItem, "Try again", MessageBoxButtons.OK,MessageBoxIcon.Warning);
                            }
                        }
                        else
                        {
                            MessageBox.Show("Duplicate Entry of NITDate", "Try again", MessageBoxButtons.RetryCancel);
                        }
                        
                    }
                    else
                    {
                        MessageBox.Show("Invalid Month", "Try again", MessageBoxButtons.RetryCancel);
                    }
                    
                }
            }
            catch (Exception exp)
            {
                MessageBox.Show(exp.ToString());
           }
            
            
            //MessageBox.Show(nitdate1);           
           // MessageBox.Show(EMDdt1);
            //MessageBox.Show(EAdt1);

        }
        #endregion
        private void btnClear_Click(object sender, EventArgs e)
        {
            cmbmonth.SelectedIndex = -1;
            cmbmonth.Text = "-Select-";
            lblManualNIT.Visible = false;
            dttm1stNIT.Visible = false;
            lbldt.Visible = false;
            lblalldtsmnth.Visible = false;
            dttmPUB.Visible = false;           
            btnShowDt.Visible = false;
            btnCalDt.Visible = false;
            dgvDates.Visible = false;
            dttm1stNIT.Value = DateTime.Today;
            lblrcrds.Visible = true;
            lblno.Visible = true;
            //tableLayoutPanelDates.Visible = false;
            //tableLayoutPanelDates.Controls.Clear();
        }
        #region SHOWDATES
        private void btnShowDt_Click(object sender, EventArgs e)
        {
            
             // dynamic controls to datagridview        
          
            DataGridViewButtonColumn editbtn = new DataGridViewButtonColumn();
            editbtn.Name = "EDIT";
            editbtn.HeaderText = "EDIT";
            editbtn.Text = "EDIT";
            editbtn.UseColumnTextForButtonValue = true;
            editbtn.FlatStyle = FlatStyle.Popup;
           
            DataGridViewButtonColumn dltbtn = new DataGridViewButtonColumn();
            dltbtn.Name = "DELETE";
            dltbtn.HeaderText = "DELETE";
            dltbtn.Text = "DELETE";
            dltbtn.UseColumnTextForButtonValue = true;
            dltbtn.FlatStyle = FlatStyle.Popup;

            DataGridViewTextBoxColumn dyncol = new DataGridViewTextBoxColumn();
            dyncol.Name = "Sl";
            dyncol.HeaderText = "Sl. No.";
            
            
            

            DataSet ds_selectdts = new DataSet();
            if (cmbmonth.SelectedIndex != -1)
            {
                nitdate1 = dttm1stNIT.Value.ToLongDateString();
                string pubdate = dttmPUB.Value.ToShortDateString();
               // nitdate = Convert.ToDateTime(nitdate1);
               // nitdate = dttm1stNIT.Value;
                string qry;
                if (cmbmonth.SelectedIndex == 0)
                {
                    qry = "Select * from Publication order by NITDate";//when select all is selected in combo
                    

                }
                
                else
                    qry = "Select * from Publication where MonthName='" + cmbmonth.SelectedItem + "' and NITDate= #" + nitdate1 + "# and PubDate='" + pubdate+ "'";//when valid month and valid date is selected in combo
               
                ds_selectdts = select_data(qry);

                if (ds_selectdts.Tables[0].Rows.Count == 0 )
                {
                        MessageBox.Show("Invalid date selected. Hence,all dates of selected month of the year displayed", "Error", MessageBoxButtons.OK);
                        qry = "Select * from Publication where MonthName='" + cmbmonth.SelectedItem + "' and Finyr='"+cmbpubfinyr.SelectedItem+"'order by NITDate";//displaying all dates of valid month
                        ds_selectdts = select_data(qry); 
                          
                              
                }

                if (ds_selectdts.Tables[0].Rows.Count != 0 )
                {

                    
                  
                    //renaming dataset column names dynamically

                   // ds_selectdts.Tables[0].Columns["Sl"].ColumnName = "Sl. No.";
                    ds_selectdts.Tables[0].Columns["MonthName"].ColumnName = "Month";
                    ds_selectdts.Tables[0].Columns["PubDate"].ColumnName = "Publication Date";
                    ds_selectdts.Tables[0].Columns["NITDate"].ColumnName = "NIT Date (Date of Commencement of Deposition of EMD)";
                    ds_selectdts.Tables[0].Columns["EMDDate"].ColumnName = "Last Date of Deposition of EMD";
                    ds_selectdts.Tables[0].Columns["EAuctionDate"].ColumnName = "Date of E-Auction";

                    // datagridview dynamic

                    dgvDates.Columns.Clear();
                    dgvDates.Visible = true;
                    dgvDates.AutoSize = false;
                    dgvDates.AllowUserToAddRows = false;
                    dgvDates.AllowUserToResizeRows = true;
                    dgvDates.AllowUserToResizeColumns = true;
                    dgvDates.AutoGenerateColumns = true;
                    dgvDates.RowHeadersVisible = true;

                    // adding sl no column in gridview display
                    //dgvDates.Columns.Add(dyncol);


                    dgvDates.DataSource = ds_selectdts;
                    dgvDates.DataMember = ds_selectdts.Tables[0].ToString();

                    /* foreach (DataGridViewRow r in dgvDates.Rows)
                    {

                        dgvDates.Rows[r.Index].Cells["Sl"].Value = (r.Index + 1).ToString();

                    }*/

                    // dgvDates.Columns["Sl. No."].ReadOnly = true;//Sl. No. column made non-editable
                    dgvDates.Columns["NIT Date (Date of Commencement of Deposition of EMD)"].ReadOnly = true;//NITDate column made non-editable
                    dgvDates.Columns.Add(editbtn);
                    dgvDates.Columns.Add(dltbtn);

                    lblrcrds.Visible = true;
                    lblno.Visible = true;
                    lblno.Text = dgvDates.RowCount.ToString();
                    
                    dgvDates.Refresh();
                }
                else
                {
                    MessageBox.Show("Oops!No Dates in DB for this selection, Calculate Dates First", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    dgvDates.Visible = false;   
                }

            }
            
                
               

        }
        #endregion

        

        // edit and delete dynamically
        # region update delete dynamically
        private void dgvdate_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            // Check if wanted column was clicked 
            if (e.ColumnIndex == 6 && e.RowIndex >= 0)
            {               
                //Perform on edit button click code
                string month = dgvDates.CurrentRow.Cells["Month"].Value.ToString();
                string pubdate = dgvDates.CurrentRow.Cells["Publication Date"].Value.ToString();
                //string nitdate = dgvDates.CurrentRow.Cells["Date of Commencement of Deposition of EMD"].Value.ToString();
                string nitdate = Convert.ToDateTime(dgvDates.CurrentRow.Cells["NIT Date (Date of Commencement of Deposition of EMD)"].Value).ToLongDateString();
                string emddate = dgvDates.CurrentRow.Cells["Last Date of Deposition of EMD"].Value.ToString();
                string eadate = dgvDates.CurrentRow.Cells["Date of E-Auction"].Value.ToString();

                //check whether publication dt & nit dt matches with month, else stop update
                string pubdate_month = Convert.ToDateTime(pubdate).Month.ToString();
                string pubdate_year = Convert.ToDateTime(pubdate).Year.ToString();
                string nitdate_month = Convert.ToDateTime(nitdate).Month.ToString();
                string nitdate_mnth=Convert.ToDateTime(nitdate).ToString("MMMM");// GETTING MONTH NAME
                string nitdate_year = Convert.ToDateTime(nitdate).Year.ToString();
               /* string emddate_month = Convert.ToDateTime(emddate).Month.ToString();
                string emddate_year = Convert.ToDateTime(emddate).Year.ToString();
                string eadate_month = Convert.ToDateTime(eadate).Month.ToString();
                string eadate_year = Convert.ToDateTime(eadate).Year.ToString();*/

                if (nitdate_year == pubdate_year && nitdate_month == pubdate_month && nitdate_mnth == month)
                {
                    string qry_updt = "update Publication set MonthName='" + month + "',PubDate='" + pubdate + "',EMDdate='" + emddate + "',EAuctionDate='"+eadate+"'where NITDate=#" + nitdate + "#";
                    insert_update_deleted(qry_updt);              
                
                }
                else
                {
                    MessageBox.Show("Mismatch of either month or year between NITdt & PUBdt or Month column or both ", "Stop Update", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                
                    
                

            }
            else if (e.ColumnIndex == 7 && e.RowIndex >= 0)
            {
                //Perform on deletes button click code

                string nitdate = Convert.ToDateTime(dgvDates.CurrentRow.Cells["NIT Date (Date of Commencement of Deposition of EMD)"].Value).ToLongDateString();
                string qry = "Delete from Publication where NITDate=#" + nitdate + "# ";
                insert_update_deleted(qry);                
                dgvDates.Rows.RemoveAt(dgvDates.CurrentRow.Index);
                


            }
            else
            {

            }
        }

        # endregion

        private void btnPDF_Click(object sender, EventArgs e)
        {
            DataSet pub = new DataSet();
            string str;
            if (cmbmonth.SelectedIndex != -1)
            {
                if (cmbmonth.SelectedIndex != 0)
                {
                    string year = cmbpubfinyr.SelectedItem.ToString();                                  
                    
                    str = "Select * from Publication where MonthName='" + cmbmonth.SelectedItem + "' and FINyr='"+year+"' order by NITDate";
                    pub = select_data(str);
                    if (pub.Tables[0].Rows.Count != 0) 
                    {
                        AD_Creation(pub);
                    }
                    else
                    {
                        MessageBox.Show("Oops!No Dates in DB for this selection, Calculate Dates First", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                    
                }
                else
                    MessageBox.Show("Select 1 particular Month", "Wrong Selection", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    
            }
            else 
            {
                MessageBox.Show("Select 1 particular Month", "Invalid Selection", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        #region AD CREATION
        private void AD_Creation(DataSet pub)
        {

            //all the static items in the page hold in variables

            string refno = "Ref. No.:- " + "Com.32/OMSS(D)/BULK/WHEAT/"+cmbpubfinyr.SelectedItem.ToString();
            string date = "Dated:" + pub.Tables[0].Rows[0].ItemArray[2].ToString();
            string notice = "NOTICE INVITING TENDER";
            string sale = "SALE OF RAW RICE GRADE-A & WHEAT UNDER OMSS(D) THROUGH E-AUCTION";
            string one_one = "GM, FCI (West Bengal) invites Financial Bid  for sale of ";
            string one_one_a = "  Raw  Rice  Grade A";
            string one_one_b = "   and  ";
            string one_one_c = "   Wheat";
            string one_one_d = "   in";
            string one_two_a = "West Bengal Region through ";
            string one_two_b = "E-Auction only";
            string one_two_c = "from the";
            string one_two_d = "FCI empanelled bulk consumers/traders";
            string one = "of rice and wheat respectively. Further, empanelment being a continuous process, interested bulk consumers/traders of rice and wheat may get their firms empanelled with FCI, procedural details of which are available on our website.";
            string two_one_a = "ONLINE Address for E-Auction is ";
            string two_one_b = "http://www.ncdexspot.com/.";
            string three_one = "Schedule   of  Tenders  for   both  Raw  Rice   Grade  A  and   Wheat  during  the  month  of";
            string three_two_a = cmbmonth.SelectedItem + " ' " + Convert.ToDateTime(pub.Tables[0].Rows[0].ItemArray[2].ToString()).Year.ToString();
            string three_two_b = "is as under:";
            string four_a = "Online bidding for the above tenders would be conducted on the day of E-Auction from " ;
            string four_b = "11:00 am to 02:00 pm.";
            string five_one_a= "Details of tenders would be available on";
            string five_one_b = "www.fciweb.nic.in, www.ncdexspot.com.";
            string six = "Sd/-";
            string seven = "For General Manager (WB)";
            string eight = "NOT TO BE PUBLISHED";
            string nine = "Distribution:-";
            string ten = "1. Area Manager,Food Corporation of India, W.B. Region";
            string eleven = "2. E.D (East) , FCI, ZO (E),Kolkata  ";
            string twelve = "3. General Manager (Sales), FCI, Hqrs., New Delhi";
            string thirteen_a = "4. AGM (Coordn.), FCI, RO, Kol... For publishing the above tender notice by" + " " + pub.Tables[0].Rows[0].ItemArray[2].ToString() +" "+"in Bengali, Hindi & ";
            string thirteen_b = "English local daily newspapers positively and ensure submission of the paper clippings in support of the";
            string thirteen_c = "publication to this section immediately without fail.";
            string fourteen_a = "5. Calcutta Flour Mills Association";
            string fourteen_b = "15, Brabourne Road, Kolkata-700001…with request to circulate among the members of your  Association";
            string fifteen = "6. AGM (Movt/QC/Stg/A/cs)/M (Cash), FCI, R.O. Kolkata .For information and necessary action";
            string sixteen = "7. File No. Com.32/OMSS(D)/BULK/RRA/" + cmbpubfinyr.SelectedItem.ToString() ;





            // Create a new PDF document
            PdfDocument document = new PdfDocument();
            document.Info.Title = "Created with PDFsharp";

            // Create an empty page
            PdfPage page = document.AddPage();

            // Get an XGraphics object for drawing
            XGraphics gfx = XGraphics.FromPdfPage(page);
            XTextFormatter tf = new XTextFormatter(gfx);
            

            // Draw FCI logo of the page
            string jpegSamplePath = "../IMAGES/logo.jpg";
            XImage image = XImage.FromFile(jpegSamplePath);
            gfx.DrawImage(image, 44, 70, 500, 80);            

            // Create a font

            XFont font1 = new XFont("Calibri (Body)", 12, XFontStyle.Regular);
            XFont font2 = new XFont("Calibri (Body)", 12, XFontStyle.Underline);            
            XFont font3 = new XFont("Calibri (Body)", 12, XFontStyle.Italic);
            XFont font4 = new XFont("Calibri (Body)", 12, XFontStyle.Bold);
            XFont font5 = new XFont("Calibri (Body)", 10, XFontStyle.Regular);
            XFont font6 = new XFont("Calibri (Body)", 11, XFontStyle.Bold);
            XFont font7 = new XFont("Calibri (Body)", 09, XFontStyle.Bold);
            XFont font8 = new XFont("Calibri (Body)", 09, XFontStyle.Regular);

            int x =44;
            int y =30;

             // Draw the text

            // file refno
            gfx.DrawString(refno, font1, XBrushes.Black, new XRect(x, y + 130, page.Width, page.Height), XStringFormats.TopLeft);
            // date of ad== date of publication
            gfx.DrawString(date, font1, XBrushes.Black, new XRect(x + 390, y + 130, page.Width, page.Height), XStringFormats.TopLeft);
            //notice line-centre alignment
            gfx.DrawString(notice, font2, XBrushes.Black, new XRect(x-44, y + 155, page.Width, page.Height), XStringFormats.TopCenter);
            //sale-centre alignment
            gfx.DrawString(sale, font2, XBrushes.Black, new XRect(x - 44, y + 180, page.Width, page.Height), XStringFormats.TopCenter);
                        
            // para1            
            gfx.DrawString(one_one, font1, XBrushes.Black, new XRect(x + 20, y + 205, page.Width, page.Height), XStringFormats.TopLeft);
            gfx.DrawString(one_one_a, font4, XBrushes.Black, new XRect(x + 312, y + 205, page.Width, page.Height), XStringFormats.TopLeft);
            gfx.DrawString(one_one_b, font1, XBrushes.Black, new XRect(x + 417, y + 205, page.Width, page.Height), XStringFormats.TopLeft);
            gfx.DrawString(one_one_c, font4, XBrushes.Black, new XRect(x + 440, y + 205, page.Width, page.Height), XStringFormats.TopLeft);
            gfx.DrawString(one_one_d, font1, XBrushes.Black, new XRect(x + 480, y + 205, page.Width, page.Height), XStringFormats.TopLeft);
            gfx.DrawString(one_two_a, font1, XBrushes.Black, new XRect(x, y + 218, page.Width, page.Height), XStringFormats.TopLeft);
            gfx.DrawString(one_two_b, font4, XBrushes.Black, new XRect(x + 157, y + 218, page.Width, page.Height), XStringFormats.TopLeft);
            gfx.DrawString(one_two_c, font1, XBrushes.Black, new XRect(x + 240, y + 218, page.Width, page.Height), XStringFormats.TopLeft);
            gfx.DrawString(one_two_d, font3, XBrushes.Black, new XRect(x + 285, y + 218, page.Width, page.Height), XStringFormats.TopLeft);

            tf.Alignment = XParagraphAlignment.Justify;
            tf.DrawString(one, font1, XBrushes.Black, new XRect(x, y + 231, 500, 80), XStringFormats.TopLeft);
            
            //para2
            gfx.DrawString(two_one_a, font1, XBrushes.Black, new XRect(x + 20, y + 286, page.Width, page.Height), XStringFormats.TopLeft);
            gfx.DrawString(two_one_b, font1, XBrushes.Black, new XRect(x + 200, y + 286, page.Width, page.Height), XStringFormats.TopLeft);
            
            //para3
            tf.DrawString(three_one, font1, XBrushes.Black, new XRect(x+20, y + 303, 500, 80), XStringFormats.TopLeft);
            //gfx.DrawString(three_one, font1, XBrushes.Black, new XRect(x + 20, y + 293, page.Width, page.Height), XStringFormats.TopLeft);
            gfx.DrawString(three_two_a, font4, XBrushes.Black, new XRect(x, y + 315, page.Width, page.Height), XStringFormats.TopLeft);
            gfx.DrawString(three_two_b, font1, XBrushes.Black, new XRect(x+95, y + 315, page.Width, page.Height), XStringFormats.TopLeft);

           // drawing table
            
           //gfx.DrawLine(pen, 45, 250, 45, 703);
           //gfx.DrawRectangle(XPens.Black, x, y + 200, 50, 100); 
           
     
            // values to be inserted in the table from db
            int x1 = x + 20;
            int y1 = y + 30;

            //drawing header of table
            string column1 = "  " + "Sl. No.";
            string column2a = "   Date of commencement ";
            string column2b = "    for Deposition of EMD";
            string column3a = "   Last Date for depositing";
            string column3b = "   EMD through E-payment";
            string column4 = "Date of E-Auction";

            //SL. NO.
            gfx.DrawRectangle(XPens.Black, x1, y1 + 305, 45, 29);
            gfx.DrawString(column1, font6, XBrushes.Black, new XRect(x1, y1 + 311, page.Width.Point, page.Height.Point), XStringFormats.TopLeft);

            //NITDATE
            gfx.DrawRectangle(XPens.Black, x1 + 44, y1 + 305, 135, 29);
            gfx.DrawString(column2a, font6, XBrushes.Black, new XRect(x1 + 44, y1 + 305, page.Width.Point, page.Height.Point), XStringFormats.TopLeft);
            gfx.DrawString(column2b, font6, XBrushes.Black, new XRect(x1 + 44, y1 + 317, page.Width.Point, page.Height.Point), XStringFormats.TopLeft);

            //EMDDATE
            gfx.DrawRectangle(XPens.Black, x1 + 180, y1 + 305, 137, 29);
            gfx.DrawString(column3a, font6, XBrushes.Black, new XRect(x1 + 182, y1 + 305, page.Width.Point, page.Height.Point), XStringFormats.TopLeft);
            gfx.DrawString(column3b, font6, XBrushes.Black, new XRect(x1 + 182, y1 + 317, page.Width.Point, page.Height.Point), XStringFormats.TopLeft);
          
            // EA DATE
            gfx.DrawRectangle(XPens.Black, x1 + 318, y1 + 305, 105, 29);
            gfx.DrawString(column4, font6, XBrushes.Black, new XRect(x1 + 326, y1 + 311, page.Width.Point, page.Height.Point), XStringFormats.TopLeft);
           

            for (int i = 0; i < pub.Tables[0].Rows.Count; i++)
                {         
               //drawing body of table

                   string sl =(i + 1).ToString();
                   string nitdt = (Convert.ToDateTime(pub.Tables[0].Rows[i].ItemArray[3])).ToShortDateString();
                   string emddt = pub.Tables[0].Rows[i].ItemArray[4].ToString();
                   string eadt = pub.Tables[0].Rows[i].ItemArray[5].ToString();

                   //SL.NO.
                   gfx.DrawRectangle(XPens.Black, x1, y1 + 335, 45, 20);
                   gfx.DrawString(sl, font5, XBrushes.Black, new XRect(x1+18, y1 + 335, page.Width.Point, page.Height.Point), XStringFormats.TopLeft);
                  
                   //NIT DT
                   gfx.DrawRectangle(XPens.Black, x1 + 44, y1 + 335, 135, 20);
                   gfx.DrawString(nitdt, font5, XBrushes.Black, new XRect(x1 + 84, y1 + 335, page.Width.Point, page.Height.Point), XStringFormats.TopLeft);
                   
                   //EMD dt
                   gfx.DrawRectangle(XPens.Black, x1 + 180, y1 + 335, 137, 20);
                   gfx.DrawString(emddt, font5, XBrushes.Black, new XRect(x1 + 224, y1 + 335, page.Width.Point, page.Height.Point), XStringFormats.TopLeft);
                   
                   //EA dt
                   gfx.DrawRectangle(XPens.Black, x1 + 318, y1 + 335, 105, 20);
                   gfx.DrawString(eadt, font5, XBrushes.Black, new XRect(x1 + 344, y1 + 335, page.Width.Point, page.Height.Point), XStringFormats.TopLeft);
                                                           
                   y1= y1+20;

                }
            int y1_1 = (y + 30) + 305 + (20 * (pub.Tables[0].Rows.Count));
            int y2 = y1_1 + 20;
            //para4
            tf.DrawString(four_a, font4, XBrushes.Black, new XRect(x+20, y2+20, 500, 80), XStringFormats.TopLeft);
            //gfx.DrawString(four_a, font4, XBrushes.Black, new XRect(x + 20, y + 435, page.Width, page.Height), XStringFormats.TopLeft);
            gfx.DrawString(four_b, font4, XBrushes.Black, new XRect(x, y2+20+17, page.Width, page.Height), XStringFormats.TopLeft);

            //para5

            gfx.DrawString(five_one_a, font1, XBrushes.Black, new XRect(x + 20, y2 + 20 + 17+20, page.Width, page.Height), XStringFormats.TopLeft);
            gfx.DrawString(five_one_b, font1, XBrushes.Black, new XRect(x + 238, y2 + 20 + 17 + 20, page.Width, page.Height), XStringFormats.TopLeft);

            //para6/7
            tf.DrawString(six, font1, XBrushes.Black, new XRect(x + 480, y2 + 20 + 17 + 20+20, 500, 80), XStringFormats.TopLeft);
            tf.DrawString(seven, font1, XBrushes.Black, new XRect(x + 356, y2 + 20+10 + 17 + 20 + 20, 500, 80), XStringFormats.TopLeft);
            //gfx.DrawString(six, font1, XBrushes.Black, new XRect(x + 500, y + 486, page.Width, page.Height), XStringFormats.TopLeft);
            //gfx.DrawString(seven, font1, XBrushes.Black, new XRect(x + 406, y + 497, page.Width, page.Height), XStringFormats.TopLeft);

            //para6/7

            gfx.DrawString(eight, font7, XBrushes.Black, new XRect(x, y2 + 20 + 10 + 17 + 20 + 20+20, page.Width, page.Height), XStringFormats.TopLeft);
            gfx.DrawString(nine, font8, XBrushes.Black, new XRect(x, y2 + 20 + 10 + 17 + 20 + 20 + 20+20, page.Width, page.Height), XStringFormats.TopLeft);

            //copies
            gfx.DrawString(ten, font8, XBrushes.Black, new XRect(x + 20, y2 + 20 + 10 + 17 + 20 + 20 + 20 + 20+16, page.Width, page.Height), XStringFormats.TopLeft);
            gfx.DrawString(eleven, font8, XBrushes.Black, new XRect(x + 20, y2 + 20 + 10 + 17 + 20 + 20 + 20 + 20 + 16+15, page.Width, page.Height), XStringFormats.TopLeft);
            gfx.DrawString(twelve, font8, XBrushes.Black, new XRect(x + 20, y2 + 20 + 10 + 17 + 20 + 20 + 20 + 20 + 16 + 15+15, page.Width, page.Height), XStringFormats.TopLeft);
            gfx.DrawString(thirteen_a, font8, XBrushes.Black, new XRect(x + 20, y2 + 20 + 10 + 17 + 20 + 20 + 20 + 20 + 16 + 15 + 15+10+5, page.Width, page.Height), XStringFormats.TopLeft);
            gfx.DrawString(thirteen_b, font8, XBrushes.Black, new XRect(x + 30, y2 + 20 + 10 + 17 + 20 + 20 + 20 + 20 + 16 + 15 + 15 + 10+10+5, page.Width, page.Height), XStringFormats.TopLeft);
            gfx.DrawString(thirteen_c, font8, XBrushes.Black, new XRect(x + 30, y2 + 20 + 10 + 17 + 20 + 20 + 20 + 20 + 16 + 15 + 15 + 10 + 10+10+5, page.Width, page.Height), XStringFormats.TopLeft);
            gfx.DrawString(fourteen_a, font8, XBrushes.Black, new XRect(x + 20, y2 + 20 + 10 + 17 + 20 + 20 + 20 + 20 + 16 + 15 + 15 + 10 + 10 + 10+15+5, page.Width, page.Height), XStringFormats.TopLeft);
            gfx.DrawString(fourteen_b, font8, XBrushes.Black, new XRect(x + 30, y2 + 20 + 10 + 17 + 20 + 20 + 20 + 20 + 16 + 15 + 15 + 10 + 10 + 10 + 15+10+5, page.Width, page.Height), XStringFormats.TopLeft);
            gfx.DrawString(fifteen, font8, XBrushes.Black, new XRect(x + 20, y2 + 20 + 10 + 17 + 20 + 20 + 20 + 20 + 16 + 15 + 15 + 10 + 10 + 10 + 15 + 10+15+5, page.Width, page.Height), XStringFormats.TopLeft);
            gfx.DrawString(sixteen, font8, XBrushes.Black, new XRect(x + 20, y2 + 20 + 10 + 17 + 20 + 20 + 20 + 20 + 16 + 15 + 15 + 10 + 10 + 10 + 15 + 10 + 15+10+5+5, page.Width, page.Height), XStringFormats.TopLeft);


            tf.DrawString(seven, font1, XBrushes.Black, new XRect(x + 356, y2 + 20 + 10 + 17 + 20 + 20 + 20 + 20 + 16 + 15 + 15 + 10 + 10 + 10 + 15 + 10 + 15 + 10 + 5+10+10+10, 500, 80), XStringFormats.TopLeft);


            // Save the document...
            string filename = "pub.pdf";
            document.Save(filename);
           
            // ...and start a viewer.
            Process.Start(filename);
        }
        #endregion
        #endregion


        # region approval
        private void tabPageAPPROVAL_Click(object sender, EventArgs e)
        {
            getCon();
            DataSet ds = new DataSet();

            string qry = "Select NITdate from Publication order by NITDate";
            ds = select_data(qry);
            comboBox1.DataSource = ds.Tables[0];
            comboBox1.ValueMember = "NITdate";
            comboBox1.DisplayMember = "NITdate";
            comboBox1.Text = "-Select-";
        }     

        private void cmbcmdty_SelectedIndexChanged_1(object sender, EventArgs e)
        {
            if (cmbcmdty.SelectedIndex != -1)
            {

                if (cmbcmdty.SelectedItem == "Wheat")//if commodity wheat
                {
                    
                    cmbrprt_wheat.Visible = true;
                    cmbrprt_RICE.Visible = false;
                    dttmrprts.Visible = true;
                    txtbrwsrprt.Visible = true;
                    btnupload.Visible = true;
                    btnapsub.Visible = true;
                    btnapedit.Visible = true;
                    lblslctrprt.Visible = true;
                    lblslctdt.Visible = true;
                    dgvrprts.Visible = true;
                    btnshwqty.Visible = true;
                    

                }

                else if (cmbcmdty.SelectedItem == "Raw Rice Grade A")// if commodity rice 
                {
                    
                    cmbrprt_wheat.Visible = false;
                    cmbrprt_RICE.Visible = true;
                    dttmrprts.Visible = true;
                    //txtrice.Visible = true;
                    //lblMT.Visible = true;
                    dttmrprts.Visible = true;
                    txtbrwsrprt.Visible = true;
                    btnupload.Visible = true;
                    btnapsub.Visible = true;
                    btnapedit.Visible = true;
                    lblslctrprt.Visible = true;
                    lblslctdt.Visible = true;
                    dgvrprts.Visible = true;
                    btnshwqty.Visible = true;
                }

                else
                {
                    
                    cmbrprt_wheat.Visible = false;
                    cmbrprt_RICE.Visible = false;
                    dttmrprts.Visible = false;
                   // txtrice.Visible = false;
                    //lblMT.Visible = false;
                    dttmrprts.Visible = true;
                    txtbrwsrprt.Visible = false;
                    btnupload.Visible = false;
                    btnapsub.Visible = false;
                    btnapedit.Visible = false;
                    lblslctrprt.Visible = false;
                    lblslctdt.Visible = false;
                    dgvrprts.Visible = false;
                    btnshwqty.Visible = false;
                    MessageBox.Show("Select valid commodity-Wheat or Rice", "Invalid selection", MessageBoxButtons.OK);
                }

            }
        }

        

        private void btnapsub_Click(object sender, EventArgs e)
        {
            string nitdate1 = comboBox1.SelectedValue.ToString(); 
            //report names
            

            //report dates
            string dtrprt = dttmrprts.Value.ToShortDateString();

            if(cmbcmdty.SelectedItem=="Wheat")      
            {
                string rprt = cmbrprt_wheat.SelectedItem.ToString();
                
                string str = "insert into Reports_Wheat(NITDate,ReportName,ReportDate) values('" + nitdate1 + "', '" + rprt + "','" + dtrprt + "')";
                insert_update_deleted(str);
            }
            else if (cmbcmdty.SelectedItem=="Raw Rice Grade A")
            {

                string rprt = cmbrprt_RICE.SelectedItem.ToString();
                string str = "insert into Reports_Rice(NITDate,ReportName,ReportDate) values('" + nitdate1 + "', '" + rprt+ "','" + dtrprt + "')";
                
                insert_update_deleted(str);
            }
            else
            {
                MessageBox.Show("Select valid commodity-Wheat or Rice", "Invalid selection", MessageBoxButtons.OK);
            }
        }



       

        private void btnupload_Click(object sender, EventArgs e)
        {
            
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1.InitialDirectory = @"C:\";
            openFileDialog1.Title = "Browse Files";
            openFileDialog1.CheckFileExists = true;
            openFileDialog1.CheckPathExists = true;

            openFileDialog1.DefaultExt = "xls";

            openFileDialog1.Filter = "Excel files (*.xls)|*.xls|All files (*.*)|*.*";
            openFileDialog1.RestoreDirectory = true;
            openFileDialog1.ReadOnlyChecked = true;
            openFileDialog1.ShowReadOnly = true;
          
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {             
              txtbrwsrprt.Text  = openFileDialog1.FileName;       
          
            }
            
        }

        private void cmbrprt_RICE_SelectedIndexChanged(object sender, EventArgs e)
        {
            /*if(cmbrprt_RICE.SelectedItem=="Email from Siliguri")
            {
                lblqtyemail.Visible = true;
                txtrice.Visible = true;
                lblMT.Visible = true;

            }

            else 
            {
                lblqtyemail.Visible = false;
                txtrice.Visible = false;
                lblMT.Visible = false;
            }*/
        }

        private void btnshwqty_Click(object sender, EventArgs e)
        {
            DataSet dsrprt = new DataSet();
            if (cmbrprt_RICE.SelectedItem == "Depotwise Rice Position")
            {
                dsrprt = getConexcel1R(txtbrwsrprt.Text);
                //dgvrprts.DataSource = dsrprt.Tables[0];
            }
            else if (cmbrprt_RICE.SelectedItem == "Email from Siliguri")
            {
                dsrprt = getConexcelemail(txtbrwsrprt.Text);
            }
            else if (cmbrprt_wheat.SelectedItem == "Total Wheat Stock")
            {
                dsrprt = getConexcel1W(txtbrwsrprt.Text);
               // dgvrprts.DataSource = dsrprt.Tables[0];
            }
            else if (cmbrprt_wheat.SelectedItem == "Wheat URS Stock" || cmbrprt_wheat.SelectedItem == "Wheat (Only)  Stock")
            {
                dsrprt = getConexcel(txtbrwsrprt.Text);
                //dgvrprts.DataSource = dsrprt.Tables[0];
            }
            else if (cmbrprt_wheat.SelectedItem == "Wheat Rakes Planned")
            {
                dsrprt = getConexcel2(txtbrwsrprt.Text);
                //dgvrprts.DataSource = dsrprt.Tables[0];
            }
            else if (cmbrprt_wheat.SelectedItem == "RATES")
            {
                dsrprt = getConexcel3(txtbrwsrprt.Text);
                //dgvrprts.DataSource = dsrprt.Tables[0];
            }
            else
            {
                MessageBox.Show("Mismatch:Report selected and Browsed");
            }
           
            dgvrprts.DataSource = dsrprt.Tables[0];
        }

        private void btnapedit_Click(object sender, EventArgs e)
        {
            string nitdate = Convert.ToDateTime(comboBox1.SelectedValue ).ToLongDateString();

            //DateTime nitdate = Convert.ToDateTime(comboBox1.SelectedValue.ToString());

            //report dates
            string dtrprt = dttmrprts.Value.ToShortDateString();

            if (cmbcmdty.SelectedItem == "Wheat")
            {
                string rprt = cmbrprt_wheat.SelectedItem.ToString();//report names

                string str = "update Reports_Wheat set ReportDate='" + dtrprt + "' where NITDate=#" + nitdate + "# and ReportName='" + rprt + "'";

                insert_update_deleted(str);
            }
            else if (cmbcmdty.SelectedItem == "Raw Rice Grade A")
            {

                string rprt = cmbrprt_RICE.SelectedItem.ToString();//report names
                string str = "update Reports_Rice set ReportName='" + rprt + "',ReportDate='" + dtrprt + "' where NITDate=#" + nitdate + "# and ReportName='" + rprt + "'";
                insert_update_deleted(str);
            }
            else
            {
                MessageBox.Show("Select valid commodity-Wheat or Rice", "Invalid selection", MessageBoxButtons.OK);
            }
        }
        // show prevnit qty in gridview
        private void btnshwprevqty_Click(object sender, EventArgs e)
        {
            // dynamic controls added

            DataGridViewButtonColumn editbtn = new DataGridViewButtonColumn();
            editbtn.Name = "ENTER";
            editbtn.HeaderText = "ENTER";
            editbtn.Text = "ENTER";
            editbtn.UseColumnTextForButtonValue = true;
            editbtn.FlatStyle = FlatStyle.Popup;

            DataGridViewButtonColumn dltbtn = new DataGridViewButtonColumn();
            dltbtn.Name = "DELETE";
            dltbtn.HeaderText = "DELETE";
            dltbtn.Text = "DELETE";
            dltbtn.UseColumnTextForButtonValue = true;
            dltbtn.FlatStyle = FlatStyle.Popup;

            DataGridViewCheckBoxColumn Select = new DataGridViewCheckBoxColumn();
            Select.Name = "Select";
            //Select.Selected = false;
            Select.FlatStyle = FlatStyle.Popup;
            


            DataSet ds=new DataSet();
            string str;
            //DateTime nitdate=Convert.ToDateTime(dttmprevnit.Value.ToLongDateString());
            string nitdate = dttmprevnit.Value.ToLongDateString();
            if(cmbcmdty.SelectedItem=="Wheat")
            {
                str = "SELECT NITdate,Depot_Selected,District_Selected,Wheat_URS_stock,QtyProposed,Rate FROM ApprovalQty_Wheat WHERE  NITdate=#" + nitdate + "# ";
                ds = select_data(str);
            }
          
            else if(cmbcmdty.SelectedItem=="Raw Rice Grade A")
            {
                  str = "SELECT NITdate,Depot_Selected,District_Selected,RRA_stock,QtyProposed,Rate FROM ApprovalQty_Rice WHERE  NITdate=#" + nitdate + "# ";
                  ds = select_data(str);
            }

            if (ds.Tables[0].Rows.Count != 0)
            {



                dgvprevqty.Columns.Clear();
                dgvprevqty.Visible = true;
                //dgvprevqty.Columns.Add(Select);//checkboxcolumnn
                //dgvprevqty.AutoGenerateColumns = true;
                dgvprevqty.DataSource = ds.Tables[0];
                // dgvprevqty.DataMember = ds.Tables[0].ToString();
                dgvprevqty.Columns.Add(editbtn);
                dgvprevqty.Refresh();
                dgvprevqty.AllowUserToAddRows = true;
                dgvDates.AllowUserToResizeColumns = true;
               
                dgvprevqty.Refresh();
            }

            else
            {
                MessageBox.Show("No values in DB for this NITDate", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                dgvprevqty.Visible = false;
            }
            
        }

        # region insert to new NITdate by editing prev qty gridview
        private void dgvprevqty_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            // Check if wanted column was clicked 
            if (e.ColumnIndex == 6 && e.RowIndex >= 0)
            {
                //Perform on edit button click code

                string nitdate = comboBox1.SelectedValue.ToString(); ;
                string depot = dgvprevqty.CurrentRow.Cells[1].Value.ToString();
                string district = dgvprevqty.CurrentRow.Cells[2].Value.ToString();
                double qty_present = Convert.ToDouble(dgvprevqty.CurrentRow.Cells[3].Value.ToString());
                double qty_proposed = Convert.ToDouble(dgvprevqty.CurrentRow.Cells[4].Value.ToString());
                double rate = Convert.ToDouble(dgvprevqty.CurrentRow.Cells[5].Value.ToString());
               
                if (cmbcmdty.SelectedItem == "Wheat")
                {
                    string qry_updt = "insert into ApprovalQty_Wheat (NITDate,Depot_Selected,District_Selected,Wheat_URS_stock,QtyProposed,Rate) values('" + nitdate + "','" + depot + "','" + district + "'," + qty_present + "," + qty_proposed + ","+rate+")";
                    insert_update_deleted(qry_updt);

                }
                else if (cmbcmdty.SelectedItem == "Raw Rice Grade A")
                {
                    string qry_updt = "insert into ApprovalQty_Rice (NITDate,Depot_Selected,District_Selected,RRA_stock,QtyProposed,Rate) values('" + nitdate + "','" + depot + "','" + district + "'," + qty_present + "," + qty_proposed + "," + rate + ")";
                    insert_update_deleted(qry_updt);
                }            

            }
            
            else
            {
                
            }
        }

        # endregion         
        //show proposed qty in gridview
        private void btnpropsedqty_Click(object sender, EventArgs e)
        {
            // dynamic controls added

            DataGridViewButtonColumn editbtn = new DataGridViewButtonColumn();
            editbtn.Name = "EDIT";
            editbtn.HeaderText = "EDIT";
            editbtn.Text = "EDIT";
            editbtn.UseColumnTextForButtonValue = true;
            editbtn.FlatStyle = FlatStyle.Popup;

            DataGridViewButtonColumn dltbtn = new DataGridViewButtonColumn();
            dltbtn.Name = "DELETE";
            dltbtn.HeaderText = "DELETE";
            dltbtn.Text = "DELETE";
            dltbtn.UseColumnTextForButtonValue = true;
            dltbtn.FlatStyle = FlatStyle.Popup;

            DataGridViewCheckBoxColumn Select = new DataGridViewCheckBoxColumn();
            Select.Name = "Select";
            //Select.Selected = false;
            Select.FlatStyle = FlatStyle.Popup;



            DataSet ds = new DataSet();
            string str;
           // DateTime nitdate = Convert.ToDateTime(comboBox1.SelectedValue.ToString());
            string nitdate = Convert.ToDateTime(comboBox1.SelectedValue).ToLongDateString(); 
            
            if (cmbcmdty.SelectedItem == "Wheat")
            {
               // str = "SELECT NITdate,Depot_Selected,District_Selected,Wheat_URS_stock,QtyProposed,Rate FROM ApprovalQty_Wheat WHERE  NITdate=#" + nitdate + "# ";
                str = "SELECT * FROM ApprovalQty_Wheat WHERE  NITdate=#" + nitdate + "# ";
               
                ds = select_data(str);
            }

            else if (cmbcmdty.SelectedItem == "Raw Rice Grade A")
            {
                //str = "SELECT NITdate,Depot_Selected,District_Selected,RRA_stock,QtyProposed,Rate FROM ApprovalQty_Rice WHERE  NITdate=#" + nitdate + "# ";
                str = "SELECT * FROM ApprovalQty_Rice WHERE  NITdate=#" + nitdate + "# ";
               
                ds = select_data(str);
            }

            if (ds.Tables[0].Rows.Count > 0)
            {



               dgvcurrentqty.Columns.Clear();
                dgvcurrentqty.Visible = true;
                //dgvprevqty.Columns.Add(Select);//checkboxcolumnn
               // dgvcurrentqty.AutoGenerateColumns = true;
                dgvcurrentqty.DataSource = ds.Tables[0];
                // dgvprevqty.DataMember = ds.Tables[0].ToString();
                dgvcurrentqty.Columns["ID"].ReadOnly = true;
                dgvcurrentqty.Columns.Add(editbtn);
                dgvcurrentqty.Refresh();
                //dgvprevqty.AllowUserToAddRows = true;
              
            }

            else
            {
                dgvprevqty.Visible = false;
                dgvcurrentqty.Visible = false;
                DialogResult dr=MessageBox.Show("No values for this NITDate,Enter values?", "Empty", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if(dr==DialogResult.Yes)
                {
                    DataGridViewButtonColumn addbtn = new DataGridViewButtonColumn();
                    addbtn.Name = "ADD";
                    addbtn.HeaderText = "ADD";
                    addbtn.Text = "ADD";
                    addbtn.UseColumnTextForButtonValue = true;
                    addbtn.FlatStyle = FlatStyle.Popup;

                    
                    dgvcurrentqty.Visible = false;
                    dgvprevqty.Visible = false;
                   
                   

                    if (cmbcmdty.SelectedItem == "Wheat")
                    {
                        str = "SELECT NITdate,Depot_Selected,District_Selected,Wheat_URS_stock,QtyProposed,Rate FROM ApprovalQty_Wheat WHERE  NITdate=#" + nitdate + "# ";
                        ds = select_data(str);
                       
                    }

                    else if (cmbcmdty.SelectedItem == "Raw Rice Grade A")
                    {
                        str = "SELECT NITdate,Depot_Selected,District_Selected,RRA_stock,QtyProposed,Rate FROM ApprovalQty_Rice WHERE  NITdate=#" + nitdate + "# ";
                        ds = select_data(str);
                    }

                    if (ds.Tables[0].Rows.Count == 0)
                    {
                       // if (cmbcmdty.SelectedItem == "Wheat")
                        //{
                            dgvcurrntentry.Refresh();
                            dgvcurrntentry.Columns.Clear();
                            dgvcurrntentry.Visible = true;
                            dgvcurrntentry.DataSource = ds.Tables[0];
                            dgvcurrntentry.Columns.Add(addbtn);
                            dgvcurrntentry.AllowUserToAddRows = true;
                            dgvRICEcrntqtyentry.Visible = false;
                            
                            
                       // }
                       /* else if (cmbcmdty.SelectedItem == "Raw Rice Grade A")
                        {
                            dgvRICEcrntqtyentry.Refresh();
                            dgvRICEcrntqtyentry.Columns.Clear();
                            dgvRICEcrntqtyentry.Visible = true;
                            
                            dgvRICEcrntqtyentry.DataSource = ds.Tables[0];
                            dgvRICEcrntqtyentry.Columns.Add(addbtn);
                            dgvRICEcrntqtyentry.AllowUserToAddRows = true;
                            dgvcurrntentry.Visible = false;
                          
                           
                        }*/
                    }
                    


               }
                else if (dr == DialogResult.No)
                {
                    dgvcurrentqty.Visible = false;
                    dgvprevqty.Visible = false;
                    dgvcurrntentry.Visible = false;
                    dgvRICEcrntqtyentry.Visible = false;
                }
            }
            
        }
        // enter proposed qty in wheat
        private void dgvcurrntentry_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            // Check if wanted column was clicked 
            if (e.ColumnIndex == 6 && e.RowIndex >= 0)
            {
                //Perform on edit button click code

                string nitdate = Convert.ToDateTime(comboBox1.SelectedValue).ToLongDateString();
                string depot = dgvcurrntentry.CurrentRow.Cells[1].Value.ToString();
                string district = dgvcurrntentry.CurrentRow.Cells[2].Value.ToString();
                double qty_present = Convert.ToDouble(dgvcurrntentry.CurrentRow.Cells[3].Value.ToString());
                double qty_proposed = Convert.ToDouble(dgvcurrntentry.CurrentRow.Cells[4].Value.ToString());
                double rate = Convert.ToDouble(dgvcurrntentry.CurrentRow.Cells[5].Value.ToString());

                if (cmbcmdty.SelectedItem == "Wheat")
                {
                    string qry_updt = "insert into ApprovalQty_Wheat (NITDate,Depot_Selected,District_Selected,Wheat_URS_stock,QtyProposed,Rate) values('" + nitdate + "','" + depot + "','" + district + "'," + qty_present + "," + qty_proposed + "," + rate + ")";
                    insert_update_deleted(qry_updt);

                }
                else if (cmbcmdty.SelectedItem == "Raw Rice Grade A")
                {
                    string qry_updt = "insert into ApprovalQty_Rice (NITDate,Depot_Selected,District_Selected,RRA_stock,QtyProposed,Rate) values('" + nitdate + "','" + depot + "','" + district + "'," + qty_present + "," + qty_proposed + "," + rate + ")";
                    insert_update_deleted(qry_updt);
                }

            }

            else
            {

            }
        }
       
        // edit propsed quantity if required 
        private void dgvcurrentqty_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            // Check if wanted column was clicked 
            if (e.ColumnIndex == 7 && e.RowIndex >= 0)
            {
                //Perform on edit button click code
                int id = Int32.Parse(dgvcurrentqty.CurrentRow.Cells[0].Value.ToString());
                string nitdate = Convert.ToDateTime(comboBox1.SelectedValue).ToLongDateString();
                string depot = dgvcurrentqty.CurrentRow.Cells[2].Value.ToString();
                string district = dgvcurrentqty.CurrentRow.Cells[3].Value.ToString();
                double qty_present = Convert.ToDouble(dgvcurrentqty.CurrentRow.Cells[4].Value.ToString());
                double qty_proposed = Convert.ToDouble(dgvcurrentqty.CurrentRow.Cells[5].Value.ToString());
                double rate = Convert.ToDouble(dgvcurrentqty.CurrentRow.Cells[6].Value.ToString());

                if (cmbcmdty.SelectedItem == "Wheat")
                {
                   // string qry_updt = "update ApprovalQty_Wheat set Depot_Selected='" + depot + "',District_Selected='" + district + "',Wheat_URS_stock=" + qty_present + ",QtyProposed=" + qty_proposed + ",Rate=" + rate + " where NITDate=#" + nitdate + "#";
                    string qry_updt = "update ApprovalQty_Wheat set Depot_Selected='" + depot + "',District_Selected='" + district + "',Wheat_URS_stock=" + qty_present + ",QtyProposed=" + qty_proposed + ",Rate=" + rate + " where ID=" + id+ "";
                   
                    insert_update_deleted(qry_updt);  

                }
                else if (cmbcmdty.SelectedItem == "Raw Rice Grade A")
                {
                   // string qry_updt = "update ApprovalQty_Rice set Depot_Selected='" + depot + "',District_Selected='" + district + "',Wheat_URS_stock=" + qty_present + ",QtyProposed=" + qty_proposed + ",Rate=" + rate + " where NITDate=#" + nitdate + "#";
                    string qry_updt = "update ApprovalQty_Rice set Depot_Selected='" + depot + "',District_Selected='" + district + "',RRA_stock =" + qty_present + ",QtyProposed=" + qty_proposed + ",Rate=" + rate + " where ID=" + id + "";
                    
                    insert_update_deleted(qry_updt); 
                }
                else
                {
                    MessageBox.Show("Select either Wheat or Rice in Commodity", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }

            }

            else
            {

            }
        }

      private void btnvalidnit_Click(object sender, EventArgs e)
        {
           /* int flag = 0;
            DataSet ds = new DataSet();
            string nitdate = comboBox1.SelectedValue.ToString();
            string  qry = "Select NITdate from Publication order by NITDate";
            ds = select_data(qry);
            for(int i=0;i<ds.Tables[0].Rows.Count;i++)
            {
                if (nitdate == Convert.ToDateTime(ds.Tables[0].Rows[i].ItemArray[0].ToString()).ToShortDateString())
                {
                    flag = flag + 1;
                }
                
            }
            if(flag==0)
            {
                MessageBox.Show("Invalid, Enter a valid NIT date from the list");
                rtbvaliddts.Visible = true;
                dynamic ColCount = ds.Tables[0].Columns.Count;
                foreach (DataRow row in ds.Tables[0].Rows)
                {
                    for (int i = 0; i <= ColCount - 1; i++)
                    {
                        rtbvaliddts.Text += Convert.ToDateTime( row[i]).ToShortDateString() ;
                        if (i == ColCount - 1)
                        {
                            rtbvaliddts.Text += "\r";
                        }
                    }
                }
            }
            else
            {
                MessageBox.Show("Valid");
            }
          */
        }

        private void cmbauthor_SelectedIndexChanged(object sender, EventArgs e)
        {
            lbl1stnit.Visible = true;
            cmb1stNIT.Visible = true;
        }

        private void cmb1stNIT_SelectedIndexChanged(object sender, EventArgs e)
        {

            lblnsprevnit.Visible = true;
            dttmnsprevnit.Visible = true;
            lblfinyr.Visible = true;
            cmbfinyr.Visible = true;
           
        }

        private void btncopyns_Click(object sender, EventArgs e)
        {
            rtbns.SelectAll();
            rtbns.Cut();
            rtbns.DeselectAll();

        }

       
        #region NIT PDF generation
        private void btngnrtNIT_Click(object sender, EventArgs e)
        {
            DataSet nit = new DataSet();
            string str;
            string nitdate = Convert.ToDateTime(comboBox1.SelectedValue).ToLongDateString();
            if (cmbcmdty.SelectedIndex != -1)
            {
                if(cmbcmdty.SelectedItem=="Wheat")
                {
                    str = "SELECT * FROM ApprovalQty_Wheat WHERE  NITDate=#" + nitdate + "# ";                  
                    nit = select_data(str);
                    if (nit.Tables[0].Rows.Count != 0)
                    {
                        NIT_Creation_W(nit);
                    }
                    else
                    {
                        MessageBox.Show("No Values in DB for this NITdate or commodity", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
                else if (cmbcmdty.SelectedItem == "Raw Rice Grade A")
                {
                    str = "SELECT * FROM ApprovalQty_Rice WHERE  NITDate=#" + nitdate + "# ";
                    nit = select_data(str);
                    if (nit.Tables[0].Rows.Count != 0)
                    {
                       NIT_Creation_R(nit);
                    }
                    else
                    {
                        MessageBox.Show("No Values in DB for this NITdate or commodity", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
                    
                    

                }
                else
                    MessageBox.Show("Select Wheat or Rice in Commodity", "Wrong Selection", MessageBoxButtons.OK, MessageBoxIcon.Error);
            
           
        }
        private void NIT_Creation_W(DataSet nit)
        {

            //all the static items in the page hold in variables

            string refno = "Ref. No.:- " + "Com.32/OMSS(D)/BULK/WHEAT/2015-16";
            string date = "Dated:" + Convert.ToDateTime(nit.Tables[0].Rows[0].ItemArray[1].ToString()).ToShortDateString();
            string notice = "T E N D E R   N O T I C E";
            string one_one = "Financial bid is hereby invited for sale of wheat from the FCI empanelled bulk consumers/";
            string one_one_a = "traders through";
            string one_one_b = "E-AUCTION";
            string one_one_c = "only. The identified depots vis-à-vis the quantity tendered is as";
            string one_one_d = "under :";
            string two_one_a = "The  minimum  quantity of  wheat for which a bid can be  submitted  would  be 100 MT  and";
            string two_one_b = "maximum would be 3500 MT. The list of DEPOT-WISE Applicable Reserve Price for West Bengal is available on www.fciweb.nic.in,www.ncdexspot.com. Further a Market fee of 0.5% is payable on the Reserve Price.";
            string three= "Tender schedule for the above quantity is as under :- 	 ";
            string four = "The EMD/COST is to be deposited in the following Account Exclusively for OMSS (D) :";
            string five= "Name of Beneficiary : FOOD CORPORATION OF INDIA";
            string five_a = "Name of Branch : SBI COMMERCIAL BRANCH, 24 PARK STREET, KOL-16";
            string five_b = "Account No : 33684806347";
            string five_c = "IFS Code : SBIN0007502";
            string six = "Sd/-";
            string seven = "For General Manager (WB)";
            string eight = "NOT TO BE PUBLISHED/DISTRIBUTION :";
            string nine = "1. The Director, DDP&S";
            string nine_a = "Deptt. Of Food & Supplies, Govt. of West Bengal";
            string nine_b = "11A, Mirza Ghalib Street, Kolkata – 700 087";
            string ten = "2. The Director of Rationing  ";
            string ten_a = "Deptt. Of Food & Supplies, Govt. of West Bengal";
            string ten_b = "11A, Mirza Ghalib Street, Kolkata – 700 087";
            string eleven = "3.	All Area Managers";
            string eleven_a = "FCI –with the request to issue stocks as per FIFO. ";
            string twelve = "4. The ED(Zone), FCI, ZO (E), Kolkata ";
            string thirteen = "5. The General Manager (Sales), FCI, Hqrs., New Delhi";
            string fourteen_a = "6.	The Calcutta Flour Mills Association,";
            string fourteen_b = "15, Brabourne Road, Kolkata-700001…with request to circulate among the members of your association";
            string fifteen = "7. The DGM(F&A), FCI, R.O. Kolkata";
            string sixteen = "8. The AGM (Movt/QC/Stg/A/cs)/M (Cash), FCI, R.O. Kolkata… for information and necessary action ";
            string seventeen = "9. Manager (Computer), FCI, RO, Kol…with request to upload the NIT and MTF ";
            string seventeen_a = "in FCI Website www.fciweb.nic.in and CPPP of GOI www.eprocure.gov.in by " + " " +Convert.ToDateTime(nit.Tables[0].Rows[0].ItemArray[1].ToString()).ToShortDateString();
            string eighteen = "10. Sri Manish Tripathi/ Sri Pallav Bhatt, NCDEX Spot Exchange Ltd.";
            string nineteen = "Manager (Comml)";





            // Create a new PDF document
            PdfDocument document = new PdfDocument();
            document.Info.Title = "Created with PDFsharp";

            // Create an empty page
            PdfPage page = document.AddPage();

            // Get an XGraphics object for drawing
            XGraphics gfx = XGraphics.FromPdfPage(page);
            XTextFormatter tf = new XTextFormatter(gfx);


            // Draw FCI logo of the page
            string jpegSamplePath = "../IMAGES/logo.jpg";
            XImage image = XImage.FromFile(jpegSamplePath);
            gfx.DrawImage(image, 44, 70, 500, 80);

            // Create a font

            XFont font1 = new XFont("Calibri (Body)", 12, XFontStyle.Regular);
            XFont font2 = new XFont("Calibri (Body)", 12, XFontStyle.Underline);
            XFont font3 = new XFont("Calibri (Body)", 12, XFontStyle.Italic);
            XFont font4 = new XFont("Calibri (Body)", 12, XFontStyle.Bold);
            XFont font5 = new XFont("Calibri (Body)", 10, XFontStyle.Regular);
            XFont font6 = new XFont("Calibri (Body)", 11, XFontStyle.Bold);
            XFont font7 = new XFont("Calibri (Body)", 09, XFontStyle.Bold);
            XFont font8 = new XFont("Calibri (Body)", 09, XFontStyle.Regular);

            int x = 44;
            int y = 30;

            // Draw the text

            // file refno
            gfx.DrawString(refno, font1, XBrushes.Black, new XRect(x, y + 130, page.Width, page.Height), XStringFormats.TopLeft);
            // file date == date of nit
            gfx.DrawString(date, font1, XBrushes.Black, new XRect(x + 390, y + 130, page.Width, page.Height), XStringFormats.TopLeft);
            //notice line-centre alignment
            gfx.DrawString(notice, font4, XBrushes.Black, new XRect(x - 44, y + 155, page.Width, page.Height), XStringFormats.TopCenter);
           

            // para1            
            gfx.DrawString(one_one, font1, XBrushes.Black, new XRect(x + 20, y + 170, page.Width, page.Height), XStringFormats.TopLeft);            
            gfx.DrawString(one_one_a, font1, XBrushes.Black, new XRect(x, y + 183, page.Width, page.Height), XStringFormats.TopLeft);
            gfx.DrawString(one_one_b, font4, XBrushes.Black, new XRect(x +85, y + 183, page.Width, page.Height), XStringFormats.TopLeft);
            gfx.DrawString(one_one_c, font1, XBrushes.Black, new XRect(x + 160, y + 183, page.Width, page.Height), XStringFormats.TopLeft);
            gfx.DrawString(one_one_d, font1, XBrushes.Black, new XRect(x, y + 196, page.Width, page.Height), XStringFormats.TopLeft);

           
           
            // drawing table1           

            // values to be inserted in the table from db
            int x1 = x + 20;
            int y1 = y + 30;

            //drawing header of table1
            string column1 = "Sl. No.";
            string column2a = "Depot";
            string column3 = "District Office";
            string column4a = "Qty put to tender";
            string column4b = "(in MT)";
            string column5 = "Depot-Wise ";
            string column5a = "Applicable RESERVE ";
            string column5b = "PRICE(Rs./Qtl)";

           

            //SL. NO.
            gfx.DrawRectangle(XPens.Black, x1, y1 + 190, 45, 32);
            gfx.DrawString(column1, font6, XBrushes.Black, new XRect(x1+5, y1 + 200, page.Width.Point, page.Height.Point), XStringFormats.TopLeft);

            //Depot
            gfx.DrawRectangle(XPens.Black, x1 + 44, y1 + 190, 105, 32);
            gfx.DrawString(column2a, font6, XBrushes.Black, new XRect(x1 + 80, y1 + 200, page.Width.Point, page.Height.Point), XStringFormats.TopLeft);
           

            //District
            gfx.DrawRectangle(XPens.Black, x1 + 150, y1 + 190, 105, 32);
            gfx.DrawString(column3, font6, XBrushes.Black, new XRect(x1 + 170, y1 + 200, page.Width.Point, page.Height.Point), XStringFormats.TopLeft);
           

            // Qty
            gfx.DrawRectangle(XPens.Black, x1 + 255, y1 + 190, 105, 32);
           gfx.DrawString(column4a, font6, XBrushes.Black, new XRect(x1 + 260, y1 + 190, page.Width.Point, page.Height.Point), XStringFormats.TopLeft);
           gfx.DrawString(column4b, font6, XBrushes.Black, new XRect(x1 + 280, y1 + 205, page.Width.Point, page.Height.Point), XStringFormats.TopLeft);
          

           // Rate
           gfx.DrawRectangle(XPens.Black, x1 + 360, y1 + 190, 115, 32);
           gfx.DrawString(column5, font6, XBrushes.Black, new XRect(x1 + 385, y1 + 190, page.Width.Point, page.Height.Point), XStringFormats.TopLeft);
           gfx.DrawString(column5a, font6, XBrushes.Black, new XRect(x1 + 365, y1 + 200, page.Width.Point, page.Height.Point), XStringFormats.TopLeft);
           gfx.DrawString(column5b, font6, XBrushes.Black, new XRect(x1 + 385, y1 + 210, page.Width.Point, page.Height.Point), XStringFormats.TopLeft);


            for (int i = 0; i < nit.Tables[0].Rows.Count; i++)
            {
                //drawing body of table

                string sl = (i + 1).ToString();                
                string depot = nit.Tables[0].Rows[i].ItemArray[2].ToString();
                string dis = nit.Tables[0].Rows[i].ItemArray[3].ToString();
                string qty = nit.Tables[0].Rows[i].ItemArray[5].ToString();
                string rate = nit.Tables[0].Rows[i].ItemArray[6].ToString();

                //SL.NO.
                gfx.DrawRectangle(XPens.Black, x1, y1 + 222, 45, 20);
                gfx.DrawString(sl, font5, XBrushes.Black, new XRect(x1+20, y1 + 222, page.Width.Point, page.Height.Point), XStringFormats.TopLeft);

                //Depot
                gfx.DrawRectangle(XPens.Black, x1 + 44, y1 + 222, 105, 20);
                gfx.DrawString(depot, font5, XBrushes.Black, new XRect(x1 + 47, y1 + 222, page.Width.Point, page.Height.Point), XStringFormats.TopLeft);

                //District
                gfx.DrawRectangle(XPens.Black, x1 + 150, y1 + 222, 105, 20);
                gfx.DrawString(dis, font5, XBrushes.Black, new XRect(x1 + 153, y1 + 222, page.Width.Point, page.Height.Point), XStringFormats.TopLeft);

                //QTY
                gfx.DrawRectangle(XPens.Black, x1 + 255, y1 + 222, 105, 20);
                gfx.DrawString(qty, font5, XBrushes.Black, new XRect(x1 + 320, y1 + 222, page.Width.Point, page.Height.Point), XStringFormats.TopLeft);

                //rate
                gfx.DrawRectangle(XPens.Black, x1 + 360, y1 + 222, 115, 20);
                gfx.DrawString(rate, font5, XBrushes.Black, new XRect(x1 + 400, y1 + 222, page.Width.Point, page.Height.Point), XStringFormats.TopLeft);

                y1 = y1 + 20;

            }

            int y1_1 = (y + 30) + 222 + (20 * (nit.Tables[0].Rows.Count));
            gfx.DrawRectangle(XPens.Black, x1, y1_1, 45, 20);
            gfx.DrawString("Total", font6, XBrushes.Black, new XRect(x1 + 47, y1_1, page.Width.Point, page.Height.Point), XStringFormats.TopLeft);
            gfx.DrawRectangle(XPens.Black, x1+44, y1_1, 105, 20);
            gfx.DrawRectangle(XPens.Black, x1+150, y1_1, 105, 20);
            gfx.DrawRectangle(XPens.Black, x1 + 255, y1_1, 105, 20);
            gfx.DrawString("1,00,000", font6, XBrushes.Black, new XRect(x1 + 300, y1_1, page.Width.Point, page.Height.Point), XStringFormats.TopLeft);
            gfx.DrawRectangle(XPens.Black, x1+360, y1_1, 115, 20);
            


           int y2 = y1_1+20;

         
            //int y2 = y + 352;

            //para2

            gfx.DrawString(two_one_a, font1, XBrushes.Black, new XRect(x + 20, y2+10+10, page.Width, page.Height), XStringFormats.TopLeft);
            //gfx.DrawString(two_one_b, font1, XBrushes.Black, new XRect(x, y2+23, page.Width, page.Height), XStringFormats.TopLeft);
            tf.Alignment = XParagraphAlignment.Justify;
            tf.DrawString(two_one_b, font1, XBrushes.Black, new XRect(x, y2+23+10, 500, 80), XStringFormats.TopLeft);

            //para3

            gfx.DrawString(three, font1, XBrushes.Black, new XRect(x + 20, y2 + 65+10, page.Width, page.Height), XStringFormats.TopLeft);
           

            //table2

            //drawing header of table2
            string column2_1 = "Date of commencement ";
            string column2__1a = "for Depositing EMD in ";
            string column2__1b = "the below mentioned";
            string column2_1c = "A/C";
            string column2_2 = "Last Date for";
            string column2_2a = "depositing EMD";
            string column2_2b = "through E-payment";
            string column2_3a = "Starting date and ";
            string column2_3b = "time for online Bidding ";
            string column2_4a = "End Date and time  ";
            string column2_4b = "for online bidding";

            int y3 = y2 + 20;
            //dt1
            gfx.DrawRectangle(XPens.Black, x1, y3 + 78, 125, 45);
            gfx.DrawString(column2_1, font6, XBrushes.Black, new XRect(x1 +5, y3 + 78, page.Width.Point, page.Height.Point), XStringFormats.TopLeft);
            gfx.DrawString(column2__1a, font6, XBrushes.Black, new XRect(x1 + 5, y3 + 88, page.Width.Point, page.Height.Point), XStringFormats.TopLeft);
            gfx.DrawString(column2__1b, font6, XBrushes.Black, new XRect(x1 + 5, y3 + 98, page.Width.Point, page.Height.Point), XStringFormats.TopLeft);
            gfx.DrawString(column2_1c, font6, XBrushes.Black, new XRect(x1 + 5, y3 + 108, page.Width.Point, page.Height.Point), XStringFormats.TopLeft);
          
            //Dt2
            gfx.DrawRectangle(XPens.Black, x1 + 125, y3 + 78, 105, 45);
            gfx.DrawString(column2_2, font6, XBrushes.Black, new XRect(x1 + 135, y3 + 78, page.Width.Point, page.Height.Point), XStringFormats.TopLeft);
            gfx.DrawString(column2_2a, font6, XBrushes.Black, new XRect(x1 + 135, y3 + 88, page.Width.Point, page.Height.Point), XStringFormats.TopLeft);
            gfx.DrawString(column2_2b, font6, XBrushes.Black, new XRect(x1 + 135, y3 + 98, page.Width.Point, page.Height.Point), XStringFormats.TopLeft);
            

            //Dt3
            gfx.DrawRectangle(XPens.Black, x1 + 230, y3 + 78, 145, 45);
            gfx.DrawString(column2_3a, font6, XBrushes.Black, new XRect(x1 + 240, y3 + 78, page.Width.Point, page.Height.Point), XStringFormats.TopLeft);
            gfx.DrawString(column2_3b, font6, XBrushes.Black, new XRect(x1 + 240, y3 + 88, page.Width.Point, page.Height.Point), XStringFormats.TopLeft);


            // dt4
            gfx.DrawRectangle(XPens.Black, x1 + 375, y3 + 78, 135, 45);
            gfx.DrawString(column2_4a, font6, XBrushes.Black, new XRect(x1 + 380, y3 + 78, page.Width.Point, page.Height.Point), XStringFormats.TopLeft);
            gfx.DrawString(column2_4b, font6, XBrushes.Black, new XRect(x1 + 380, y3 + 88, page.Width.Point, page.Height.Point), XStringFormats.TopLeft);


            

            //body of table2

            string column2_b_1 = Convert.ToDateTime(nit.Tables[0].Rows[0].ItemArray[1].ToString()).ToShortDateString();
            string column2__b_2 = Convert.ToDateTime(column2_b_1).AddDays(2).ToShortDateString();
            string column2__b_3 = Convert.ToDateTime(column2_b_1).AddDays(3).ToShortDateString() + "from 11:00 a.m.";
            string column2__b_4 = Convert.ToDateTime(column2_b_1).AddDays(3).ToShortDateString() + "upto 02:00 pm";

            //dt1
            gfx.DrawRectangle(XPens.Black, x1, y3 + 123, 125, 20);
            gfx.DrawString(column2_b_1, font6, XBrushes.Black, new XRect(x1 + 5, y3 + 127, page.Width.Point, page.Height.Point), XStringFormats.TopLeft);
            
            //Dt2
            gfx.DrawRectangle(XPens.Black, x1 + 125, y3 + 123, 105, 20);
            gfx.DrawString(column2__b_2, font6, XBrushes.Black, new XRect(x1 + 135, y3 + 127, page.Width.Point, page.Height.Point), XStringFormats.TopLeft);
            

            //Dt3
            gfx.DrawRectangle(XPens.Black, x1 + 230, y3 + 123, 145, 20);
            gfx.DrawString(column2__b_3, font6, XBrushes.Black, new XRect(x1 + 240, y3 + 127, page.Width.Point, page.Height.Point), XStringFormats.TopLeft);
          

            // dt4
            gfx.DrawRectangle(XPens.Black, x1 + 375, y3 + 123, 135, 20);
            gfx.DrawString(column2__b_4, font6, XBrushes.Black, new XRect(x1 + 380, y3 + 127, page.Width.Point, page.Height.Point), XStringFormats.TopLeft);
            

            //para4
            gfx.DrawString(four, font4, XBrushes.Black, new XRect(x, y3 + 140+10, page.Width, page.Height), XStringFormats.TopLeft);
            gfx.DrawString(five, font1, XBrushes.Black, new XRect(x, y3 + 153+10, page.Width, page.Height), XStringFormats.TopLeft);
            gfx.DrawString(five_a, font1, XBrushes.Black, new XRect(x, y3 + 166+10, page.Width, page.Height), XStringFormats.TopLeft);
            gfx.DrawString(five_b, font1, XBrushes.Black, new XRect(x, y3 + 179+10, page.Width, page.Height), XStringFormats.TopLeft);
            gfx.DrawString(five_c, font1, XBrushes.Black, new XRect(x, y3 + 192+10, page.Width, page.Height), XStringFormats.TopLeft);
        
            //para5/6
            tf.DrawString(six, font1, XBrushes.Black, new XRect(x + 480, y3 + 222, 500, 80), XStringFormats.TopLeft);
            tf.DrawString(seven, font1, XBrushes.Black, new XRect(x + 356, y3 + 237, 500, 80), XStringFormats.TopLeft);

            // Create an empty page
            PdfPage page2 = document.AddPage();

            // Get an XGraphics object for drawing
            XGraphics gfx2= XGraphics.FromPdfPage(page2);
            XTextFormatter tf2 = new XTextFormatter(gfx2);
            //para7/8

            gfx2.DrawString(eight, font7, XBrushes.Black, new XRect(x, y, page2.Width, page2.Height), XStringFormats.TopLeft);
            
            
            //copies
            gfx2.DrawString(nine, font8, XBrushes.Black, new XRect(x + 20, y+20, page2.Width, page2.Height), XStringFormats.TopLeft);
            gfx2.DrawString(nine_a, font8, XBrushes.Black, new XRect(x + 30, y + 30, page2.Width, page2.Height), XStringFormats.TopLeft);
            gfx2.DrawString(nine_b, font8, XBrushes.Black, new XRect(x + 30, y + 40, page2.Width, page2.Height), XStringFormats.TopLeft);
            gfx2.DrawString(ten, font8, XBrushes.Black, new XRect(x + 20, y + 60, page2.Width, page2.Height), XStringFormats.TopLeft);
            gfx2.DrawString(ten_a, font8, XBrushes.Black, new XRect(x + 30, y + 70, page2.Width, page2.Height), XStringFormats.TopLeft);
            gfx2.DrawString(ten_b, font8, XBrushes.Black, new XRect(x + 30, y + 80, page2.Width, page2.Height), XStringFormats.TopLeft);
            gfx2.DrawString(eleven, font8, XBrushes.Black, new XRect(x + 20, y + 100, page2.Width, page2.Height), XStringFormats.TopLeft);
            gfx2.DrawString(eleven_a, font8, XBrushes.Black, new XRect(x + 30, y + 110, page2.Width, page2.Height), XStringFormats.TopLeft);
            gfx2.DrawString(twelve, font8, XBrushes.Black, new XRect(x + 20, y + 130, page2.Width, page2.Height), XStringFormats.TopLeft);
            gfx2.DrawString(thirteen, font8, XBrushes.Black, new XRect(x + 20, y + 150, page2.Width, page2.Height), XStringFormats.TopLeft);
            gfx2.DrawString(fourteen_a, font8, XBrushes.Black, new XRect(x + 20, y + 170, page2.Width, page2.Height), XStringFormats.TopLeft);
            gfx2.DrawString(fourteen_b, font8, XBrushes.Black, new XRect(x + 30, y + 180, page2.Width, page2.Height), XStringFormats.TopLeft);
            gfx2.DrawString(fifteen, font8, XBrushes.Black, new XRect(x + 20, y + 200, page2.Width, page2.Height), XStringFormats.TopLeft);
            gfx2.DrawString(sixteen, font8, XBrushes.Black, new XRect(x + 20, y + 220, page2.Width, page2.Height), XStringFormats.TopLeft);
            gfx2.DrawString(seventeen, font8, XBrushes.Black, new XRect(x + 20, y + 240, page2.Width, page2.Height), XStringFormats.TopLeft);
            gfx2.DrawString(seventeen_a, font8, XBrushes.Black, new XRect(x + 30, y + 250, page2.Width, page2.Height), XStringFormats.TopLeft);
            gfx2.DrawString(eighteen, font8, XBrushes.Black, new XRect(x + 20, y + 270, page2.Width, page2.Height), XStringFormats.TopLeft);
            

            tf2.DrawString(nineteen, font1, XBrushes.Black, new XRect(x + 356, y+ 300, 500, 80), XStringFormats.TopLeft);

            tf2.DrawString(seven, font1, XBrushes.Black, new XRect(x + 356, y+ 315, 500, 80), XStringFormats.TopLeft);


            // Save the document...
            string filename = "NIT_WHEAT_" + Convert.ToDateTime(nit.Tables[0].Rows[0].ItemArray[1].ToString()).ToShortDateString() + ".pdf";
            document.Save(filename);

            // ...and start a viewer.
            Process.Start(filename);
        }
        private void NIT_Creation_R(DataSet nit)
        {

            //all the static items in the page hold in variables

            string refno = "Ref. No.:- " + "Com.32/OMSS(D)/BULK/RRA/2015-16";
            string date = "Dated:" + Convert.ToDateTime(nit.Tables[0].Rows[0].ItemArray[1].ToString()).ToShortDateString();
            string notice = "T E N D E R   N O T I C E";
            string one_one = "Financial bid is hereby invited for sale of Raw Rice Grade A  from the FCI empanelled bulk ";
            string one_one_a = "consumers/traders  through";
            string one_one_b = " E-AUCTION";
            string one_one_c = "  only.  The  identified  depots vis-à-vis  the  quantity ";
            string one_one_d = "tendered is as under :";
            string two_one_a = "The  minimum  quantity of  Raw Rice Grade A  for which a bid can be  submitted  would  be  ";
            string two_one_b = "50 MT and maximum would be 3500 MT.Further a Market fee of 0.5% is payable on the Reserve Price.";
            string three = "Tender schedule for the above quantity is as under :- 	 ";
            string four_a = "The relevant MTF is available on www.fciweb.nic.in, www.ncdexspot.com. ";
            string four_b = "E-Auction would be held on 04.06.2015 on www.ncdexspot.com";
            string four = "The EMD/COST is to be deposited in the following Account Exclusively for OMSS (D) :";
            string five = "Name of Beneficiary : FOOD CORPORATION OF INDIA";
            string five_a = "Name of Branch : SBI COMMERCIAL BRANCH, 24 PARK STREET, KOL-16";
            string five_b = "Account No : 33684806347";
            string five_c = "IFS Code : SBIN0007502";
            string six = "Sd/-";
            string seven = "For General Manager (WB)";
            string eight = "NOT TO BE PUBLISHED/DISTRIBUTION :";
            string nine = "1. The Director, DDP&S";
            string nine_a = "Deptt. Of Food & Supplies, Govt. of West Bengal";
            string nine_b = "11A, Mirza Ghalib Street, Kolkata – 700 087";
            string ten = "2. The Director of Rationing  ";
            string ten_a = "Deptt. Of Food & Supplies, Govt. of West Bengal";
            string ten_b = "11A, Mirza Ghalib Street, Kolkata – 700 087";
            string eleven = "3.	All Area Managers";
            string eleven_a = "FCI –with the request to issue stocks as per FIFO. ";
            string twelve = "4. The ED(Zone), FCI, ZO (E), Kolkata ";
            string thirteen = "5. The General Manager (Sales), FCI, Hqrs., New Delhi";
            string fourteen_a = "6.	The Calcutta Flour Mills Association,";
            string fourteen_b = "15, Brabourne Road, Kolkata-700001…with request to circulate among the members of your association";
            string fifteen = "7. The DGM(F&A), FCI, R.O. Kolkata";
            string sixteen = "8. The AGM (Movt/QC/Stg/A/cs)/M (Cash), FCI, R.O. Kolkata… for information and necessary action ";
            string seventeen = "9. Manager (Computer), FCI, RO, Kol…with request to upload the NIT and MTF ";
            string seventeen_a = "in FCI Website www.fciweb.nic.in and CPPP of GOI www.eprocure.gov.in by" + Convert.ToDateTime(nit.Tables[0].Rows[0].ItemArray[1].ToString()).ToShortDateString();
            string eighteen = "10. Sri Manish Tripathi/ Sri Pallav Bhatt, NCDEX Spot Exchange Ltd.";
            string nineteen = "Manager (Comml)";





            // Create a new PDF document
            PdfDocument document = new PdfDocument();
            document.Info.Title = "Created with PDFsharp";

            // Create an empty page
            PdfPage page = document.AddPage();

            // Get an XGraphics object for drawing
            XGraphics gfx = XGraphics.FromPdfPage(page);
            XTextFormatter tf = new XTextFormatter(gfx);


            // Draw FCI logo of the page
            string jpegSamplePath = "../IMAGES/logo.jpg";
            XImage image = XImage.FromFile(jpegSamplePath);
            gfx.DrawImage(image, 44, 70, 500, 80);

            // Create a font

            XFont font1 = new XFont("Calibri (Body)", 12, XFontStyle.Regular);
            XFont font2 = new XFont("Calibri (Body)", 12, XFontStyle.Underline);
            XFont font3 = new XFont("Calibri (Body)", 12, XFontStyle.Italic);
            XFont font4 = new XFont("Calibri (Body)", 12, XFontStyle.Bold);
            XFont font5 = new XFont("Calibri (Body)", 10, XFontStyle.Regular);
            XFont font6 = new XFont("Calibri (Body)", 11, XFontStyle.Bold);
            XFont font7 = new XFont("Calibri (Body)", 09, XFontStyle.Bold);
            XFont font8 = new XFont("Calibri (Body)", 09, XFontStyle.Regular);

            int x = 44;
            int y = 30;

            // Draw the text

            // file refno
            gfx.DrawString(refno, font1, XBrushes.Black, new XRect(x, y + 130, page.Width, page.Height), XStringFormats.TopLeft);
            // file date == date of nit
            gfx.DrawString(date, font1, XBrushes.Black, new XRect(x + 390, y + 130, page.Width, page.Height), XStringFormats.TopLeft);
            //notice line-centre alignment
            gfx.DrawString(notice, font4, XBrushes.Black, new XRect(x - 44, y + 155, page.Width, page.Height), XStringFormats.TopCenter);


            // para1            
            gfx.DrawString(one_one, font1, XBrushes.Black, new XRect(x + 20, y + 170, page.Width, page.Height), XStringFormats.TopLeft);
            gfx.DrawString(one_one_a, font1, XBrushes.Black, new XRect(x, y + 183, page.Width, page.Height), XStringFormats.TopLeft);
            gfx.DrawString(one_one_b, font4, XBrushes.Black, new XRect(x + 150, y + 183, page.Width, page.Height), XStringFormats.TopLeft);
            gfx.DrawString(one_one_c, font1, XBrushes.Black, new XRect(x + 220, y + 183, page.Width, page.Height), XStringFormats.TopLeft);
            gfx.DrawString(one_one_d, font1, XBrushes.Black, new XRect(x, y + 196, page.Width, page.Height), XStringFormats.TopLeft);



            // drawing table1           

            // values to be inserted in the table from db
            int x1 = x + 20;
            int y1 = y + 30;

            //drawing header of table1
            string column1 = "Sl. No.";
            string column2a = "Depot";
            string column3 = "District Office";
            string column4a = "Qty put to tender";
            string column4b = "(in MT)";
            string column5 = "Depot-Wise ";
            string column5a = "Applicable RESERVE ";
            string column5b = "PRICE(Rs./Qtl)";



            //SL. NO.
            gfx.DrawRectangle(XPens.Black, x1, y1 + 190, 45, 32);
            gfx.DrawString(column1, font6, XBrushes.Black, new XRect(x1 + 5, y1 + 200, page.Width.Point, page.Height.Point), XStringFormats.TopLeft);

            //Depot
            gfx.DrawRectangle(XPens.Black, x1 + 44, y1 + 190, 105, 32);
            gfx.DrawString(column2a, font6, XBrushes.Black, new XRect(x1 + 80, y1 + 200, page.Width.Point, page.Height.Point), XStringFormats.TopLeft);


            //District
            gfx.DrawRectangle(XPens.Black, x1 + 150, y1 + 190, 105, 32);
            gfx.DrawString(column3, font6, XBrushes.Black, new XRect(x1 + 170, y1 + 200, page.Width.Point, page.Height.Point), XStringFormats.TopLeft);


            // Qty
            gfx.DrawRectangle(XPens.Black, x1 + 255, y1 + 190, 105, 32);
            gfx.DrawString(column4a, font6, XBrushes.Black, new XRect(x1 + 260, y1 + 190, page.Width.Point, page.Height.Point), XStringFormats.TopLeft);
            gfx.DrawString(column4b, font6, XBrushes.Black, new XRect(x1 + 280, y1 + 205, page.Width.Point, page.Height.Point), XStringFormats.TopLeft);


            // Rate
            gfx.DrawRectangle(XPens.Black, x1 + 360, y1 + 190, 115, 32);
            gfx.DrawString(column5, font6, XBrushes.Black, new XRect(x1 + 385, y1 + 190, page.Width.Point, page.Height.Point), XStringFormats.TopLeft);
            gfx.DrawString(column5a, font6, XBrushes.Black, new XRect(x1 + 365, y1 + 200, page.Width.Point, page.Height.Point), XStringFormats.TopLeft);
            gfx.DrawString(column5b, font6, XBrushes.Black, new XRect(x1 + 385, y1 + 210, page.Width.Point, page.Height.Point), XStringFormats.TopLeft);


            for (int i = 0; i < nit.Tables[0].Rows.Count; i++)
            {
                //drawing body of table

                string sl = (i + 1).ToString();
                string depot = nit.Tables[0].Rows[i].ItemArray[2].ToString();
                string dis = nit.Tables[0].Rows[i].ItemArray[3].ToString();
                string qty = nit.Tables[0].Rows[i].ItemArray[5].ToString();
                string rate = nit.Tables[0].Rows[i].ItemArray[6].ToString();

                //SL.NO.
                gfx.DrawRectangle(XPens.Black, x1, y1 + 222, 45, 20);
                gfx.DrawString(sl, font5, XBrushes.Black, new XRect(x1 + 20, y1 + 222, page.Width.Point, page.Height.Point), XStringFormats.TopLeft);

                //Depot
                gfx.DrawRectangle(XPens.Black, x1 + 44, y1 + 222, 105, 20);
                gfx.DrawString(depot, font5, XBrushes.Black, new XRect(x1 + 47, y1 + 222, page.Width.Point, page.Height.Point), XStringFormats.TopLeft);

                //District
                gfx.DrawRectangle(XPens.Black, x1 + 150, y1 + 222, 105, 20);
                gfx.DrawString(dis, font5, XBrushes.Black, new XRect(x1 + 153, y1 + 222, page.Width.Point, page.Height.Point), XStringFormats.TopLeft);

                //QTY
                gfx.DrawRectangle(XPens.Black, x1 + 255, y1 + 222, 105, 20);
                gfx.DrawString(qty, font5, XBrushes.Black, new XRect(x1 + 320, y1 + 222, page.Width.Point, page.Height.Point), XStringFormats.TopLeft);

                //rate
                gfx.DrawRectangle(XPens.Black, x1 + 360, y1 + 222, 115, 20);
                gfx.DrawString(rate, font5, XBrushes.Black, new XRect(x1 + 400, y1 + 222, page.Width.Point, page.Height.Point), XStringFormats.TopLeft);

                y1 = y1 + 20;

            }

            int y1_1 = (y + 30) + 222 + (20 * (nit.Tables[0].Rows.Count));
            gfx.DrawRectangle(XPens.Black, x1, y1_1, 45, 20);
            gfx.DrawString("Total", font6, XBrushes.Black, new XRect(x1 + 47, y1_1, page.Width.Point, page.Height.Point), XStringFormats.TopLeft);
            gfx.DrawRectangle(XPens.Black, x1 + 44, y1_1, 105, 20);
            gfx.DrawRectangle(XPens.Black, x1 + 150, y1_1, 105, 20);
            gfx.DrawRectangle(XPens.Black, x1 + 255, y1_1, 105, 20);
            gfx.DrawString("5,000", font6, XBrushes.Black, new XRect(x1 + 320, y1_1, page.Width.Point, page.Height.Point), XStringFormats.TopLeft);
            gfx.DrawRectangle(XPens.Black, x1 + 360, y1_1, 115, 20);



            int y2 = y1_1 + 20;


            //int y2 = y + 352;

            //para2

            gfx.DrawString(two_one_a, font1, XBrushes.Black, new XRect(x + 20, y2 + 10 + 10, page.Width, page.Height), XStringFormats.TopLeft);
            //gfx.DrawString(two_one_b, font1, XBrushes.Black, new XRect(x, y2+23, page.Width, page.Height), XStringFormats.TopLeft);
            tf.Alignment = XParagraphAlignment.Justify;
            tf.DrawString(two_one_b, font1, XBrushes.Black, new XRect(x, y2 + 23 + 10, 500, 80), XStringFormats.TopLeft);

            //para3

            gfx.DrawString(three, font1, XBrushes.Black, new XRect(x + 20, y2 + 65 + 10, page.Width, page.Height), XStringFormats.TopLeft);


            //table2

            //drawing header of table2
            string column2_1 = "Date of commencement ";
            string column2__1a = "for Depositing EMD in ";
            string column2__1b = "the below mentioned";
            string column2_1c = "A/C";
            string column2_2 = "Last Date for";
            string column2_2a = "depositing EMD";
            string column2_2b = "through E-payment";
            string column2_3a = "Starting date and ";
            string column2_3b = "time for online Bidding ";
            string column2_4a = "End Date and time  ";
            string column2_4b = "for online bidding";

            int y3 = y2 + 20;
            //dt1
            gfx.DrawRectangle(XPens.Black, x1, y3 + 78, 125, 45);
            gfx.DrawString(column2_1, font6, XBrushes.Black, new XRect(x1 + 5, y3 + 78, page.Width.Point, page.Height.Point), XStringFormats.TopLeft);
            gfx.DrawString(column2__1a, font6, XBrushes.Black, new XRect(x1 + 5, y3 + 88, page.Width.Point, page.Height.Point), XStringFormats.TopLeft);
            gfx.DrawString(column2__1b, font6, XBrushes.Black, new XRect(x1 + 5, y3 + 98, page.Width.Point, page.Height.Point), XStringFormats.TopLeft);
            gfx.DrawString(column2_1c, font6, XBrushes.Black, new XRect(x1 + 5, y3 + 108, page.Width.Point, page.Height.Point), XStringFormats.TopLeft);

            //Dt2
            gfx.DrawRectangle(XPens.Black, x1 + 125, y3 + 78, 105, 45);
            gfx.DrawString(column2_2, font6, XBrushes.Black, new XRect(x1 + 135, y3 + 78, page.Width.Point, page.Height.Point), XStringFormats.TopLeft);
            gfx.DrawString(column2_2a, font6, XBrushes.Black, new XRect(x1 + 135, y3 + 88, page.Width.Point, page.Height.Point), XStringFormats.TopLeft);
            gfx.DrawString(column2_2b, font6, XBrushes.Black, new XRect(x1 + 135, y3 + 98, page.Width.Point, page.Height.Point), XStringFormats.TopLeft);


            //Dt3
            gfx.DrawRectangle(XPens.Black, x1 + 230, y3 + 78, 145, 45);
            gfx.DrawString(column2_3a, font6, XBrushes.Black, new XRect(x1 + 240, y3 + 78, page.Width.Point, page.Height.Point), XStringFormats.TopLeft);
            gfx.DrawString(column2_3b, font6, XBrushes.Black, new XRect(x1 + 240, y3 + 88, page.Width.Point, page.Height.Point), XStringFormats.TopLeft);


            // dt4
            gfx.DrawRectangle(XPens.Black, x1 + 375, y3 + 78, 135, 45);
            gfx.DrawString(column2_4a, font6, XBrushes.Black, new XRect(x1 + 380, y3 + 78, page.Width.Point, page.Height.Point), XStringFormats.TopLeft);
            gfx.DrawString(column2_4b, font6, XBrushes.Black, new XRect(x1 + 380, y3 + 88, page.Width.Point, page.Height.Point), XStringFormats.TopLeft);




            //body of table2

            string column2_b_1 = Convert.ToDateTime(nit.Tables[0].Rows[0].ItemArray[1].ToString()).ToShortDateString();
            string column2__b_2 = Convert.ToDateTime(column2_b_1).AddDays(2).ToShortDateString();
            string column2__b_3 = Convert.ToDateTime(column2_b_1).AddDays(3).ToShortDateString() + "from 11:00 a.m.";
            string column2__b_4 = Convert.ToDateTime(column2_b_1).AddDays(3).ToShortDateString() + "upto 02:00 pm";

            //dt1
            gfx.DrawRectangle(XPens.Black, x1, y3 + 123, 125, 20);
            gfx.DrawString(column2_b_1, font6, XBrushes.Black, new XRect(x1 + 5, y3 + 127, page.Width.Point, page.Height.Point), XStringFormats.TopLeft);

            //Dt2
            gfx.DrawRectangle(XPens.Black, x1 + 125, y3 + 123, 105, 20);
            gfx.DrawString(column2__b_2, font6, XBrushes.Black, new XRect(x1 + 135, y3 + 127, page.Width.Point, page.Height.Point), XStringFormats.TopLeft);


            //Dt3
            gfx.DrawRectangle(XPens.Black, x1 + 230, y3 + 123, 145, 20);
            gfx.DrawString(column2__b_3, font6, XBrushes.Black, new XRect(x1 + 240, y3 + 127, page.Width.Point, page.Height.Point), XStringFormats.TopLeft);


            // dt4
            gfx.DrawRectangle(XPens.Black, x1 + 375, y3 + 123, 135, 20);
            gfx.DrawString(column2__b_4, font6, XBrushes.Black, new XRect(x1 + 380, y3 + 127, page.Width.Point, page.Height.Point), XStringFormats.TopLeft);

            //para before 4
            gfx.DrawString(four_a, font1, XBrushes.Black, new XRect(x, y3 + 140 + 10, page.Width, page.Height), XStringFormats.TopLeft);
            gfx.DrawString(four_b, font1, XBrushes.Black, new XRect(x, y3 + 153 + 10, page.Width, page.Height), XStringFormats.TopLeft);
           

            //para4
            gfx.DrawString(four, font4, XBrushes.Black, new XRect(x, y3 + 153 + +20+10+10, page.Width, page.Height), XStringFormats.TopLeft);
            gfx.DrawString(five, font1, XBrushes.Black, new XRect(x, y3 + 166 + +20+10+10, page.Width, page.Height), XStringFormats.TopLeft);
            gfx.DrawString(five_a, font1, XBrushes.Black, new XRect(x, y3 + 179 + 20+10+10, page.Width, page.Height), XStringFormats.TopLeft);
            gfx.DrawString(five_b, font1, XBrushes.Black, new XRect(x, y3 + 192 + 20+10+10, page.Width, page.Height), XStringFormats.TopLeft);
            gfx.DrawString(five_c, font1, XBrushes.Black, new XRect(x, y3 + 205 +20+ 10+10, page.Width, page.Height), XStringFormats.TopLeft);

            //para5/6
            tf.DrawString(six, font1, XBrushes.Black, new XRect(x + 480, y3 + 222+20, 500, 80), XStringFormats.TopLeft);
            tf.DrawString(seven, font1, XBrushes.Black, new XRect(x + 356, y3 + 237+20, 500, 80), XStringFormats.TopLeft);

            // Create an empty page
            PdfPage page2 = document.AddPage();

            // Get an XGraphics object for drawing
            XGraphics gfx2 = XGraphics.FromPdfPage(page2);
            XTextFormatter tf2 = new XTextFormatter(gfx2);
            //para7/8

            gfx2.DrawString(eight, font7, XBrushes.Black, new XRect(x, y, page2.Width, page2.Height), XStringFormats.TopLeft);


            //copies
            gfx2.DrawString(nine, font8, XBrushes.Black, new XRect(x + 20, y + 20, page2.Width, page2.Height), XStringFormats.TopLeft);
            gfx2.DrawString(nine_a, font8, XBrushes.Black, new XRect(x + 30, y + 30, page2.Width, page2.Height), XStringFormats.TopLeft);
            gfx2.DrawString(nine_b, font8, XBrushes.Black, new XRect(x + 30, y + 40, page2.Width, page2.Height), XStringFormats.TopLeft);
            gfx2.DrawString(ten, font8, XBrushes.Black, new XRect(x + 20, y + 60, page2.Width, page2.Height), XStringFormats.TopLeft);
            gfx2.DrawString(ten_a, font8, XBrushes.Black, new XRect(x + 30, y + 70, page2.Width, page2.Height), XStringFormats.TopLeft);
            gfx2.DrawString(ten_b, font8, XBrushes.Black, new XRect(x + 30, y + 80, page2.Width, page2.Height), XStringFormats.TopLeft);
            gfx2.DrawString(eleven, font8, XBrushes.Black, new XRect(x + 20, y + 100, page2.Width, page2.Height), XStringFormats.TopLeft);
            gfx2.DrawString(eleven_a, font8, XBrushes.Black, new XRect(x + 30, y + 110, page2.Width, page2.Height), XStringFormats.TopLeft);
            gfx2.DrawString(twelve, font8, XBrushes.Black, new XRect(x + 20, y + 130, page2.Width, page2.Height), XStringFormats.TopLeft);
            gfx2.DrawString(thirteen, font8, XBrushes.Black, new XRect(x + 20, y + 150, page2.Width, page2.Height), XStringFormats.TopLeft);
            gfx2.DrawString(fourteen_a, font8, XBrushes.Black, new XRect(x + 20, y + 170, page2.Width, page2.Height), XStringFormats.TopLeft);
            gfx2.DrawString(fourteen_b, font8, XBrushes.Black, new XRect(x + 30, y + 180, page2.Width, page2.Height), XStringFormats.TopLeft);
            gfx2.DrawString(fifteen, font8, XBrushes.Black, new XRect(x + 20, y + 200, page2.Width, page2.Height), XStringFormats.TopLeft);
            gfx2.DrawString(sixteen, font8, XBrushes.Black, new XRect(x + 20, y + 220, page2.Width, page2.Height), XStringFormats.TopLeft);
            gfx2.DrawString(seventeen, font8, XBrushes.Black, new XRect(x + 20, y + 240, page2.Width, page2.Height), XStringFormats.TopLeft);
            gfx2.DrawString(seventeen_a, font8, XBrushes.Black, new XRect(x + 30, y + 250, page2.Width, page2.Height), XStringFormats.TopLeft);
            gfx2.DrawString(eighteen, font8, XBrushes.Black, new XRect(x + 20, y + 270, page2.Width, page2.Height), XStringFormats.TopLeft);


            tf2.DrawString(nineteen, font1, XBrushes.Black, new XRect(x + 356, y + 300, 500, 80), XStringFormats.TopLeft);

            tf2.DrawString(seven, font1, XBrushes.Black, new XRect(x + 356, y + 315, 500, 80), XStringFormats.TopLeft);


            // Save the document...
            string filename = "NIT_WHEAT_" + Convert.ToDateTime(nit.Tables[0].Rows[0].ItemArray[1].ToString()).ToShortDateString() + ".pdf";
            document.Save(filename);

            // ...and start a viewer.
            Process.Start(filename);
        }
        
        #endregion
        #region MAKE NOTESHEET
        private void cmbfinyr_SelectedIndexChanged(object sender, EventArgs e)
        {
            rtbns.Visible = true;
            btncopyns.Visible = true;
            btngnrtNIT.Visible = true;
            btnmake.Visible = true;
            
        }
        private void btnmake_Click(object sender, EventArgs e)

        {
            // string note, noofprevioustender, noofpresenttender, prevnitdate, preveadate, currentnitdt, currenteadt, nitdate_mnth, nitdate_year;
            nitdate1 = Convert.ToDateTime(comboBox1.SelectedValue).ToLongDateString();
            string prevnitdate1 = dttmnsprevnit.Value.ToLongDateString();
            DataSet pub1 = new DataSet();
            DataSet pub1a = new DataSet();
            DataSet pub2 = new DataSet();
            DataSet pub2a = new DataSet();
            DataSet pub3 = new DataSet();

            DataSet Qty = new DataSet();
            DataSet Qty1 = new DataSet();
            DataSet result = new DataSet();
            DataSet rprts = new DataSet();
            


            //DATES

            string nitdate = Convert.ToDateTime(comboBox1.SelectedValue).ToLongDateString();
            nitdate_mnth = Convert.ToDateTime(nitdate).ToString("MMMM");// GETTING MONTH NAME
            nitdate_year = Convert.ToDateTime(nitdate).Year.ToString();
            string qry1 = "Select * from Publication order by NITDate";
            string qry1a = "Select NITDate,EMDDate,EAuctionDate from Publication where  MonthName='" + nitdate_mnth + "' order by NITDate";
            string qry2 = "Select * from Publication where  NITDate= #" + nitdate1 + "#";
            string qry2a = "Select NITDate,EMDDate,EAuctionDate from Publication where  NITDate= #" + nitdate1 + "#";
            string qry3 = "SELECT * FROM Publication WHERE (((Publication.NITDate)<# " + nitdate1 + " #)) order by NITDate";
            pub1 = select_data(qry1);
            pub1a = select_data(qry1a);
            pub2 = select_data(qry2);
            pub2a = select_data(qry2a);
            pub3 = select_data(qry3);

            if (pub3.Tables[0].Rows.Count > 0)
            {
                noofprevioustender = pub3.Tables[0].Rows.Count.ToString();
                noofpresenttender = (pub3.Tables[0].Rows.Count + 1).ToString();
                prevnitdate = Convert.ToDateTime(pub3.Tables[0].Rows[(pub3.Tables[0].Rows.Count - 1)].ItemArray[2]).ToShortDateString();
                preveadate = pub3.Tables[0].Rows[(pub3.Tables[0].Rows.Count - 1)].ItemArray[4].ToString();
            }
            else
            {
                MessageBox.Show("No Data");
            }
            if (pub2.Tables[0].Rows.Count > 0)
            {
                currentnitdt = Convert.ToDateTime(pub2.Tables[0].Rows[0].ItemArray[2]).ToShortDateString();
                currenteadt = Convert.ToDateTime(pub2.Tables[0].Rows[0].ItemArray[4]).ToShortDateString();
                pubdate = Convert.ToDateTime(pub2.Tables[0].Rows[0].ItemArray[1]).ToShortDateString();
                
            }

            else
            {
                MessageBox.Show("No Data");
            }

            // get finanacial year
            finyr = cmbfinyr.SelectedItem.ToString();
            // MessageBox.Show(nitdate_mnth + nitdate_year);

            //Report Dates, QUANTITY put, result qty al vary as per wheat or rice
            if (cmbcmdty.SelectedItem == "Wheat")
            {

                //report dates
                string qry4 = "Select ReportName,ReportDate from Reports_Wheat where  NITDate= #" + nitdate1 + "#";
                rprts = select_data(qry4);
                if (rprts.Tables[0].Rows.Count > 0)
                {
                    rprtrks = rprts.Tables[0].Rows[0].ItemArray[1].ToString();
                    rprtot = rprts.Tables[0].Rows[1].ItemArray[1].ToString();
                    rprurs = rprts.Tables[0].Rows[2].ItemArray[1].ToString();
                    rpronly = rprts.Tables[0].Rows[3].ItemArray[1].ToString();
                   // MessageBox.Show(rprtrks + rprtot + rprurs + rpronly);
                }
                else
                {
                    MessageBox.Show("No Data");
                }
                //QUANTITY PUT
                string qry5 = "Select Depot_Selected,District_Selected,Wheat_URS_stock,QtyProposed from ApprovalQty_Wheat where  NITDate= #" + nitdate + "#";
                Qty = select_data(qry5);
                string qry6 = "Select sum(QtyProposed) from ApprovalQty_Wheat where  NITDate= #" + nitdate + "#";
                Qty1 = select_data(qry6);
                if (Qty.Tables[0].Rows.Count > 0)
                {
                   /*for(int i=0;i<Qty.Tables[0].Rows.Count;i++)
                    {
                        dep = Qty.Tables[0].Rows[i].ItemArray[0].ToString();
                        dis = Qty.Tables[0].Rows[i].ItemArray[1].ToString();
                        stock = Qty.Tables[0].Rows[i].ItemArray[2].ToString();
                        qtyprpsd = Qty.Tables[0].Rows[i].ItemArray[3].ToString();
                    }*/
                   if (Qty1.Tables[0].Rows.Count > 0)
                    totqtyput = Qty1.Tables[0].Rows[0].ItemArray[0].ToString();
                   else
                   {
                       MessageBox.Show("No Data");
                   }
                   // MessageBox.Show(totqtyput);
                }
                else
                {
                    MessageBox.Show("No Data");
                }

                //RESULT
                string qry7 = "Select sum(Qty) from Result where  NITDate= #" + prevnitdate1 + "# and Commodity='" + cmbcmdty.SelectedItem.ToString()+"'";
                result= select_data(qry7);           
               
                    if (result.Tables[0].Rows.Count > 0)
                    { 
                        totresut = result.Tables[0].Rows[0].ItemArray[0].ToString();
                        if(totresut=="")
                        {
                            totresut = "0";

                        }
                      
                    }
                        
                    else
                    {
                        totresut = "0";
                    }
                 
              // MessageBox.Show(totresut);
                


            }

            else if (cmbcmdty.SelectedItem == "Raw Rice Grade A")
            {
                // report dates
                string qry4 = "Select ReportName,ReportDate from Reports_Rice where  NITDate= #" + nitdate1 + "#";
                rprts = select_data(qry4);
                if (rprts.Tables[0].Rows.Count > 0)
                {
                    rpremail = rprts.Tables[0].Rows[0].ItemArray[1].ToString();
                    rprtot = rprts.Tables[0].Rows[1].ItemArray[1].ToString();
                   // MessageBox.Show(rpremail + rprtot);
                }
                else
                {
                    MessageBox.Show("No Reports of Rice in DB");
                }

                //QUANTITY PUT
                string qry5 = "Select Depot_Selected,District_Selected,RRA_stock,QtyProposed from ApprovalQty_Rice where  NITDate= #" + nitdate + "#";
                Qty = select_data(qry5);
                string qry6 = "Select sum(QtyProposed) from ApprovalQty_Rice where  NITDate= #" + nitdate + "#";
                Qty1 = select_data(qry6);
               // if (Qty.Tables[0].Rows.Count > 0)
                //{
                   /* for (int i = 0; i < Qty.Tables[0].Rows.Count; i++)
                    {
                        dep = Qty.Tables[0].Rows[i].ItemArray[0].ToString();
                        dis = Qty.Tables[0].Rows[i].ItemArray[1].ToString();
                        stock = Qty.Tables[0].Rows[i].ItemArray[2].ToString();
                        qtyprpsd = Qty.Tables[0].Rows[i].ItemArray[3].ToString();
                    }*/
                    if (Qty1.Tables[0].Rows.Count > 0)
                        totqtyput = Qty1.Tables[0].Rows[0].ItemArray[0].ToString();
                    else
                    {
                        MessageBox.Show("No ");
                    }
               // }
               // else
               // {
                //    MessageBox.Show("No Data");
                //}


                //RESULT
                    string qry7 = "Select sum(Qty) from Result where  NITDate= #" + prevnitdate1 + "# and Commodity='" + cmbcmdty.SelectedItem.ToString() + "'";
                    result = select_data(qry7);

                if (result.Tables[0].Rows.Count > 0)
                    totresut = result.Tables[0].Rows[0].ItemArray[0].ToString();
                if (totresut == "")
                {
                    totresut = "0";

                }
                else
                {
                    totresut = "0";
                }
             //   MessageBox.Show(totresut);
            }




//ENTER DATA TO NOTESHEET

            if (cmbauthor.SelectedItem == "Dealing Assistant")
            {
                if (cmb1stNIT.SelectedItem == "YES")
                {
                    if (cmbcmdty.SelectedItem == "Wheat")
                    {
                        
                        note = "\r  The" + noofprevioustender + "th" + " " + "tender for the year" + " " + finyr + " " + "  for Sale of " + " " + totqtyput + "MT" +"wheat was floated on " + " " + prevnitdate + " " + ", e-Auctions for which were held on " + preveadate + ".";
                        // if result 0 or not
                        if(totresut=="0")
                            note = note + "No Bids were received in the mentioned e-Auction.";
                        else
                            note = note + totresut + "MT" + " " +"were sold in the mentioned e-Auction.";
                        
                        note = note + "\r\r  The"+" " + noofpresenttender+ " th" + " " +"tender for the year"+ " " + finyr + " " +"is to be floated on" +" " + currentnitdt +" " +", e-Auction for which is to be conducted on "+" " + currenteadt +" " +"." ;
                        note=note+"\rThe depot-wise stock position of TOTAL WHEAT(including URS) as available on "+" " +rprtot+" " +" (CP.....),the depot-wise stock position of WHEAT only as well as WHEAT URS generated as separate reports on" + " "+ rprurs+" "+" (CP... and ....), so as to dispose of Wheat URS stock in the current season on a priority basis as per Hqrs instructions in letter no. J1(1)/2015/OMSS(D)/B/C/S.III dated: 29.05.2015 and 09.06.2015 (CP 90-96), Movement plan indicating rakes on run (CP.....) are placed for your ready reference. ";
                        note = note + "\rAccordingly, the depot-wise quantity may kindly be proposed and decided for NIT dated" + " " + currentnitdt + "as deemed fit.";
                        note = note + "\r\r  Further, a common Newspaper Advertisement for tenders for Sale of both Raw Rice Grade A and Wheat during the month of" +" " + nitdate_mnth +"'" + nitdate_year +" " +"may be published together in order to save on cost component.";
                        note = note + "\r\r  Schedule of Tenders for both Raw Rice Grade A and Wheat during the month of" + " " + nitdate_mnth + "'" + nitdate_year + " " + "is as under:";
                        note += "\r\r\r";
                        note += "Date of commencement" + "  Last Date for depositing " + "Date of E-Auction ";
                        note += "\rfor Deposition of EMD" + " EMD through E-payment ";
                        note +="\r\r";
                        dynamic ColCount = pub1a.Tables[0].Columns.Count;
                        if (pub1a.Tables[0].Rows.Count > 0)
                        {
                            
                            for (int j=0;j< pub1a.Tables[0].Rows.Count;j++)
                            {
                                for (int i =0 ; i <= ColCount - 1; i++)
                                {
                                    note = note +Convert.ToDateTime( pub1a.Tables[0].Rows[j].ItemArray[i].ToString()).ToShortDateString() + "                        ";
                                    if (i == ColCount - 1)
                                    {
                                        note += "\r\r";
                                    }
                                }
                            }
                        }
                        
                        note = note + "\r\r  Draft of newspaper publication placed alongside." + "\r\r अवलोकन तथा आगे की निर्णय हेतु प्रस्तुत है। ";
                        note = note + "\r\r\r  सहायक श्रेणी-III ";
                        note = note + "\r\r  प्रबंधक (वाणिज्य)";


                        rtbns.SelectAll();
                        rtbns.SelectionAlignment = HorizontalAlignment.Left;
                        rtbns.Text = note;
                    }
                    else if (cmbcmdty.SelectedItem == "Raw Rice Grade A")
                    {

                        note = "\r  The" + noofprevioustender + "th" + " " + "tender for the year" + " " + finyr + " " + "  for Sale of Raw Rice Grade A  was floated on " + " " + prevnitdate + " " + ", e-Auctions for which were held on " + preveadate + " "+"for the quantity of" + " "+ totqtyput + "MT" + "from CSD Dabgram under FCI, D.O. Siliguri.";
                        // if result 0 or not
                        if (totresut == "0")
                            note = note + "No Bids were received in the mentioned e-Auction.";
                        else
                            note = note + totresut + "MT" + " " + "were sold in the mentioned e-Auction.";

                        note = note + "\r\r  Accordingly,the" + " " + noofpresenttender + " th" + " " + "tender for the year" + " " + finyr + " " + "is to be floated on" + " " + currentnitdt + " " + ", e-Auction for which is to be conducted on " + " " + currenteadt + " " + ".";
                        note = note + "\rWe may float tender from CSD Dabgram only, which at present is having"+" "+stock +" MT"+"  "+"("+rpremail+", OB) of Raw Rice Grade A, as confirmed by FCI D.O. Siliguri vide e-mail dtd."+" "+ rpremail+" "+" (CP .....).";
                        note = note + "\rOverall depot-wise, commodity-wise stock positionas on" + " " + rprtot + " " + "is placed at CP .... for your ready reference.";
                        note = note + "\rAs such, the quantity for Sale of Raw Rice Grade A under OMSS (D) through e-Auction to be held on " + " " + currenteadt + "may kindly be proposed and decided as deemed fit.";
                        note = note + "\r\r  Further, a common Newspaper Advertisement for tenders for Sale of both Raw Rice Grade A and Wheat during the month of" + " " + nitdate_mnth + "'" + nitdate_year + " " + "may be published together in order to save on cost component.";
                        note = note + "\r\r  Schedule of Tenders for both Raw Rice Grade A and Wheat during the month of" + " " + nitdate_mnth + "'" + nitdate_year + " " + "is as under:";
                        note += "\r\r\r";
                        note += "Date of commencement" + "  Last Date for depositing " + "Date of E-Auction ";
                        note += "\rfor Deposition of EMD" + " EMD through E-payment ";
                        note += "\r\r";
                        dynamic ColCount = pub1a.Tables[0].Columns.Count;
                        if (pub1a.Tables[0].Rows.Count > 0)
                        {

                            for (int j = 0; j < pub1a.Tables[0].Rows.Count; j++)
                            {
                                for (int i = 0; i <= ColCount - 1; i++)
                                {
                                    note = note + Convert.ToDateTime(pub1a.Tables[0].Rows[j].ItemArray[i].ToString()).ToShortDateString() + "                        ";
                                    if (i == ColCount - 1)
                                    {
                                        note += "\r\r";
                                    }
                                }
                            }
                        }

                        note = note + "\r\r  Draft of newspaper publication placed alongside." + "\r\r अवलोकन तथा आगे की निर्णय हेतु प्रस्तुत है। ";
                        note = note + "\r\r\r  सहायक श्रेणी-III ";
                        note = note + "\r\r  प्रबंधक (वाणिज्य)";


                        rtbns.SelectAll();
                        rtbns.SelectionAlignment = HorizontalAlignment.Left;
                        rtbns.Text = note;
                    }
                }
                else if (cmb1stNIT.SelectedItem == "NO")
                {
                    if (cmbcmdty.SelectedItem == "Wheat")
                    {
                         note = "\r  The" + noofprevioustender + "th" + " " + "tender for the year" + " " + finyr + " " + "  for Sale of " + " " + totqtyput + "MT" +"wheat was floated on " + " " + prevnitdate + " " + ", e-Auctions for which were held on " + preveadate + ".";
                        // if result 0 or not
                        if(totresut=="0")
                            note = note + "No Bids were received in the mentioned e-Auction.";
                        else
                            note = note + totresut + "MT" + " " +"were sold in the mentioned e-Auction.";
                        
                        note = note + "\r\r  The"+" " + noofpresenttender+ "th" + " "+"tender for the year"+ " " + finyr + " " +"is to be floated on" +" " + currentnitdt +" " +", e-Auction for which is to be conducted on "+" " + currenteadt +" " +"." ;
                        note=note+"\rThe depot-wise stock position of TOTAL WHEAT(including URS) as available on "+" " +rprtot+" " +" (CP.....),the depot-wise stock position of WHEAT only as well as WHEAT URS generated as separate reports on" + " "+ rprurs+" "+" (CP... and ....), so as to dispose of Wheat URS stock in the current season on a priority basis as per Hqrs instructions in letter no. J1(1)/2015/OMSS(D)/B/C/S.III dated: 29.05.2015 and 09.06.2015 (CP 90-96), Movement plan indicating rakes on run (CP.....) are placed for your ready reference. ";
                        note = note + "\rAccordingly, the depot-wise quantity may kindly be proposed and decided for NIT dated" + " " + currentnitdt + "as deemed fit.";
                        note = note + "\r\r  Further, newspaper advertisement for all tenders of Sale of Wheat during the month of" +" " + nitdate_mnth +"'" + nitdate_year +" " +"has already been published on " +" " +pubdate+" " +"(CP .......). ";
                        note = note + "\r\r  Besides, the following is the tender schedule for NIT dated: " + " " + currentnitdt+ " " + ":-";
                        note += "\r\r\r";
                        note += "Date of commencement" + "  Last Date for depositing " + "Date of E-Auction ";
                        note += "\rfor Deposition of EMD" + " EMD through E-payment ";
                        note +="\r\r";
                        dynamic ColCount = pub2a.Tables[0].Columns.Count;
                        if (pub2a.Tables[0].Rows.Count > 0)
                        {

                            for (int j = 0; j < pub2a.Tables[0].Rows.Count; j++)
                            {
                                for (int i =0 ; i <= ColCount - 1; i++)
                                {
                                    note = note + Convert.ToDateTime(pub2a.Tables[0].Rows[j].ItemArray[i].ToString()).ToShortDateString() + "                        ";
                                    if (i == ColCount - 1)
                                    {
                                        note += "\r\r";
                                    }
                                }
                            }
                        }
                        
                        note = note + "\r\r   अवलोकन तथा आगे की निर्णय हेतु प्रस्तुत है। ";
                       note = note + "\r\r\r  सहायक श्रेणी-III ";
                        note = note + "\r\r  प्रबंधक (वाणिज्य)";
                        
                        rtbns.SelectAll();
                        rtbns.SelectionAlignment = HorizontalAlignment.Left;
                        rtbns.Text = note;
                    }
                    }
                    else if (cmbcmdty.SelectedItem == "Raw Rice Grade A")
                    {
                        note = "\r  The" + noofprevioustender + "th" + " " + "tender for the year" + " " + finyr + " " + "  for Sale of Raw Rice Grade A  was floated on " + " " + prevnitdate + " " + ", e-Auctions for which were held on " + preveadate + " " + "for the quantity of" + " " + totqtyput + "MT" + "from CSD Dabgram under FCI, D.O. Siliguri.";
                        // if result 0 or not
                        if (totresut == "0")
                            note = note + "No Bids were received in the mentioned e-Auction.";
                        else
                            note = note + totresut + "MT" + " " + "were sold in the mentioned e-Auction.";

                        note = note + "\r\r  Accordingly,the" + " " + noofpresenttender + " th" + " " + "tender for the year" + " " + finyr + " " + "is to be floated on" + " " + currentnitdt + " " + ", e-Auction for which is to be conducted on " + " " + currenteadt + " " + ".";
                        note = note + "\rWe may float tender from CSD Dabgram only, which at present is having" + " " + stock + " MT" + "  " + "(" + rpremail + ", OB) of Raw Rice Grade A, as confirmed by FCI D.O. Siliguri vide e-mail dtd." + " " + rpremail + " " + " (CP .....).";
                        note = note + "\rOverall depot-wise, commodity-wise stock positionas on" + " " + rprtot + " " + "is placed at CP .... for your ready reference.";
                        note = note + "\rAs such, the quantity for Sale of Raw Rice Grade A under OMSS (D) through e-Auction to be held on " + " " + currenteadt + "may kindly be proposed and decided as deemed fit.";
                        note = note + "\r\r  Further, newspaper advertisement for all tenders of Sale of RawRice during the month of" + " " + nitdate_mnth + "'" + nitdate_year + " " + "has already been published on " + " " + pubdate + " " + "(CP .......). ";
                        note = note + "\r\r  Besides, the following is the tender schedule for NIT dated: " + " " + currentnitdt + " " + ":-";
                        note += "\r\r\r";
                        note += "Date of commencement" + "  Last Date for depositing " + "Date of E-Auction ";
                        note += "\rfor Deposition of EMD" + " EMD through E-payment ";
                        note += "\r\r";
                        dynamic ColCount = pub2a.Tables[0].Columns.Count;
                        if (pub2a.Tables[0].Rows.Count > 0)
                        {

                            for (int j = 0; j < pub2a.Tables[0].Rows.Count; j++)
                            {
                                for (int i = 0; i <= ColCount - 1; i++)
                                {
                                    note = note + Convert.ToDateTime(pub2a.Tables[0].Rows[j].ItemArray[i].ToString()).ToShortDateString() + "                        ";
                                    if (i == ColCount - 1)
                                    {
                                        note += "\r\r";
                                    }
                                }
                            }
                        }

                        note = note + "\r\r   अवलोकन तथा आगे की निर्णय हेतु प्रस्तुत है। ";
                        note = note + "\r\r\r  सहायक श्रेणी-III ";
                        note = note + "\r\r  प्रबंधक (वाणिज्य)";


                        rtbns.SelectAll();
                        rtbns.SelectionAlignment = HorizontalAlignment.Left;
                        rtbns.Text = note;
                    }
                }
            
    
            else if (cmbauthor.SelectedItem == "Manager")
            {
                if (cmb1stNIT.SelectedItem == "YES")
                {
                    if (cmbcmdty.SelectedItem == "Wheat")
                    {
                        note = "\r Following is the proposed depot-wise quantity against NIT dated:"+ currentnitdt;
                        note += "\r\r\r";
                        note += "       Depot " + "         District Office " + "       Stock " + "    Qty proposed ";
                      
                        note += "\r\r";
                        dynamic ColCount = Qty.Tables[0].Columns.Count;
                        if (Qty.Tables[0].Rows.Count > 0)
                        {

                            for (int j = 0; j < Qty.Tables[0].Rows.Count; j++)
                            {
                                for (int i = 0; i <= ColCount - 1; i++)
                                {
                                    note = note + Qty.Tables[0].Rows[j].ItemArray[i].ToString() + "           ";
                                    if (i == ColCount - 1)
                                    {
                                        note += "\r\r";
                                    }
                                }
                            }
                        }
                        note = note + "\r\r  Newspaper Advertisement for all tenders for Sale Wheat during the month of "+ nitdate_mnth + "'"+nitdate_year +" may be published together on"+" " +pubdate+"draft of which is placed alongside.";
                        note = note + "\r\r  The detailed depot wise position for NIT dated "+ " " + currentnitdt+" " +"may be uploaded on the FCI website, www.fciweb.nic.in & ncdexspot.com, draft NIT for which is placed alongside as per the above note.";
                        note += "\r\r";                      

                        note = note + "\r\r अवलोकन तथा आगे की निर्णय हेतु प्रस्तुत है। ";
                        note = note + "\r\r प्रबंधक (वाणिज्य)";
                        note = note + "\r\r सहायक महा प्रबंधक (वाणिज्य)";
                        note = note + "\r\r उप महा पबंधक (क्षेत्र)";


                        rtbns.SelectAll();
                        rtbns.SelectionAlignment = HorizontalAlignment.Left;
                        rtbns.Text = note;
                    }
                    else if (cmbcmdty.SelectedItem == "Raw Rice Grade A")
                    {
                        note = "\r Following is the quantity of “Raw Rice Grade A” proposed for Sale under OMSS(D) through e-Auction to be held on " + currenteadt;
                        note += "\r\r\r";
                        note += "       Depot " + "         District Office " + "       Stock " + "    Qty proposed ";

                        note += "\r\r";
                        dynamic ColCount = Qty.Tables[0].Columns.Count;
                        if (Qty.Tables[0].Rows.Count > 0)
                        {

                            for (int j = 0; j < Qty.Tables[0].Rows.Count; j++)
                            {
                                for (int i = 0; i <= ColCount - 1; i++)
                                {
                                    note = note + Qty.Tables[0].Rows[j].ItemArray[i].ToString() + "           ";
                                    if (i == ColCount - 1)
                                    {
                                        note += "\r\r";
                                    }
                                }
                            }
                        }
                        note = note + "\r\r  Further, in respect of Sikkim, it is stated that there is .........MT Raw Rice in Sikkim that is not sufficient to meet its own TPDS/OWS requirement and hence we may not float any tender for Sikkim.";
                        note = note + "\r\r  Newspaper Advertisement for all tenders for Sale Raw Rice Grade-A  during the month of " + nitdate_mnth + "'" + nitdate_year + " may be published together on" + " " + pubdate + "draft of which is placed alongside.";
                        note = note + "\r\r  The detailed depot wise position for NIT dated " + " " + currentnitdt + " " + "may be uploaded on the FCI website, www.fciweb.nic.in & ncdexspot.com, draft NIT for which is placed alongside as per the above note.";
                        note += "\r\r";

                        note = note + "\r\r अवलोकन तथा आगे की निर्णय हेतु प्रस्तुत है। तथानुसार, दिनांक"+" "+currentnitdt+" "+" के लिये निविदा आमंत्रण सूचना का मसौदा प्रस्तुत है।";
                        note = note + "\r\r प्रबंधक (वाणिज्य)";
                        note = note + "\r\r सहायक महा प्रबंधक (वाणिज्य)";
                        note = note + "\r\r उप महा पबंधक (क्षेत्र)";


                        rtbns.SelectAll();
                        rtbns.SelectionAlignment = HorizontalAlignment.Left;
                        rtbns.Text = note;
                    }
                }
                else if (cmb1stNIT.SelectedItem == "NO")
                {
                    if (cmbcmdty.SelectedItem == "Wheat")
                    {
                        note = "\r Following is the proposed depot-wise quantity against NIT dated:" + currentnitdt;
                        note += "\r\r\r";
                        note += "       Depot " + "         District Office " + "       Stock " + "    Qty proposed ";

                        note += "\r\r";
                        dynamic ColCount = Qty.Tables[0].Columns.Count;
                        if (Qty.Tables[0].Rows.Count > 0)
                        {

                            for (int j = 0; j < Qty.Tables[0].Rows.Count; j++)
                            {
                                for (int i = 0; i <= ColCount - 1; i++)
                                {
                                    note = note + Qty.Tables[0].Rows[j].ItemArray[i].ToString() + "           ";
                                    if (i == ColCount - 1)
                                    {
                                        note += "\r\r";
                                    }
                                }
                            }
                        }
                        note = note + "\r\r  The detailed depot wise position for NIT dated " + " " + currentnitdt + " " + "may be uploaded on the FCI website, www.fciweb.nic.in & ncdexspot.com, draft NIT for which is placed alongside as per the above note.";
                        note += "\r\r";

                        note = note + "\r\r अवलोकन तथा आगे की निर्णय हेतु प्रस्तुत है। ";
                        note = note + "\r\r प्रबंधक (वाणिज्य)";
                        note = note + "\r\r सहायक महा प्रबंधक (वाणिज्य)";
                        note = note + "\r\r उप महा पबंधक (क्षेत्र)";


                        rtbns.SelectAll();
                        rtbns.SelectionAlignment = HorizontalAlignment.Left;
                        rtbns.Text = note;
                    }
                    else if (cmbcmdty.SelectedItem == "Raw Rice Grade A")
                    {
                        note = "\r Following is the quantity of “Raw Rice Grade A” proposed for Sale under OMSS(D) through e-Auction to be held on " + currenteadt;
                        note += "\r\r\r";
                        note += "       Depot " + "         District Office " + "       Stock " + "    Qty proposed ";

                        note += "\r\r";
                        dynamic ColCount = Qty.Tables[0].Columns.Count;
                        if (Qty.Tables[0].Rows.Count > 0)
                        {

                            for (int j = 0; j < Qty.Tables[0].Rows.Count; j++)
                            {
                                for (int i = 0; i <= ColCount - 1; i++)
                                {
                                    note = note + Qty.Tables[0].Rows[j].ItemArray[i].ToString() + "           ";
                                    if (i == ColCount - 1)
                                    {
                                        note += "\r\r";
                                    }
                                }
                            }
                        }
                        note = note + "\r\r  Further, in respect of Sikkim, it is stated that there is .........MT Raw Rice in Sikkim that is not sufficient to meet its own TPDS/OWS requirement and hence we may not float any tender for Sikkim.";
                        note = note + "\r\r  The detailed depot wise position for NIT dated " + " " + currentnitdt + " " + "may be uploaded on the FCI website, www.fciweb.nic.in & ncdexspot.com, draft NIT for which is placed alongside as per the above note.";
                        note += "\r\r";

                        note = note + "\r\r अवलोकन तथा आगे की निर्णय हेतु प्रस्तुत है। तथानुसार, दिनांक" + " " + currentnitdt + " " + " के लिये निविदा आमंत्रण सूचना का मसौदा प्रस्तुत है।";
                        note = note + "\r\r प्रबंधक (वाणिज्य)";
                        note = note + "\r\r सहायक महा प्रबंधक (वाणिज्य)";
                        note = note + "\r\r उप महा पबंधक (क्षेत्र)";


                        rtbns.SelectAll();
                        rtbns.SelectionAlignment = HorizontalAlignment.Left;
                        rtbns.Text = note;
                    }
                }

            }
        }
        #endregion

#endregion 
        #region Result
      
        private void btnbrwsrslt_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1.InitialDirectory = @"C:\";
            openFileDialog1.Title = "Browse Files";
            openFileDialog1.CheckFileExists = true;
            openFileDialog1.CheckPathExists = true;

            openFileDialog1.DefaultExt = "xls";

            openFileDialog1.Filter = "Excel files (*.xls)|*.xls|All files (*.*)|*.*";
            openFileDialog1.RestoreDirectory = true;
            openFileDialog1.ReadOnlyChecked = true;
            openFileDialog1.ShowReadOnly = true;

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                
                txtbrwsresult.Text = openFileDialog1.FileName;

            }
        }

        private void btnshwrslt_Click(object sender, EventArgs e)
        {
            DataSet dsrprt = new DataSet();

            dsrprt = getConexcel_result(txtbrwsresult.Text);             
            

            dgvresult.DataSource = dsrprt.Tables[0];
        }
        private void tabPageResult_Click(object sender, EventArgs e)
        {
            
            DataSet ds = new DataSet();
            string qry = "Select NITDate from Publication order by NITDate";
            ds = select_data(qry);
            cmbrsltnit.DataSource = ds.Tables[0];
            cmbrsltnit.ValueMember = "NITdate";
            cmbrsltnit.DisplayMember = "NITdate";
            cmbrsltnit.Text = "-Select-";
        }
        private void btnrsltadd_Click(object sender, EventArgs e)
        {
            string nitdt = (Convert.ToDateTime(cmbrsltnit.SelectedValue)).ToLongDateString();
            string cmdty=cmbrsltcmdty.SelectedItem.ToString();

            DataSet ds1 = new DataSet();
            DataSet ds2 = new DataSet();
            DataSet ds3 = new DataSet();

            if(cmbof.SelectedItem=="EMD")
            {
                pnlemd.Visible = true;
                pnledtemd.Visible = false;

                lblrmncost.Visible = true;
                lblrmncost.Text = "Status:";
                txtrmncost.Visible = false;
                cmbstatus.Visible = true;
                lbldep2.Visible = false;
                txtdep2.Visible = false;
                lblutrno2.Visible = false;
                txtutrno2.Visible = false;
                lblutrdt2.Visible = false;
                dttmutrdt2.Visible = false;

                string qry3 = "Select  Party_Name from Party order by Party_Name";
                ds3 = select_data(qry3);
                if (ds3.Tables[0].Rows.Count > 0)
                {
                    cmbparty.Refresh();
                    // cmbparty.Items.Clear();
                    cmbparty.DataSource = ds3.Tables[0];
                    cmbparty.ValueMember = "Party_Name";
                    cmbparty.DisplayMember = "Party_Name";
                    cmbparty.Text = "-Select-";

                }

                else
                {
                    cmbparty.Refresh();
                    // cmbparty.Items.Clear();
                    // cmbparty.SelectedText  = "-Select-";
                }
                string qry2 = "Select  distinct Depot from District_Depot order by Depot";
                ds2 = select_data(qry2);
                if (ds2.Tables[0].Rows.Count > 0)
                {
                    cmbep.Refresh();
                    //cmbep.Items.Clear();
                    cmbep.DataSource = ds2.Tables[0];
                    cmbep.ValueMember = "Depot";
                    cmbep.DisplayMember = "Depot";
                    cmbep.Text = "-Select-";

                }
                else
                {
                    cmbep.Refresh();
                    //  cmbep.Items.Clear();
                    // cmbep.SelectedText = "-Select-";
                }
                string qry = "Select distinct District from District_Depot order by District";
                ds1 = select_data(qry);


                if (ds1.Tables[0].Rows.Count > 0)
                {
                    cmbdis.Refresh();
                    // cmbdis.Items.Clear();
                    cmbdis.DataSource = ds1.Tables[0];
                    cmbdis.ValueMember = "District";
                    cmbdis.DisplayMember = "District";
                    cmbdis.Text = "-Select-";
                }
                else
                {
                    cmbdis.Refresh();
                    //cmbdis.Items.Clear();
                    // cmbdis.SelectedText = "-Select-";
                }
            }
            else if(cmbof.SelectedItem=="COST")
            {
                pnlemd.Visible = true;                
                pnledtemd.Visible = false;
                lblrmncost.Visible = true;
                txtrmncost.Visible = true;
                lblrmncost.Text = "To Be Paid";
                cmbstatus.Visible = false;
                lbldep2.Visible = true;
                txtdep2.Visible = true;
                lblutrno2.Visible = true;
                txtutrno2.Visible = true;
                lblutrdt2.Visible = true;
                dttmutrdt2.Visible = true;


                string qry3 = "Select  distinct Party from Result where NITDate=#" + nitdt + "# and Commodity='" + cmdty + "' and EMDDep is NOT NULL and COSTDep1 is NULL and Status='Winner' order by Party";
                ds3 = select_data(qry3);
                if (ds3.Tables[0].Rows.Count > 0)
                {
                    cmbparty.Refresh();
                    // cmbparty.Items.Clear();
                    cmbparty.DataSource = ds3.Tables[0];
                    cmbparty.ValueMember = "Party";
                    cmbparty.DisplayMember = "Party";
                    cmbparty.Text = "-Select-";
                    txtqty.Enabled = true;
                    txtrate.Enabled = true;
                    txtrmncost.Enabled = true;
                    txtemd.Enabled = true;
                    txtutr.Enabled = true;
                    dttmutrdt.Enabled = true;
                    txtdep2.Enabled = true;
                    txtutrno2.Enabled = true;
                    dttmutrdt2.Enabled = true;

                }

                else if (ds3.Tables[0].Rows.Count == 0)
                {
                    txtqty.Enabled = false;
                    txtrate.Enabled = false;
                    txtrmncost.Enabled = false;
                    txtemd.Enabled = false;
                    txtutr.Enabled = false;
                    dttmutrdt.Enabled = false;
                    txtdep2.Enabled = false;
                    txtutrno2.Enabled = false;
                    dttmutrdt2.Enabled = false;
                    MessageBox.Show("Either Cost Already Deposited or No Winning Parties");
                    cmbparty.Refresh();

                    //cmbparty.Items.Clear();
                   // cmbparty.SelectedText  = "-Select-";
                }
                string qry2 = "Select  distinct Depot from Result where NITDate=#" + nitdt + "# and Commodity='" + cmdty + "'and EMDDep is NOT NULL and COSTDep1 is NULL and Status='Winner'order by Depot";
               
                ds2 = select_data(qry2);
                if (ds2.Tables[0].Rows.Count > 0)
                {
                    cmbep.Refresh();
                    //cmbep.Items.Clear();
                    cmbep.DataSource = ds2.Tables[0];
                    cmbep.ValueMember = "Depot";
                    cmbep.DisplayMember = "Depot";
                    cmbep.Text = "-Select-";

                }
                else
                {
                    cmbep.Refresh();
                   // cmbep.Items.Clear();
                   // cmbep.SelectedText = "-Select-";
                }
                string qry = "Select  distinct District from Result where NITDate=#" + nitdt + "# and Commodity='" + cmdty + "'and EMDDep is NOT NULL and COSTDep1 is NULL and Status='Winner' order by District";
                ds1 = select_data(qry);


                if (ds1.Tables[0].Rows.Count > 0)
                {
                    cmbdis.Refresh();
                    // cmbdis.Items.Clear();
                    cmbdis.DataSource = ds1.Tables[0];
                    cmbdis.ValueMember = "District";
                    cmbdis.DisplayMember = "District";
                    cmbdis.Text = "-Select-";
                }
                else
                {
                    cmbdis.Refresh();
                   // cmbdis.Items.Clear();
                   // cmbdis.SelectedText = "-Select-";
                }
              
            }
        }
        private void btnrsltedit_Click(object sender, EventArgs e)
        {
            dgvedemd.Visible = false;
            if (cmbof.SelectedItem == "EMD")
            {
                pnledtemd.Visible = true;
                pnlemd.Visible = false;

                btnrefnd.Text = "NON-WINNER ION";
               
                string nitdate = Convert.ToDateTime(cmbrsltnit.SelectedValue).ToLongDateString();
                string cmdty = cmbrsltcmdty.SelectedItem.ToString();
                string finyr = cmbrsltfinyr.SelectedItem.ToString();

                DataSet ds1 = new DataSet();
                DataSet ds2 = new DataSet();
                DataSet ds3 = new DataSet();
                
                string qry3 = "Select  distinct Party from Result where NITDate=#"+nitdate+"# and Commodity='"+cmdty+"' order by Party";
                ds3 = select_data(qry3);
                if (ds3.Tables[0].Rows.Count > 0)
                {
                    cmbedprty.Refresh();
                    // cmbedprty.Items.Clear();
                    cmbedprty.DataSource = ds3.Tables[0];
                    cmbedprty.ValueMember = "Party";
                    cmbedprty.DisplayMember = "Party";
                    cmbedprty.Text = "-Select-";

                }

                else
                {
                    cmbedprty.Refresh();
                    // cmbedprty.Items.Clear();
                    // cmbedprty.SelectedText  = "-Select-";
                }
                string qry2 = "Select  distinct Depot from Result where NITDate=#" + nitdate + "# and Commodity='" + cmdty + "' order by Depot";
                ds2 = select_data(qry2);
                if (ds2.Tables[0].Rows.Count > 0)
                {
                   cmbeddep.Refresh();
                   //cmbeddep.Items.Clear();
                   cmbeddep.DataSource = ds2.Tables[0];
                   cmbeddep.ValueMember = "Depot";
                   cmbeddep.DisplayMember = "Depot";
                   cmbeddep.Text = "-Select-";

                }
                else
                {
                    cmbeddep.Refresh();
                    //  cmbeddep.Items.Clear();
                    // cmbeddep.SelectedText = "-Select-";
                }
                string qry = "Select distinct District from Result where NITDate=#" + nitdate + "# and Commodity='" + cmdty + "' order by District";
                ds1 = select_data(qry);


                if (ds1.Tables[0].Rows.Count > 0)
                {
                   cmbeddis.Refresh();
                   // cmbeddis.Items.Clear();
                   cmbeddis.DataSource = ds1.Tables[0];
                   cmbeddis.ValueMember = "District";
                   cmbeddis.DisplayMember = "District";
                   cmbeddis.Text = "-Select-";
                }
                else
                {
                    cmbeddis.Refresh();
                    //cmbeddis.Items.Clear();
                    // cmbeddis.SelectedText = "-Select-";
                }


            }
            else if (cmbof.SelectedItem == "COST")
            {
               
                pnledtemd.Visible = true;
                pnlemd.Visible = false;

                btnrefnd.Text = "REFUND ION";

                string nitdt = (Convert.ToDateTime(cmbrsltnit.SelectedValue)).ToLongDateString();
                string cmdty = cmbrsltcmdty.SelectedItem.ToString();

                DataSet ds1 = new DataSet();
                DataSet ds2 = new DataSet();
                DataSet ds3 = new DataSet();


                string qry3 = "Select  distinct Party from Result where NITDate=#" + nitdt + "# and Commodity='" + cmdty + "'  and Status='Winner' order by Party";
                ds3 = select_data(qry3);
                if (ds3.Tables[0].Rows.Count > 0)
                {
                    cmbedprty.Refresh();
                    // cmbedprty.Items.Clear();
                    cmbedprty.DataSource = ds3.Tables[0];
                    cmbedprty.ValueMember = "Party";
                    cmbedprty.DisplayMember = "Party";
                    cmbedprty.Text = "-Select-";
                    
                }

                else 
                {

                    cmbedprty.Refresh();

                    //cmbedprty.Items.Clear();
                    // cmbedprty.SelectedText  = "-Select-";
                }
                string qry2 = "Select  distinct Depot from Result where NITDate=#" + nitdt + "# and Commodity='" + cmdty + "' and Status='Winner' order by Depot";

                ds2 = select_data(qry2);
                if (ds2.Tables[0].Rows.Count > 0)
                {
                    cmbeddep.Refresh();
                    //cmbeddep.Items.Clear();
                    cmbeddep.DataSource = ds2.Tables[0];
                    cmbeddep.ValueMember = "Depot";
                    cmbeddep.DisplayMember = "Depot";
                    cmbeddep.Text = "-Select-";

                }
                else
                {
                    cmbeddep.Refresh();
                    // cmbeddep.Items.Clear();
                    // cmbeddep.SelectedText = "-Select-";
                }
                string qry = "Select  distinct District from Result where NITDate=#" + nitdt + "# and Commodity='" + cmdty + "' and Status='Winner' order by District";
                ds1 = select_data(qry);


                if (ds1.Tables[0].Rows.Count > 0)
                {
                    cmbeddis.Refresh();
                    // cmbeddis.Items.Clear();
                    cmbeddis.DataSource = ds1.Tables[0];
                    cmbeddis.ValueMember = "District";
                    cmbeddis.DisplayMember = "District";
                    cmbeddis.Text = "-Select-";
                }
                else
                {
                    cmbeddis.Refresh();
                    // cmbeddis.Items.Clear();
                    // cmbeddis.SelectedText = "-Select-";
                }
              
               
            }

        }

        private void btnemdentr_Click(object sender, EventArgs e)
        {
            string nitdate =Convert.ToDateTime(cmbrsltnit.SelectedValue).ToLongDateString();
            string cmdty = cmbrsltcmdty.SelectedItem.ToString();
            string finyr = cmbrsltfinyr.SelectedItem.ToString();
            string dep = cmbep.SelectedValue.ToString();
            string dis = cmbdis.SelectedValue.ToString();
            string prty = cmbparty.SelectedValue.ToString();
            
            if (cmbparty.SelectedIndex != -1)
            {
                emdprty = cmbparty.SelectedValue.ToString();
            }

            if (cmbep.SelectedIndex != -1)
            {
                emddep = cmbep.SelectedValue.ToString();
            }
            if (cmbdis.SelectedIndex != -1)
            {
                emddis = cmbdis.SelectedValue.ToString();
            }
            // enter from textboxes
            string emdqty = txtqty.Text;
            string emdrate = txtrate.Text;
            string emdrs = txtemd.Text;
            double emdrs1 = Convert.ToDouble(txtemd.Text);
            string emdutrno = txtutr.Text;
            string emdutrdt = dttmutrdt.Value.ToLongDateString();
          
            

           
            
            if(cmbrsltcmdty.SelectedIndex!=-1)
            {
                if(cmbof.SelectedItem=="EMD")
                {
                    string status = cmbstatus.SelectedItem.ToString();
                   if(status=="Winner")
                   {  //calculated entries
                       double totcost = (Convert.ToDouble(emdqty) * Convert.ToDouble(emdrate) * 10) + (Convert.ToDouble(emdqty) * Convert.ToDouble(emdrate) * 10 * 0.005);
                       double remaincost = (totcost - (Convert.ToDouble(emdrs)));

                       string qry = "insert into Result(Commodity,NITDate,FinYr,Party,Depot,District,Qty,Rate,EMDDep,EMDUTRno,EMDUTRdt,Status,TotalCost,Remain) values('" + cmdty + "','" + nitdate + "','" + finyr + "','" + emdprty + "','" + emddep + "','" + emddis + "','" + emdqty + "','" + emdrate + "'," + emdrs1 + ",'" + emdutrno + "','" + emdutrdt + "','" + status + "'," + totcost + "," + remaincost + ")";
                       insert_update_deleted(qry);
                   }
                   else if(status=="Non-Winner")
                   {
                       

                       string qry = "insert into Result(Commodity,NITDate,FinYr,Party,Depot,District,Qty,Rate,EMDDep,EMDUTRno,EMDUTRdt,Status,TotalCost,Remain) values('" + cmdty + "','" + nitdate + "','" + finyr + "','" + emdprty + "','" + emddep + "','" + emddis + "',NULL,NULL," + emdrs1 + ",'" + emdutrno + "','" + emdutrdt + "','" + status + "',NULL,NULL)";
                       insert_update_deleted(qry);
                   }

                }
                else if (cmbof.SelectedItem == "COST")
                {
                    if (txtdep2.Text != "")
                    {
                        refund = (Convert.ToDouble(emdrs) + Convert.ToDouble(txtdep2.Text)) - (Convert.ToDouble(txtrmncost.Text));
                        dep2 = Convert.ToDouble(txtdep2.Text);
                        string qry = "update Result set COSTDep1=" + emdrs1 + ",COSTUTRno1='" + emdutrno + "',COSTUTRdt1='" + emdutrdt + "',COSTDep2=" + dep2 + ",COSTUTRno2='" + txtutrno2.Text + "',COSTUTRdt2='" + dttmutrdt2.Value.ToShortDateString() + "',RefundAmnt=" + refund + " where NITDate=#" + nitdate + "# and Commodity='" + cmdty + "' and EMDDep is NOT NULL and Status='Winner'and Party='" + prty + "' and Depot='" + dep + "' and District='" + dis + "'";
                        insert_update_deleted(qry);
                    }
                    else
                    {
                        refund = (Convert.ToDouble(emdrs)) - (Convert.ToDouble(txtrmncost.Text));
                        string qry = "update Result set COSTDep1=" + emdrs1 + ",COSTUTRno1='" + emdutrno + "',COSTUTRdt1='" + emdutrdt + "',RefundAmnt=" + refund + " where NITDate=#" + nitdate + "# and Commodity='" + cmdty + "' and EMDDep is NOT NULL and Status='Winner'and Party='" + prty + "' and Depot='" + dep + "' and District='" + dis + "'";
                        insert_update_deleted(qry);
                    }
                   //insert_update_deleted(qry);
                }    
              
            }
            

        }
        private void btnshowemd_Click(object sender, EventArgs e)
        {
            DataGridViewButtonColumn editbtn = new DataGridViewButtonColumn();
            editbtn.Name = "EDIT";
            editbtn.HeaderText = "EDIT";
            editbtn.Text = "EDIT";
            editbtn.UseColumnTextForButtonValue = true;
            editbtn.FlatStyle = FlatStyle.Popup;

            DataGridViewButtonColumn dltbtn = new DataGridViewButtonColumn();
            dltbtn.Name = "DELETE";
            dltbtn.HeaderText = "DELETE";
            dltbtn.Text = "DELETE";
            dltbtn.UseColumnTextForButtonValue = true;
            dltbtn.FlatStyle = FlatStyle.Popup;


            DataSet ds = new DataSet();

            if (cmbedprty.SelectedIndex != -1 || cmbeddep.SelectedIndex != -1 || cmbeddis.SelectedIndex != -1)
            {
                string nitdt = (Convert.ToDateTime(cmbrsltnit.SelectedValue)).ToLongDateString();
               if(cmbof.SelectedItem=="EMD")
               {
                   string qry = "Select ID,Commodity,NITDate,FinYr,Party,Depot,District,Status,Qty,Rate,EMDDep,EMDUTRno,EMDUTRdt,TotalCost,Remain from Result where NITdate=#" + nitdt + "# and Commodity='" + cmbrsltcmdty.SelectedItem.ToString() + "' and FinYr='" + cmbrsltfinyr.SelectedItem.ToString() + "' and (Party='" + cmbedprty.SelectedValue + "' or Depot='" + cmbeddep.SelectedValue + "' or District='" + cmbeddis.SelectedValue + "')";
                   ds = select_data(qry);
                   if (ds.Tables[0].Rows.Count > 0)
                   {
                       dgvedemd.Columns.Clear();
                       dgvedemd.Visible = true;
                       //dgvedemd.AutoGenerateColumns = true;
                       dgvedemd.DataSource = ds.Tables[0];
                       dgvedemd.Columns["ID"].ReadOnly = true;
                       dgvedemd.Columns["Status"].ReadOnly = true;
                       dgvedemd.Columns["TotalCost"].ReadOnly = true;
                       dgvedemd.Columns["Remain"].ReadOnly = true;
                       dgvedemd.Columns.Add(editbtn);
                       dgvedemd.Columns.Add(dltbtn);
                       dgvedemd.Refresh();


                   }

                   else
                   {
                       dgvedemd.Visible = false;
                       MessageBox.Show("No data");
                   }
               }
               else if (cmbof.SelectedItem == "COST")
               {
                   string qry = "Select ID,Commodity,NITDate,FinYr,Party,Depot,District,Qty,Rate,Remain,COSTDep1,COSTUTRno1,COSTUTRdt1,COSTDep2,COSTUTRno2,COSTUTRdt2,RefundAmnt from Result where NITdate=#" + nitdt + "# and Commodity='" + cmbrsltcmdty.SelectedItem.ToString() + "' and FinYr='" + cmbrsltfinyr.SelectedItem.ToString() + "' and Status='Winner' and (Party='" + cmbedprty.SelectedValue + "' or Depot='" + cmbeddep.SelectedValue + "' or District='" + cmbeddis.SelectedValue + "')";
                   ds = select_data(qry);
                   if (ds.Tables[0].Rows.Count > 0)
                   {
                       dgvedemd.Columns.Clear();
                       dgvedemd.Visible = true;
                       //dgvedemd.AutoGenerateColumns = true;
                       dgvedemd.DataSource = ds.Tables[0];
                       dgvedemd.Columns["ID"].ReadOnly = true;
                      // dgvedemd.Columns["TotalCost"].ReadOnly = true;
                       dgvedemd.Columns["Remain"].Name = "Amount To Be paid";
                       dgvedemd.Columns["Amount To Be paid"].ReadOnly = true;
                       dgvedemd.Columns["RefundAmnt"].ReadOnly = true;
                       dgvedemd.Columns["Qty"].ReadOnly = true;
                       dgvedemd.Columns["Rate"].ReadOnly = true;
                       dgvedemd.Columns.Add(editbtn);
                       dgvedemd.Columns.Add(dltbtn);
                       dgvedemd.Refresh();


                   }

                   else
                   {
                       MessageBox.Show("No data");
                   }
               }
                
            }


            

        }

        private void dgvedemd_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            if(cmbof.SelectedItem=="EMD")
            {
                // Check if wanted column was clicked 
                if (e.ColumnIndex == 15 && e.RowIndex >= 0)
                {
                    //Perform on edit button click code   
                    //  Int64 id =Convert.ToInt64 (dgvedtfcidis.CurrentRow.Cells[0].Value);
                    int id = Int32.Parse(dgvedemd.CurrentRow.Cells[0].Value.ToString());
                    string emdedcmdty = dgvedemd.CurrentRow.Cells[1].Value.ToString();
                    string emdednit = dgvedemd.CurrentRow.Cells[2].Value.ToString();
                    string emdedfinyr = dgvedemd.CurrentRow.Cells[3].Value.ToString();
                    string emdedprty = dgvedemd.CurrentRow.Cells[4].Value.ToString();
                    string emdeddep = dgvedemd.CurrentRow.Cells[5].Value.ToString();
                    string emdeddis = dgvedemd.CurrentRow.Cells[6].Value.ToString();
                    string status = dgvedemd.CurrentRow.Cells[7].Value.ToString();
                    string emdedqty = dgvedemd.CurrentRow.Cells[8].Value.ToString();
                    string emdedrt = dgvedemd.CurrentRow.Cells[9].Value.ToString();
                    string emdedrs = dgvedemd.CurrentRow.Cells[10].Value.ToString();
                    string emdedutrno = dgvedemd.CurrentRow.Cells[11].Value.ToString();
                    string emdedutrdt = dgvedemd.CurrentRow.Cells[12].Value.ToString();
                    // string emdedtotcost = dgvedemd.CurrentRow.Cells[12].Value.ToString();
                    //string emdedremn = dgvedemd.CurrentRow.Cells[13].Value.ToString();
                    if(status=="Winner" && emdedqty!=""&& emdedrt!="")
                    { 
                    double totcost = (Convert.ToDouble(emdedqty) * Convert.ToDouble(emdedrt) * 10) + (Convert.ToDouble(emdedqty) * Convert.ToDouble(emdedrt) * 10 * 0.005);
                    double remaincost = (totcost - (Convert.ToDouble(emdedrs)));

                    // MessageBox.Show(id.ToString());
                    string qry_updt = "update Result set Commodity='" + emdedcmdty + "',NITDate='" + emdednit + "',FinYr='" + emdedfinyr + "',Party='" + emdedprty + "',Depot='" + emdeddep + "',District='" + emdeddis + "',Qty='" + emdedqty + "',Rate='" + emdedrt + "',EMDdep='" + emdedrs + "',EMDUTRno='" + emdedutrno + "',EMDUTRdt='" + emdedutrdt + "',TotalCost=" + totcost + ",Remain=" + remaincost + " where ID=" + id + " and Status='Winner'";
                    insert_update_deleted(qry_updt);
                    }
                    if (status == "Non-Winner" && emdedqty == "" && emdedrt == "")
                    {
                        
                        // MessageBox.Show(id.ToString());
                        string qry_updt = "update Result set Commodity='" + emdedcmdty + "',NITDate='" + emdednit + "',FinYr='" + emdedfinyr + "',Party='" + emdedprty + "',Depot='" + emdeddep + "',District='" + emdeddis + "',Qty=NULL,Rate=NULL,EMDdep='" + emdedrs + "',EMDUTRno='" + emdedutrno + "',EMDUTRdt='" + emdedutrdt + "',TotalCost= NULL,Remain=NULL where ID=" + id + " and Status='Non-Winner'";
                        insert_update_deleted(qry_updt);
                    }
                }

                else if (e.ColumnIndex == 16 && e.RowIndex >= 0)
                {
                    //Perform on edit button click code          

                    int id = Int32.Parse(dgvedemd.CurrentRow.Cells[0].Value.ToString());
                    string qry_updt = "delete from Result  WHERE ID=" + id + " ";
                    insert_update_deleted(qry_updt);
                    dgvedemd.Rows.RemoveAt(dgvedemd.CurrentRow.Index);

                }

            }
            else if(cmbof.SelectedItem=="COST")
            {
                // Check if wanted column was clicked 
                if (e.ColumnIndex == 17 && e.RowIndex >= 0)
                {
                    //Perform on edit button click code   
                   
                    int id = Int32.Parse(dgvedemd.CurrentRow.Cells[0].Value.ToString());
                    string emdedcmdty = dgvedemd.CurrentRow.Cells[1].Value.ToString();
                    string emdednit = dgvedemd.CurrentRow.Cells[2].Value.ToString();
                    string emdedfinyr = dgvedemd.CurrentRow.Cells[3].Value.ToString();
                    string emdedprty = dgvedemd.CurrentRow.Cells[4].Value.ToString();
                    string emdeddep = dgvedemd.CurrentRow.Cells[5].Value.ToString();
                    string emdeddis = dgvedemd.CurrentRow.Cells[6].Value.ToString();
                    string emdedqty = dgvedemd.CurrentRow.Cells[7].Value.ToString();
                    string emdedrt = dgvedemd.CurrentRow.Cells[8].Value.ToString();
                    string remain = dgvedemd.CurrentRow.Cells[9].Value.ToString();                   
                    string emdedrs = dgvedemd.CurrentRow.Cells[10].Value.ToString();
                    double emdrs1 = Convert.ToDouble(dgvedemd.CurrentRow.Cells[10].Value.ToString());
                    string emdedutrno = dgvedemd.CurrentRow.Cells[11].Value.ToString();
                    string emdedutrdt = dgvedemd.CurrentRow.Cells[12].Value.ToString();
                    string dep2=dgvedemd.CurrentRow.Cells[13].Value.ToString();
                    string utrno2=dgvedemd.CurrentRow.Cells[14].Value.ToString();
                    string utrdt2=dgvedemd.CurrentRow.Cells[15].Value.ToString();

                    if(remain!="" && dep2!="")
                    {
                        remn=Convert.ToDouble(remain);
                        double depcost2 = Convert.ToDouble(dep2);
                        refund = (emdrs1 + depcost2) - remn;
                        string qry_updt = "update Result set Commodity='" + emdedcmdty + "',NITDate='" + emdednit + "',FinYr='" + emdedfinyr + "',Party='" + emdedprty + "',Depot='" + emdeddep + "',District='" + emdeddis + "',Qty='" + emdedqty + "',Rate='" + emdedrt + "',COSTDep1=" + emdedrs + ",COSTUTRno1='" + emdedutrno + "',COSTUTRdt1='" + emdedutrdt + "',COSTDep2=" + depcost2 + ",COSTUTRno2='" + utrno2 + "',COSTUTRdt2='" + utrdt2 + "',RefundAmnt=" + refund + " where ID=" + id + "";
                        insert_update_deleted(qry_updt);
                    }
                    else if (remain != "" && dep2 == "")
                    {
                        
                        remn = Convert.ToDouble(remain);
                       // double depcost2 = 0.00;
                        refund = (emdrs1) - remn;
                        string qry_updt = "update Result set Commodity='" + emdedcmdty + "',NITDate='" + emdednit + "',FinYr='" + emdedfinyr + "',Party='" + emdedprty + "',Depot='" + emdeddep + "',District='" + emdeddis + "',Qty='" + emdedqty + "',Rate='" + emdedrt + "',COSTDep1=" + emdedrs + ",COSTUTRno1='" + emdedutrno + "',COSTUTRdt1='" + emdedutrdt + "',COSTDep2=NULL,COSTUTRno2='" + utrno2 + "',COSTUTRdt2='" + utrdt2 + "',RefundAmnt=" + refund + " where ID=" + id + "";
                        insert_update_deleted(qry_updt);
                    }

                    else if (remain == "" && dep2 == "")
                    {
                        refund = (emdrs1);
                        string qry_updt = "update Result set Commodity='" + emdedcmdty + "',NITDate='" + emdednit + "',FinYr='" + emdedfinyr + "',Party='" + emdedprty + "',Depot='" + emdeddep + "',District='" + emdeddis + "',Qty='" + emdedqty + "',Rate='" + emdedrt + "',COSTDep1=" + emdedrs + ",COSTUTRno1='" + emdedutrno + "',COSTUTRdt1='" + emdedutrdt + "',COSTDep2=NULL,COSTUTRno2='" + utrno2 + "',COSTUTRdt2='" + utrdt2 + "',RefundAmnt=" + refund + " where ID=" + id + "";
                        insert_update_deleted(qry_updt);
                        MessageBox.Show("No Remaining Amnt to be paid, so no cost dep");
      
                    }
                    

                }

                else if (e.ColumnIndex == 18 && e.RowIndex >= 0)
                {
                    //Perform on edit button click code          

                    int id = Int32.Parse(dgvedemd.CurrentRow.Cells[0].Value.ToString());
                    string qry_updt = "delete from Result  WHERE ID=" + id + " ";
                    insert_update_deleted(qry_updt);
                    dgvedemd.Rows.RemoveAt(dgvedemd.CurrentRow.Index);


                }
            }
            
        }

        private void cmbdis_SelectedIndexChanged(object sender, EventArgs e)
        {

            if (cmbof.SelectedIndex != -1 && cmbrsltnit.SelectedIndex != -1 && cmbrsltcmdty.SelectedIndex != -1 && cmbrsltfinyr.SelectedIndex != -1)
            {
                if (cmbof.SelectedItem == "COST")
                {
                    DataSet ds_1 = new DataSet();
                    string nitdt = (Convert.ToDateTime(cmbrsltnit.SelectedValue)).ToLongDateString();
                    string cmdty = cmbrsltcmdty.SelectedItem.ToString();
                    if (cmbparty.SelectedIndex != -1 && cmbdis.SelectedIndex != -1 && cmbep.SelectedIndex != -1)
                    {
                        string prty = cmbparty.SelectedValue.ToString();
                        string dep = cmbep.SelectedValue.ToString();
                        string dis = cmbdis.SelectedValue.ToString();


                        string qry = "Select  Qty,Rate,Remain from Result where NITDate=#" + nitdt + "# and Commodity='" + cmdty + "' and EMDDep is NOT NULL and COSTDep1 is NULL and Party='" + prty + "' and Depot='" + dep + "' and District='" + dis + "'";
                        ds_1 = select_data(qry);
                        if (ds_1.Tables[0].Rows.Count > 0)
                        {
                            txtqty.Enabled = true;
                            txtrate.Enabled = true;
                            txtrmncost.Enabled = true;
                            txtemd.Enabled = true;
                            txtutr.Enabled = true;
                            dttmutrdt.Enabled = true;
                            txtdep2.Enabled = true;
                            txtutrno2.Enabled = true;
                            dttmutrdt2.Enabled = true;
                            txtqty.Text = ds_1.Tables[0].Rows[0].ItemArray[0].ToString();
                            txtrate.Text = ds_1.Tables[0].Rows[0].ItemArray[1].ToString();
                            txtrmncost.Text = ds_1.Tables[0].Rows[0].ItemArray[2].ToString();
                            txtqty.Enabled = true;
                            txtrate.Enabled = true;
                            txtrmncost.Enabled = true;
                            txtemd.Enabled = true;
                            txtutr.Enabled = true;
                            dttmutrdt.Enabled = true;
                            txtdep2.Enabled = true;
                            txtutrno2.Enabled = true;
                            dttmutrdt2.Enabled = true;


                        }
                        else if (ds_1.Tables[0].Rows.Count==0)
                        {
                            txtqty.Text = "";
                            txtrate.Text = "";
                            txtrmncost.Text ="";
                            txtqty.Enabled = false;
                            txtrate.Enabled = false;
                            txtrmncost.Enabled = false;
                            txtemd.Enabled = false;
                            txtutr.Enabled = false;
                            dttmutrdt.Enabled = false;
                            txtdep2.Enabled = false;
                            txtutrno2.Enabled = false;
                            dttmutrdt2.Enabled = false;
                            MessageBox.Show("Either Cost Already Deposited or No Winning Parties");

                        }

                    }


                }
                else
                {
                    txtqty.Text = "";
                    txtrate.Text = "";
                    txtrmncost.Text = "";
                }
            }
        }
        //IONS
        #region EMD_COST ION
        private void btnemdcostion_Click(object sender, EventArgs e)
        {
            DataSet emd_cost = new DataSet();
            string str;
            if (cmbrsltnit.SelectedIndex != -1 && cmbrsltcmdty.SelectedIndex!=-1 && cmbrsltfinyr.SelectedIndex!=-1 && cmbof.SelectedIndex!=-1)
            {
                string nitdt = Convert.ToDateTime(cmbrsltnit.SelectedValue).ToLongDateString();
                string cmdty = cmbrsltcmdty.SelectedItem.ToString();
                string finyr =cmbrsltfinyr.SelectedItem.ToString();
                if (cmbof.SelectedItem == "EMD")
                {
                    str = "Select * from Result where NITDate=#" + nitdt + "# and Commodity='" + cmdty + "' and FinYr='" + finyr + "'and Status='Winner'";
                    emd_cost = select_data(str);
                }
                else if (cmbof.SelectedItem == "COST")
                {
                    str = "Select * from Result where NITDate=#" + nitdt + "# and Commodity='" + cmdty + "' and FinYr='" + finyr + "'and COSTdep1 is NOT NULL and Status='Winner'";
                    emd_cost = select_data(str);
                }
               if (emd_cost.Tables[0].Rows.Count != 0)
                    {
                        EMD_COST_ION_CREATION(emd_cost);
                    }
                    else
                    {
                        MessageBox.Show("No data for these selections", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }

                
            }
            else
            {
                MessageBox.Show("Enter valid Commodity or NITDate or Financial Year or all ", "Invalid Selection", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        // func for EMD_COST ION
        private void EMD_COST_ION_CREATION(DataSet emd_cost)
        {
            string of = cmbof.SelectedItem.ToString();
            string cmdty = cmbrsltcmdty.SelectedItem.ToString();
            string finyr = cmbrsltfinyr.SelectedItem.ToString();   //all the static items in the page hold in variables
                //COMMON
                //string hindi_ref = "संदर्भ सं.: वाणिज्य.32/ओ.एम.एस.एस(डी)/बल्क/आई.ओ.जी/2014-15";
               // string hindi_date = "दिनांक:" + " " + DateTime.Today.ToShortDateString();
             string refno = "Ref. No.:- " + "Com.32/OMSS(D)/BULK/" + cmdty + "/ " + finyr;
            string date = "Dated:" + DateTime.Today.ToShortDateString();
                string ion = "अंतर कार्यालय नोट/I.O.N";
                string nitdate = Convert.ToDateTime(cmbrsltnit.SelectedValue).ToShortDateString();
                string EAdate = Convert.ToDateTime(cmbrsltnit.SelectedValue).AddDays(3).ToShortDateString();
               
            //EMD  

                string one_one = "Please find enclosed herewith the information regarding the EMD deposited by the participating parties in Sale of " + " " + cmdty + " " + " under OMSS (D) through E-Auctioning held on" + " " + EAdate + " " + "corresponding to NIT dated: " + " " + nitdate + " " + ",duly verified with the Commercial Bank Statements of the dedicated Cash Credit A/C, as well as cross-verified with the correspondences received from NCDEX. ";
               
            //COST

                string c_one_one = "Please find enclosed herewith the consolidated information regarding the party-wise,district-wise Cost Details and EMD details (EMD details have already been intimated to Cash section vide I.O.N no." + " "+refno + " "+")who participated in Sale of" +" "+cmdty+" "+" under OMSS (D) through E-Auctioning held on"+ " " +EAdate+" "+"corresponding to NIT"+ " " +nitdate+" "+" You are requested to make necessary A/C book entries, besides sending the IOGA to the concerned District Office for the amounts mentioned therein at the earliest please.";
             
             
                // Create a new PDF document
                PdfDocument document = new PdfDocument();
                document.Info.Title = "Created with PDFsharp";

                // Create an empty page
                PdfPage page = document.AddPage();

                // Get an XGraphics object for drawing
                XGraphics gfx = XGraphics.FromPdfPage(page);
                XTextFormatter tf = new XTextFormatter(gfx);


                // Draw FCI logo of the page
                string jpegSamplePath = "../IMAGES/logo.jpg";
                XImage image = XImage.FromFile(jpegSamplePath);
                gfx.DrawImage(image, 44, 70, 500, 80);

                // Create a font

                XFont font1 = new XFont("Calibri (Body)", 12, XFontStyle.Regular);
                XFont font2 = new XFont("Calibri (Body)", 12, XFontStyle.Underline);
                XFont font3 = new XFont("Calibri (Body)", 12, XFontStyle.Italic);
                XFont font4 = new XFont("Calibri (Body)", 12, XFontStyle.Bold);
                XFont font5 = new XFont("Calibri (Body)", 10, XFontStyle.Regular);
                XFont font6 = new XFont("Calibri (Body)", 11, XFontStyle.Bold);
                XFont font7 = new XFont("Calibri (Body)", 08, XFontStyle.Bold);
                XFont font8 = new XFont("Calibri (Body)", 08, XFontStyle.Regular);

                int x = 44;
                int y = 30;

                // Draw the text

                // file refno
               // gfx.DrawString(hindi_ref, font1, XBrushes.Black, new XRect(x, y + 130, page.Width, page.Height), XStringFormats.TopLeft);
                gfx.DrawString(refno, font1, XBrushes.Black, new XRect(x, y + 130, page.Width, page.Height), XStringFormats.TopLeft);
               
               // date 
               // gfx.DrawString(hindi_date, font1, XBrushes.Black, new XRect(x + 390, y + 130, page.Width, page.Height), XStringFormats.TopLeft);
                gfx.DrawString(date, font1, XBrushes.Black, new XRect(x + 390, y + 130, page.Width, page.Height), XStringFormats.TopLeft);
               
            //ION
               // gfx.DrawString(ion, font4, XBrushes.Black, new XRect(x - 44, y + 155, page.Width, page.Height), XStringFormats.TopCenter);
               
                string jpegSamplePath1= "../IMAGES/hindi4.jpg";
                XImage imageion = XImage.FromFile(jpegSamplePath1);
                gfx.DrawImage(imageion, x + 130, y + 165, 250, 15);
               
                // para1    
                if (of == "EMD")
                {
                    
                    tf.Alignment = XParagraphAlignment.Justify;
                    tf.DrawString(one_one, font1, XBrushes.Black, new XRect(x, y + 205+18, 500, 120), XStringFormats.TopLeft);
                    //para 3
                    string jpegSamplePath3 = "../IMAGES/hindi2.jpg";
                    XImage imagepara3 = XImage.FromFile(jpegSamplePath3);
                    gfx.DrawImage(imagepara3, x, y + 205 + 18 + 200, 520, 250); 
                
                }
                else if (of == "COST")
                {
                    tf.Alignment = XParagraphAlignment.Justify;
                    tf.DrawString(c_one_one, font1, XBrushes.Black, new XRect(x, y + 205+18, 500, 120), XStringFormats.TopLeft);
                    //para 3
                    string jpegSamplePath3 = "../IMAGES/hindi1.jpg";
                    XImage imagepara3 = XImage.FromFile(jpegSamplePath3);
                    gfx.DrawImage(imagepara3,x, y + 205 + 18 + 200, 520, 250);  

                }
               
                //para2
                string jpegSamplePath2 = "../IMAGES/hindi3.jpg";
                XImage imagepara2 = XImage.FromFile(jpegSamplePath2);
                gfx.DrawImage(imagepara2, x, y + 205 +100+18, 350, 30);



                int x1 = 44;
                int y1 = 30;
                // drawing table

                // Create an empty page
                PdfPage page2 = document.AddPage();

                // Get an XGraphics object for drawing
                XGraphics gfx2 = XGraphics.FromPdfPage(page2);
                XTextFormatter tf2 = new XTextFormatter(gfx2);


               

                //drawing header of table for EMD

                string title = "EMD DETAILS OF SUCCESSFUL BIDDERS AGAINST NIT DTD:" +" "+nitdate+", e-AUCTION FOR WHICH WAS HELD ON"+" "+EAdate+" "+"(WEST BENGAL REGION)";
                string column1 = "Sl. No.";
                string column2 = "Name of Party ";
                string column3 = "Depot";
                string column4 = "F.C.I D.O";
                string column5 = "Qty (MT)";
                string column6 = "Rate(Rs/Qtl)";
               string column7 = "EMD DEPOSITED (Rs)";
               string column8= "EMD UTR NO";
                string column9 = "EMD DATE";

                //drawing header of table for COST

                string c_title = "EMD and Cost Details for NIT dated :" + " " + nitdate + " " + "E-Auctioning held on" + " " + EAdate;
                string c_column1 = "Sl. No.";
                string c_column2 = "Name of Party ";
                string c_column3 = "Depot";
                string c_column4 = "F.C.I D.O";
                string c_column5 = "Qty(MT)";
                string c_column6 = "Rate(Rs/Qtl)";
               // string c_column7 = "Total Cost";
                //string c_column8 = "Remaining Cost";
                string c_column9 = "Cost(Rs.)";
                string c_column10 = "UTR No.";
                string c_column11 = "Deposit Dt.";
                string c_column12 = "Cost2(Rs.)";
                string c_column13 = "UTR No.";
                string c_column14 = "Deposit Dt.";

            if(of=="EMD")
            {
                tf2.Alignment = XParagraphAlignment.Justify;
                tf2.DrawString(title, font4, XBrushes.Black, new XRect(x1, y1, 500, 120), XStringFormats.TopLeft);
                    
                //SL. NO.
                gfx2.DrawRectangle(XPens.Black, x1 - 42, y1 + 100, 30, 29);
                gfx2.DrawString(column1, font7, XBrushes.Black, new XRect(x1 - 42, y1 + 105, page.Width.Point, page.Height.Point), XStringFormats.TopLeft);

                //Party
                gfx2.DrawRectangle(XPens.Black, x1 - 42 + 30, y1 + 100, 85, 29);
                gfx2.DrawString(column2, font7, XBrushes.Black, new XRect(x1 - 42 + 30 + 3, y1 + 105, page.Width.Point, page.Height.Point), XStringFormats.TopLeft);
               
                //Depot
                gfx2.DrawRectangle(XPens.Black, x1 - 42 + 30 + 85, y1 + 100, 65, 29);
                gfx2.DrawString(column3, font7, XBrushes.Black, new XRect(x1 - 42 + 30 + 85 + 5, y1 + 105, page.Width.Point, page.Height.Point), XStringFormats.TopLeft);

                //District
                gfx2.DrawRectangle(XPens.Black, x1 - 42 + 30 + 85 + 65, y1 + 100, 65, 29);
                gfx2.DrawString(column4, font7, XBrushes.Black, new XRect(x1 - 42 + 30 + 75 + 65 + 12, y1 + 105, page.Width.Point, page.Height.Point), XStringFormats.TopLeft);
                
                // qty
                gfx2.DrawRectangle(XPens.Black, x1 - 42 + 30 + 85 + 65 + 65, y1 + 100, 50, 29);
                gfx2.DrawString(column5, font7, XBrushes.Black, new XRect(x1 - 42 + 30 + 75 + 65 + 65 + 10, y1 + 105, page.Width.Point, page.Height.Point), XStringFormats.TopLeft);

                // rt
                gfx2.DrawRectangle(XPens.Black, x1 - 42 + 30 + 85 + 65 + 65 + 50, y1 + 100, 50, 29);
                gfx2.DrawString(column6, font7, XBrushes.Black, new XRect(x1 - 42 + 30 + 75 + 65 + 65 + 50 + 10, y1 + 105, page.Width.Point, page.Height.Point), XStringFormats.TopLeft);

                // emd
                gfx2.DrawRectangle(XPens.Black, x1 - 42 + 30 + 85 + 65 + 65 + 50 + 50, y1 + 100, 90, 29);
                gfx2.DrawString(column7, font7, XBrushes.Black, new XRect(x1 - 42 + 30 + 75 + 65 + 65 + 50 + 50 + 10, y1 + 105, page.Width.Point, page.Height.Point), XStringFormats.TopLeft);

                // emd utr no
                gfx2.DrawRectangle(XPens.Black, x1 - 42 + 30 + 85 + 65 + 65 + 50 + 50 + 90, y1 + 100, 100, 29);
                gfx2.DrawString(column8, font7, XBrushes.Black, new XRect(x1 - 42 + 30 + 85 + 65 + 65 + 50 + 50 + 90 + 3, y1 + 105, page.Width.Point, page.Height.Point), XStringFormats.TopLeft);

                // emd utr dt
                gfx2.DrawRectangle(XPens.Black, x1 - 42 + 30 + 85 + 65 + 65 + 50 + 50 + 90 + 100, y1 + 100, 50, 29);
                gfx2.DrawString(column9, font7, XBrushes.Black, new XRect(x1 - 42 + 30 + 85 + 65 + 65 + 50 + 50 + 90 + 100 + 3, y1 + 105, page.Width.Point, page.Height.Point), XStringFormats.TopLeft);

                for (int i = 0; i < emd_cost.Tables[0].Rows.Count; i++)
                {
                    //drawing body of table

                    string sl = (i + 1).ToString();
                    string party = emd_cost.Tables[0].Rows[i].ItemArray[4].ToString();
                    string depot = emd_cost.Tables[0].Rows[i].ItemArray[5].ToString();
                    string dis = emd_cost.Tables[0].Rows[i].ItemArray[6].ToString();
                    string qty = emd_cost.Tables[0].Rows[i].ItemArray[7].ToString();
                    string rate = emd_cost.Tables[0].Rows[i].ItemArray[8].ToString();
                    string emd = emd_cost.Tables[0].Rows[i].ItemArray[9].ToString();
                    string emdutrno = emd_cost.Tables[0].Rows[i].ItemArray[10].ToString();
                    string emddt = Convert.ToDateTime( emd_cost.Tables[0].Rows[i].ItemArray[11]).ToShortDateString();

                    //SL. NO.
                    gfx2.DrawRectangle(XPens.Black, x1 - 42, y1 + 105+25, 30, 15);
                    gfx2.DrawString(sl, font7, XBrushes.Black, new XRect(x1 - 42, y1 + 105 + 29, page.Width.Point, page.Height.Point), XStringFormats.TopLeft);

                    //Party
                    gfx2.DrawRectangle(XPens.Black, x1 - 42 + 30, y1 + 105 + 25, 85, 15);
                    gfx2.DrawString(party, font7, XBrushes.Black, new XRect(x1 - 42 + 30 + 3, y1 + 105 + 29, page.Width.Point, page.Height.Point), XStringFormats.TopLeft);

                    //Depot
                    gfx2.DrawRectangle(XPens.Black, x1 - 42 + 30 + 85, y1 + 105 + 25, 65, 15);
                    gfx2.DrawString(depot, font7, XBrushes.Black, new XRect(x1 - 42 + 30 + 85 + 3, y1 + 105 + 29, page.Width.Point, page.Height.Point), XStringFormats.TopLeft);

                    //District
                    gfx2.DrawRectangle(XPens.Black, x1 - 42 + 30 + 85 + 65, y1 + 105 + 25, 65, 15);
                    gfx2.DrawString(dis, font7, XBrushes.Black, new XRect(x1 - 42 + 30 + 85 + 65 + 3, y1 + 105 + 29, page.Width.Point, page.Height.Point), XStringFormats.TopLeft);

                    // qty
                    gfx2.DrawRectangle(XPens.Black, x1 - 42 + 30 + 85 + 65 + 65, y1 + 105 + 25, 50, 15);
                    gfx2.DrawString(qty, font7, XBrushes.Black, new XRect(x1 - 42 + 30 + 85 + 65 + 65 + 3, y1 + 105 + 29, page.Width.Point, page.Height.Point), XStringFormats.TopLeft);

                    // rt
                    gfx2.DrawRectangle(XPens.Black, x1 - 42 + 30 + 85 + 65 + 65 + 50, y1 + 105 + 25, 50, 15);
                    gfx2.DrawString(rate, font7, XBrushes.Black, new XRect(x1 - 42 + 30 + 85 + 65 + 65 + 50 + 3, y1 + 105 + 29, page.Width.Point, page.Height.Point), XStringFormats.TopLeft);

                    // emd
                    gfx2.DrawRectangle(XPens.Black, x1 - 42 + 30 + 85 + 65 + 65 + 50 + 50, y1 + 105 + 25, 90, 15);
                    gfx2.DrawString(emd, font7, XBrushes.Black, new XRect(x1 - 42 + 30 + 85 + 65 + 65 + 50 + 50 + 3, y1 + 105 + 29, page.Width.Point, page.Height.Point), XStringFormats.TopLeft);

                    // emd utr no
                    gfx2.DrawRectangle(XPens.Black, x1 - 42 + 30 + 85 + 65 + 65 + 50 + 50 + 90, y1 + 105 + 25, 100, 15);
                    gfx2.DrawString(emdutrno, font7, XBrushes.Black, new XRect(x1 - 42 + 30 + 85 + 65 + 65 + 50 + 50 + 90 + 3, y1 + 105 + 29, page.Width.Point, page.Height.Point), XStringFormats.TopLeft);

                    // emd utr dt
                    gfx2.DrawRectangle(XPens.Black, x1 - 42 + 30 + 85 + 65 + 65 + 50 + 50 + 90 + 100, y1 + 105 + 25, 50, 15);
                    gfx2.DrawString(emddt, font7, XBrushes.Black, new XRect(x1 - 42 + 30 + 85 + 65 + 65 + 50 + 50 + 90 + 100 + 3, y1 + 105 + 29, page.Width.Point, page.Height.Point), XStringFormats.TopLeft);

                    y1 = y1 + 15;

                }
            }
            else if(of=="COST")
            {
                tf2.Alignment = XParagraphAlignment.Justify;
                tf2.DrawString(c_title, font4, XBrushes.Black, new XRect(x1, y1, 500, 120), XStringFormats.TopLeft);



                //SL. NO.
                gfx2.DrawRectangle(XPens.Black, x1 - 42, y1 + 100, 30, 29);
                gfx2.DrawString(c_column1, font7, XBrushes.Black, new XRect(x1 - 42, y1 + 105, page.Width.Point, page.Height.Point), XStringFormats.TopLeft);

                //Party
                gfx2.DrawRectangle(XPens.Black, x1 - 42 + 30, y1 + 100, 85, 29);
                gfx2.DrawString(c_column2, font7, XBrushes.Black, new XRect(x1 - 42 + 30 + 3, y1 + 105, page.Width.Point, page.Height.Point), XStringFormats.TopLeft);

                //Depot
                gfx2.DrawRectangle(XPens.Black, x1 - 42 + 30 + 85, y1 + 100, 65, 29);
                gfx2.DrawString(c_column3, font7, XBrushes.Black, new XRect(x1 - 42 + 30 + 85 + 5, y1 + 105, page.Width.Point, page.Height.Point), XStringFormats.TopLeft);

                //District
                gfx2.DrawRectangle(XPens.Black, x1 - 42 + 30 + 85 + 65, y1 + 100, 65, 29);
                gfx2.DrawString(c_column4, font7, XBrushes.Black, new XRect(x1 - 42 + 30 + 75 + 65 + 12, y1 + 105, page.Width.Point, page.Height.Point), XStringFormats.TopLeft);

                // qty
                gfx2.DrawRectangle(XPens.Black, x1 - 42 + 30 + 85 + 65 + 65, y1 + 100, 50, 29);
                gfx2.DrawString(c_column5, font7, XBrushes.Black, new XRect(x1 - 42 + 30 + 75 + 65 + 65 + 10, y1 + 105, page.Width.Point, page.Height.Point), XStringFormats.TopLeft);

                // rt
                gfx2.DrawRectangle(XPens.Black, x1 - 42 + 30 + 85 + 65 + 65 + 50, y1 + 100, 50, 29);
                gfx2.DrawString(c_column6, font7, XBrushes.Black, new XRect(x1 - 42 + 30 + 75 + 65 + 65 + 50 + 10, y1 + 105, page.Width.Point, page.Height.Point), XStringFormats.TopLeft);

                // cost
                gfx2.DrawRectangle(XPens.Black, x1 - 42 + 30 + 85 + 65 + 65 + 50 + 50, y1 + 100, 70, 29);
                gfx2.DrawString(c_column9, font7, XBrushes.Black, new XRect(x1 - 42 + 30 + 85 + 65 + 65 + 50 + 50 + 10, y1 + 105, page.Width.Point, page.Height.Point), XStringFormats.TopLeft);

                // cost utr no
                gfx2.DrawRectangle(XPens.Black, x1 - 42 + 30 + 85 + 65 + 65 + 50 + 50 + 70, y1 + 100, 100, 29);
                gfx2.DrawString(c_column10, font7, XBrushes.Black, new XRect(x1 - 42 + 30 + 85 + 65 + 65 + 50 + 50 + 70 + 3, y1 + 105, page.Width.Point, page.Height.Point), XStringFormats.TopLeft);

                // cost utr dt
                gfx2.DrawRectangle(XPens.Black, x1 - 42 + 30 + 85 + 65 + 65 + 50 + 50 + 70 + 100, y1 + 100, 50, 29);
                gfx2.DrawString(c_column11, font7, XBrushes.Black, new XRect(x1 - 42 + 30 + 85 + 65 + 65 + 50 + 50 + 70 + 100 + 3, y1 + 105, page.Width.Point, page.Height.Point), XStringFormats.TopLeft);

               /* // cost2
                gfx2.DrawRectangle(XPens.Black, x1 - 42 + 30 + 85 + 65 + 65 + 50 + 50 + 55 + 80 + 50, y1 + 100, 55, 29);
                gfx2.DrawString(c_column12, font7, XBrushes.Black, new XRect(x1 - 42 + 30 + 85 + 65 + 65 + 50 + 50 + 55 + 80 + 50 + 10, y1 + 105, page.Width.Point, page.Height.Point), XStringFormats.TopLeft);

                // cost2 utr no
                gfx2.DrawRectangle(XPens.Black, x1 - 42 + 30 + 85 + 65 + 65 + 50 + 50 + 55 + 80 + 50 + 55, y1 + 100, 80, 29);
                gfx2.DrawString(c_column13, font7, XBrushes.Black, new XRect(x1 - 42 + 30 + 85 + 65 + 65 + 50 + 50 + 65 + 90 + 50 + 55 + 3, y1 + 105, page.Width.Point, page.Height.Point), XStringFormats.TopLeft);

                // cost2 utr dt
                gfx2.DrawRectangle(XPens.Black, x1 - 42 + 30 + 85 + 65 + 65 + 50 + 50 + 65 + 80 + 50 + 80 + 80, y1 + 100, 50, 29);
                gfx2.DrawString(c_column14, font7, XBrushes.Black, new XRect(x1 - 42 + 30 + 85 + 65 + 65 + 50 + 50 + 65 + 80 + 50 + 80 + 80 + 3, y1 + 105, page.Width.Point, page.Height.Point), XStringFormats.TopLeft);
                */
                for (int i = 0; i < emd_cost.Tables[0].Rows.Count; i++)
                {
                    //drawing body of table

                    string sl = (i + 1).ToString();
                    string party = emd_cost.Tables[0].Rows[i].ItemArray[4].ToString();
                    string depot = emd_cost.Tables[0].Rows[i].ItemArray[5].ToString();
                    string dis = emd_cost.Tables[0].Rows[i].ItemArray[6].ToString();
                    string qty = emd_cost.Tables[0].Rows[i].ItemArray[7].ToString();
                    string rate = emd_cost.Tables[0].Rows[i].ItemArray[8].ToString();
                    string cost = emd_cost.Tables[0].Rows[i].ItemArray[15].ToString();
                    string costutrno = emd_cost.Tables[0].Rows[i].ItemArray[17].ToString();
                    string costdt = Convert.ToDateTime(emd_cost.Tables[0].Rows[i].ItemArray[18]).ToShortDateString();
                    string cost2 = emd_cost.Tables[0].Rows[i].ItemArray[16].ToString();
                    string cost2utrno = emd_cost.Tables[0].Rows[i].ItemArray[19].ToString();
                    if(emd_cost.Tables[0].Rows[i].ItemArray[20].ToString()!="")
                    {
                        cost2dt = Convert.ToDateTime(emd_cost.Tables[0].Rows[i].ItemArray[20]).ToShortDateString();
                    }
                    else
                    {
                        cost2dt = emd_cost.Tables[0].Rows[i].ItemArray[20].ToString();
                
                    }

                    //SL. NO.
                    gfx2.DrawRectangle(XPens.Black, x1 - 42, y1 + 105+25, 30, 29);
                    gfx2.DrawString(sl, font7, XBrushes.Black, new XRect(x1 - 42, y1 + 105 + 29, page.Width.Point, page.Height.Point), XStringFormats.TopLeft);

                    //Party
                    gfx2.DrawRectangle(XPens.Black, x1 - 42 + 30, y1 + 105 + 25, 85, 29);
                    gfx2.DrawString(party, font7, XBrushes.Black, new XRect(x1 - 42 + 30 + 3, y1 + 105 + 29, page.Width.Point, page.Height.Point), XStringFormats.TopLeft);

                    //Depot
                    gfx2.DrawRectangle(XPens.Black, x1 - 42 + 30 + 85, y1 + 105 + 25, 65, 29);
                    gfx2.DrawString(depot, font7, XBrushes.Black, new XRect(x1 - 42 + 30 + 85 + 5, y1 + 105 + 29, page.Width.Point, page.Height.Point), XStringFormats.TopLeft);

                    //District
                    gfx2.DrawRectangle(XPens.Black, x1 - 42 + 30 + 85 + 65, y1 + 105 + 25, 65, 29);
                    gfx2.DrawString(dis, font7, XBrushes.Black, new XRect(x1 - 42 + 30 + 75 + 65 + 12, y1 + 105 + 29, page.Width.Point, page.Height.Point), XStringFormats.TopLeft);

                    // qty
                    gfx2.DrawRectangle(XPens.Black, x1 - 42 + 30 + 85 + 65 + 65, y1 + 105 + 25, 50, 29);
                    gfx2.DrawString(qty, font7, XBrushes.Black, new XRect(x1 - 42 + 30 + 75 + 65 + 65 + 10, y1 + 105 + 29, page.Width.Point, page.Height.Point), XStringFormats.TopLeft);

                    // rt
                    gfx2.DrawRectangle(XPens.Black, x1 - 42 + 30 + 85 + 65 + 65 + 50, y1 + 105 + 25, 50, 29);
                    gfx2.DrawString(rate, font7, XBrushes.Black, new XRect(x1 - 42 + 30 + 75 + 65 + 65 + 50 + 10, y1 + 105 + 29, page.Width.Point, page.Height.Point), XStringFormats.TopLeft);

                    // cost
                    gfx2.DrawRectangle(XPens.Black, x1 - 42 + 30 + 85 + 65 + 65 + 50 + 50, y1 + 105 + 25, 70, 29);
                    gfx2.DrawString(cost, font7, XBrushes.Black, new XRect(x1 - 42 + 30 + 85 + 65 + 65 + 50 + 50 + 10, y1 + 105 + 29, page.Width.Point, page.Height.Point), XStringFormats.TopLeft);

                    // cost utr no
                    gfx2.DrawRectangle(XPens.Black, x1 - 42 + 30 + 85 + 65 + 65 + 50 + 50 + 70, y1 + 105 + 25, 100, 29);
                    gfx2.DrawString(costutrno, font7, XBrushes.Black, new XRect(x1 - 42 + 30 + 85 + 65 + 65 + 50 + 50 + 70 + 3, y1 + 105 + 29, page.Width.Point, page.Height.Point), XStringFormats.TopLeft);

                    // cost utr dt
                    gfx2.DrawRectangle(XPens.Black, x1 - 42 + 30 + 85 + 65 + 65 + 50 + 50 + 70 + 100, y1 + 105 + 25, 50, 29);
                    gfx2.DrawString(costdt, font7, XBrushes.Black, new XRect(x1 - 42 + 30 + 85 + 65 + 65 + 50 + 50 + 70 + 100 + 3, y1 + 105 + 29, page.Width.Point, page.Height.Point), XStringFormats.TopLeft);

                 /*   // cost
                    gfx2.DrawRectangle(XPens.Black, x1 - 42 + 30 + 85 + 65 + 65 + 50 + 50, y1 + 105 + 25, 90, 15);
                    gfx2.DrawString(cost2, font7, XBrushes.Black, new XRect(x1 - 42 + 30 + 85 + 65 + 65 + 50 + 50 + 3, y1 + 105 + 29, page.Width.Point, page.Height.Point), XStringFormats.TopLeft);

                    // cost utr no
                    gfx2.DrawRectangle(XPens.Black, x1 - 42 + 30 + 85 + 65 + 65 + 50 + 50 + 90, y1 + 105 + 25, 100, 15);
                    gfx2.DrawString(cost2utrno, font7, XBrushes.Black, new XRect(x1 - 42 + 30 + 85 + 65 + 65 + 50 + 50 + 90 + 3, y1 + 105 + 29, page.Width.Point, page.Height.Point), XStringFormats.TopLeft);

                    // cost utr dt
                    gfx2.DrawRectangle(XPens.Black, x1 - 42 + 30 + 85 + 65 + 65 + 50 + 50 + 90 + 100, y1 + 105 + 25, 50, 15);
                    gfx2.DrawString(cost2dt, font7, XBrushes.Black, new XRect(x1 - 42 + 30 + 85 + 65 + 65 + 50 + 50 + 90 + 100 + 3, y1 + 105 + 29, page.Width.Point, page.Height.Point), XStringFormats.TopLeft);
*/

                    y1 = y1 + 15;

                }
                    
            }
 

                
       
                // Save the document...
                string filename = "ION.pdf";
                document.Save(filename);

                // ...and start a viewer.
                Process.Start(filename);

         }

        #endregion
        #region REFUND ION
        private void btnrefnd_Click(object sender, EventArgs e)
        {
            DataSet refund = new DataSet();
            string str;
            if (cmbrsltnit.SelectedIndex != -1 && cmbrsltcmdty.SelectedIndex != -1 && cmbrsltfinyr.SelectedIndex != -1 && cmbof.SelectedIndex != -1)
            {
                string nitdt = Convert.ToDateTime(cmbrsltnit.SelectedValue).ToLongDateString();
                string cmdty = cmbrsltcmdty.SelectedItem.ToString();
                string finyr = cmbrsltfinyr.SelectedItem.ToString();
                if (cmbof.SelectedItem == "EMD")
                {
                    str = "Select * from Result where NITDate=#" + nitdt + "# and Commodity='" + cmdty + "' and FinYr='" + finyr + "'and Status='Non-Winner'";
                    refund = select_data(str);
                }
                else if (cmbof.SelectedItem == "COST")
                {
                    str = "Select * from Result where NITDate=#" + nitdt + "# and Commodity='" + cmdty + "' and FinYr='" + finyr + "'and RefundAmnt is NOT NULL  and Status='Winner'";
                    refund = select_data(str);
                }
                if (refund.Tables[0].Rows.Count != 0)
                {
                    REFUND_ION_CREATION(refund);
                }
                else
                {
                    MessageBox.Show("No data for these selections", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }


            }
            else
            {
                MessageBox.Show("Enter valid Commodity or NITDate or Financial Year or all ", "Invalid Selection", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        // func for REFUND ION
        private void REFUND_ION_CREATION(DataSet refund)
        {
            string of = cmbof.SelectedItem.ToString();
            string cmdty = cmbrsltcmdty.SelectedItem.ToString();
            string finyr = cmbrsltfinyr.SelectedItem.ToString();   //all the static items in the page hold in variables
            //COMMON
            string refno = "Ref. No.:- " + "Comml/OMSS(D)/Bulk/"+of+" "+"REFUND"+ "/ " + finyr;
            string date = "Dated:" + DateTime.Today.ToShortDateString();
            string nitdate = Convert.ToDateTime(cmbrsltnit.SelectedValue).ToShortDateString();
            string EAdate = Convert.ToDateTime(cmbrsltnit.SelectedValue).AddDays(3).ToShortDateString();

          

            string one_one = " The  amount of money given below deposited by the following parties towards "+ "  "+of+"regarding Sale of"+ " "+cmdty+" under OMSS(D) through E-Auction held on"+ " "+EAdate+" "+",details of which are given here under,may please be refunded as approved by the competetent authority.The Bank particulars of tenderer and the approval of the competent authority are attached herewith for necessary action at your end please.";
         
            

            // Create a new PDF document
            PdfDocument document = new PdfDocument();
            document.Info.Title = "Created with PDFsharp";

            // Create an empty page
            PdfPage page = document.AddPage();

            // Get an XGraphics object for drawing
            XGraphics gfx = XGraphics.FromPdfPage(page);
            XTextFormatter tf = new XTextFormatter(gfx);


            // Draw FCI logo of the page
            string jpegSamplePath = "../IMAGES/logo.jpg";
            XImage image = XImage.FromFile(jpegSamplePath);
            gfx.DrawImage(image, 44, 70, 500, 80);

            // Create a font

            XFont font1 = new XFont("Calibri (Body)", 12, XFontStyle.Regular);
            XFont font2 = new XFont("Calibri (Body)", 12, XFontStyle.Underline);
            XFont font3 = new XFont("Calibri (Body)", 12, XFontStyle.Italic);
            XFont font4 = new XFont("Calibri (Body)", 12, XFontStyle.Bold);
            XFont font5 = new XFont("Calibri (Body)", 10, XFontStyle.Regular);
            XFont font6 = new XFont("Calibri (Body)", 11, XFontStyle.Bold);
            XFont font7 = new XFont("Calibri (Body)", 08, XFontStyle.Bold);
            XFont font8 = new XFont("Calibri (Body)", 08, XFontStyle.Regular);

            int x = 44;
            int y = 30;

            // Draw the text

            // file refno
            // gfx.DrawString(hindi_ref, font1, XBrushes.Black, new XRect(x, y + 130, page.Width, page.Height), XStringFormats.TopLeft);
            gfx.DrawString(refno, font1, XBrushes.Black, new XRect(x, y + 130, page.Width, page.Height), XStringFormats.TopLeft);

            // date 
            // gfx.DrawString(hindi_date, font1, XBrushes.Black, new XRect(x + 390, y + 130, page.Width, page.Height), XStringFormats.TopLeft);
            gfx.DrawString(date, font1, XBrushes.Black, new XRect(x + 390, y + 130, page.Width, page.Height), XStringFormats.TopLeft);

            //ION
            // gfx.DrawString(ion, font4, XBrushes.Black, new XRect(x - 44, y + 155, page.Width, page.Height), XStringFormats.TopCenter);

            string jpegSamplePath1 = "../IMAGES/hindi4.jpg";
            XImage imageion = XImage.FromFile(jpegSamplePath1);
            gfx.DrawImage(imageion, x + 130, y + 165, 250, 15);

            // para1    
            

                tf.Alignment = XParagraphAlignment.Justify;
                tf.DrawString(one_one, font1, XBrushes.Black, new XRect(x, y + 205 + 18, 500, 120), XStringFormats.TopLeft);
               
            


            
            // drawing table       

                int x1 = x + 44;
                int y1 = y + 30;


            //drawing header of table for EMD

            string column1 = "Sl. No.";
            string column2 = "Name of Party ";           
            string column3 = "EMD DEPOSITED (Rs)";
            string column4 = "EMD UTR NO";
            string column5 = "EMD DATE";

            //drawing header of table for COST

            string c_column1 = "Sl. No.";
            string c_column2 = "Name of Party ";
            string c_column3 = "Depot";
            string c_column4 = "F.C.I D.O";
            string c_column5 = "Qty(MT)";
            string c_column6 = "Rate(Rs/Qtl)";
            string c_column7 = "Total Cost";            
            string c_column8 = "Refund";

            if (of == "EMD")
            {
               
                //SL. NO.
                gfx.DrawRectangle(XPens.Black, x, y + 205 + 100, 30, 29);
                gfx.DrawString(column1, font7, XBrushes.Black, new XRect(x, y + 205 + 100+3, page.Width.Point, page.Height.Point), XStringFormats.TopLeft);

                //Party
                gfx.DrawRectangle(XPens.Black, x + 30, y + 205 + 100, 100, 29);
                gfx.DrawString(column2, font7, XBrushes.Black, new XRect(x + 30 + 3, y + 205 + 100 + 3, page.Width.Point, page.Height.Point), XStringFormats.TopLeft);
                
                // emd
                gfx.DrawRectangle(XPens.Black, x + 30 + 100, y + 205 + 100, 105, 29);
                gfx.DrawString(column3, font7, XBrushes.Black, new XRect(x + 30 + 100 + 10, y + 205 + 100 + 3, page.Width.Point, page.Height.Point), XStringFormats.TopLeft);

                // emd utr no
                gfx.DrawRectangle(XPens.Black, x + 30 + 100 + 105, y + 205 + 100, 100, 29);
                gfx.DrawString(column4, font7, XBrushes.Black, new XRect(x + 30 + 100 + 105 + 3, y + 205 + 100 + 3, page.Width.Point, page.Height.Point), XStringFormats.TopLeft);

                // emd utr dt
                gfx.DrawRectangle(XPens.Black, x + 30 + 100 + 105 + 100, y + 205 + 100, 50, 29);
                gfx.DrawString(column5, font7, XBrushes.Black, new XRect(x + 30 + 100 + 105 + 100 + 3, y + 205 + 100 + 3, page.Width.Point, page.Height.Point), XStringFormats.TopLeft);

                for (int i = 0; i < refund.Tables[0].Rows.Count; i++)
                {
                    //drawing body of table

                    string sl = (i + 1).ToString();
                    string party = refund.Tables[0].Rows[i].ItemArray[4].ToString();
                    string emd = refund.Tables[0].Rows[i].ItemArray[9].ToString();
                    string emdutrno = refund.Tables[0].Rows[i].ItemArray[10].ToString();
                    string emddt = Convert.ToDateTime(refund.Tables[0].Rows[i].ItemArray[11]).ToShortDateString();

                    //SL. NO.
                    gfx.DrawRectangle(XPens.Black, x, y + 205 + 100+29, 30, 29);
                    gfx.DrawString(sl, font7, XBrushes.Black, new XRect(x, y + 205 + 100 + 29 + 3, page.Width.Point, page.Height.Point), XStringFormats.TopLeft);

                    //Party
                    gfx.DrawRectangle(XPens.Black, x + 30, y + 205 + 100 + 29, 100, 29);
                    gfx.DrawString(party, font7, XBrushes.Black, new XRect(x + 30 + 3, y + 205 + 100 + 29 + 3, page.Width.Point, page.Height.Point), XStringFormats.TopLeft);

                    // emd
                    gfx.DrawRectangle(XPens.Black, x + 30 + 100, y + 205 + 100 + 29, 105, 29);
                    gfx.DrawString(emd, font7, XBrushes.Black, new XRect(x + 30 + 100 + 10, y + 205 + 100 + 29 + 3, page.Width.Point, page.Height.Point), XStringFormats.TopLeft);

                    // emd utr no
                    gfx.DrawRectangle(XPens.Black, x + 30 + 100 + 105, y + 205 + 100 + 29, 100, 29);
                    gfx.DrawString(emdutrno, font7, XBrushes.Black, new XRect(x + 30 + 100 + 105 + 3, y + 205 + 100 + 29 + 3, page.Width.Point, page.Height.Point), XStringFormats.TopLeft);

                    // emd utr dt
                    gfx.DrawRectangle(XPens.Black, x + 30 + 100 + 105 + 100, y + 205 + 100 + 29, 50, 29);
                    gfx.DrawString(emddt, font7, XBrushes.Black, new XRect(x + 30 + 100 + 105 + 100 + 3, y + 205 + 100 + 29 + 3, page.Width.Point, page.Height.Point), XStringFormats.TopLeft);

                    y = y + 15;

                }
            }
            else if (of == "COST")
            {
                
               //SL. NO.
                gfx.DrawRectangle(XPens.Black, x, y + 205 + 100 , 30, 29);
                gfx.DrawString(c_column1, font7, XBrushes.Black, new XRect(x, y + 205 + 100 +3, page.Width.Point, page.Height.Point), XStringFormats.TopLeft);

                //Party
                gfx.DrawRectangle(XPens.Black, x + 30, y + 205 + 100, 85, 29);
                gfx.DrawString(c_column2, font7, XBrushes.Black, new XRect(x + 30 + 3, y + 205+3 + 100, page.Width.Point, page.Height.Point), XStringFormats.TopLeft);

                //Depot
                gfx.DrawRectangle(XPens.Black, x + 30 + 85, y + 205 + 100, 65, 29);
                gfx.DrawString(c_column3, font7, XBrushes.Black, new XRect(x + 30 + 85 + 5, y + 205 + 3 + 100, page.Width.Point, page.Height.Point), XStringFormats.TopLeft);

                //District
                gfx.DrawRectangle(XPens.Black, x + 30 + 85 + 65, y + 205 + 100, 65, 29);
                gfx.DrawString(c_column4, font7, XBrushes.Black, new XRect(x + 30 + 75 + 65 + 12, y + 205 + 100+3, page.Width.Point, page.Height.Point), XStringFormats.TopLeft);

                // qty
                gfx.DrawRectangle(XPens.Black, x + 30 + 85 + 65 + 65, y + 205 + 100, 50, 29);
                gfx.DrawString(c_column5, font7, XBrushes.Black, new XRect(x + 30 + 75 + 65 + 65 + 10, y + 205 + 100+3, page.Width.Point, page.Height.Point), XStringFormats.TopLeft);

                // rt
                gfx.DrawRectangle(XPens.Black, x + 30 + 85 + 65 + 65 + 50, y + 205 + 100, 50, 29);
                gfx.DrawString(c_column6, font7, XBrushes.Black, new XRect(x + 30 + 75 + 65 + 65 + 50 + 10, y + 205 + 100+3, page.Width.Point, page.Height.Point), XStringFormats.TopLeft);

                // total cost
                gfx.DrawRectangle(XPens.Black, x + 30 + 85 + 65 + 65 + 50 + 50, y + 205 + 100, 100, 29);
                gfx.DrawString(c_column7, font7, XBrushes.Black, new XRect(x + 30 + 85 + 65 + 65 + 50 + 50 + 10, y + 205 + 100+3, page.Width.Point, page.Height.Point), XStringFormats.TopLeft);

                // refund
                gfx.DrawRectangle(XPens.Black, x + 30 + 85 + 65 + 65 + 50 + 50 + 100, y + 205 + 100, 100, 29);
                gfx.DrawString(c_column8, font7, XBrushes.Black, new XRect(x + 30 + 85 + 65 + 65 + 50 + 50 + 100 + 3, y + 205 + 100+3, page.Width.Point, page.Height.Point), XStringFormats.TopLeft);

                
                for (int i = 0; i < refund.Tables[0].Rows.Count; i++)
                {
                    //drawing body of table

                    string sl = (i + 1).ToString();
                    string party = refund.Tables[0].Rows[i].ItemArray[4].ToString();
                    string depot = refund.Tables[0].Rows[i].ItemArray[5].ToString();
                    string dis = refund.Tables[0].Rows[i].ItemArray[6].ToString();
                    string qty = refund.Tables[0].Rows[i].ItemArray[7].ToString();
                    string rate = refund.Tables[0].Rows[i].ItemArray[8].ToString();
                    string totcost = refund.Tables[0].Rows[i].ItemArray[13].ToString();
                    

                    if (refund.Tables[0].Rows[i].ItemArray[21].ToString() != "")
                    {
                        refnd = refund.Tables[0].Rows[i].ItemArray[21].ToString();
                    }
                    else
                    {
                        refnd = " ";
                    }

                    //SL. NO.
                    gfx.DrawRectangle(XPens.Black, x, y + 205 + 100+29, 30, 29);
                    gfx.DrawString(sl, font7, XBrushes.Black, new XRect(x, y + 205 + 100 + 29 + 3, page.Width.Point, page.Height.Point), XStringFormats.TopLeft);

                    //Party
                    gfx.DrawRectangle(XPens.Black, x + 30, y + 205 + 100 + 29, 85, 29);
                    gfx.DrawString(party, font7, XBrushes.Black, new XRect(x + 30 + 3,y + 205 + 100+29+ 3 , page.Width.Point, page.Height.Point), XStringFormats.TopLeft);

                    //Depot
                    gfx.DrawRectangle(XPens.Black, x + 30 + 85, y + 205 + 100+29, 65, 29);
                    gfx.DrawString(depot, font7, XBrushes.Black, new XRect(x + 30 + 85 + 5, y + 205 + 100 + 29+3, page.Width.Point, page.Height.Point), XStringFormats.TopLeft);

                    //District
                    gfx.DrawRectangle(XPens.Black, x + 30 + 85 + 65, y + 205 + 100 + 29, 65, 29);
                    gfx.DrawString(dis, font7, XBrushes.Black, new XRect(x + 30 + 75 + 65 + 12, y + 205 + 100 + 29 + 3, page.Width.Point, page.Height.Point), XStringFormats.TopLeft);

                    // qty
                    gfx.DrawRectangle(XPens.Black, x + 30 + 85 + 65 + 65, y + 205 + 100 + 29, 50, 29);
                    gfx.DrawString(qty, font7, XBrushes.Black, new XRect(x + 30 + 75 + 65 + 65 + 10, y + 205 + 100 + 29 + 3, page.Width.Point, page.Height.Point), XStringFormats.TopLeft);

                    // rt
                    gfx.DrawRectangle(XPens.Black, x + 30 + 85 + 65 + 65 + 50, y + 205 + 100 + 29, 50, 29);
                    gfx.DrawString(rate, font7, XBrushes.Black, new XRect(x + 30 + 75 + 65 + 65 + 50 + 10, y + 205 + 100 + 29 + 3, page.Width.Point, page.Height.Point), XStringFormats.TopLeft);

                    // total cost
                    gfx.DrawRectangle(XPens.Black, x + 30 + 85 + 65 + 65 + 50 + 50, y + 205 + 100 + 29, 100, 29);
                    gfx.DrawString(totcost, font7, XBrushes.Black, new XRect(x + 30 + 85 + 65 + 65 + 50 + 50 + 10, y + 205 + 100 + 29 + 3, page.Width.Point, page.Height.Point), XStringFormats.TopLeft);

                    // refund
                    gfx.DrawRectangle(XPens.Black, x + 30 + 85 + 65 + 65 + 50 + 50 + 100, y + 205 + 100 + 29, 100, 29);
                    gfx.DrawString(refnd, font7, XBrushes.Black, new XRect(x + 30 + 85 + 65 + 65 + 50 + 50 + 100 + 3, y + 205 + 100 + 29 + 3, page.Width.Point, page.Height.Point), XStringFormats.TopLeft);

                
                    
                   

                    y= y + 15;

                }

            }

             int y2=(y + 205 + 150 + 25)+(15 * refund.Tables[0].Rows.Count);
            //para2
            string jpegSamplePath2 = "../IMAGES/pic1.jpg";
            XImage imagepara2 = XImage.FromFile(jpegSamplePath2);
            gfx.DrawImage(imagepara2, x, y2, 500, 150);


            // Save the document...
            string filename = "ION.pdf";
            document.Save(filename);

            // ...and start a viewer.
            Process.Start(filename);

        }
        #endregion
        #endregion

        #region Database
        private void btnadd_Click(object sender, EventArgs e)
        {
            if(cmbdbwhat.SelectedIndex!=-1)
            {
                if(cmbdbwhat.SelectedItem=="Party")
                {
                    pnlentrprty.Visible = true;
                    pnleditprty.Visible = false;
                    pnleditfcidis.Visible = false;
                }
                else if (cmbdbwhat.SelectedItem == "FCI District")
                {
                    pnlentryfcidis.Visible = true;
                    pnleditprty.Visible = false;
                    pnleditfcidis.Visible = false;

                }
            }
        }
       

        private void btnedit_Click(object sender, EventArgs e)
        {
            if (cmbdbwhat.SelectedIndex != -1)
            {
                if (cmbdbwhat.SelectedItem == "Party")
                {
                    pnleditprty.Visible = true;
                    pnlentrprty.Visible = false;                   
                    pnleditfcidis.Visible = false;
                }
                else if (cmbdbwhat.SelectedItem == "FCI District")
                {
                    pnleditfcidis.Visible = true;
                   pnlentryfcidis.Visible = false;
                   pnleditprty.Visible = false;
                    
                    
                }
            }
        }
             

        private void btndbsub_Click(object sender, EventArgs e)
        {
            string partyname = txtprtynm.Text;
            string partyad1 = txtaddrs1.Text;
            string partyad2 = txtaddrss2.Text;
            string partydis= txtdis.Text;
            string partystate = txtstate.Text;
            string partypin = txtpin.Text;
            string partyemail = txtemail.Text;
            string partycntct = txtprtycntct.Text;
            string empnldt = dttmempnldt.Value.ToLongDateString();
            string finyr = cmbdbfinyr.SelectedItem.ToString();

            string qry_insert = "insert into Party(Party_Name,Party_Ad1,Party_Ad2,Party_Dis,Party_State,Party_Pin,Party_Email,Party_Contact,Party_Finyr,Party_Empnldt) values('" + partyname + "','" + partyad1 + "','" + partyad2 + "','" + partydis + "','" + partystate + "','" + partypin + "','" + partyemail + "','" + partycntct + "','" + finyr + "','" + empnldt + "')";
            insert_update_deleted(qry_insert);

        }       

        private void btnfcidissub_Click(object sender, EventArgs e)
        {
            string fcidis = txtfcidis.Text;
            string fcidep = txtdepot.Text;
            string qry_insert = "insert into District_Depot(District,Depot) values('" + fcidis + "','" + fcidep + "')";
            insert_update_deleted(qry_insert);
        }
       

        private void btnedtdissrch_Click(object sender, EventArgs e)
        {
            DataGridViewButtonColumn editbtn = new DataGridViewButtonColumn();
            editbtn.Name = "EDIT";
            editbtn.HeaderText = "EDIT";
            editbtn.Text = "EDIT";
            editbtn.UseColumnTextForButtonValue = true;
            editbtn.FlatStyle = FlatStyle.Popup;

            DataGridViewButtonColumn dltbtn = new DataGridViewButtonColumn();
            dltbtn.Name = "DELETE";
            dltbtn.HeaderText = "DELETE";
            dltbtn.Text = "DELETE";
            dltbtn.UseColumnTextForButtonValue = true;
            dltbtn.FlatStyle = FlatStyle.Popup;

           
            DataSet ds = new DataSet();
          
            if (cmbfcidis1.SelectedIndex != -1 || cmbfcidep1.SelectedIndex != -1)
            {
                string qry = "Select * from District_Depot where District='" + cmbfcidis1.SelectedValue + "' or Depot='" + cmbfcidep1.SelectedValue + "'";
                ds = select_data(qry);
            }
            
            
            if (ds.Tables[0].Rows.Count > 0)
            {
                dgvedtfcidis.Columns.Clear();
                dgvedtfcidis.Visible = true;                
                //dgvedtfcidis.AutoGenerateColumns = true;
                dgvedtfcidis.DataSource = ds.Tables[0];
                dgvedtfcidis.Columns["ID"].ReadOnly = true;
                dgvedtfcidis.Columns.Add(editbtn);
                dgvedtfcidis.Columns.Add(dltbtn);
                dgvedtfcidis.Refresh();           
                         

            }
            
            else
            {
                MessageBox.Show("No data");
            }
         
        }

       

        private void dgvedtfcidis_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            // Check if wanted column was clicked 
            if (e.ColumnIndex == 3 && e.RowIndex >= 0)
            {
                //Perform on edit button click code   
             //  Int64 id =Convert.ToInt64 (dgvedtfcidis.CurrentRow.Cells[0].Value);
               int id = Int32.Parse(dgvedtfcidis.CurrentRow.Cells[0].Value.ToString());
                string district = dgvedtfcidis.CurrentRow.Cells[1].Value.ToString();
                string depot = dgvedtfcidis.CurrentRow.Cells[2].Value.ToString();

                if(cmbfcidis1.SelectedIndex!=-1)
                {
                   // MessageBox.Show(id.ToString());
                    string qry_updt = "update District_Depot set District='" + district + "',Depot='" + depot + "' WHERE ID="+id+"";
                insert_update_deleted(qry_updt);
                }

            }

            else  if (e.ColumnIndex == 4 && e.RowIndex >= 0)
            {
                //Perform on edit button click code          

                string district = dgvedtfcidis.CurrentRow.Cells[1].Value.ToString();
                string depot = dgvedtfcidis.CurrentRow.Cells[2].Value.ToString();
                string qry_updt = "delete from District_Depot where District='" + district + "'and Depot='" + depot + "' ";
                insert_update_deleted(qry_updt);
                dgvedtfcidis.Rows.RemoveAt(dgvedtfcidis.CurrentRow.Index);


            }   
        }

       

        
        private void Database_Click(object sender, EventArgs e)
        {
            
            DataSet ds1 = new DataSet();
            DataSet ds2 = new DataSet();
            DataSet ds3 = new DataSet();

            string qry3 = "Select  Party_Name from Party order by Party_Name";
            ds3 = select_data(qry3);
            if (ds3.Tables[0].Rows.Count > 0)
            {
                cmbedtprty.Refresh();
               // cmbedtprty.Items.Clear();
                cmbedtprty.DataSource = ds3.Tables[0];
                cmbedtprty.ValueMember = "Party_Name";
                cmbedtprty.DisplayMember = "Party_Name";
                cmbedtprty.Text = "-Select-";

            }

            else
            {
                cmbedtprty.Refresh();
               // cmbedtprty.Items.Clear();
               // cmbedtprty.SelectedText  = "-Select-";
            }
             string qry2 = "Select  distinct Depot from District_Depot order by Depot";
             ds2 = select_data(qry2);
             if (ds2.Tables[0].Rows.Count > 0)
             {
                 cmbfcidep1.Refresh();
                 //cmbfcidep1.Items.Clear();
                 cmbfcidep1.DataSource = ds2.Tables[0];
                 cmbfcidep1.ValueMember = "Depot";
                 cmbfcidep1.DisplayMember = "Depot";
               cmbfcidep1.Text = "-Select-";

             }
             else
             {
                 cmbfcidep1.Refresh();
               //  cmbfcidep1.Items.Clear();
                // cmbfcidep1.SelectedText = "-Select-";
             }
            string qry = "Select distinct District from District_Depot order by District";
            ds1 = select_data(qry);
            

            if (ds1.Tables[0].Rows.Count > 0)
            {
                cmbfcidis1.Refresh();
               // cmbfcidis1.Items.Clear();
                cmbfcidis1.DataSource = ds1.Tables[0];
                cmbfcidis1.ValueMember = "District";
                cmbfcidis1.DisplayMember = "District";
                cmbfcidis1.Text = "-Select-";
            }
            else
            {
               cmbfcidis1.Refresh();
                //cmbfcidis1.Items.Clear();
               // cmbfcidis1.SelectedText = "-Select-";
            }
            
        }



        

        private void dgvedtprty_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            // Check if wanted column was clicked 
            if (e.ColumnIndex == 10 && e.RowIndex >= 0)
            {
                //Perform on edit button click code  

                
                string prtynm = dgvedtprty.CurrentRow.Cells[0].Value.ToString();
                string prtyad1 = dgvedtprty.CurrentRow.Cells[1].Value.ToString();
                string prtyad2 = dgvedtprty.CurrentRow.Cells[2].Value.ToString();
                string prtydis = dgvedtprty.CurrentRow.Cells[3].Value.ToString();
                string prtystate = dgvedtprty.CurrentRow.Cells[4].Value.ToString();
                string prtypin = dgvedtprty.CurrentRow.Cells[5].Value.ToString();
                string prtyemail = dgvedtprty.CurrentRow.Cells[6].Value.ToString();
                string prtycontct = dgvedtprty.CurrentRow.Cells[7].Value.ToString();
                string prtyfinyr = dgvedtprty.CurrentRow.Cells[8].Value.ToString();
                string prtyempnl = dgvedtprty.CurrentRow.Cells[9].Value.ToString();

                if (cmbedtprty.SelectedIndex != -1)
                {

                    string qry_updt = "update Party set Party_name='" + prtynm + "',Party_Ad1='" + prtyad1 + "',Party_Ad2='" + prtyad2 + "',Party_Dis='" + prtydis + "',Party_State='" + prtystate + "',Party_Pin='" + prtypin + "',Party_Email='" + prtyemail + "',Party_Contact='" + prtycontct + "',Party_Finyr='" + prtyfinyr + "',Party_Empnldt='" + prtyempnl + "' where Party_Name='" + cmbedtprty.SelectedValue + "'";
                    insert_update_deleted(qry_updt);
                }

            }
            else if (e.ColumnIndex == 11 && e.RowIndex >= 0)
            {
                //Perform on deletebutton click code  
                string prtynm = dgvedtprty.CurrentRow.Cells[0].Value.ToString();
              
                string qry_updt = "delete from Party where Party_Name='" + prtynm + "'";
                insert_update_deleted(qry_updt);
                dgvedtprty.Rows.RemoveAt(dgvedtprty.CurrentRow.Index);

            }   
        }

        private void btndbsrch_Click(object sender, EventArgs e)
        {
            DataGridViewButtonColumn editbtn = new DataGridViewButtonColumn();
            editbtn.Name = "EDIT";
            editbtn.HeaderText = "EDIT";
            editbtn.Text = "EDIT";
            editbtn.UseColumnTextForButtonValue = true;
            editbtn.FlatStyle = FlatStyle.Popup;

            DataGridViewButtonColumn dltbtn = new DataGridViewButtonColumn();
            dltbtn.Name = "DELETE";
            dltbtn.HeaderText = "DELETE";
            dltbtn.Text = "DELETE";
            dltbtn.UseColumnTextForButtonValue = true;
            dltbtn.FlatStyle = FlatStyle.Popup;
           

           

            DataSet ds = new DataSet();
            string qry = "Select Party_Name,Party_Ad1,Party_Ad2,Party_Dis,Party_State,Party_Pin,Party_Email,Party_Contact,Party_Finyr,Party_Empnldt from Party where Party_Name='" + cmbedtprty.SelectedValue + "' ";
            ds = select_data(qry);
            if (ds.Tables[0].Rows.Count > 0)
            {
                dgvedtprty.Columns.Clear();
                dgvedtprty.Visible = true;
                //dgvedtprty.AutoGenerateColumns = true;
                dgvedtprty.DataSource = ds.Tables[0];
               // dgvedtprty.Columns["Party_Name"].ReadOnly = true;
                dgvedtprty.Columns.Add(editbtn);
                dgvedtprty.Columns.Add(dltbtn);
                dgvedtprty.Refresh();
                
            }

            else
            {
                MessageBox.Show("No data");
            }
            
        }









        #endregion
        # region Reports
        private void cmbrprtkind_SelectedIndexChanged(object sender, EventArgs e)
        {
            dgvrprtsrch.Visible = false;
            dgvrprtsrch2.Visible = false;
            //string finyr=cmbrprtfinyr.SelectedItem.ToString();
            DataSet ds = new DataSet();
            if(cmbrprtkind.SelectedIndex!=-1)
            {
                if(cmbrprtkind.SelectedIndex ==0)
                {
                    lblrprtparty.Visible = true;
                    cmbrprtparty.Visible = true;
                    lblrprtdis.Visible = false;
                    cmbrprtdis.Visible = false;
                    lblrprtdep.Visible = false;
                    cmbrprtdep.Visible = false;
                  /* string qry = "Select  Party_Name from Party where Party_Finyr='"+finyr+"' order by Party_Name";
                    ds = select_data(qry);
                    if (ds.Tables[0].Rows.Count > 0)
                    {
                        cmbrprtparty.Refresh();
                        // cmbrprtparty.Items.Clear();
                        cmbrprtparty.DataSource = ds.Tables[0];
                        cmbrprtparty.ValueMember = "Party_Name";
                        cmbrprtparty.DisplayMember = "Party_Name";
                        cmbrprtparty.Text = "-Select-";

                    }

                    else
                    {
                        cmbrprtparty.Refresh();
                        MessageBox.Show("No parties exist for this financial year");
                        // cmbrprtparty.Items.Clear();
                        // cmbrprtparty.SelectedText  = "-Select-";
                    }*/
                }
                else if (cmbrprtkind.SelectedIndex == 1)
                {
                    lblrprtdis.Visible = true;
                    cmbrprtdis.Visible = true;
                    lblrprtparty.Visible = false;
                    cmbrprtparty.Visible = false;                   
                    lblrprtdep.Visible = false;
                    cmbrprtdep.Visible = false;

                  string qry = "Select distinct District from District_Depot order by District";

                    ds = select_data(qry);
                    if (ds.Tables[0].Rows.Count > 0)
                    {
                        cmbrprtdis.Refresh();
                        // cmbrprtdis.Items.Clear();
                        cmbrprtdis.DataSource = ds.Tables[0];
                        cmbrprtdis.ValueMember = "District";
                        cmbrprtdis.DisplayMember = "District";
                        cmbrprtdis.Text = "-Select-";

                    }

                    else
                    {
                        cmbrprtdis.Refresh();

                        // cmbrprtdis.Items.Clear();
                        // cmbrprtdis.SelectedText  = "-Select-";
                    }
                   
                }
                else  if (cmbrprtkind.SelectedIndex == 2)
                {
                    lblrprtdep.Visible = true;
                    cmbrprtdep.Visible = true;
                    lblrprtparty.Visible = false;
                    cmbrprtparty.Visible = false;
                    lblrprtdis.Visible = false;
                    cmbrprtdis.Visible = false;

                    string qry = "Select  distinct Depot from District_Depot order by Depot";
                     ds = select_data(qry);
                    if (ds.Tables[0].Rows.Count > 0)
                    {
                        cmbrprtdep.Refresh();
                        // cmbrprtdep.Items.Clear();
                        cmbrprtdep.DataSource = ds.Tables[0];
                        cmbrprtdep.ValueMember = "Depot";
                        cmbrprtdep.DisplayMember = "Depot";
                        cmbrprtdep.Text = "-Select-";

                    }

                    else
                    {
                        cmbrprtdep.Refresh();
                       
                        // cmbrprtdep.Items.Clear();
                        // cmbrprtdep.SelectedText  = "-Select-";
                    }
                   
                }            
                
                    
                 
                  
                   
                    
                
            }
        }
        
        private void btnrprtsrch_Click(object sender, EventArgs e)
        {
            string cmdty = cmbrprtcmdty.SelectedItem.ToString();
            string finyr = cmbrprtfinyr.SelectedItem.ToString();
            DataSet ds = new DataSet();
            DataSet ds1 = new DataSet();
            string qry, qry1;
            if(cmbrprtkind.SelectedIndex!=-1)
            {
                if(cmbrprtkind.SelectedIndex==0)
                {
                    if(cmbrprtparty.SelectedIndex!=-1) 
                    {
                    string prty=cmbrprtparty.SelectedValue.ToString();
                    qry = "Select FinYr,Commodity,NITDate,Party,Depot,District,Qty from Result where Commodity='" + cmdty + "' and FinYr='" + finyr + "' and Party='" + prty + "' and Status='Winner'";
                    qry1 = "Select Depot,sum(Qty) As Qty_Per_Depot from Result where Commodity='" + cmdty + "' and FinYr='" + finyr + "' and Party='" + prty + "'and Status='Winner' group by Depot";
                    
                     ds = select_data(qry);
                     ds1 = select_data(qry1);
                        ExportDataSetToExcel(ds);
                     if (ds.Tables[0].Rows.Count > 0 && ds1.Tables[0].Rows.Count > 0)
                    {
                        dgvrprtsrch.Columns.Clear();
                        dgvrprtsrch.Visible = true;
                        //dgvrprtsrch.AutoGenerateColumns = true;
                        dgvrprtsrch.DataSource = ds.Tables[0];
                        dgvrprtsrch.Columns["FinYr"].ReadOnly = true;
                        dgvrprtsrch.Columns["Commodity"].ReadOnly = true;
                        dgvrprtsrch.Columns["NITDate"].ReadOnly = true;
                        dgvrprtsrch.Columns["Party"].ReadOnly = true;
                        dgvrprtsrch.Columns["Depot"].ReadOnly = true;
                        dgvrprtsrch.Columns["District"].ReadOnly = true;
                        dgvrprtsrch.Columns["Qty"].ReadOnly = true;
                        dgvrprtsrch.Refresh();

                        dgvrprtsrch2.Columns.Clear();
                        dgvrprtsrch2.Visible = true;
                        //dgvrprtsrch2.AutoGenerateColumns = true;
                        dgvrprtsrch2.DataSource = ds1.Tables[0];
                        dgvrprtsrch2.Columns[0].ReadOnly = true;
                        dgvrprtsrch2.Columns[1].ReadOnly = true;
                        dgvrprtsrch2.Refresh();
                    }

                    else
                    {
                        dgvrprtsrch.Visible = false;
                        dgvrprtsrch2.Visible = false;
                        MessageBox.Show("No data");
                    }

                     }
                    else
                    {
                        dgvrprtsrch.Visible = false;
                        dgvrprtsrch2.Visible = false;
                        MessageBox.Show("Select Valid Party.");
                    }
                }
                else if(cmbrprtkind.SelectedIndex==2)
                {
                    if (cmbrprtdep.SelectedIndex != -1)
                    {
                        string depot = cmbrprtdep.SelectedValue.ToString();
                        qry = "Select FinYr,Commodity,NITDate,Party,Depot,District,Qty from Result where Commodity='" + cmdty + "' and FinYr='" + finyr + "' and Depot='" + depot + "'and Status='Winner'";
                        qry1 = "Select Party,sum(Qty) As Qty_Per_Party from Result where Commodity='" + cmdty + "' and FinYr='" + finyr + "' and Depot='" + depot + "'and Status='Winner' group by Party";

                        ds = select_data(qry);
                        ds1 = select_data(qry1);

                        if (ds.Tables[0].Rows.Count > 0 && ds1.Tables[0].Rows.Count > 0)
                        {
                            dgvrprtsrch.Columns.Clear();
                            dgvrprtsrch.Visible = true;
                            //dgvrprtsrch.AutoGenerateColumns = true;
                            dgvrprtsrch.DataSource = ds.Tables[0];
                            dgvrprtsrch.Columns["FinYr"].ReadOnly = true;
                            dgvrprtsrch.Columns["Commodity"].ReadOnly = true;
                            dgvrprtsrch.Columns["NITDate"].ReadOnly = true;
                            dgvrprtsrch.Columns["Party"].ReadOnly = true;
                            dgvrprtsrch.Columns["Depot"].ReadOnly = true;
                            dgvrprtsrch.Columns["District"].ReadOnly = true;
                            dgvrprtsrch.Columns["Qty"].ReadOnly = true;
                            dgvrprtsrch.Refresh();

                            dgvrprtsrch2.Columns.Clear();
                            dgvrprtsrch2.Visible = true;
                            //dgvrprtsrch2.AutoGenerateColumns = true;
                            dgvrprtsrch2.DataSource = ds1.Tables[0];
                            dgvrprtsrch2.Columns[0].ReadOnly = true;
                            dgvrprtsrch2.Columns[1].ReadOnly = true;
                            dgvrprtsrch2.Refresh();
                        }

                        else
                        {
                            dgvrprtsrch.Visible = false;
                            dgvrprtsrch2.Visible = false;
                            MessageBox.Show("No data");
                        }

                    }
                    else
                    {
                        dgvrprtsrch.Visible = false;
                        dgvrprtsrch2.Visible = false;
                        MessageBox.Show("Select Valid Depot");
                    }
                }
                else if (cmbrprtkind.SelectedIndex == 1)
                {
                    if (cmbrprtdis.SelectedIndex != -1)
                    {
                        string district = cmbrprtdis.SelectedValue.ToString();
                        qry = "Select FinYr,Commodity,NITDate,Party,Depot,District,Qty from Result where Commodity='" + cmdty + "' and FinYr='" + finyr + "' and District='" + district + "'and Status='Winner'";
                        qry1 = "Select Party,sum(Qty) As Qty_Per_Party from Result where Commodity='" + cmdty + "' and FinYr='" + finyr + "' and District='" + district + "' and Status='Winner'group by Party";

                        ds = select_data(qry);
                        ds1 = select_data(qry1);

                        if (ds.Tables[0].Rows.Count > 0 && ds1.Tables[0].Rows.Count > 0)
                        {
                            dgvrprtsrch.Columns.Clear();
                            dgvrprtsrch.Visible = true;
                            //dgvrprtsrch.AutoGenerateColumns = true;
                            dgvrprtsrch.DataSource = ds.Tables[0];
                            dgvrprtsrch.Columns["FinYr"].ReadOnly = true;
                            dgvrprtsrch.Columns["Commodity"].ReadOnly = true;
                            dgvrprtsrch.Columns["NITDate"].ReadOnly = true;
                            dgvrprtsrch.Columns["Party"].ReadOnly = true;
                            dgvrprtsrch.Columns["Depot"].ReadOnly = true;
                            dgvrprtsrch.Columns["District"].ReadOnly = true;
                            dgvrprtsrch.Columns["Qty"].ReadOnly = true;
                            dgvrprtsrch.Refresh();

                            dgvrprtsrch2.Columns.Clear();
                            dgvrprtsrch2.Visible = true;
                            //dgvrprtsrch2.AutoGenerateColumns = true;
                            dgvrprtsrch2.DataSource = ds1.Tables[0];
                            dgvrprtsrch2.Columns[0].ReadOnly = true;
                            dgvrprtsrch2.Columns[1].ReadOnly = true;
                            dgvrprtsrch2.Refresh();
                        }

                        else
                        {
                            dgvrprtsrch.Visible = false;
                            dgvrprtsrch2.Visible = false;
                            MessageBox.Show("No data");
                        }

                    }
                    else
                    {
                        dgvrprtsrch.Visible = false;
                        dgvrprtsrch2.Visible = false;
                        MessageBox.Show("Select Valid District");
                    }
                }
            }
            else
            {
                dgvrprtsrch.Visible = false;
                dgvrprtsrch2.Visible = false;
                MessageBox.Show("Select the kind of report to view!");
            }

        }

        private void cmbrprtfinyr_SelectedIndexChanged(object sender, EventArgs e)
        {
            string finyr=cmbrprtfinyr.SelectedItem.ToString();
            DataSet ds = new DataSet();

            string qry = "Select  Party_Name from Party where Party_Finyr='"+finyr+"' order by Party_Name";
                   ds = select_data(qry);
                   if (ds.Tables[0].Rows.Count > 0)
                   {
                       cmbrprtparty.Refresh();
                       // cmbrprtparty.Items.Clear();
                       cmbrprtparty.DataSource = ds.Tables[0];
                       cmbrprtparty.ValueMember = "Party_Name";
                       cmbrprtparty.DisplayMember = "Party_Name";
                       cmbrprtparty.Text = "-Select-";

                   }

                   else
                   {
                       cmbrprtparty.Refresh();
                       MessageBox.Show("No parties exist for this financial year");
                       // cmbrprtparty.Items.Clear();
                       // cmbrprtparty.SelectedText  = "-Select-";
                   }
        }
        private void btnrprtsexcel_Click(object sender, EventArgs e)
        {

        }
        private void ExportDataSetToExcel(DataSet ds)
        {
            //Creae an Excel application instance
            Excel.Application excelApp = new Excel.Application();

            //Create an Excel workbook instance and open it from the predefined location
            Excel.Workbook excelWorkBook = excelApp.Workbooks.Open(@"G:\db1.xlsx");

            foreach (DataTable table in ds.Tables)
            {
                //Add a new worksheet to workbook with the Datatable name
                Excel.Worksheet excelWorkSheet = excelWorkBook.Sheets.Add();
                excelWorkSheet.Name = table.TableName;

                for (int i = 1; i < table.Columns.Count + 1; i++)
                {
                    excelWorkSheet.Cells[1, i] = table.Columns[i - 1].ColumnName;
                }

                for (int j = 0; j < table.Rows.Count; j++)
                {
                    for (int k = 0; k < table.Columns.Count; k++)
                    {
                        excelWorkSheet.Cells[j + 2, k + 1] = table.Rows[j].ItemArray[k].ToString();
                    }
                }
            }

            excelWorkBook.Save();
            excelWorkBook.Close();
            excelApp.Quit();

        }
        # endregion

       

        










        // serch for qty block of approval, now deleted

      /*  private void cmbdis_SelectedIndexChanged(object sender, EventArgs e)
        {
            
            if(cmbdis.SelectedIndex==0)
            {                           
               
                cmbdep.Enabled = false;
                cmbdep.Items.Clear();
            }
            else if(cmbdis.SelectedItem.ToString()=="24 PARGNAS")
            {
                
                cmbdep.Enabled = true;
                cmbdep.Items.Clear();               
                cmbdep.Items.Add("ASHOKNAGAR");
                cmbdep.Items.Add("BARASAT");
                cmbdep.Items.Add("DOHARIA");
            }
            else if(cmbdis.SelectedItem.ToString()=="BANKURA")
            {
                cmbdep.Enabled = true;
                cmbdep.Items.Clear();
                cmbdis.Items.RemoveAt(cmbdis.SelectedIndex = -1);
                cmbdep.Items.Add("ADRA");
                cmbdep.Items.Add("BIKNA");
                
            }
            else if (cmbdis.SelectedItem.ToString() == "BIRBHUM")
            {
                cmbdis.Text = "";
                cmbdis.SelectedIndex = -1;
                cmbdis.SelectedItem = null;

                cmbdep.Enabled = true;
                cmbdep.Items.Clear();
                cmbdep.Items.Add("ABDARPUR");
               

            }
            else if (cmbdis.SelectedItem.ToString() == "BURDWAN")
            {
                cmbdep.Enabled = true;
                cmbdep.Items.Clear();
                cmbdep.Items.Add("ALAMGANJ");
                

            }
            else if (cmbdis.SelectedItem.ToString() == "CALCUTTA NPD")
            {
                cmbdep.Enabled = true;
                cmbdep.Items.Clear();
                cmbdep.Items.Add("KALYANI");
                cmbdep.Items.Add("OJM BUDGE BUDGE");
            }

            else if (cmbdis.SelectedItem.ToString() == "CALCUTTA PORT")
            {
                cmbdep.Enabled = true;
                cmbdep.Items.Clear();
                cmbdep.Items.Add("J J P");               
            }
            else if (cmbdis.SelectedItem.ToString() == "COOCH BIHAR")
            {
                cmbdep.Enabled = true;
                cmbdep.Items.Clear();
                cmbdep.Items.Add("BABURHAT");
                cmbdep.Items.Add("DINHATA");
                cmbdep.Items.Add("KHAGRABARI");
            }
            else if (cmbdis.SelectedItem.ToString() == "DURGAPUR")
            {
                cmbdep.Enabled = true;
                cmbdep.Items.Clear();
                cmbdep.Items.Add("DURGAPUR");
                cmbdep.Items.Add("SITARAMPUR");
            }
            else if (cmbdis.SelectedItem.ToString() == "HOOGHLY")
            {
                cmbdep.Enabled = true;
                cmbdep.Items.Clear();
                cmbdep.Items.Add("CHINSURAH");
                cmbdep.Items.Add("DANKUNI");
                cmbdep.Items.Add("SERAMPORE");
                cmbdep.Items.Add("SILO BANDEL (ADANI)");
            }
            else if (cmbdis.SelectedItem.ToString() == "MALDA")
            {
                cmbdep.Enabled = true;
                cmbdep.Items.Clear();
                cmbdep.Items.Add("MANGALBARI");
            }

            else if (cmbdis.SelectedItem.ToString() == "MIDNAPUR")
            {
                cmbdep.Enabled = true;
                cmbdep.Items.Clear();
                cmbdep.Items.Add("MIDNAPUR");
                cmbdep.Items.Add("NIMPURA");
            }
            else if (cmbdis.SelectedItem.ToString() == "MURSHIDABAD")
            {
                cmbdep.Enabled = true;
                cmbdep.Items.Clear();
                cmbdep.Items.Add("BERHAMPUR");                
            }
            else if (cmbdis.SelectedItem.ToString() == "NADIA")
            {
                cmbdep.Enabled = true;
                cmbdep.Items.Clear();
                cmbdep.Items.Add("BHATJUNGLA");
                cmbdep.Items.Add("KALIRHAT");
            }
            else if (cmbdis.SelectedItem.ToString() == "PURULIA")
            {
                cmbdep.Enabled = true;
                cmbdep.Items.Clear();
                cmbdep.Items.Add("BALRAMPUR");
                cmbdep.Items.Add("CHARRAH");

            }
            else if (cmbdis.SelectedItem.ToString() == "SIKKIM(UT)")
            {
                cmbdep.Enabled = true;
                cmbdep.Items.Clear();
                cmbdep.Items.Add("JORTHANG");
                cmbdep.Items.Add("RANGPOO");
            }
            else if (cmbdis.SelectedItem.ToString() == "SILIGURI")
            {
                cmbdep.Enabled = true;
                cmbdep.Items.Clear();
                cmbdep.Items.Add("CSD DABGRAM");                
            }
            else if (cmbdis.SelectedItem.ToString() == "WEST DINAJPUR")
            {
                cmbdep.Enabled = true;
                cmbdep.Items.Clear();
                cmbdep.Items.Add("BUNIADPUR");
                cmbdep.Items.Add("RAIGANJ");
            }            
        }*/
       

       


        

        














    }
}
