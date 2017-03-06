using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Web;
using System.IO;
using System.Configuration;
using System.Data.SqlClient;
using System.Data.SqlTypes;
using System.Reflection;
using Excel = Microsoft.Office.Interop.Excel;


namespace WindowsFormsApplication1
{
    public partial class NorthernWhite_PQR : Form
    {
     //   string ConnectionString = "Data Source=DEEPTHI-PC\\DEEPTHI;Initial Catalog=db_NorthernWhite;User id=sa;Password=$Test123";
        string ConnectionString = "Data Source=WN7-JB956BS\\DEEPTHI;Initial Catalog=db_NorthernWhite ;User id=sa;Password=$Dell2011";

        public NorthernWhite_PQR()
        {
            InitializeComponent();
        }

        private void NorthernWhite_PQR_Load(object sender, EventArgs e)
        {

            SqlConnection sc = new SqlConnection();
            sc.ConnectionString = ConnectionString;

            DataTable tableDust = new DataTable();

            using (SqlDataAdapter daDust = new SqlDataAdapter(@"SELECT DustID,DustType from tbvDust", sc))
                daDust.Fill(tableDust);
            cmbDustData.DataSource = new BindingSource(tableDust, null);
            cmbDustData.DisplayMember = "DustType"; //colum you want to show in comboBox
            cmbDustData.ValueMember = "DustID";

            DataTable tableStatic = new DataTable();
            using (SqlDataAdapter daStatic = new SqlDataAdapter(@"SELECT StaticID,Description from tbvStatic", sc))

                daStatic.Fill(tableStatic);
            cmbStaticData.DataSource = new BindingSource(tableStatic, null);
            cmbStaticData.DisplayMember = "Description"; //colum you want to show in comboBox
            cmbStaticData.ValueMember = "StaticID";
            //cmbClumps
            DataTable tableClumps = new DataTable();
            using (SqlDataAdapter daClumps = new SqlDataAdapter(@"SELECT ClumpsID,Description from Clumps", sc))

                daClumps.Fill(tableClumps);
            cmbClumps.DataSource = new BindingSource(tableClumps, null);
            cmbClumps.DisplayMember = "Description"; //colum you want to show in comboBox
            cmbClumps.ValueMember = "ClumpsID";

            DataTable tableFMYesNo = new DataTable();
            using (SqlDataAdapter daFMYesNo = new SqlDataAdapter(@"SELECT YesNo,_Description from ForeignMaterials", sc))

                daFMYesNo.Fill(tableFMYesNo);
            cmbFMYesNo.DataSource = new BindingSource(tableFMYesNo, null);
            cmbFMYesNo.DisplayMember = "_Description"; //colum you want to show in comboBox
            cmbFMYesNo.ValueMember = "YesNo";

            //cmbFMat
            DataTable tableFMat = new DataTable();
            using (SqlDataAdapter daFMat = new SqlDataAdapter(@"SELECT MaterialID,Material from tbvMaterials", sc))
                daFMat.Fill(tableFMat);
            cmbFMat.DataSource = new BindingSource(tableFMat, null);
            cmbFMat.DisplayMember = "Material"; //colum you want to show in comboBox
            cmbFMat.ValueMember = "MaterialID";

            //cmbProdRate
            DataTable tableProdRate = new DataTable();
            using (SqlDataAdapter daProdRate = new SqlDataAdapter(@"SELECT RateID,Description from ProductRate", sc))

                daProdRate.Fill(tableProdRate);
            cmbProdRate.DataSource = new BindingSource(tableProdRate, null);
            cmbProdRate.DisplayMember = "Description"; //colum you want to show in comboBox
            cmbProdRate.ValueMember = "RateID";


            //cmbAuthority
            DataTable tableAuthority = new DataTable();
            using (SqlDataAdapter daAuthority = new SqlDataAdapter(@"SELECT AuthorityID,Authority_Name from PassingAuthority", sc))

                daAuthority.Fill(tableAuthority);
            cmbAuthority.DataSource = new BindingSource(tableAuthority, null);
            cmbAuthority.DisplayMember = "Authority_Name"; //colum you want to show in comboBox
            cmbAuthority.ValueMember = "AuthorityID";

            // Dropdowns for Shift and Clumps
            //comboBox1.Items.Add("1");
            //comboBox1.Items.Add("2");
            //comboBox1.Items.Add("3");
            //comboBox1.Items.Add("4");
            //comboBox6.Items.Add("0");
            //comboBox6.Items.Add("1");
            //comboBox6.Items.Add("2");
            //comboBox6.Items.Add("Over 2");
            //// Dropdowns for Static and Dust
            //comboBox4.Items.Add("Low");
            //comboBox4.Items.Add("Medium");
            // Dropdown for Passing Authority
            //comboBox2.Items.Add("Srikanth R");
            //comboBox2.Items.Add("Mike L");
            //comboBox2.Items.Add("Rick Young");
            //comboBox2.Items.Add("Supervisor");
            //comboBox2.Items.Add("Process Operator");
            //comboBox2.Items.Add("Others");
            //comboBox4.Items.Add("High");
            //comboBox5.Items.Add("Low");
            //comboBox5.Items.Add("Medium");
            //comboBox5.Items.Add("High");
            //comboBox7.Items.Add("Yes");
            //comboBox7.Items.Add("No");
            //comboBox8.Items.Add("Glass Shards");
            //comboBox8.Items.Add("Binder");
            //comboBox8.Items.Add("Metal");
            //comboBox9.Items.Add("Good");
            //comboBox9.Items.Add("Passable");
            //comboBox9.Items.Add("Bad/Scrap");
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            SqlConnection sc = new SqlConnection();
            // SqlCommand com = new SqlCommand();


            sc.ConnectionString = ConnectionString;
             //  com.CommandText = (@"INSERT INTO Line3_data_input ([Pull(Position9)],[Pull(Position10)]) VALUES (" + txtPull9.Text + "," + txtPull10.Text + ")");
            //  com.CommandText = (@"INSERT INTO Line3_data_input ([Pull(Position9)],[Pull(Position10)],[Silicone(Position9)],[Silicone(Position10)],[Oil(Position9)],[Oil(Position10)],[Temp(Plenum9)],[Temp(Plenum10)],[Air(Position9)],[Air(Position10)],[Gas(Position9)],[Gas(Position10)],[Ratio(Position9)],[Ratio(Position10)],[RefIndex ],Solidcontent,LOI,LOD,ReLhumidity,[Temp(ambient)],Comments,Added_On) VALUES (" + txtPull9.Text + "," + txtPull10.Text + "," + txtSilicon9.Text + "," + txtSilicon10.Text + "," + txtOilFlow9.Text + "," + txtOilFlow10.Text + "," + txtTemp9.Text + "," + txtTemp10.Text + "," + txtAir9.Text + "," + txtAir10.Text + "," + txtGas9.Text + "," + txtGas10.Text + "," + txtRatio9.Text + "," + txtRatio10.Text + "," + txtRI.Text + "," + txtSolidContent.Text + "," + txtRelhumidity.Text + "," + txtLOI.Text + "," + txtLOD.Text + "," + txtTempAmbi.Text + ",'" + txtComments.Text + "',GetDate())");

            SqlCommand cmd = new SqlCommand();
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.Connection = sc;
            cmd.CommandText = "SP_InsertData";

            decimal Pull9, Pull10, Silicone9, Silicone10,TempPlenum9,TempPlenum10,Oil9,Oil10,Air9,Air10,Gas9,Gas10,Ratio9,Ratio10,RefIndex,SolidContent,
                LOI, LOD, ReLhumidity, TempAmbi, Comments, Density, BagTemp, Micronare;
            
            decimal.TryParse(txtPull9.Text, out Pull9); cmd.Parameters.Add("@Pull9", SqlDbType.Float).Value = Pull9;
            decimal.TryParse(txtPull10.Text, out Pull10);cmd.Parameters.Add("@Pull10", SqlDbType.Float).Value = Pull10;

           decimal.TryParse(txtSilicon9.Text, out Silicone9);cmd.Parameters.Add("@Silicone9", SqlDbType.Float).Value = Silicone9;
            decimal.TryParse(txtSilicon10.Text, out Silicone10); cmd.Parameters.Add("@Silicone10", SqlDbType.Float).Value = Silicone9;

            decimal.TryParse(txtOilFlow9.Text, out Oil9);cmd.Parameters.Add("@Oil9", SqlDbType.Float).Value = Oil9;
            decimal.TryParse(txtOilFlow10.Text, out Oil10);cmd.Parameters.Add("@Oil10", SqlDbType.Float).Value = Oil10;

            decimal.TryParse(txtTemp9.Text, out TempPlenum9); cmd.Parameters.Add("@TempPlenum9", SqlDbType.Float).Value = TempPlenum9;
              decimal.TryParse(txtTemp10.Text, out TempPlenum10); cmd.Parameters.Add("@TempPlenum10", SqlDbType.Float).Value = TempPlenum10;

            decimal.TryParse(txtAir9.Text, out Air9); cmd.Parameters.Add("@Air9", SqlDbType.Float).Value = Air9;
            decimal.TryParse(txtAir10.Text, out Air10); cmd.Parameters.Add("@Air10", SqlDbType.Float).Value = Air10;

            decimal.TryParse(txtAir9.Text, out Gas9);cmd.Parameters.Add("@Gas9", SqlDbType.Float).Value = Gas9;
            decimal.TryParse(txtAir10.Text, out Gas10); cmd.Parameters.Add("@Gas10", SqlDbType.Float).Value = Gas10;

            decimal.TryParse(txtRatio9.Text, out Ratio9); cmd.Parameters.Add("@Ratio9", SqlDbType.Float).Value = Ratio9;
            decimal.TryParse(txtRatio10.Text, out Ratio10); cmd.Parameters.Add("@Ratio10", SqlDbType.Float).Value = Ratio10;

            decimal.TryParse(txtRI.Text, out RefIndex); cmd.Parameters.Add("@RefIndex", SqlDbType.Float).Value = RefIndex;
            decimal.TryParse(txtSolidContent.Text, out SolidContent);cmd.Parameters.Add("@Solidcontent", SqlDbType.Float).Value = SolidContent;
            decimal.TryParse(txtLOI.Text, out LOI);  cmd.Parameters.Add("@LOI", SqlDbType.Float).Value = LOI;
            decimal.TryParse(txtLOD.Text, out LOD);cmd.Parameters.Add("@LOD", SqlDbType.Float).Value = LOD;
            decimal.TryParse(txtRelhumidity.Text, out ReLhumidity); cmd.Parameters.Add("@ReLhumidity",SqlDbType.Float).Value = ReLhumidity;
            decimal.TryParse(txtTempAmbi.Text, out TempAmbi); cmd.Parameters.Add("@TempAmbi", SqlDbType.Float).Value =TempAmbi;
            decimal.TryParse(txtComments.Text, out Comments);  cmd.Parameters.Add("@Comments", SqlDbType.NVarChar).Value =Comments;
            decimal.TryParse(txtDensity.Text, out Density);cmd.Parameters.Add("@Density", SqlDbType.Float).Value = Density;
            decimal.TryParse(txtBagTemp.Text, out BagTemp); cmd.Parameters.Add("@BagTemp", SqlDbType.Float).Value = BagTemp;
            decimal.TryParse(txtMicronaire.Text, out Micronare); cmd.Parameters.Add("@Micronare", SqlDbType.Float).Value = Micronare;

         cmd.Parameters.AddWithValue("@DustID",cmbDustData.SelectedValue);
         cmd.Parameters.AddWithValue("@StaticID", cmbStaticData.SelectedValue);
            //Bag Weights
            decimal BW1,BW2,BW3,BW4,BW5,BW6,BW7,BW8,BW9,BW10;
           decimal BW11,BW12,BW13,BW14,BW15,BW16,BW17,BW18,BW19,BW20;
           decimal BW21, BW22, BW23, BW24, BW25;

         decimal.TryParse(txtBW1.Text, out BW1); cmd.Parameters.Add("@BW1", SqlDbType.Float).Value = BW1;
         decimal.TryParse(txtBW2.Text, out BW2); cmd.Parameters.Add("@BW2", SqlDbType.Float).Value = BW2;
         decimal.TryParse(txtBW3.Text, out BW3); cmd.Parameters.Add("@BW3", SqlDbType.Float).Value = BW3;
         decimal.TryParse(txtBW4.Text, out BW4); cmd.Parameters.Add("@BW4", SqlDbType.Float).Value = BW4;
         decimal.TryParse(txtBW5.Text, out BW5); cmd.Parameters.Add("@BW5", SqlDbType.Float).Value = BW5;

       decimal.TryParse(txtBW6.Text, out BW6); cmd.Parameters.Add("@BW6", SqlDbType.Float).Value = BW6;
       decimal.TryParse(txtBW7.Text, out BW7); cmd.Parameters.Add("@BW7", SqlDbType.Float).Value = BW7;
       decimal.TryParse(txtBW8.Text, out BW8); cmd.Parameters.Add("@BW8", SqlDbType.Float).Value = BW8;
       decimal.TryParse(txtBW9.Text, out BW9); cmd.Parameters.Add("@BW9", SqlDbType.Float).Value =BW9;
       decimal.TryParse(txtBW10.Text, out BW10); cmd.Parameters.Add("@BW10", SqlDbType.Float).Value = BW10;

       decimal.TryParse(txtBW11.Text, out BW11); cmd.Parameters.Add("@BW11", SqlDbType.Float).Value = BW11;
       decimal.TryParse(txtBW12.Text, out BW12); cmd.Parameters.Add("@BW12", SqlDbType.Float).Value = BW12;
       decimal.TryParse(txtBW13.Text, out BW13); cmd.Parameters.Add("@BW13", SqlDbType.Float).Value = BW13;
       decimal.TryParse(txtBW14.Text, out BW14); cmd.Parameters.Add("@BW14", SqlDbType.Float).Value = BW14;
       decimal.TryParse(txtBW15.Text, out BW15); cmd.Parameters.Add("@BW15", SqlDbType.Float).Value =BW15;

       decimal.TryParse(txtBW16.Text, out BW16); cmd.Parameters.Add("@BW16", SqlDbType.Float).Value = BW16;
       decimal.TryParse(txtBW17.Text, out BW17); cmd.Parameters.Add("@BW17", SqlDbType.Float).Value = BW17;
       decimal.TryParse(txtBW18.Text, out BW18); cmd.Parameters.Add("@BW18", SqlDbType.Float).Value = BW18;
       decimal.TryParse(txtBW19.Text, out BW19); cmd.Parameters.Add("@BW19", SqlDbType.Float).Value = BW19;
       decimal.TryParse(txtBW20.Text, out BW20); cmd.Parameters.Add("@BW20", SqlDbType.Float).Value = BW20;

       decimal.TryParse(txtBW21.Text, out BW21); cmd.Parameters.Add("@BW21", SqlDbType.Float).Value = BW21;
       decimal.TryParse(txtBW22.Text, out BW22); cmd.Parameters.Add("@BW22", SqlDbType.Float).Value = BW22;
       decimal.TryParse(txtBW23.Text, out BW23); cmd.Parameters.Add("@BW23", SqlDbType.Float).Value = BW23;
       decimal.TryParse(txtBW24.Text, out BW24); cmd.Parameters.Add("@BW24", SqlDbType.Float).Value = BW24;
       decimal.TryParse(txtBW25.Text, out BW25); cmd.Parameters.Add("@BW25", SqlDbType.Float).Value = BW25;
            

         cmd.Parameters.AddWithValue("@ClumpsID", cmbClumps.SelectedValue);
         cmd.Parameters.AddWithValue("@YesNo", cmbFMYesNo.SelectedValue);
         cmd.Parameters.AddWithValue("@FM", cmbFMat.SelectedValue);
         cmd.Parameters.AddWithValue("@RateID", cmbProdRate.SelectedValue);
         cmd.Parameters.AddWithValue("@AuthorityID", cmbAuthority.SelectedValue);
         sc.Open();
         cmd.ExecuteNonQuery();
         MessageBox.Show("Entry Added!");
            sc.Close();
        }

        private void button3_Click(object sender, EventArgs e)
        {

            SqlConnection sc = new SqlConnection();
            sc.ConnectionString = ConnectionString;
            // string sql = null;
            string data = null;
            int i, j;
            // int j = 0;

            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            Excel.Range oRange;
            object misValue = System.Reflection.Missing.Value;

            xlApp = new Excel.Application();
            xlWorkBook = xlApp.Workbooks.Add(misValue);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);


            //xlWorkSheet.Cells.Style.EntireColumn.AutoFit();
            // Automatically set the width on columns B and C. 
            // xlWorkSheet.Cells["A:AZ"].Columns.AutoFit(); 
            //  xlWorkSheet.Cells.Style.Font.Name = "Arial";
            //xlWorkSheet.get_Range(excelSheetPrint.Cells[1, 1], excelSheetPrint.Cells[1, maxCol]).Font.Bold = true;
            //xlWorkSheet.get_Range(excelSheetPrint.Cells[1, 1], excelSheetPrint.Cells[1, maxCol]).Font.Size = 10;
            //xlWorkSheet.get_Range(excelSheetPrint.Cells[1, 1], excelSheetPrint.Cells[maxRow + 1, maxCol]).Borders.LineStyle = 1;

            Microsoft.Office.Interop.Excel.Application app = new Microsoft.Office.Interop.Excel.Application();
            app.Visible = false;
            sc.Open();
            SqlCommand cmd = new SqlCommand();
            SqlDataAdapter dscmd = new SqlDataAdapter();
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.Connection = sc;
            cmd.CommandText = "DisplayData";
            // sql = "SELECT * FROM Line3_data_input ORDER BY Added_On DESC";
            // SqlDataAdapter dscmd = cmd;
            dscmd.SelectCommand = cmd;
            DataSet dsDisplay = new DataSet();
            dscmd.Fill(dsDisplay);
            //aRange = (Excel.Range)xlWorkSheet.get_Range(xlWorkSheet.Cells[1, 1] as Excel._OLEObject, xlWorkSheet.Cells[4, 4] as Excel.OLEObject);
            oRange = (Excel.Range)xlWorkSheet.Range[xlWorkSheet.Cells[1, 1], xlWorkSheet.Cells[200, 200]];
            oRange.EntireColumn.AutoFit();



            //Excel.Range headerRange = xlWorkSheet.get_Range("A1", "V1");
            //headerRange.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
            //headerRange.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
            //headerRange.Value = "Header text 1";
            //xlWorkSheet.Cells[1, "A"] = "Name";
            //xlWorkSheet.Cells[1, "B"] = "Color";
            //xlWorkSheet.Cells[1, "C"] = "Maximum speed";

            for (i = 0; i < dsDisplay.Tables[0].Columns.Count - 1; i++)
            {
                xlWorkSheet.Cells[1, i + 1] = dsDisplay.Tables[0].Columns[i].ColumnName;
                xlWorkSheet.Cells[1, i + 1].Style.Font.Bold = true;

            }

            for (i = 0; i <= dsDisplay.Tables[0].Rows.Count - 1; i++)
            {
                for (j = 0; j <= dsDisplay.Tables[0].Columns.Count - 1; j++)
                {
                    data = dsDisplay.Tables[0].Rows[i].ItemArray[j].ToString();
                    xlWorkSheet.Cells[i + 2, j + 1] = data;
                }
            }


            xlWorkBook.SaveAs("NorthernWhite.xls", Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
            oRange = (Excel.Range)xlWorkSheet.Range[xlWorkSheet.Cells[1, 1], xlWorkSheet.Cells[200, 200]];
            oRange.EntireColumn.AutoFit();
            xlWorkBook.Close(true, misValue, misValue);
            xlApp.Quit();

            releaseObject(xlWorkSheet);
            releaseObject(xlWorkBook);
            releaseObject(xlApp);

            MessageBox.Show("Excel file created , you can find the file  in  Documents folder : NorthernWhite.xls");
        }

        private void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                MessageBox.Show("Unable to release the Object " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }



    }

}



