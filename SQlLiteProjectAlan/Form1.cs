using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Data.SQLite;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop;
using System.Data.OleDb;

namespace SQlLiteProjectAlan
{
    
    
    public partial class Form1 : Form
    {
        int my_time;
        int time_;
        public Form1()
        {
            InitializeComponent();
        }

        private void tabPage2_Click(object sender, EventArgs e)
        {

        }

        private void Form1_Load(object sender, EventArgs e)
        {
            // TODO: данная строка кода позволяет загрузить данные в таблицу "developmentDataSet.consts". При необходимости она может быть перемещена или удалена.
            this.constsTableAdapter.Fill(this.developmentDataSet.consts);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "developmentDataSet.ocenkas". При необходимости она может быть перемещена или удалена.
            this.ocenkasTableAdapter.Fill(this.developmentDataSet.ocenkas);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "developmentDataSet.nds". При необходимости она может быть перемещена или удалена.
            this.ndsTableAdapter.Fill(this.developmentDataSet.nds);
            
        }

        private void bindingNavigator1_RefreshItems_1(object sender, EventArgs e)
        {

        }

        private void toolStripButton1_Click(object sender, EventArgs e)
        {


            try
            {

                
                this.ndsTableAdapter.Update(developmentDataSet.nds);
                
               //this.ndsTableAdapter.Update((DataTable)dataGridView1.DataSource);
                MessageBox.Show("Изменения в базе данных выполнены!",
                  "Уведомление о результатах", MessageBoxButtons.OK);
            }
            catch (Exception)
            {
                MessageBox.Show("Изменения в базе данных выполнить не удалось!",
                  "Уведомление о результатах", MessageBoxButtons.OK);
            }
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            ndsTableAdapter.Ozenka0();
            this.ndsTableAdapter.Fill(this.developmentDataSet.nds);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            ndsTableAdapter.time0();
            this.ndsTableAdapter.Fill(this.developmentDataSet.nds);
        }

        private void button3_Click(object sender, EventArgs e)
        {
            ndsTableAdapter.DeleteAll();
            this.ndsTableAdapter.Fill(this.developmentDataSet.nds);
        }

        private void button4_Click(object sender, EventArgs e)
        {
            if (radioButton1.Checked ==  true)
            {
                ndsTableAdapter.Update_MyIs(true);
                this.ndsTableAdapter.Fill(this.developmentDataSet.nds);
            }
            else if (radioButton2.Checked == true)
            {
                ndsTableAdapter.Update_MyIs(false);
                this.ndsTableAdapter.Fill(this.developmentDataSet.nds);
            }
            
        }

        private void groupBox2_RegionChanged(object sender, EventArgs e)
        {
           
        }

        private void groupBox2_EnabledChanged(object sender, EventArgs e)
        {
            MessageBox.Show("Test");
        }

        private void button5_Click(object sender, EventArgs e)
        {
            string my_arg = "id";
            foreach (Control cnt in groupBox2.Controls)
            {
                
                RadioButton tb = cnt as RadioButton;
                if (tb != null)
                {
                    if (tb.Checked == true)
                        my_arg = tb.Name;
                    
                }
              

            }
            

            DataTable contacts =  developmentDataSet.Tables["nds"];
            DataView view = contacts.AsDataView();
            view.Sort = my_arg + " asc";
            ndsBindingSource.DataSource = view;
            //dataGridView1.AutoResizeColumns();

        }

        private void toolStripButton2_Click(object sender, EventArgs e)
        {
           // В References нужно дабавить Microsoft Excel 11.0 Object Library и Microsoft Office 11.0 Object Library
// не обязательно версия 11.0 может отличаться. взависимости какая версия Excel

            Microsoft.Office.Interop.Excel.Application ThisApplication = new Microsoft.Office.Interop.Excel.Application();

            Microsoft.Office.Interop.Excel.Workbook MyBook = ThisApplication.Workbooks.Add(Type.Missing);

            Microsoft.Office.Interop.Excel.Worksheet MySheet = (Microsoft.Office.Interop.Excel.Worksheet)MyBook.Sheets[1];

            ThisApplication.Visible = true;

            Microsoft.Office.Interop.Excel.Range myRange;

            int iEndValue = 0;


            string[,] dtValue = new string[dataGridView1.Rows.Count-1, dataGridView1.ColumnCount];

            for (int i = 0; i < dataGridView1.Rows.Count-1; i++)
            {
                for (int j = 0; j < dataGridView1.ColumnCount; j++)
                {
                    dtValue[i, j] = dataGridView1.Rows[i].Cells[j].Value.ToString();
                }
            }
            MySheet.get_Range("A1", "A1").Value2 = "ID";
            MySheet.get_Range("B1", "B1").Value2 = "Имя";
            MySheet.get_Range("C1", "C1").Value2 = "Оценка1";
            MySheet.get_Range("D1", "D1").Value2 = "Оценка2";
            MySheet.get_Range("E1", "E1").Value2 = "Оценка3";
            MySheet.get_Range("F1", "F1").Value2 = "Оценка4";
            MySheet.get_Range("G1", "G1").Value2 = "Оценка5";
            MySheet.get_Range("H1", "H1").Value2 = "Оценка6";
            MySheet.get_Range("I1", "I1").Value2 = "Оценка7";
            MySheet.get_Range("J1", "J1").Value2 = "Оценка8";
            MySheet.get_Range("K1", "K1").Value2 = "Оценка9";
            MySheet.get_Range("L1", "L1").Value2 = "Участие";
            MySheet.get_Range("M1", "M1").Value2 = "Время";

            MySheet.get_Range("A3", "M" + (dataGridView1.Rows.Count + 1).ToString()).Value2 = dtValue;
            MySheet.get_Range("A3", "M" + (dataGridView1.Rows.Count + 1).ToString()).NumberFormat = "@";
            MySheet.get_Range("A3", "M" + (dataGridView1.Rows.Count + 1).ToString()).Font.Name = "Microsoft Sans Serif";
            MySheet.get_Range("A3", "M" + (dataGridView1.Rows.Count + 1).ToString()).Font.Size = 10;
            //for (int i = 1; i <= 5; i++)

            //{

            //myRange = MySheet.get_Range("A" + i.ToString(), Type.Missing).Cells;

            //myRange.Value2 = i.ToString();

            //myRange.Font.Bold = true;

            //myRange.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;

            //iEndValue = i;

            //}

            //myRange = MySheet.get_Range("A1", "A" + iEndValue).Cells;

            //myRange.Columns.AutoFit();

            //myRange.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;

            //myRange.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;

            //myRange.Borders.ColorIndex = Microsoft.Office.Interop.Excel.XlColorIndex.xlColorIndexAutomatic;

}

        private void button6_Click(object sender, EventArgs e)
        {
            OpenFileDialog opf = new OpenFileDialog();
            opf.Filter = "Excel (*.XLS)|*.XLS";
            opf.ShowDialog();
            DataTable tb = new DataTable();
            string filename = opf.FileName;
            //if (filename == "")  break;
            string ConStr = String.Format("Provider=Microsoft.Jet.OLEDB.4.0; Data Source={0}; Extended Properties=Excel 8.0;", filename);
            System.Data.DataSet ds = new System.Data.DataSet("EXCEL");
            OleDbConnection cn = new OleDbConnection(ConStr);
            cn.Open();
            DataTable schemaTable = cn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new object[] { null, null, null, "TABLE" });
            string sheet1 = (string)schemaTable.Rows[0].ItemArray[2];
            string select = String.Format("SELECT * FROM [{0}]", sheet1);
            OleDbDataAdapter ad = new OleDbDataAdapter(select, cn);
            ad.Fill(ds);
            tb = ds.Tables[0];
            cn.Close();
            dataGridView2.DataSource = tb;
            string new_name;
            DateTime date1 = new DateTime(2008, 5, 1, 8, 30, 52);
            DateTime date2 = new DateTime(2008, 5, 1, 8, 30, 52);
            ndsTableAdapter.DeleteAll();
            for (int i = 0; i < dataGridView2.Rows.Count - 1; i++)
            {
                new_name = dataGridView2.Rows[i].Cells[0].Value.ToString();
                ndsTableAdapter.InsertFromExcel(new_name, date1, date2);




            }
            this.ndsTableAdapter.Fill(this.developmentDataSet.nds);
        }

        private void button7_Click(object sender, EventArgs e)
        {
            
        }

        private void bindingSource1_CurrentChanged(object sender, EventArgs e)
        {

        }

        private void button7_Click_1(object sender, EventArgs e)
        {
            try
            {


                //this.ndsTableAdapter.Update(developmentDataSet.nds);
                this.ocenkasTableAdapter.Update(developmentDataSet.ocenkas);
               
                MessageBox.Show("Изменения в базе данных выполнены!",
                  "Уведомление о результатах", MessageBoxButtons.OK);
            }
            catch (Exception)
            {
                MessageBox.Show("Изменения в базе данных выполнить не удалось!",
                  "Уведомление о результатах", MessageBoxButtons.OK);
            }
        }

        private void button8_Click(object sender, EventArgs e)
        {
            try
            {


                //this.ndsTableAdapter.Update(developmentDataSet.nds);
                this.constsTableAdapter.Update(developmentDataSet.consts);
                
                MessageBox.Show("Изменения в базе данных выполнены!",
                  "Уведомление о результатах", MessageBoxButtons.OK);
            }
            catch (Exception)
            {
                MessageBox.Show("Изменения в базе данных выполнить не удалось!",
                  "Уведомление о результатах", MessageBoxButtons.OK);
            }
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            time_ = 0;
            my_time += 1;
            progressBar1.Value = my_time * 10;
            label2.Text = my_time.ToString(); 
            int temp_c=0;
            for (int i = 0; i < dataGridView4.Rows.Count; i++)
            {
               temp_c =  Convert.ToInt32(dataGridView4.Rows[i].Cells[2].Value);
               if (temp_c  > 0)
               {
                   temp_c -= 1;
                   
                   dataGridView4.Rows[i].Cells[2].Value = temp_c;
                   int temp_id = Convert.ToInt32(dataGridView4.Rows[i].Cells[0].Value);
                   this.ndsTableAdapter.UpdateTime(temp_c, temp_id);
                   
                   if (temp_c == 0)
                   {
                       string name = dataGridView4.Rows[i].Cells[1].Value.ToString();
                       MessageBox.Show("Время участника " + name + " израсходовалось !");
                   }

                   
               }

               
               

            }
            
        }

        private void button9_Click(object sender, EventArgs e)
        {
            my_time = 0;
            time_ = 0;
            timer1.Start();
            timer2.Start();
            button9.Enabled = false;
            button10.Enabled = true;
        }

        private void button10_Click(object sender, EventArgs e)
        {
            timer1.Stop();
            timer2.Stop();
            time_ = 0;
            progressBar1.Value = 0;
            button9.Enabled = true;
            button10.Enabled = false;
        }

        private void timer2_Tick(object sender, EventArgs e)
        {
            time_ += 1;
            progressBar1.Value = time_;
        }

       
        }

       

   
    
}
