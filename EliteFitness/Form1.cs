using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
/**
 * Parametros que pode ser configurados
 * 
 * ConnectionString
 * Tempo de atualizacao
 * Emails
 * Pasta de Logs
 * 
 */
namespace EliteFitness
{
    public partial class Form1 : Form
    {
        int counter = 0;
        int refreshTime = 10 * 60;
        List<int> lstChecked = new List<int>();
        double daysOut = 7;
        string strEmailSend= "souelitefitness@gmail.com,informatica@souelitefitness.pt";


        public Form1()
        {
            InitializeComponent();
            timer.Interval = 1000;
            timer.Start();
            infoData();
        }

        private void infoData()
        {
            DataTable dtTable = searchData();
            //DataTable dtTable = mockData();

            List<string> lstSleepers = new List<string>();
            foreach (DataRow dr in dtTable.Rows)
            {
                //verificar data
                DateTime parsedDate = DateTime.Now;
                try
                {
                    parsedDate = DateTime.Parse(dr[2].ToString());
                }
                catch { }

                double lastDay = (DateTime.Now - parsedDate).TotalDays;

                if (lastDay >= daysOut)
                {
                    try
                    {
                        parsedDate = DateTime.Parse(dr[2].ToString());
                    }
                    catch { }
                    double totalDays = (DateTime.Now - parsedDate).TotalDays;
                    lstSleepers.Add(dr[0].ToString() + " - " + dr[1].ToString() + " - " + Convert.ToInt32(totalDays).ToString());
                }
            }

            bool isToSendEmail = false;
            try
            {
                isToSendEmail = saveLogFile(lstSleepers);
            }
            catch  (Exception e)
            {
                lbError.Text = "ERRO a gravar logs. Contatar Informatico.";
            }

            if (isToSendEmail)
            {
                try
                {
                    sendEmail();
                }
                catch (Exception e)
                {
                    lbError.Text = "ERRO no registo. Contatar Informatico.";
                }
            }
            
        }

        private void Count()
        {
            counter++;
            if (counter == refreshTime)
            {
                counter = 0;
                infoData();
            }
        }

        private void Timer_Tick(object sender, EventArgs e)
        {
            Count();
        }

        private void BtRefreshTimer_Click(object sender, EventArgs e)
        {
            if (timer.Enabled)
            {
                timer.Stop();
                btRefreshTimer.Text = "Atualizar";
                counter = 0;
            }
            else
            {
                timer.Start();
                btRefreshTimer.Text = "STOP";
                checkAvisados();
                infoData();
            }
        }


        private DataTable searchData()
        {
            getWindowToFront();
            dataGridView1.DataSource = null;
            dataGridView1.Refresh();
            DataTable dt = new DB().Select("SELECT atl.numcliente, atl.nome, atl.ultimavisita,  atl.codigo,  atl.telemovel, us.nome, atl.sexo FROM ATLETAS atl INNER JOIN(SELECT DISTINCT atleta FROM ENTRADAS_TEMP) ent ON atl.codigo = ent.atleta INNER JOIN Users us ON atl.id_userprof = us.id_user");
            dataGridView1.DataSource = dt;
            dateTableToListView(dt);
            return dt;
        }

        private void getWindowToFront()
        {
            this.WindowState = FormWindowState.Minimized;
            this.Show();
            this.WindowState = FormWindowState.Normal;
        }

        private void dateTableToListView(DataTable data)
        {
            listView1.Items.Clear();
            int male = 0;
            int female = 0;
            foreach (DataRow row in data.Rows)
            {
                if (row[6].ToString().ToUpper().Equals("M"))
                {
                    male += 1;
                }
                if (row[6].ToString().ToUpper().Equals("F"))
                {
                    female += 1;
                }

                ListViewItem item = new ListViewItem(row[0].ToString());

                foreach (int id in lstChecked)
                {
                    if (id == int.Parse(row[0].ToString()))
                    {
                        item.Checked = true;
                    }
                }

                //verificar data
                DateTime parsedDate = DateTime.Now;
                try
                {
                    parsedDate = DateTime.Parse(row[2].ToString());
                }
                catch { }

                double lastDay = (DateTime.Now - parsedDate).TotalDays;

                if (lastDay >= daysOut)
                {
                    for (int i = 1; i < data.Columns.Count; i++)
                    {
                        item.SubItems.Add(row[i].ToString());
                    }
                    item.BackColor = changeColors(lastDay, daysOut);
                    listView1.Items.Add(item);
                }
            }
            lbGender.Text = "Hora: "+ getTwoDigits(DateTime.Now.Hour.ToString())+":"+ getTwoDigits(DateTime.Now.Minute.ToString()) + " estavam: Homens: " + male + " e Mulheres: " + female;
        }

        private void TxInterval_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !char.IsDigit(e.KeyChar) && !char.IsControl(e.KeyChar);
        }


        private void checkAvisados()
        {
            lstChecked.Clear();
            for (int i = 0; i < listView1.Items.Count; i++)
            {
                if (listView1.Items[i].Checked)
                {
                    lstChecked.Add(int.Parse(listView1.Items[i].SubItems[0].Text));
                    //MessageBox.Show("Listview items " + listView1.Items[i].SubItems[0].Text + " is checked");
                }
            }
        }

        private Color changeColors(Double totalDay, Double defined)
        {
            Double total = totalDay - defined;
            if (total > 21)
                return Color.Red;
            if (total > 14)
                return Color.Orange;
            if (total > 7)
                return Color.Yellow;
            else return Color.Green;
        }


        private void ListView1_DoubleClick(object sender, EventArgs e)
        {
            if (listView1.SelectedItems.Count == 1)
            {
                ListViewItem lvItem = listView1.SelectedItems[0];

                //MessageBox.Show("Listview Double Click " + lvItem.Text);
                Popup frm2 = new Popup(lvItem);
                frm2.ShowDialog();
            }
        }

        private String getTwoDigits(String strTime)
        {
            if (strTime.Length == 1)
            {
                return "0" + strTime;
            }
            return strTime;
        }

        private void sendEmail()
        {
            DateTime dtYesterday = DateTime.Now.AddDays(-1);
            string filename = dtYesterday.Year + "" + dtYesterday.Month + "" + dtYesterday.Day + ".txt";
            string pathFile = @"Logs\" + filename;
            if (File.Exists(pathFile))
            {
                string[] lines = System.IO.File.ReadAllLines(pathFile);

                string strBody = "";
                foreach (string s in lines)
                {
                    strBody += s + Environment.NewLine;
                }

                using (MailMessage mm = new MailMessage("informatica@souelitefitness.pt", strEmailSend))
                {
                    mm.Subject = "Listagem de sleepers GINASIO " + new DB().Ginasio.ToString();
                    mm.Body = strBody;

                    mm.IsBodyHtml = false;
                    SmtpClient smtp = new SmtpClient();
                    smtp.Host = "mail.souelitefitness.pt";
                    smtp.EnableSsl = true;
                    smtp.UseDefaultCredentials = false;
                    NetworkCredential NetworkCred = new NetworkCredential("informatica@souelitefitness.pt", "Ser19Elitec*");
                
                    smtp.Credentials = NetworkCred;
                    smtp.Port = 587;
                    smtp.Send(mm);
                    //MessageBox.Show("Email sent.", "Message");
                }
            }
        }


        private Boolean saveLogFile(List<String> lstSleepers)
        {
            String filename = DateTime.Now.Year+""+DateTime.Now.Month+""+ DateTime.Now.Day+".txt";
            String pathFile = @"Logs\" + filename;
            if (File.Exists(pathFile)) 
            {
                //Se existir ler do ficheiro e comparar, depois guarda no mesmo ficheiro
                
                // Read each line of the file into a string array. Each element of the array is one line of the file.
                String[] lines = System.IO.File.ReadAllLines(pathFile);

                // Compare the file and the db info and excludes what already exists.
                List<String> newSleepers = lstSleepers.Except(lines).ToList();

                // The using statement automatically flushes AND CLOSES the stream and calls IDisposable.Dispose on the stream object.
                using (System.IO.StreamWriter file = new System.IO.StreamWriter(pathFile, true))
                {
                    foreach (String strSleepers in newSleepers)
                    {
                        file.WriteLine(strSleepers);
                    }
                }
                return false;
            }
            else
            {
                //Guardar logo a primeira vez com o nome do ficheiro
                File.WriteAllLines(pathFile, lstSleepers.ToArray());
                return true;
            }
        }





        private DataTable mockData()
        {
            getWindowToFront();

            DataTable table = new DataTable("Table");
            DataColumn numcliente = new DataColumn("atl.numcliente", typeof(int));
            DataColumn amountColumn = new DataColumn("atl.nome", typeof(String));
            DataColumn dateColumn = new DataColumn("atl.ultimavisita", typeof(String));
            DataColumn idColumn = new DataColumn("atl.codigo", typeof(int));
            DataColumn telemovel = new DataColumn("atl.telemovel", typeof(String));
            DataColumn userprof = new DataColumn("us.nome", typeof(String));
            DataColumn sexo = new DataColumn("atl.sexo", typeof(String));

            table.Columns.Add(numcliente);
            table.Columns.Add(amountColumn);
            table.Columns.Add(dateColumn);
            table.Columns.Add(idColumn);
            table.Columns.Add(telemovel);
            table.Columns.Add(userprof);
            table.Columns.Add(sexo);

            DataRow newRow = table.NewRow();
            newRow["atl.numcliente"] = 100101;
            newRow["atl.nome"] = "Rui Ribeiro";
            newRow["atl.ultimavisita"] = new DateTime(2019, 09, 20).ToString();
            newRow["atl.codigo"] = 4;
            newRow["atl.telemovel"] = "123456";
            newRow["us.nome"] = "PT Alverca";
            newRow["atl.sexo"] = "M";
            table.Rows.Add(newRow);

            newRow = table.NewRow();
            newRow["atl.numcliente"] = 100201;
            newRow["atl.nome"] = "Pedro Ferreira";
            newRow["atl.ultimavisita"] = new DateTime(2019, 09, 12).ToString();
            newRow["atl.codigo"] = 5;
            newRow["atl.telemovel"] = "123456";
            newRow["us.nome"] = "PT Alverca 2";
            newRow["atl.sexo"] = "M";
            table.Rows.Add(newRow);

            newRow = table.NewRow();
            newRow["atl.numcliente"] = 100301;
            newRow["atl.nome"] = "Carina Santos";
            newRow["atl.ultimavisita"] = new DateTime(2019, 09, 06).ToString();
            newRow["atl.codigo"] = 6;
            newRow["atl.telemovel"] = "123456";
            newRow["us.nome"] = "Alverca PT 1";
            newRow["atl.sexo"] = "F";
            table.Rows.Add(newRow);

            newRow = table.NewRow();
            newRow["atl.numcliente"] = 100401;
            newRow["atl.nome"] = "Matilde Constança";
            newRow["atl.ultimavisita"] = new DateTime(2019, 09, 28).ToString();
            newRow["atl.codigo"] = 5;
            newRow["atl.telemovel"] = "123456";
            newRow["us.nome"] = "PT Alverca 2";
            newRow["atl.sexo"] = "F";
            table.Rows.Add(newRow);

            newRow = table.NewRow();
            newRow["atl.numcliente"] = 100501;
            newRow["atl.nome"] = "Carlota Francisca";
            newRow["atl.ultimavisita"] = new DateTime(2019, 10, 01).ToString();
            newRow["atl.codigo"] = 6;
            newRow["atl.telemovel"] = "123456";
            newRow["us.nome"] = "Alverca PT 1";
            newRow["atl.sexo"] = "F";
            table.Rows.Add(newRow);

            dataGridView1.DataSource = table;
            dateTableToListView(table);

            //DisplayInExcel(table);
            return table;
        }
        

        static void DisplayInExcel(DataTable data)
        {
            var excelApp = new Excel.Application();
            // Make the object visible.
            excelApp.Visible = true;

            // Create a new, empty workbook and add it to the collection returned 
            // by property Workbooks. The new workbook becomes the active workbook.
            // Add has an optional parameter for specifying a praticular template. 
            // Because no argument is sent in this example, Add creates a new workbook. 
            excelApp.Workbooks.Add();

            // This example uses a single workSheet. The explicit type casting is
            // removed in a later procedure.
            Excel._Worksheet workSheet = (Excel.Worksheet)excelApp.ActiveSheet;

            // Establish column headings in cells A1 and B1.
            workSheet.Cells[1, "A"] = "ID Number";
            workSheet.Cells[1, "B"] = "Current Balance";

            var irow = 1;
            foreach (DataRow row in data.Rows)
            {
                irow++;
                workSheet.Cells[irow, "A"] = row[0].ToString();
                workSheet.Cells[irow, "B"] = row[1].ToString();
            }

            workSheet.Columns[1].AutoFit();
            workSheet.Columns[2].AutoFit();

            ((Excel.Range)workSheet.Columns[1]).AutoFit();
            ((Excel.Range)workSheet.Columns[2]).AutoFit();
        }

        private void createExcel() { 
        Excel.Application xlApp;
        Excel.Workbook xlWorkBook;
        Excel.Worksheet xlWorkSheet;
        object misValue = System.Reflection.Missing.Value;

        xlApp = new Excel.Application();
        xlWorkBook = xlApp.Workbooks.Add(misValue);

        xlWorkSheet = (Excel.Worksheet) xlWorkBook.Worksheets.get_Item(1);
        xlWorkSheet.Cells[1, 1] = "http://csharp.net-informations.com";

        xlWorkBook.SaveAs("csharp-Excel.xls", Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
        xlWorkBook.Close(true, misValue, misValue);
        xlApp.Quit();

        releaseObject(xlWorkSheet);
        releaseObject(xlWorkBook);
        releaseObject(xlApp);

        MessageBox.Show("Excel file created , you can find the file c:\\csharp-Excel.xls");
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
                MessageBox.Show("Exception Occured while releasing object " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }

        private void BtnMinimizar_Click(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Minimized;
        }
    }
}
