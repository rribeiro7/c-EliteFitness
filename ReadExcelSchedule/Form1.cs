using EliteFitness;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace ReadExcelSchedule
{
    public partial class Form1 : Form
    {
        private const string IniPath = "Audio.ini";
        IniFile MyIni = null;
        
        public Form1()
        {
            InitializeComponent();
        }

        private void readExcel(String filePath)
        {
            MyIni = new IniFile(IniPath);

            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            Excel.Range range;

            string strTitulo;
            string strData;
            string strHora;
            int rCnt;
            int rw = 0;
            int cl = 0;

            xlApp = new Excel.Application();
            xlWorkBook = xlApp.Workbooks.Open(filePath, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

            range = xlWorkSheet.UsedRange;
            rw = range.Rows.Count;
            cl = range.Columns.Count;
          
            for (rCnt = 2; rCnt <= rw; rCnt++)
            {
                strTitulo = (string)(range.Cells[rCnt, 3] as Excel.Range).Value2;
                strData = (string)(range.Cells[rCnt, 5] as Excel.Range).Value2;
                strHora = (string)(range.Cells[rCnt, 6] as Excel.Range).Value2;
                regexString(strData, strHora, strTitulo);
            }

            xlWorkBook.Close(true, null, null);
            xlApp.Quit();

            Marshal.ReleaseComObject(xlWorkSheet);
            Marshal.ReleaseComObject(xlWorkBook);
            Marshal.ReleaseComObject(xlApp);

        }

        private void BtnOpenFile_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog
            {
                InitialDirectory = @"P:\",
                Title = "Browse Text Files",

                CheckFileExists = true,
                CheckPathExists = true,

                DefaultExt = "xls",
                Filter = "xls files (*.xls*)|*.xls*",
                FilterIndex = 2,
                RestoreDirectory = true,

                ReadOnlyChecked = true,
                ShowReadOnly = true
            };

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                if (File.Exists(IniPath))
                {
                    File.Delete(IniPath);
                }
                readExcel(openFileDialog1.FileName);
            }
        }

        private void regexString (String strData, String strHora, String strTitulo)
        {
            try
            {
                strHora = ifExists(strHora, strData);

                MyIni.Write(strHora, strTitulo.ToUpper(), strData);
            }
            catch
            {
                MessageBox.Show(strData + " - " + strHora + " - " + strTitulo);
            }
        }

        private String ifExists(String hora, String data)
        {
            if (MyIni.KeyExists(hora, data))
            {
                DateTime oDate = DateTime.ParseExact(hora, "HH:mm", null);

                hora = oDate.AddMinutes(1).ToShortTimeString();

                ifExists(hora, data);
            }
            return hora;
        }
    }
}
