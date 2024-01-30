using System;
using System.Drawing;
using System.IO;
using System.Windows.Forms;

namespace EliteFitness
{
    public partial class Popup : Form
    {
        string extension = ".jpg";

        public Popup(ListViewItem lvItem)
        {
            InitializeComponent();
            string filename = lvItem.SubItems[3].Text;
            this.Text = lvItem.SubItems[0].Text + " "+ lvItem.SubItems[1].Text;
            FileInfo myFile = new FileInfo(getPath() + filename+ extension);
            if (myFile.Exists) { 
                pictureBox1.Image = Image.FromFile(myFile.FullName);
            }
            else
            {
                pictureBox1.Image = null;
            }
            lbContato.Text = "Contato: " + lvItem.SubItems[4].Text;
            lbInstrutor.Text = "Instrutor: " + lvItem.SubItems[5].Text;
            
            //verificar data
            DateTime parsedDate = DateTime.Now;
            try
            {
                parsedDate = DateTime.Parse(lvItem.SubItems[2].Text.ToString());
            }
            catch { }
            double totalDays = (DateTime.Now - parsedDate).TotalDays;
            lbDays.Text = "Nº dias ausente: " + Convert.ToInt32(totalDays).ToString();
        }

        private string getPath()
        {
            int gym = new DB().Ginasio;
            if (gym == 1) { 
                string path = @"\\riomaiorsrv\hc\fotos\";
                return path;
            }
            if (gym == 2)
            {
                string path = @"\\servidoralv\hc\fotos\";
                return path;
            }
            return null;
        }
    }
}
