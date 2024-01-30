using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using EliteFitness;
using Excel = Microsoft.Office.Interop.Excel;

namespace EF_Alarme
{
    public partial class Form1 : Form
    {
        int counter = 0;
        int refreshTime = 1 * 60; //minutos * 60 segundos
        int pauseSpotify = 8; //minutos * 60 segundos
        double iMinutesBeforeClass = 5;
        Boolean isPlaying = false;

        //read from Excel
        string strTitulo1;
        string strTitulo2;
        string strTitulo3;
        string strData;
        string strHora;
        int rCnt;
        int rw = 0;
        int cl = 0;

        public const int KEYEVENTF_EXTENTEDKEY = 1;
        public const int KEYEVENTF_KEYUP = 0;
        public const int VK_MEDIA_NEXT_TRACK = 0xB0;
        public const int VK_MEDIA_PLAY_PAUSE = 0xB3;
        public const int VK_MEDIA_PREV_TRACK = 0xB1;

        private const string IniPath = "Audio.ini";
        IniFile MyIni = null;

        [DllImport("user32.dll")]
        public static extern void keybd_event(byte virtualKey, byte scanCode, uint flags, IntPtr extraInfo);

        public Form1()
        {
            InitializeComponent();
            isPlaying = false;
            timer.Interval = 1000;
            timer.Start();
            DateTime dtNow = DateTime.Now;
            //int day1 = (int)dtNow.DayOfWeek;

            String[] aux = IniFile.ReadKeyValuePairs(dtNow.ToString("yyyy-MM-dd"), System.IO.Path.GetFullPath(@"Audio.ini"));
            if (aux == null)
            {
                MessageBox.Show("Dia sem horário. Por favor carregar o ficheiro excel com o horario do OVG.");
            }
            else
            {
                string joined = string.Join("\n", aux);
                lblhorario.Text = joined;
                readIniFile();
            }
        }

        //Le os dados que estao no ficheiro Audio.ini e se tiver valor,  ReproduzirSom(value);
        private void readIniFile()
        {
            DateTime dtNow = DateTime.Now.AddMinutes(iMinutesBeforeClass);
            //int day1 = (int)dtNow.DayOfWeek;
            //String dayWeek = day1.ToString();
            String strTime = getTwoDigits(dtNow.Hour.ToString())+ ":" +getTwoDigits(dtNow.Minute.ToString());
            String value = IniFile.ReadValue(dtNow.ToString("yyyy-MM-dd"), strTime, System.IO.Path.GetFullPath(@"Audio.ini"));

            if (!String.IsNullOrEmpty(value))
            {
                ReproduzirSom(value);
            } 
        }

        //
        /// <summary>
        /// Simula a tecla de pause do teclado
        /// Vai executar no wmplayer o ficheiro mp3 correspondente ao nome do ficheiro ini
        /// 
        /// </summary>
        /// <param name="value">Nome do ficheiro a ser lido</param>
        private void ReproduzirSom(String value)
        {
            keybd_event(VK_MEDIA_PLAY_PAUSE, 0, KEYEVENTF_EXTENTEDKEY, IntPtr.Zero);
            isPlaying = true;
            WMPLib.WindowsMediaPlayer wplayer = new WMPLib.WindowsMediaPlayer();
            wplayer.URL = @"Alarmes/" + value + ".mp4";
            wplayer.controls.play();

            //WMPLib.IWMPMedia media = wplayer.newMedia(wplayer.URL);
            //pauseSpotify = Convert.ToInt32(media.duration) + restSilence;

            pauseSpotify = int.Parse(IniFile.ReadValue("setup", "pause", System.IO.Path.GetFullPath(@"setup.ini")));
        }

        private void Count()
        {
            counter++;
            if (counter == refreshTime)
            {
                counter = 0;
                readIniFile();
            }
            if (counter == pauseSpotify && isPlaying)//get back music
            {
                keybd_event(VK_MEDIA_PLAY_PAUSE, 0, KEYEVENTF_EXTENTEDKEY, IntPtr.Zero);
                isPlaying = false;
            }
        }

        private void Timer_Tick(object sender, EventArgs e)
        {
            Count();
        }

        private String getTwoDigits(String strTime)
        {
            if (strTime.Length == 1)
            {
                return "0" + strTime;
            }
            return strTime;
        }

        
        /// <summary>
        /// le o ficheiro setup.ini e atualiza o ficheiro Audio.ini e pasta de Sons
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BtnAtualizar_Click(object sender, EventArgs e)
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
                MessageBox.Show("Atualização concluída.");
            }
        }

        private void readExcel(String filePath)
        {
            MyIni = new IniFile(IniPath);

            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            Excel.Range range;

 

            xlApp = new Excel.Application();
            xlWorkBook = xlApp.Workbooks.Open(filePath, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

            range = xlWorkSheet.UsedRange;
            rw = range.Rows.Count;
            cl = range.Columns.Count;

            segunda(range);
            terca(range);
            quarta(range);
            quinta(range);
            sexta(range);
            sabado(range);
            domingo(range);

            xlWorkBook.Close(true, null, null);
            xlApp.Quit();

            Marshal.ReleaseComObject(xlWorkSheet);
            Marshal.ReleaseComObject(xlWorkBook);
            Marshal.ReleaseComObject(xlApp);

        }

        private void segunda(Excel.Range range)
        {
            strData = (string)(range.Cells[3, 2] as Excel.Range).Value2;
            for (rCnt = 4; rCnt <= rw; rCnt++)
            {
                strTitulo1 = (string)(range.Cells[rCnt, 2] as Excel.Range).Value2;
                strTitulo2 = (string)(range.Cells[rCnt, 3] as Excel.Range).Value2;
                strTitulo3 = (string)(range.Cells[rCnt, 4] as Excel.Range).Value2;
                //strData = (string)(range.Cells[rCnt, 5] as Excel.Range).Value2;
                //strHora = Convert.ToString((range.Cells[rCnt, 1] as Excel.Range).Value2);
                double d = System.Convert.ToDouble((range.Cells[rCnt, 1] as Excel.Range).Value2);
                DateTime conv = DateTime.FromOADate(d);
                strHora = conv.ToString("HH:mm");

                if (!String.IsNullOrEmpty(strTitulo1))
                {
                    regexString(strData, strHora, strTitulo1);
                }
                if (!String.IsNullOrEmpty(strTitulo2))
                {
                    regexString(strData, strHora, strTitulo2);
                }
                if (!String.IsNullOrEmpty(strTitulo3))
                {
                    regexString(strData, strHora, strTitulo3);
                }

            }
        }

        private void terca(Excel.Range range)
        {
            string strTitulo4;
            strData = (string)(range.Cells[3, 5] as Excel.Range).Value2;
            for (rCnt = 4; rCnt <= rw; rCnt++)
            {
                strTitulo1 = (string)(range.Cells[rCnt, 5] as Excel.Range).Value2;
                strTitulo2 = (string)(range.Cells[rCnt, 6] as Excel.Range).Value2;
                strTitulo3 = (string)(range.Cells[rCnt, 7] as Excel.Range).Value2;
                //strTitulo4 = (string)(range.Cells[rCnt, 8] as Excel.Range).Value2;
                
                double d = System.Convert.ToDouble((range.Cells[rCnt, 1] as Excel.Range).Value2);
                DateTime conv = DateTime.FromOADate(d);
                strHora = conv.ToString("HH:mm");

                if (!String.IsNullOrEmpty(strTitulo1))
                {
                    regexString(strData, strHora, strTitulo1);
                }
                if (!String.IsNullOrEmpty(strTitulo2))
                {
                    regexString(strData, strHora, strTitulo2);
                }
                if (!String.IsNullOrEmpty(strTitulo3))
                {
                    regexString(strData, strHora, strTitulo3);
                }
                /*if (!String.IsNullOrEmpty(strTitulo4))
                {
                    regexString(strData, strHora, strTitulo4);
                }*/
            }
        }

        private void quarta(Excel.Range range)
        {
            strData = (string)(range.Cells[3, 8] as Excel.Range).Value2;
            for (rCnt = 4; rCnt <= rw; rCnt++)
            {
                strTitulo1 = (string)(range.Cells[rCnt, 8] as Excel.Range).Value2;
                strTitulo2 = (string)(range.Cells[rCnt, 9] as Excel.Range).Value2;
                strTitulo3 = (string)(range.Cells[rCnt,10] as Excel.Range).Value2;
                //strData = (string)(range.Cells[rCnt, 5] as Excel.Range).Value2;
                //strHora = Convert.ToString((range.Cells[rCnt, 1] as Excel.Range).Value2);
                double d = System.Convert.ToDouble((range.Cells[rCnt, 1] as Excel.Range).Value2);
                DateTime conv = DateTime.FromOADate(d);
                strHora = conv.ToString("HH:mm");

                if (!String.IsNullOrEmpty(strTitulo1))
                {
                    regexString(strData, strHora, strTitulo1);
                }
                if (!String.IsNullOrEmpty(strTitulo2))
                {
                    regexString(strData, strHora, strTitulo2);
                }
                if (!String.IsNullOrEmpty(strTitulo3))
                {
                    regexString(strData, strHora, strTitulo3);
                }

            }
        }

        private void quinta(Excel.Range range)
        {
            strData = (string)(range.Cells[3, 11] as Excel.Range).Value2;
            for (rCnt = 4; rCnt <= rw; rCnt++)
            {
                strTitulo1 = (string)(range.Cells[rCnt, 11] as Excel.Range).Value2;
                strTitulo2 = (string)(range.Cells[rCnt, 12] as Excel.Range).Value2;
                strTitulo3 = (string)(range.Cells[rCnt, 13] as Excel.Range).Value2;
                //strData = (string)(range.Cells[rCnt, 5] as Excel.Range).Value2;
                //strHora = Convert.ToString((range.Cells[rCnt, 1] as Excel.Range).Value2);
                double d = System.Convert.ToDouble((range.Cells[rCnt, 1] as Excel.Range).Value2);
                DateTime conv = DateTime.FromOADate(d);
                strHora = conv.ToString("HH:mm");

                if (!String.IsNullOrEmpty(strTitulo1))
                {
                    regexString(strData, strHora, strTitulo1);
                }
                if (!String.IsNullOrEmpty(strTitulo2))
                {
                    regexString(strData, strHora, strTitulo2);
                }
                if (!String.IsNullOrEmpty(strTitulo3))
                {
                    regexString(strData, strHora, strTitulo3);
                }

            }
        }

        private void sexta(Excel.Range range)
        {
            string strTitulo4;
            strData = (string)(range.Cells[3, 14] as Excel.Range).Value2;
            for (rCnt = 4; rCnt <= rw; rCnt++)
            {
                strTitulo1 = (string)(range.Cells[rCnt, 14] as Excel.Range).Value2;
                strTitulo2 = (string)(range.Cells[rCnt, 15] as Excel.Range).Value2;
                strTitulo3 = (string)(range.Cells[rCnt, 16] as Excel.Range).Value2;
                //strTitulo4 = (string)(range.Cells[rCnt, 18] as Excel.Range).Value2;

                double d = System.Convert.ToDouble((range.Cells[rCnt, 1] as Excel.Range).Value2);
                DateTime conv = DateTime.FromOADate(d);
                strHora = conv.ToString("HH:mm");

                if (!String.IsNullOrEmpty(strTitulo1))
                {
                    regexString(strData, strHora, strTitulo1);
                }
                if (!String.IsNullOrEmpty(strTitulo2))
                {
                    regexString(strData, strHora, strTitulo2);
                }
                if (!String.IsNullOrEmpty(strTitulo3))
                {
                    regexString(strData, strHora, strTitulo3);
                }
                /*if (!String.IsNullOrEmpty(strTitulo4))
                {
                    regexString(strData, strHora, strTitulo4);
                }*/
            }
        }

        private void sabado(Excel.Range range)
        {
            strData = (string)(range.Cells[3, 17] as Excel.Range).Value2;
            for (rCnt = 4; rCnt <= rw; rCnt++)
            {
                strTitulo1 = (string)(range.Cells[rCnt, 17] as Excel.Range).Value2;
                strTitulo2 = (string)(range.Cells[rCnt, 18] as Excel.Range).Value2;
                strTitulo3 = (string)(range.Cells[rCnt, 19] as Excel.Range).Value2;
                //strData = (string)(range.Cells[rCnt, 5] as Excel.Range).Value2;
                //strHora = Convert.ToString((range.Cells[rCnt, 1] as Excel.Range).Value2);
                double d = System.Convert.ToDouble((range.Cells[rCnt, 1] as Excel.Range).Value2);
                DateTime conv = DateTime.FromOADate(d);
                strHora = conv.ToString("HH:mm");

                if (!String.IsNullOrEmpty(strTitulo1))
                {
                    regexString(strData, strHora, strTitulo1);
                }
                if (!String.IsNullOrEmpty(strTitulo2))
                {
                    regexString(strData, strHora, strTitulo2);
                }
                if (!String.IsNullOrEmpty(strTitulo3))
                {
                    regexString(strData, strHora, strTitulo3);
                }

            }
        }

        private void domingo(Excel.Range range)
        {
            strData = (string)(range.Cells[3, 20] as Excel.Range).Value2;
            for (rCnt = 4; rCnt <= rw; rCnt++)
            {
                strTitulo1 = (string)(range.Cells[rCnt, 20] as Excel.Range).Value2;
                strTitulo2 = (string)(range.Cells[rCnt, 21] as Excel.Range).Value2;
                strTitulo3 = (string)(range.Cells[rCnt, 22] as Excel.Range).Value2;
                //strData = (string)(range.Cells[rCnt, 5] as Excel.Range).Value2;
                //strHora = Convert.ToString((range.Cells[rCnt, 1] as Excel.Range).Value2);
                double d = System.Convert.ToDouble((range.Cells[rCnt, 1] as Excel.Range).Value2);
                DateTime conv = DateTime.FromOADate(d);
                strHora = conv.ToString("HH:mm");

                if (!String.IsNullOrEmpty(strTitulo1))
                {
                    regexString(strData, strHora, strTitulo1);
                }
                if (!String.IsNullOrEmpty(strTitulo2))
                {
                    regexString(strData, strHora, strTitulo2);
                }
                if (!String.IsNullOrEmpty(strTitulo3))
                {
                    regexString(strData, strHora, strTitulo3);
                }

            }
        }

        

        private void regexString(String strData, String strHora, String strTitulo)
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

        /**
         * 
         * Codigo antigo para atualizar da pasta partilhada para a pasta do projeto
         * 
         */
        private void AtualizarFicheiros()
        {
            String value = IniFile.ReadValue("setup", "local", System.IO.Path.GetFullPath(@"setup.ini"));

            if (!String.IsNullOrEmpty(value))
            {
                string sourceDirectory = @"P:\" + value;
                string targetDirectory = System.IO.Path.GetDirectoryName(Assembly.GetEntryAssembly().Location);
                var diSource = new DirectoryInfo(sourceDirectory);
                var diTarget = new DirectoryInfo(targetDirectory);

                CopyAll(diSource, diTarget);
            }
        }

        /// <summary>
        /// Copia e substitui todos os ficheiros para a diretoria onde está a ser executado
        /// </summary>
        /// <param name="source"></param>
        /// <param name="target"></param>
        public static void CopyAll(DirectoryInfo source, DirectoryInfo target)
        {
            Directory.CreateDirectory(target.FullName);

            // Copy each file into the new directory.
            foreach (FileInfo fi in source.GetFiles())
            {
                Console.WriteLine(@"Copying {0}\{1}", target.FullName, fi.Name);
                fi.CopyTo(Path.Combine(target.FullName, fi.Name), true);
            }

            // Copy each subdirectory using recursion.
            foreach (DirectoryInfo diSourceSubDir in source.GetDirectories())
            {
                DirectoryInfo nextTargetSubDir =
                    target.CreateSubdirectory(diSourceSubDir.Name);
                CopyAll(diSourceSubDir, nextTargetSubDir);
            }
        }


        /***
 * 
 * Controlo do volume
 * 
 */

        private const int APPCOMMAND_VOLUME_MUTE = 0x80000;
        private const int APPCOMMAND_VOLUME_UP = 0xA0000;
        private const int APPCOMMAND_VOLUME_DOWN = 0x90000;
        private const int WM_APPCOMMAND = 0x319;

        [DllImport("user32.dll")]
        public static extern IntPtr SendMessageW(IntPtr hWnd, int Msg,
            IntPtr wParam, IntPtr lParam);

        private void Mute()
        {
            SendMessageW(this.Handle, WM_APPCOMMAND, this.Handle,
                (IntPtr)APPCOMMAND_VOLUME_MUTE);
        }

        private void VolDown()
        {
            SendMessageW(this.Handle, WM_APPCOMMAND, this.Handle,
                (IntPtr)APPCOMMAND_VOLUME_DOWN);
        }

        private void VolUp()
        {
            SendMessageW(this.Handle, WM_APPCOMMAND, this.Handle,
                (IntPtr)APPCOMMAND_VOLUME_UP);
        }

        private void BtnMais_Click(object sender, EventArgs e)
        {
            VolUp();
        }

        private void BtnMenos_Click(object sender, EventArgs e)
        {
            VolDown();
        }

        private void BtnMute_Click(object sender, EventArgs e)
        {
            const string app = "Mozilla Firefox";

            foreach (string name in EnumerateApplications())
            {
                Console.WriteLine("name:" + name);
                if (name == app)
                {
                    // display mute state & volume level (% of master)
                    Console.WriteLine("Mute:" + GetApplicationMute(app));
                    Console.WriteLine("Volume:" + GetApplicationVolume(app));

                    // mute the application
                    SetApplicationMute(app, true);

                    // set the volume to half of master volume (50%)
                    SetApplicationVolume(app, 10);
                }
            }
            //Mute();
        }


        public static float? GetApplicationVolume(string name)
        {
            ISimpleAudioVolume volume = GetVolumeObject(name);
            if (volume == null)
                return null;

            float level;
            volume.GetMasterVolume(out level);
            return level * 100;
        }

        public static bool? GetApplicationMute(string name)
        {
            ISimpleAudioVolume volume = GetVolumeObject(name);
            if (volume == null)
                return null;

            bool mute;
            volume.GetMute(out mute);
            return mute;
        }

        public static void SetApplicationVolume(string name, float level)
        {
            ISimpleAudioVolume volume = GetVolumeObject(name);
            if (volume == null)
                return;

            Guid guid = Guid.Empty;
            volume.SetMasterVolume(level / 100, ref guid);
        }

        public static void SetApplicationMute(string name, bool mute)
        {
            ISimpleAudioVolume volume = GetVolumeObject(name);
            if (volume == null)
                return;

            Guid guid = Guid.Empty;
            volume.SetMute(mute, ref guid);
        }

        public static IEnumerable<string> EnumerateApplications()
        {
            // get the speakers (1st render + multimedia) device
            IMMDeviceEnumerator deviceEnumerator = (IMMDeviceEnumerator)(new MMDeviceEnumerator());
            IMMDevice speakers;
            deviceEnumerator.GetDefaultAudioEndpoint(EDataFlow.eRender, ERole.eMultimedia, out speakers);

            // activate the session manager. we need the enumerator
            Guid IID_IAudioSessionManager2 = typeof(IAudioSessionManager2).GUID;
            object o;
            speakers.Activate(ref IID_IAudioSessionManager2, 0, IntPtr.Zero, out o);
            IAudioSessionManager2 mgr = (IAudioSessionManager2)o;

            // enumerate sessions for on this device
            IAudioSessionEnumerator sessionEnumerator;
            mgr.GetSessionEnumerator(out sessionEnumerator);
            int count;
            sessionEnumerator.GetCount(out count);

            for (int i = 0; i < count; i++)
            {
                IAudioSessionControl ctl;
                sessionEnumerator.GetSession(i, out ctl);
                string dn;
                ctl.GetDisplayName(out dn);
                yield return dn;
                Marshal.ReleaseComObject(ctl);
            }
            Marshal.ReleaseComObject(sessionEnumerator);
            Marshal.ReleaseComObject(mgr);
            Marshal.ReleaseComObject(speakers);
            Marshal.ReleaseComObject(deviceEnumerator);
        }

        private static ISimpleAudioVolume GetVolumeObject(string name)
        {
            // get the speakers (1st render + multimedia) device
            IMMDeviceEnumerator deviceEnumerator = (IMMDeviceEnumerator)(new MMDeviceEnumerator());
            IMMDevice speakers;
            deviceEnumerator.GetDefaultAudioEndpoint(EDataFlow.eRender, ERole.eMultimedia, out speakers);

            // activate the session manager. we need the enumerator
            Guid IID_IAudioSessionManager2 = typeof(IAudioSessionManager2).GUID;
            object o;
            speakers.Activate(ref IID_IAudioSessionManager2, 0, IntPtr.Zero, out o);
            IAudioSessionManager2 mgr = (IAudioSessionManager2)o;

            // enumerate sessions for on this device
            IAudioSessionEnumerator sessionEnumerator;
            mgr.GetSessionEnumerator(out sessionEnumerator);
            int count;
            sessionEnumerator.GetCount(out count);

            // search for an audio session with the required name
            // NOTE: we could also use the process id instead of the app name (with IAudioSessionControl2)
            ISimpleAudioVolume volumeControl = null;
            for (int i = 0; i < count; i++)
            {
                IAudioSessionControl ctl;
                sessionEnumerator.GetSession(i, out ctl);
                string dn;
                ctl.GetDisplayName(out dn);
                if (string.Compare(name, dn, StringComparison.OrdinalIgnoreCase) == 0)
                {
                    volumeControl = ctl as ISimpleAudioVolume;
                    break;
                }
                Marshal.ReleaseComObject(ctl);
            }
            Marshal.ReleaseComObject(sessionEnumerator);
            Marshal.ReleaseComObject(mgr);
            Marshal.ReleaseComObject(speakers);
            Marshal.ReleaseComObject(deviceEnumerator);
            return volumeControl;
        }
    }

    [ComImport]
    [Guid("BCDE0395-E52F-467C-8E3D-C4579291692E")]
    internal class MMDeviceEnumerator
    {
    }

    internal enum EDataFlow
    {
        eRender,
        eCapture,
        eAll,
        EDataFlow_enum_count
    }

    internal enum ERole
    {
        eConsole,
        eMultimedia,
        eCommunications,
        ERole_enum_count
    }

    [Guid("A95664D2-9614-4F35-A746-DE8DB63617E6"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface IMMDeviceEnumerator
    {
        int NotImpl1();

        [PreserveSig]
        int GetDefaultAudioEndpoint(EDataFlow dataFlow, ERole role, out IMMDevice ppDevice);

        // the rest is not implemented
    }

    [Guid("D666063F-1587-4E43-81F1-B948E807363F"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface IMMDevice
    {
        [PreserveSig]
        int Activate(ref Guid iid, int dwClsCtx, IntPtr pActivationParams, [MarshalAs(UnmanagedType.IUnknown)] out object ppInterface);

        // the rest is not implemented
    }

    [Guid("77AA99A0-1BD6-484F-8BC7-2C654C9A9B6F"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface IAudioSessionManager2
    {
        int NotImpl1();
        int NotImpl2();

        [PreserveSig]
        int GetSessionEnumerator(out IAudioSessionEnumerator SessionEnum);

        // the rest is not implemented
    }

    [Guid("E2F5BB11-0570-40CA-ACDD-3AA01277DEE8"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface IAudioSessionEnumerator
    {
        [PreserveSig]
        int GetCount(out int SessionCount);

        [PreserveSig]
        int GetSession(int SessionCount, out IAudioSessionControl Session);
    }

    [Guid("F4B1A599-7266-4319-A8CA-E70ACB11E8CD"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface IAudioSessionControl
    {
        int NotImpl1();

        [PreserveSig]
        int GetDisplayName([MarshalAs(UnmanagedType.LPWStr)] out string pRetVal);

        // the rest is not implemented
    }

    [Guid("87CE5498-68D6-44E5-9215-6DA47EF883D8"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface ISimpleAudioVolume
    {
        [PreserveSig]
        int SetMasterVolume(float fLevel, ref Guid EventContext);

        [PreserveSig]
        int GetMasterVolume(out float pfLevel);

        [PreserveSig]
        int SetMute(bool bMute, ref Guid EventContext);

        [PreserveSig]
        int GetMute(out bool pbMute);
    }

}
