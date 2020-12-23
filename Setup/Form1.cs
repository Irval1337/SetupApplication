using System;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.IO;
using System.Net;
using Ionic.Zip;
using System.Diagnostics;
using System.Drawing.Text;
using NotificationManager;
using System.ComponentModel;

namespace Setup
{
    public partial class Form1 : Form
    {
        Manager manager;

        public const int WM_NCLBUTTONDOWN = 0xA1;
        public const int HT_CAPTION = 0x2;

        [DllImportAttribute("user32.dll")]
        public static extern int SendMessage(IntPtr hWnd, int Msg, int wParam, int lParam);
        [DllImportAttribute("user32.dll")]
        public static extern bool ReleaseCapture();

        string path = "";

        public Form1()
        {
            manager = new Manager();
            InitializeComponent();
            LoadFonts();
            foreach (Control control in this.Controls)
                control.Font = new Font(private_fonts.Families[0], control.Font.Size);
            label6.Font = new Font(private_fonts.Families[0], label6.Font.Size);
            bunifuTextBox1.DefaultFont = new Font(private_fonts.Families[0].Name, (float)bunifuTextBox1.DefaultFont.Size);
        }

        private void panel1_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                ReleaseCapture();
                SendMessage(Handle, WM_NCLBUTTONDOWN, HT_CAPTION, 0);
            }
        }

        private void bunifuImageButton1_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void bunifuImageButton2_Click(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Minimized;
        }

        private void bunifuButton1_Click(object sender, EventArgs e)
        {
            using (FolderBrowserDialog folderBrowser = new FolderBrowserDialog())
            {
                if (folderBrowser.ShowDialog() == DialogResult.OK)
                {
                    var Filename = folderBrowser.SelectedPath;
                    bunifuTextBox1.Text = Filename;
                }
            }
        }

        private void bunifuButton2_Click(object sender, EventArgs e)
        {
            if (!Directory.Exists(bunifuTextBox1.Text))
                manager.Alert("Указанный путь не найден", NotificationType.Error);
            else
            {
                path = bunifuTextBox1.Text;
                tabControl1.SelectedIndex++;
                pictureBox3.Location = new Point(pictureBox3.Location.X, pictureBox3.Location.Y + 50);
            }
        }

        private void bunifuButton4_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void bunifuButton3_Click(object sender, EventArgs e)
        {
            tabControl1.SelectedIndex++;
        }

        private void bunifuButton5_Click(object sender, EventArgs e)
        {
            tabControl1.SelectedIndex--;
            pictureBox3.Location = new Point(pictureBox3.Location.X, pictureBox3.Location.Y - 50);
        }

        private void bunifuButton6_Click(object sender, EventArgs e)
        {
            tabControl1.SelectedIndex++;
            pictureBox3.Location = new Point(pictureBox3.Location.X, pictureBox3.Location.Y + 50);
        }

        private void bunifuButton7_Click(object sender, EventArgs e)
        {
            tabControl1.SelectedIndex--;
            pictureBox3.Location = new Point(pictureBox3.Location.X, pictureBox3.Location.Y - 50);
        }

        private void bunifuButton8_Click(object sender, EventArgs e)
        {
            richTextBox2.Text = $"Папка установки:\n\t{path}";
            if (bunifuCheckBox1.Checked || bunifuCheckBox2.Checked)
            {
                richTextBox2.Text += "\n\nДополнительные задачи:\n";
                richTextBox2.Text += bunifuCheckBox1.Checked ? "\tСоздать ярлык на рабочем столе\n" : "";
                richTextBox2.Text += bunifuCheckBox2.Checked ? "\tОткрыть официальную тему программы на DataStock.biz" : "";
            }
            tabControl1.SelectedIndex++;
            pictureBox3.Location = new Point(pictureBox3.Location.X, pictureBox3.Location.Y + 45);
        }

        private void bunifuButton9_Click(object sender, EventArgs e)
        {
            tabControl1.SelectedIndex--;
            pictureBox3.Location = new Point(pictureBox3.Location.X, pictureBox3.Location.Y - 45);
        }

        private void bunifuButton10_Click(object sender, EventArgs e)
        {
            try
            {
                string link = @"Ваша ссылка на архив с программой";
                Directory.CreateDirectory(path + @"\DataStock BOT");
                WebClient webClient = new WebClient();
                webClient.DownloadFileCompleted += ((object se, AsyncCompletedEventArgs ev) => {
                    using (Ionic.Zip.ZipFile zip = Ionic.Zip.ZipFile.Read(path + @"\DataStock BOT\datastock.zip"))
                    {
                        foreach (ZipEntry evu in zip)
                        {
                            evu.Extract(path + @"\DataStock BOT", ExtractExistingFileAction.OverwriteSilently);
                        }
                        zip.Dispose();
                    }
                    File.Delete(path + @"\DataStock BOT\datastock.zip");

                    if (bunifuCheckBox1.Checked)
                    {
                        Type t = Type.GetTypeFromCLSID(new Guid("72C24DD5-D70A-438B-8A42-98424B88AFB8")); //Windows Script Host Shell Object
                        dynamic shell = Activator.CreateInstance(t);
                        try
                        {
                            var lnk = shell.CreateShortcut(Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory) + @"\DataStock BOT.lnk");
                            try
                            {
                                lnk.TargetPath = path + @"\DataStock BOT\DataStock.exe";
                                lnk.IconLocation = path + @"\DataStock BOT\logo.ico";
                                lnk.Save();
                            }
                            finally
                            {
                                Marshal.FinalReleaseComObject(lnk);
                            }
                        }
                        finally
                        {
                            Marshal.FinalReleaseComObject(shell);
                        }
                    }

                    tabControl1.SelectedIndex++;
                    pictureBox3.Location = new Point(pictureBox3.Location.X, pictureBox3.Location.Y + 40);
                });
                webClient.DownloadFileAsync(new Uri(link), path + @"\DataStock BOT\datastock.zip");
                manager.Alert("Установка начата", NotificationType.Success);
            }
            catch { manager.Alert("Ошибка во время установки", NotificationType.Error); }          
        }

        private void bunifuButton11_Click(object sender, EventArgs e)
        {
            if (bunifuCheckBox2.Checked)
                Process.Start("https://datastock.biz/threads/1980/");
            if (bunifuCheckBox3.Checked)
                Process.Start(path + @"\DataStock BOT\DataStock.exe");
            Application.Exit();
        }

        PrivateFontCollection private_fonts = new PrivateFontCollection();

        private void LoadFonts()
        {
            using (MemoryStream fontStream = new MemoryStream(Properties.Resources.CenturyGothic))
            {
                System.IntPtr data = Marshal.AllocCoTaskMem((int)fontStream.Length);
                byte[] fontdata = new byte[fontStream.Length];
                fontStream.Read(fontdata, 0, (int)fontStream.Length);
                Marshal.Copy(fontdata, 0, data, (int)fontStream.Length);
                private_fonts.AddMemoryFont(data, (int)fontStream.Length);
                Marshal.FreeCoTaskMem(data);
            }
        }
    }
}
