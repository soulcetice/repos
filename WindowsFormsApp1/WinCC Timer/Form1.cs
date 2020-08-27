using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SqlClient;
using WinCC_Timer.Properties;

namespace WinCC_Timer
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        public DataTable dataTable = new DataTable();

        private void button1_Click(object sender, EventArgs e)
        {
            string connectionString = "Data Source=TCMHMID01\\WINCC;Initial Catalog=SMS_RTDesign;Integrated Security=SSPI";
            string query = "SELECT * FROM RT_Menu where RefId = 1";
            SqlConnection cnn = new SqlConnection(connectionString);
            try
            {
                cnn.Open();
                //MessageBox.Show("Connection Open ! ");
                SqlCommand cmd = new SqlCommand(query, cnn);
                SqlDataAdapter da = new SqlDataAdapter(cmd);

                da.Fill(dataTable);

                for (int i = 0; i < dataTable.Rows.Count; i++)
                {
                    var row = dataTable.Rows[i].ItemArray;
                    string data = "";
                    foreach (var c in row)
                    {
                        data += c.ToString() + " ";
                    }
                    LogToFile(data, "\\Screen.logger");
                }

                cnn.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Can not open connection ! " + ex.Message);
            }

            FindObjectInHMI();
        }

        private void FindObjectInHMI()
        {
            DateTime start = DateTime.UtcNow;
            LogToFile("", "\\Screen.logger");

            Bitmap bmp = TakeScreenShot();
            Bitmap cmp = (Bitmap)Resources.ResourceManager.GetObject("s");

            DateTime end = DateTime.UtcNow;
            LogToFile(CompareMemCmp(bmp, cmp).ToString(), "\\Screen.logger");
            LogToFile((end.Second - start.Second).ToString() + "." + (end.Millisecond - start.Millisecond).ToString(), "\\Screen.logger");
        }

        [DllImport("msvcrt.dll")]
        private static extern int memcmp(IntPtr b1, IntPtr b2, long count);

        public static bool CompareMemCmp(Bitmap b1, Bitmap b2)
        {
            if ((b1 == null) != (b2 == null)) return false;
            if (b1.Size != b2.Size) return false;

            var bd1 = b1.LockBits(new Rectangle(new Point(0, 0), b1.Size), ImageLockMode.ReadOnly, PixelFormat.Format32bppArgb);
            var bd2 = b2.LockBits(new Rectangle(new Point(0, 0), b2.Size), ImageLockMode.ReadOnly, PixelFormat.Format32bppArgb);

            try
            {
                IntPtr bd1scan0 = bd1.Scan0;
                IntPtr bd2scan0 = bd2.Scan0;

                int stride = bd1.Stride;
                int len = stride * b1.Height;

                return memcmp(bd1scan0, bd2scan0, len) == 0;
            }
            finally
            {
                b1.UnlockBits(bd1);
                b2.UnlockBits(bd2);
            }
        }

        private static void LogToFile(string content, string fname)
        {
            using (var fileWriter = new StreamWriter(Application.StartupPath + fname, true))
            {
                DateTime date = DateTime.UtcNow;
                fileWriter.WriteLine(date.Year + "/" + date.Month + "/" + date.Day + " " + date.Hour + ":" + date.Minute + ":" + date.Second + ":" + date.Millisecond + " UTC: " + content);
                fileWriter.Close();
            }
        }

        private Bitmap TakeScreenShot()
        {
            int width = 5;
            int height = 5;
            int top = 989 + 15;
            int left = 1666 + 15;
            Bitmap bmp = new Bitmap(width, height);
            using (Graphics g = Graphics.FromImage(bmp))
            {
                Size s = new Size()
                {
                    Width = width, //Screen.PrimaryScreen.Bounds.Size.Width - width,
                    Height = height //Screen.PrimaryScreen.Bounds.Size.Height - height
                };
                g.CopyFromScreen(left, top, 0, 0, s);
                bmp.Save("s.png");  // saves the image
            }
            return bmp;
        }
    }
}
