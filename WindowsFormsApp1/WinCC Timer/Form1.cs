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
using Interoperability;

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

            List<MenuRow> lists = GetSQLMenu("1");

            var tier1 = lists.Where(c => c.Layer == "1");

            foreach (var t in tier1)
            {
                LogToFile(GetTextWidth(t.Caption, "Arial", 10.0f).ToString() + " " + t.Caption, "\\Screen.logger"); ;
            }

            IntPtr rt = PInvokeLibrary.FindWindow("PDLRTisAliveAndWaitsForYou", "WinCC-Runtime - ");

            //"General is 76 * 30
            int x = 25/2;
            int y = 15;

            foreach (MenuRow m in tier1)
            {
                x += GetTextWidth(m.Caption, "Arial", 10.0f) / 2;
                ClickInWindowAtXY(rt, x, y, 1);
                System.Threading.Thread.Sleep(2000);
                ClickInWindowAtXY(rt, x, y, 1);
                x += GetTextWidth(m.Caption, "Arial", 10.0f) / 2;
                x += 25;
                //still only almost, gets too offset in the end
            }

            bool v = FindObjectInHMI();
        }

        private static int GetTextWidth(string text, string fontName, Single fontSize)
        {
            var font = new Font(fontName, fontSize, FontStyle.Regular);
            int textWidth = TextRenderer.MeasureText(text, font).Width;
            return textWidth;
        }

        private void ClickInWindowAtXY(IntPtr handle, int? x, int? y, int repeat)
        {
            for (int i = 0; i < repeat; i++)
            {
                PInvokeLibrary.SetForegroundWindow(handle);

                MouseOperations.SetCursorPosition(x.Value, y.Value); //have to use the found minus/plus coordinates here
                MouseOperations.MouseEvent(MouseOperations.MouseEventFlags.LeftDown);
                MouseOperations.MouseEvent(MouseOperations.MouseEventFlags.LeftUp);
            }
        }

        public class MenuRow
        {
            public string ID;
            public string RefId;
            public string Layer;
            public string Pos;
            public string LCID;
            public string ParentId;
            public string Caption;
            public string Flags;
            public string Pdl;
            public string Parameter;
        }

        private List<MenuRow> GetSQLMenu(string id)
        {
            string connectionString = "Data Source=TCMHMID01\\WINCC;Initial Catalog=SMS_RTDesign;Integrated Security=SSPI";
            string query = "SELECT * FROM RT_Menu where RefId = " + id;
            SqlConnection cnn = new SqlConnection(connectionString);
            try
            {
                cnn.Open();
                //MessageBox.Show("Connection Open ! ");
                SqlCommand cmd = new SqlCommand(query, cnn);
                SqlDataAdapter da = new SqlDataAdapter(cmd);
                da.Fill(dataTable);
                cnn.Close();

                List<MenuRow> myData = new List<MenuRow>();

                for (int i = 0; i < dataTable.Rows.Count; i++)
                {
                    var row = dataTable.Rows[i].ItemArray;
                    myData.Add(new MenuRow()
                    {
                        ID = row?.ElementAt(0).ToString(),
                        RefId = row?.ElementAt(1).ToString(),
                        Layer = row?.ElementAt(2).ToString(),
                        Pos = row?.ElementAt(3).ToString(),
                        LCID = row?.ElementAt(4).ToString(),
                        ParentId = row?.ElementAt(5).ToString(),
                        Caption = row?.ElementAt(6).ToString(),
                        Flags = row?.ElementAt(7).ToString(),
                        Pdl = row?.ElementAt(8).ToString(),
                        Parameter = row?.ElementAt(9).ToString(),
                    });
                }

                return myData;
            }
            catch (Exception ex)
            {
                LogToFile("Can not open connection ! " + ex.Message, "\\Screen.logger");
                return null;
            }
        }

        private bool FindObjectInHMI()
        {
            DateTime start = DateTime.UtcNow;
            LogToFile("", "\\Screen.logger");

            Bitmap bmp = TakeScreenShot();
            Bitmap cmp = (Bitmap)Resources.ResourceManager.GetObject("s");

            DateTime end = DateTime.UtcNow;

            bool t = CompareMemCmp(bmp, cmp);
            LogToFile(t.ToString(), "\\Screen.logger");
            LogToFile((end.Second - start.Second).ToString() + "." + (end.Millisecond - start.Millisecond).ToString(), "\\Screen.logger");

            return t;
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
            int left = 1666 + 15; //1506 is 1681 (175 offset)
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
