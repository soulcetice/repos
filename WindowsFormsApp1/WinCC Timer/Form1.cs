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
        public string logName = "\\Screen.logger";

        private void button1_Click(object sender, EventArgs e)
        {

            List<MenuRow> lists = GetSQLMenu("1");

            var tier1 = lists.Where(c => c.Layer == "1");

            IntPtr rt = PInvokeLibrary.FindWindow("PDLRTisAliveAndWaitsForYou", "WinCC-Runtime - ");
            Bitmap cmp = (Bitmap)Resources.ResourceManager.GetObject("s");

            //"General is 76 * 30
            int x = 25 / 2;
            int y = 0;
            foreach (MenuRow m in tier1)
            {
                Size size = GetTextSize(m.Caption, "Arial", 10.0f);

                x += size.Width / 2;
                y = 15;
                y += 28;

                var ChildrenTier1 = lists.Where(c => c.ParentId == m.ID);
                var tier1Width = GetMenuDropWidth(ChildrenTier1);
                foreach (var tier2 in ChildrenTier1)
                {
                    if (tier2.Pdl != "")
                    {
                        ClickInWindowAtXY(rt, x, 15, 1); System.Threading.Thread.Sleep(500); //expand tier1 menu
                        LogToFile("For " + m.Caption + " expand menu, clicked at " + x + " x, " + 15 + " y", logName);
                        ClickInWindowAtXY(rt, x, y, 1); System.Threading.Thread.Sleep(3000); //open tier2 page
                        LogToFile("For " + tier2.Caption + " expand menu, clicked at " + x + " x, " + y + " y", logName);
                        LogToFile(tier2.Pdl, logName);
                        _ = FindObjectInHMI(cmp);
                        label1.Text = tier2.Pdl;
                        label1.Refresh();
                    }
                    else
                    {
                        var ChildrenTier2 = lists.Where(c => c.ParentId == tier2.ID);
                        ClickInWindowAtXY(rt, GetMenuDropWidth(ChildrenTier2), y, 1); System.Threading.Thread.Sleep(500); //expand tier2 menu
                        LogToFile("For " + tier2.Caption + " expand menu, clicked at " + x + " x, " + y + " y", logName);

                        int yTier3 = y;
                        int maxWidthTier2 = GetMenuDropWidth(ChildrenTier2);
                        LogToFile(maxWidthTier2 + " is maxwidthtier2", logName);
                        foreach (var tier3 in ChildrenTier2)
                        {
                            int xTier3 = x + tier1Width; // + longest element in tier2's width + some 40 pixels
                            if (tier3.Pdl != "")
                            {
                                ClickInWindowAtXY(rt, x, 15, 1); System.Threading.Thread.Sleep(500); //expand tier1 menu
                                LogToFile("For " + m.Caption + " expand menu, clicked at " + x + " x, " + 15 + " y", logName);
                                ClickInWindowAtXY(rt, x, y, 1); System.Threading.Thread.Sleep(500); //expand tier2 menu or open page
                                LogToFile("For " + tier2.Caption + " expand menu, clicked at " + x + " x, " + y + " y", logName);
                                ClickInWindowAtXY(rt, xTier3, yTier3, 1); System.Threading.Thread.Sleep(3000); //expand tier2 menu or open page
                                LogToFile("For " + tier3.Caption + " expand menu, clicked at " + xTier3 + " x, " + yTier3 + " y", logName);

                                LogToFile(tier3.Pdl, logName);
                                _ = FindObjectInHMI(cmp);
                                label1.Text = tier3.Pdl;
                                label1.Refresh();
                            }
                            else
                            {
                                //there is also a tier4....
                                var ChildrenTier3 = lists.Where(c => c.ParentId == tier3.ID);
                                ClickInWindowAtXY(rt, x, y, 1); System.Threading.Thread.Sleep(500); //expand tier2 menu or open page

                                int yTier4 = yTier3;
                                foreach (var tier4 in ChildrenTier3)
                                {
                                    int xTier4 = xTier3 + GetMenuDropWidth(ChildrenTier3) + 20; // + longest element in tier2's width + some 40 pixels
                                    if (tier4.Pdl != "")
                                    {
                                        ClickInWindowAtXY(rt, x, 15, 1); System.Threading.Thread.Sleep(500); //expand tier1 menu
                                        LogToFile("For " + m.Caption + " expand menu, clicked at " + x + " x, " + 15 + " y", logName);
                                        ClickInWindowAtXY(rt, x, y, 1); System.Threading.Thread.Sleep(500); //expand tier2 menu or open page
                                        LogToFile("For " + tier2.Caption + " expand menu, clicked at " + x + " x, " + y + " y", logName);
                                        ClickInWindowAtXY(rt, xTier3, yTier3, 1); System.Threading.Thread.Sleep(500); //expand tier2 menu or open page
                                        LogToFile("For " + tier3.Caption + " expand menu, clicked at " + xTier3 + " x, " + yTier3 + " y", logName);
                                        ClickInWindowAtXY(rt, xTier4, yTier4, 1); System.Threading.Thread.Sleep(3000); //expand tier2 menu or open page
                                        LogToFile("For " + tier4.Caption + " expand menu, clicked at " + xTier4 + " x, " + yTier4 + " y", logName);
                                        LogToFile(tier4.Pdl, logName);
                                        _ = FindObjectInHMI(cmp);
                                        label1.Text = tier4.Pdl;
                                        label1.Refresh();
                                    }
                                    else
                                    {
                                        //no tier 5 thank god
                                    }
                                    yTier4 += 25; //if 26 here, then it becomes too much
                                }
                            }
                            yTier3 += 25;
                        }
                    }
                    y += 25; //increment tier2 menu or open page
                }

                x += size.Width / 2;
                x += 25;

                //bool v = FindObjectInHMI(cmp);
                //still only almost, gets too offset in the end
            }
        }

        private int GetMenuDropWidth(IEnumerable<MenuRow> ChildrenTier2)
        {
            int maxWidthTier2 = 0;
            foreach (var z in ChildrenTier2)
            {
                var w = GetTextSize(z.Caption, "Arial", 10.0f).Width;
                if (w > maxWidthTier2)
                {
                    LogToFile(z.Caption + " is the caption for which we check width", logName);
                    maxWidthTier2 = w;
                }
            }

            LogToFile(maxWidthTier2 + " is x width for childrentier2", logName);
            return maxWidthTier2 + 40;
        }

        private Size GetTextSize(string text, string fontName, Single fontSize)
        {
            Font font = new Font(fontName, fontSize, FontStyle.Regular);
            Size size = TextRenderer.MeasureText(text, font);
            size.Width -= 8; //stupid calculation
            LogToFile(size.Width.ToString() + "width, " + size.Height.ToString() + "height, " + text, logName);
            //26 height for submenus, 30 height for the tier1 menu
            return size;
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
                LogToFile("Can not open connection ! " + ex.Message, logName);
                return null;
            }
        }

        private bool FindObjectInHMI(Bitmap cmp)
        {
            DateTime start = DateTime.UtcNow;
            Bitmap bmp = TakeScreenShot(1004, 1681);
            bool t = CompareMemCmp(bmp, cmp);
            LogToFile((DateTime.UtcNow - start).TotalMilliseconds.ToString() + "ms", logName);
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

        private Bitmap TakeScreenShot(int top, int left)
        {
            int width = 5;
            int height = 5;
            Bitmap bmp = new Bitmap(width, height);
            using (Graphics g = Graphics.FromImage(bmp))
            {
                Size s = new Size()
                {
                    Width = width, //Screen.PrimaryScreen.Bounds.Size.Width - width,
                    Height = height //Screen.PrimaryScreen.Bounds.Size.Height - height
                };
                g.CopyFromScreen(left, top, 0, 0, s);
                //bmp.Save("s.png");  // saves the image
            }
            return bmp;
        }
    }
}
