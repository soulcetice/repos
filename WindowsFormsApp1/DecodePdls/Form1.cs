using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace DecodePdls
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            button1_Click(button1, EventArgs.Empty);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            List<FileInfo> files = GetFiles();

            GetData(files, out List<string> lefts2, contain: "Position X", encoding: Encoding.Unicode);
            GetData(files, out List<string> pdlsCalled, contain: "pdlName = ", notContain: "pdlName = \"\"", encoding: Encoding.Unicode);

            Console.WriteLine("end");
        }

        private void GetData(List<FileInfo> files, out List<string> outputLines, Encoding encoding, string contain, string notContain = "")
        {
            outputLines = new List<string>();

            foreach (var f in files)
            {
                //var lines = File.ReadAllLines(f.FullName, encoding).ToList();
                var lines1 = File.ReadAllLines(f.FullName, Encoding.UTF8).ToList();
                //var lines2 = File.ReadAllLines(f.FullName, Encoding.ASCII).ToList();
                //var lines3 = File.ReadAllLines(f.FullName, Encoding.BigEndianUnicode).ToList();
                //var lines4 = File.ReadAllLines(f.FullName, Encoding.UTF32).ToList();
                //var lines5 = File.ReadAllLines(f.FullName, Encoding.UTF7).ToList();


                for (int i = 0; i < lines1.Count; i++)
                {
                    lines1[i] = lines1[i].Replace("\0", "");
                }

                var pdl = lines1.Where(c => c.Contains(contain)).ToList();

                if (pdl.Count > 0)
                {
                    Console.WriteLine("yep");
                }

                if (notContain != "")
                {
                    pdl = pdl.Where(c => !c.Contains(notContain)).ToList();

                    if (pdl.Count > 0)
                    {
                        Console.WriteLine("yep");
                    }
                }

                var filter = pdl.Where(c => c.Contains(textBox1.Text)).ToList();

                outputLines.AddRange(filter);
            }

            //foreach (var c in outputLines)
            //{
            //    var text = c.Substring(0, c.Length > 100 ? 100 : c.Length).Replace("\t", "");
            //    listBox1.Items.Add(text);
            //}
        }

        private List<FileInfo> GetFiles()
        {
            var path = "";

            var cdir = new DirectoryInfo(@"C:\Project\");
            var ddir = new DirectoryInfo(@"D:\Project\");
            DirectoryInfo cltDir = null;
            bool dirFlag = false;
            while (dirFlag == false)
            {
                if (cdir.Exists)
                {
                    FindClientFolder(cdir, ref cltDir, ref dirFlag);
                }
                else if (ddir.Exists)
                {
                    FindClientFolder(ddir, ref cltDir, ref dirFlag);
                }
            }

            if (cltDir != null)
            {

                foreach (var dir in cltDir.GetDirectories())
                {
                    if (dir.Name.ToLower() == "GraCS".ToLower())
                    {
                        path = dir.FullName;
                    }
                }
            }

            var files = new DirectoryInfo(path).GetFiles().Where(c => c.Extension.ToUpper() == ".pdl".ToUpper()).ToList();

            files = files.Where(c => !c.Name.StartsWith("@") && !c.Name.Contains("i8")).ToList();
            return files;
        }

        private static void FindClientFolder(DirectoryInfo c, ref DirectoryInfo cltDir, ref bool dirFlag)
        {
            var dirs = c.GetDirectories();
            foreach (var dir in dirs)
            {
                if (dir.Name.ToLower().EndsWith("_clt"))
                {
                    cltDir = dir;
                    dirFlag = true;
                }
            }
        }
    }
}
