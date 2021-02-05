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

            var lines = File.ReadAllLines(Path.Combine(Application.StartupPath, "TCM-DeflectorRollData.pdl"), Encoding.Unicode).ToList();

            var find = lines.Where(c => c.Contains("@V3_SMS_DatCls2_ActVal72_14")).ToList();

            Console.WriteLine(lines.Count);
        }
    }
}
