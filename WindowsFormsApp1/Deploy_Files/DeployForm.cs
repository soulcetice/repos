using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace Deploy_Files
{
    public partial class DeployForm : Form
    {
        public DeployForm()
        {
            InitializeComponent();

            GetIpsLmHosts(false);

            if (File.Exists(Application.ExecutablePath + @".ini"))
            {
                List<string> data = File.ReadLines(Application.ExecutablePath + @".ini").Skip(19).Take(7).ToList();
                //var ips = data[3].Split(Convert.ToChar(","));
                //foreach (var item in ips)
                //{
                    //if (item != string.Empty) ipList.Items.Add(item);
                //}


                textBox1.Text = data[4];
                textBox3.Text = data[5];
                //textBox4.Text = data[0];
                //textBox6.Text = data[1];
            }
        }

        public List<dynamic> fullFiles = new List<dynamic>();
        public List<dynamic> safeFiles = new List<dynamic>();

        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1.Multiselect = true;
            openFileDialog1.InitialDirectory = textBox1.Text;
            openFileDialog1.Title = "Select files to copy to selected clients";

            System.IO.Stream myStream;
            var fileList = new List<string>();

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                foreach (String file in openFileDialog1.FileNames)
                {
                    try
                    {
                        if ((myStream = openFileDialog1.OpenFile()) != null)
                        {
                            using (myStream)
                            {
                                fileList.Add(file);
                                checkedListBox1.Items.Add(file);
                                checkedListBox1.SetItemCheckState(checkedListBox1.Items.Count - 1, CheckState.Checked);
                            }
                        }
                    }

                    catch (Exception ex)
                    {
                        MessageBox.Show("Error: Could not read file from disk. Original error: " + ex.Message);
                    }
                }
                checkedListBox1.Refresh();
            }
            Console.WriteLine("test");
        }

        private void GetIpsLmHosts(bool include)
        {
            string lmhostPath = Path.Combine(Application.StartupPath, "lmhosts");
            if (File.Exists(lmhostPath))
            {
                var lmHosts = File.ReadLines(lmhostPath);
                foreach (var item in lmHosts)
                {
                    if (item.IndexOf("HMIC") > 0
                        || item.IndexOf("HmiC") > 0
                        || ((item.IndexOf("HmiE01") < 0 && item.IndexOf("HmiE0") > 0)
                        || (item.IndexOf("HMIE01") < 0 && item.IndexOf("HMIE0") > 0))
                        && item.StartsWith("#") == false && item != "")
                    {
                        clientData.Items.Add(item);
                    }
                    if ((item.IndexOf("HMID01") > 0 ||
                        item.IndexOf("HmiD01") > 0 ||
                        item.IndexOf("HMIE01") > 0 ||
                        item.IndexOf("HmiE01") > 0 ||
                        item.IndexOf("HmiS") > 0 ||
                        item.IndexOf("HMIS") > 0) &&
                        item.StartsWith("#") == false && item != "")
                    {
                        if (include)
                        {
                            clientData.Items.Add(item);
                        }
                    }
                }
            }
        }

        protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
        {
            if (keyData == (Keys.P | Keys.Control))
            {
                //do my prank
                for (int i = 0; i < 100; i++)
                {
                    Size = new Size { Height = Size.Height + 1, Width = Size.Width + 1 };

                    listBox1.Items.Add("Made the app a little bit bigger");
                    listBox1.Refresh();
                }

                return true;
            }
            return base.ProcessCmdKey(ref msg, keyData);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Start();
        }

        public void Start()
        {
            //"C:\Project\ALCCRM\wincproj\CRM_CLT" to "\\10.6.151.41\D$\Project\CRM_CLT\"
            currentClient = "";
            currentFile = "";

            if (checkedListBox1.CheckedItems.Count == 0)
            {
                return;
            }
            if (clientData.CheckedItems.Count == 0)
            {
                return;
            }
            string sep = "\t";
            string partialPath = textBox3.Text;
            if (partialPath.Last().ToString() != @"\")
            {
                partialPath = partialPath + @"\";
            }

            foreach (string client in clientData.CheckedItems)
            {
                currentClient = client;
                string[] splitContent = client.Split(sep.ToCharArray());
                string ip = splitContent[0];
                string machineName = splitContent[1];
                foreach (var file in checkedListBox1.CheckedItems)
                {
                    currentFile = file.ToString();
                    string originFile = file.ToString();
                    var extraPath = file.ToString().Replace(textBox1.Text, "".ToString());
                    var fileData = new FileInfo((string)file);
                    var destinationFile = @"\\" + machineName + @"\" + partialPath + extraPath;

                    var destinationDirectory = new DirectoryInfo(destinationFile.Replace(fileData.Name, ""));
                    if (destinationDirectory.Exists)
                    {
                        try
                        {
                            RobustCopy(originFile, destinationFile);
                        }
                        catch (Exception exc)
                        {
                            listBox1.Items.Add("For " + currentClient + " and file " + currentFile + " - " + exc.Message);
                        }

                    }
                    else
                    {
                        listBox1.Items.Add("The directory " + destinationDirectory.FullName + " did not exist or was not accessible");
                    }
                    listBox1.Refresh();
                }
            }

        }

        private string currentClient;
        private string currentFile;

        public void BufferCopy(String oldPath, String newPath)
        {
            FileStream input = null;
            FileStream output = null;
            try
            {
                input = new FileStream(oldPath, FileMode.Open);
                output = new FileStream(newPath, FileMode.Create, FileAccess.ReadWrite);

                byte[] buffer = new byte[32768];
                int read;
                while ((read = input.Read(buffer, 0, buffer.Length)) > 0)
                {
                    output.Write(buffer, 0, read);
                }
            }
            catch (Exception e)
            {
                listBox1.Items.Add("For " + currentClient + " and file " + currentFile + " - " + e.Message);
            }
            finally
            {
                input.Close();
                input.Dispose();
                output.Close();
                output.Dispose();
            }
        }

        private void RobustCopy(string originFile, string destinationFile)
        {
            listBox1.Items.Add("Attempting to copy " + originFile + " to " + destinationFile);
            File.Copy(originFile, destinationFile, true);
            if (File.Exists(destinationFile))
            {
                listBox1.Items.Add("Copied " + originFile + " to " + destinationFile);
            }
            else
            {
                listBox1.Items.Add("Attempting to copy via buffer method " + originFile + " to " + destinationFile);
                BufferCopy(originFile, destinationFile);
                if (File.Exists(destinationFile))
                {
                    listBox1.Items.Add("Copied via buffer method " + originFile + " to " + destinationFile);
                }
                else
                {
                    listBox1.Items.Add("Could not copy with either method " + originFile + " to " + destinationFile);
                }
            }
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            for (int i = 0; i < clientData.Items.Count; i++)
            {
                clientData.SetItemCheckState(i, checkBox1.CheckState);
            }
        }

        private void listBox1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Control == true && e.KeyCode == Keys.C)
            {
                string s = listBox1.SelectedItem.ToString();
                Clipboard.SetData(DataFormats.StringFormat, s);
            }
        }

        private void clientData_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void button3_Click(object sender, EventArgs e)
        {

            checkedListBox1.Items.Clear();


        }

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {
            button4.Visible = checkBox2.Checked;
            //DialogResult result = DialogResult.None;
            if (checkBox2.Checked == true)
            {
                //result = MessageBox.Show("Do you want to make the app 'little bit bigger'?", "", MessageBoxButtons.OKCancel );

                //if (result == DialogResult.Yes)
                //{
                Size = new Size() { Height = 518, Width = 603 };
                //}
                //else if (result == DialogResult.No)
                //{
                //    //...
                //}
                //else
                //{
                //    //...
                //}
            }
            else
            {
                Size = new Size() { Height = 294, Width = 603 };
            }

        }

        private void button4_Click(object sender, EventArgs e)
        {
            foreach (var item in listBox1.Items)
            {
                LogToFile(item.ToString());
            }
            listBox1.Items.Clear();
        }

        private static void LogToFile(string content)
        {
            using (var fileWriter = new StreamWriter(Path.Combine(Application.StartupPath, "DeployForm.logger"), true))
            {
                DateTime date = DateTime.UtcNow;
                fileWriter.WriteLine(date.Year + "/" + date.Month + "/" + date.Day + " " + date.Hour + ":" + date.Minute + ":" + date.Second + ":" + date.Millisecond + " UTC: " + content);
                fileWriter.Close();
            }
        }
    }
}