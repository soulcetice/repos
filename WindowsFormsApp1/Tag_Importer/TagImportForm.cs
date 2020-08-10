﻿using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using WindowHandle;
using WindowsUtilities;
using Interoperability;
using PInvoke;
using System.IO;
using System.Drawing.Imaging;
using System.Resources;
using Tag_Importer.Properties;
//using CCConfigStudio;

namespace Tag_Importer
{
    public partial class TagImportForm : Form
    {
        public TagImportForm()
        {
            InitializeComponent();

            string configFile = Application.StartupPath + "\\Tag_Importer.ini";
            GetInitialValues(configFile);

            textBox1_TextChanged(textBox1, new EventArgs());
        }

        private bool isFirstImporting;

        private void button1_Click(object sender, EventArgs e)
        {
            KeepConfig();
            ImportTags();
            RemoveTemporaryOutputs();
        }

        private IntPtr OpenTagMgmtMenu(IntPtr tagMgmt)
        {
            PInvokeLibrary.SetForegroundWindow(tagMgmt);

            System.Threading.Thread.Sleep(100);

            IntPtr importPopup = PInvokeLibrary.FindWindow("#32770", "Import");

            if (importPopup == IntPtr.Zero)
            {
                SendKeyHandled(tagMgmt, "(%)");
                SendKeyHandled(tagMgmt, "{RIGHT}");
                SendKeyHandled(tagMgmt, "{ENTER}");
                SendKeyHandled(tagMgmt, "{DOWN}");
                SendKeyHandled(tagMgmt, "{DOWN}");
                SendKeyHandled(tagMgmt, "{DOWN}");
                SendKeyHandled(tagMgmt, "{DOWN}");
                SendKeyHandled(tagMgmt, "{ENTER}");

                importPopup = PInvokeLibrary.FindWindow("#32770", "Import");
                do
                {
                    try { importPopup = PInvokeLibrary.FindWindow("#32770", "Import"); }
                    catch (Exception exc)
                    {
                        Console.WriteLine(exc.Message);
                    }
                    System.Threading.Thread.Sleep(100);
                } while (importPopup == IntPtr.Zero);

                //changing address...
                SendKeyHandled(importPopup, "{TAB}");
                SendKeyHandled(importPopup, "{TAB}");
                SendKeyHandled(importPopup, "{TAB}");
                SendKeyHandled(importPopup, "{TAB}");
                SendKeyHandled(importPopup, "{TAB}");
                SendKeyHandled(importPopup, "{ENTER}");
                string path = textBox1.Text;
                foreach (char c in path)
                    SendKeys.SendWait(c.ToString());
                SendKeyHandled(importPopup, "{ENTER}");

            }
            return importPopup;
        }

        private void KeepConfig()
        {
            string configPath = Application.StartupPath + "\\Tag_Importer.ini";
            var configFile = File.CreateText(configPath);
            configFile.WriteLine(textBox1.Text);
            configFile.WriteLine(textBox2.Text);
            configFile.WriteLine(checkBox1.Checked);
            configFile.WriteLine(checkBox2.Checked);
            configFile.WriteLine("Authored by Muresan Radu-Adrian (MURA02)");
            configFile.Close();
        }

        private void GetInitialValues(string path)
        {
            if (path != string.Empty)
            {
                if (new FileInfo(path).Exists)
                {
                    var configFile = File.ReadLines(path);
                    var fileLen = configFile.Count();

                    textBox1.Text = configFile.ElementAt(0);
                    if (fileLen >= 2) textBox2.Text = configFile.ElementAt(1);
                    if (fileLen >= 3) checkBox1.Checked = bool.Parse(configFile.ElementAt(2));
                    if (fileLen >= 4) checkBox2.Checked = bool.Parse(configFile.ElementAt(3));
                }
            }
        }

        private void SendKeyHandled(IntPtr windowHandle, string key)
        {
            bool success;
            do
            {
                try
                {
                    PInvokeLibrary.SetForegroundWindow(windowHandle);
                    SendKeys.SendWait(key);
                    success = true;
                }
                catch (Exception e)
                {
                    Console.WriteLine(e.Message);
                    LogToFile(e.Message);
                    success = false;
                }
            } while (success == false);
        }

        #region currently unused
        public string GetControlText(IntPtr hWnd)
        {
            // Get the size of the string required to hold the window title (including trailing null.) 
            Int32 titleSize = PInvokeLibrary.SendMessage(hWnd, (int)WindowsMessages.WM_GETTEXTLENGTH, 0, IntPtr.Zero).ToInt32();

            // If titleSize is 0, there is no title so return an empty string (or null)
            if (titleSize == 0)
                return String.Empty;

            _ = PInvokeLibrary.SendMessage(hWnd, (int)WindowsMessages.WM_GETTEXT, new StringBuilder(titleSize + 1).Capacity, new StringBuilder(titleSize + 1));

            return new StringBuilder(titleSize + 1).ToString();
        }

        public IntPtr GetChildBySubstring(string str, IntPtr handle)
        {
            var extractedText = new List<string>();
            List<IntPtr> childObjects = new WindowHandleInfo(handle).GetAllChildHandles();
            for (int i = 0; i < childObjects.Count; i++)
            {
                var text = GetControlText(childObjects[i]);
                if (text.Contains(str))
                {
                    return childObjects[i];
                }
            }
            return IntPtr.Zero;
        }

        private void Testing()
        {
            #region grafexe
            //grafexe.HMIObject t;
            //grafexe.HMIObjects ts;

            //grafexe.Application app;

            //app = grafexe.ApplicationClass;

            //ts = app.ActiveDocument.HMIObjects;

            //foreach (var obj in ts)
            //{

            //}
            #endregion
        }

        #endregion

        private static void RemoveTemporaryOutputs()
        {
            FileInfo output = new FileInfo(Application.StartupPath + "\\output.png");
            if (output.Exists)
            {
                output.Delete();
            }
        }

        private static void LogToFile(string content)
        {
            using (var fileWriter = new StreamWriter(Application.StartupPath + "\\Tag_Importer.logger", true))
            {
                fileWriter.WriteLine(DateTime.UtcNow.ToLongDateString() + " " + DateTime.UtcNow.ToLongTimeString() + " UTC : " + content);
                fileWriter.Close();
            }
        }

        private void ImportTags()
        {
            LogToFile("Started actions ******************************************");

            #region tree view handle capture
            IntPtr tag = PInvokeLibrary.FindWindow("WinCC ConfigurationStudio MainWindow", "Tag Management - WinCC Configuration Studio");
            IntPtr ccAx = PInvokeLibrary.FindWindowEx(tag, IntPtr.Zero, "CCAxControlContainerWindow", null);
            IntPtr navBar = PInvokeLibrary.FindWindowEx(ccAx, IntPtr.Zero, "WinCC NavigationBarControl Window", null);
            IntPtr treeView = PInvokeLibrary.FindWindowEx(navBar, IntPtr.Zero, "WinCC ConfigurationStudio NavigationBarTreeView", null);
            IntPtr trHandle = PInvokeLibrary.FindWindowEx(treeView, IntPtr.Zero, "SysTreeView32", "");
            _ = PInvokeLibrary.GetWindowRect(trHandle, out RECT trRect);



            IntPtr ccAxt = GetccAxParentOfDataGrid(tag);
            _ = PInvokeLibrary.GetWindowRect(ccAxt, out RECT ccAxtRect);
            IntPtr dataGridHandle = PInvokeLibrary.FindWindowEx(ccAxt, IntPtr.Zero, "WinCC DataGridControl Window", null);
            if (dataGridHandle == IntPtr.Zero || dataGridHandle == new IntPtr(0x00000000))
                LogToFile("dataGridHandle was IntPtr.Zero");
            #endregion

            Bitmap scrollUp = (Bitmap)Resources.ResourceManager.GetObject("scrollUp");
            Bitmap scrollDn = (Bitmap)Resources.ResourceManager.GetObject("scrollDown");

            #region lookInDirectory

            DirectoryInfo dinfo = new DirectoryInfo(textBox1.Text);
            FileInfo[] Files = dinfo.GetFiles("*.txt");

            #endregion

            #region image character recognition

            //ExpandOrHideVisibleTree(tag, trHandle, expand: true);

            //find required tag group in treeview, expand and scroll until found
            //if structure tags is not visible even after moving a little bit then it's not in the field of view at all - that's ok
            List<WordWithLocation> rowData = GetWordsInHandle(trHandle);

            foreach (var c in rowData)
            {
                LogToFile(c.word);
            }

            #region Expand Only Tag Management
            ScrollAllTheWayUp(trHandle, trRect, scrollUp, scrollDn, true);
            if (FindClosestMatch(rowData, "TagManagement") == null)
            {
                SendKeyHandled(trHandle, "{DOWN}"); SendKeyHandled(trHandle, "{DOWN}");
            }
            ExpandTreeItem(trHandle, "TagManagement", true, trRect);
            #endregion

            //now get structures
            #endregion

            #region delete variables

            //expand all tree elements and scroll until no expand is seen - in tag management only

            //hide all tree elements from down

            //get image and coordinate of following text
            //expand all tag management, omit structure tags if possible

            //get image and coordinate of following text
            //expand OPCUA

            //get image and coordinate of following text
            //expand OPC Unified Architecture

            //find text files to be imported in import folder, to get what required tag group and tag resource to click on

            //get image and coordinate of following text
            //expand required resource

            //get image and coordinate of following text
            //click on required tag group

            //focus on spreadsheet, ctrl+A , delete

            #endregion

            #region open import dialog

            isFirstImporting = true;
            for (int fileNo = 0; fileNo < checkedListBox1.CheckedItems.Count; fileNo++)
            {
                FileInfo file = Files.FirstOrDefault(f => f.Name.Equals(checkedListBox1.CheckedItems[fileNo]));

                //MED2_OPC_UA1    OPCUA OPC UnifiedArchitecture opc.tcp://10.80.92.245:4890|;#None;<>;<>;1;0;0;1;2;1
                var grup = File.ReadLines(file.FullName).Skip(3).Take(1).ToList()[0].Split(Convert.ToChar("\t"));
                for (int i = 0; i < grup.Length; i++)
                {
                    grup[i] = grup[i].Replace(" ", string.Empty);
                }

                List<string> line = File.ReadLines(file.FullName).Skip(8).Take(6).ToList();

                List<NameConnection> nc = new List<NameConnection>();
                for (int i = 0; i < line.Count; i++)
                {
                    string c = line[i];
                    if (c == "") break;
                    var conn = c.Split(Convert.ToChar("\t"));
                    nc.Add(new NameConnection() { name = conn[0], connection = conn[1] });
                    LogToFile(nc[i].connection + ", " + nc[i].name);
                }

                ExpandTreeItem(trHandle, grup[1], true, trRect);
                // have to implement a better way - in case it is not found, scroll down a bit - scroll to find lowest levenshtein distance
                //get whole tree model at least for the current "internal tags" 
                while (ExpandTreeItem(trHandle, grup[1], true, trRect) == false)
                {
                    LogToFile("Could not find " + grup[1]);
                    ScrollAllTheWayUp(trHandle, trRect, scrollUp, scrollDn, false);
                }


                ExpandTreeItem(trHandle, grup[2], true, trRect);

                while (ExpandTreeItem(trHandle, grup[2], true, trRect) == false)
                {
                    LogToFile("Could not find " + grup[2]);
                    ScrollAllTheWayUp(trHandle, trRect, scrollUp, scrollDn, false);
                }


                for (int i = 0; i < nc.Count; i++)
                {
                    DeleteExistingTags(trHandle, trRect, ccAxtRect, dataGridHandle, nc, i);
                    System.Threading.Thread.Sleep(5000);
                    LogToFile("Deleted variables for " + nc[i].connection + " " + nc[i].name);
                }

                ExpandTreeItem(trHandle, grup[2], false, trRect);
                ExpandTreeItem(trHandle, grup[1], false, trRect);
                ExpandTreeItem(trHandle, "TagManagement", false, trRect);

                ImportTagFile(tag, file);

                if (checkBox2.CheckState == CheckState.Checked)
                {
                    MoveFilesPair(file);
                    LogToFile("Moved files pair for " + file.Name);
                }
            }

            scrollUp.Dispose();
            scrollDn.Dispose();

            #endregion

            LogToFile("Finished importing from folder");
            MessageBox.Show(new Form { TopMost = true }, "Finished importing from specified folder");
        }

        private static IntPtr GetccAxParentOfDataGrid(IntPtr tag)
        {
            List<IntPtr> ccAxs = GetAllChildrenWindowHandles(tag, 4);
            var widthsccAx = new List<int>();
            for (int i = 0; i < ccAxs.Count; i++)
            {
                _ = PInvokeLibrary.GetWindowRect(ccAxs[i], out RECT rect);
                int width = rect.right - rect.left;
                widthsccAx.Add(width);
            }
            int no = widthsccAx.FindIndex(a => a == widthsccAx.OrderByDescending(c => c).Skip(2).FirstOrDefault());
            IntPtr ccAxt = ccAxs[no + 1]; //hope this holds together - second largest width of ccax
            LogToFile((no + 1).ToString());

            //ccAxt = ccAxs[2]; //hope this holds together - second largest width of ccax //in case previous stuff don't work

            return ccAxt;
        }

        private void DeleteExistingTags(IntPtr trHandle, RECT trRect, RECT ccAxtRect, IntPtr dataGridHandle, List<NameConnection> nc, int i)
        {
            List<WordWithLocation> rowData;
            NameConnection connData = nc[i];

            ExpandTreeItem(trHandle, connData.connection, true, trRect);

            rowData = GetWordsInHandle(trHandle);
            var myGroup = FindClosestMatch(rowData, connData.name);
            ClickInWindowAtXY(trHandle, trRect.left + myGroup.x, trRect.top + myGroup.y);
            //now delete the tags
            ClickInWindowAtXY(dataGridHandle, ccAxtRect.left + 100, ccAxtRect.top + 100); //click in data grid
            System.Threading.Thread.Sleep(100);
            SendKeyHandled(dataGridHandle, "^(a)");
            System.Threading.Thread.Sleep(100); //necessary to sleep because it was trying to delete before selecting
            SendKeyHandled(dataGridHandle, "{DELETE}");
            System.Threading.Thread.Sleep(100);
            ExpandTreeItem(trHandle, connData.connection, false, trRect);
        }

        private void MoveFilesPair(FileInfo file)
        {
            //\\10.16.80.31\d$\Project Data\002 - Tag Import\C1\G_CVC\2020 - 06\29
            string moveToPath = textBox2.Text;
            if (textBox2.Text.EndsWith(@"\") == false)
            {
                moveToPath = textBox2.Text + "\\";
            }

            var dateFolder = "\\" + DateTime.Now.Year + " - "
                + ((DateTime.Now.Month < 10) ? "0" + DateTime.Now.Month.ToString() : DateTime.Now.Month.ToString())
                + "\\" + (DateTime.Now.Day < 10 ? "0" + DateTime.Now.Day.ToString() : DateTime.Now.Day.ToString())
                + "\\";

            string[] myStruct = file.Name.Replace(".txt", "").Split(Convert.ToChar("_"));
            switch (myStruct.Count())
            {
                case 3:
                    moveToPath = moveToPath + myStruct[0] + "\\" + myStruct[1] + "_" + myStruct[2] + dateFolder;
                    break;
                case 2:
                    moveToPath = moveToPath + myStruct[0] + "\\" + myStruct[1] + dateFolder;
                    break;
            }

            if (file.Exists)
            {
                file.MoveTo(moveToPath + file.Name + ".txt");
            }

            FileInfo xml = new FileInfo(file.FullName.Replace(".txt", "_VarData.XML"));
            if (xml.Exists)
            {
                file.MoveTo(moveToPath + xml.Name);
            }
        }

        private void ClickInWindowAtXY(IntPtr handle, int? x, int? y)
        {
            PInvokeLibrary.SetForegroundWindow(handle);

            MouseOperations.SetCursorPosition(x.Value, y.Value); //have to use the found minus/plus coordinates here
            MouseOperations.MouseEvent(MouseOperations.MouseEventFlags.LeftDown);
            MouseOperations.MouseEvent(MouseOperations.MouseEventFlags.LeftUp);
        }

        private int GetHandleScrollState(Bitmap img)
        {
            Bitmap scrollUp = (Bitmap)Resources.ResourceManager.GetObject("scrollUp");
            Bitmap scrollDn = (Bitmap)Resources.ResourceManager.GetObject("scrollDown");
            var hasUp = Find(img, scrollUp);
            var hasDn = Find(img, scrollDn);
            if (hasUp.Count == 0 && hasDn.Count != 0) return 1; //Console.WriteLine("is already scrolled all the way up");
            if (hasDn.Count == 0 && hasUp.Count != 0) return 2; //Console.WriteLine("is already scrolled all the way down");
            if (hasDn.Count != 0 && hasDn.Count != 0) return 3; //Console.WriteLine("is in between");
            if (hasDn.Count == 0 && hasDn.Count == 0) return 4; //Console.WriteLine("there is no scrolling available, window contens are all seen. expand?");

            return 0; //inconclusive
        }

        private List<MyFunctions.PosLetter> GetCharactersInHandle(IntPtr handle)
        {

            var t = (Bitmap)Resources.ResourceManager.GetObject("lower_a");

            //only works for white background for now!!!
            MouseOperations.SetCursorPosition(1, 1); //have to use the found minus/plus coordinates here
            Bitmap img = MyFunctions.GetPngByHandle(handle);
            List<MyFunctions.PosLetter> allCharsFound = new List<MyFunctions.PosLetter>();
            allCharsFound.AddRange(Find(img, MyFunctions.MakeExistingTransparent((Bitmap)Resources.ResourceManager.GetObject("upper_A")), "A"));
            allCharsFound.AddRange(Find(img, MyFunctions.MakeExistingTransparent((Bitmap)Resources.ResourceManager.GetObject("upper_B")), "B"));
            allCharsFound.AddRange(Find(img, MyFunctions.MakeExistingTransparent((Bitmap)Resources.ResourceManager.GetObject("upper_C")), "C"));
            allCharsFound.AddRange(Find(img, MyFunctions.MakeExistingTransparent((Bitmap)Resources.ResourceManager.GetObject("upper_D")), "D"));
            allCharsFound.AddRange(Find(img, MyFunctions.MakeExistingTransparent((Bitmap)Resources.ResourceManager.GetObject("upper_E")), "E"));
            allCharsFound.AddRange(Find(img, MyFunctions.MakeExistingTransparent((Bitmap)Resources.ResourceManager.GetObject("upper_F")), "F"));
            allCharsFound.AddRange(Find(img, MyFunctions.MakeExistingTransparent((Bitmap)Resources.ResourceManager.GetObject("upper_G")), "G"));
            allCharsFound.AddRange(Find(img, MyFunctions.MakeExistingTransparent((Bitmap)Resources.ResourceManager.GetObject("upper_H")), "H"));
            allCharsFound.AddRange(Find(img, MyFunctions.MakeExistingTransparent((Bitmap)Resources.ResourceManager.GetObject("upper_I")), "I"));
            allCharsFound.AddRange(Find(img, MyFunctions.MakeExistingTransparent((Bitmap)Resources.ResourceManager.GetObject("upper_J")), "J"));
            allCharsFound.AddRange(Find(img, MyFunctions.MakeExistingTransparent((Bitmap)Resources.ResourceManager.GetObject("upper_K")), "K"));
            allCharsFound.AddRange(Find(img, MyFunctions.MakeExistingTransparent((Bitmap)Resources.ResourceManager.GetObject("upper_L")), "L"));
            allCharsFound.AddRange(Find(img, MyFunctions.MakeExistingTransparent((Bitmap)Resources.ResourceManager.GetObject("upper_M")), "M"));
            allCharsFound.AddRange(Find(img, MyFunctions.MakeExistingTransparent((Bitmap)Resources.ResourceManager.GetObject("upper_N")), "N"));
            allCharsFound.AddRange(Find(img, MyFunctions.MakeExistingTransparent((Bitmap)Resources.ResourceManager.GetObject("upper_O")), "O"));
            allCharsFound.AddRange(Find(img, MyFunctions.MakeExistingTransparent((Bitmap)Resources.ResourceManager.GetObject("upper_P")), "P"));
            allCharsFound.AddRange(Find(img, MyFunctions.MakeExistingTransparent((Bitmap)Resources.ResourceManager.GetObject("upper_Q")), "Q"));
            allCharsFound.AddRange(Find(img, MyFunctions.MakeExistingTransparent((Bitmap)Resources.ResourceManager.GetObject("upper_R")), "R"));
            allCharsFound.AddRange(Find(img, MyFunctions.MakeExistingTransparent((Bitmap)Resources.ResourceManager.GetObject("upper_S")), "S"));
            allCharsFound.AddRange(Find(img, MyFunctions.MakeExistingTransparent((Bitmap)Resources.ResourceManager.GetObject("upper_T")), "T"));
            allCharsFound.AddRange(Find(img, MyFunctions.MakeExistingTransparent((Bitmap)Resources.ResourceManager.GetObject("upper_U")), "U"));
            allCharsFound.AddRange(Find(img, MyFunctions.MakeExistingTransparent((Bitmap)Resources.ResourceManager.GetObject("upper_V")), "V"));
            allCharsFound.AddRange(Find(img, MyFunctions.MakeExistingTransparent((Bitmap)Resources.ResourceManager.GetObject("upper_W")), "W"));
            allCharsFound.AddRange(Find(img, MyFunctions.MakeExistingTransparent((Bitmap)Resources.ResourceManager.GetObject("upper_X")), "X"));
            allCharsFound.AddRange(Find(img, MyFunctions.MakeExistingTransparent((Bitmap)Resources.ResourceManager.GetObject("upper_Y")), "Y"));
            allCharsFound.AddRange(Find(img, MyFunctions.MakeExistingTransparent((Bitmap)Resources.ResourceManager.GetObject("upper_Z")), "Z"));

            allCharsFound.AddRange(Find(img, MyFunctions.MakeExistingTransparent((Bitmap)Resources.ResourceManager.GetObject("lower_a")), "a"));
            allCharsFound.AddRange(Find(img, MyFunctions.MakeExistingTransparent((Bitmap)Resources.ResourceManager.GetObject("lower_b")), "b"));
            allCharsFound.AddRange(Find(img, MyFunctions.MakeExistingTransparent((Bitmap)Resources.ResourceManager.GetObject("lower_c")), "c"));
            allCharsFound.AddRange(Find(img, MyFunctions.MakeExistingTransparent((Bitmap)Resources.ResourceManager.GetObject("lower_d")), "d"));
            allCharsFound.AddRange(Find(img, MyFunctions.MakeExistingTransparent((Bitmap)Resources.ResourceManager.GetObject("lower_e")), "e"));
            allCharsFound.AddRange(Find(img, MyFunctions.MakeExistingTransparent((Bitmap)Resources.ResourceManager.GetObject("lower_f")), "f"));
            allCharsFound.AddRange(Find(img, MyFunctions.MakeExistingTransparent((Bitmap)Resources.ResourceManager.GetObject("lower_g")), "g"));
            allCharsFound.AddRange(Find(img, MyFunctions.MakeExistingTransparent((Bitmap)Resources.ResourceManager.GetObject("lower_h")), "h"));
            allCharsFound.AddRange(Find(img, MyFunctions.MakeExistingTransparent((Bitmap)Resources.ResourceManager.GetObject("lower_i")), "i"));
            allCharsFound.AddRange(Find(img, MyFunctions.MakeExistingTransparent((Bitmap)Resources.ResourceManager.GetObject("lower_j")), "j"));
            allCharsFound.AddRange(Find(img, MyFunctions.MakeExistingTransparent((Bitmap)Resources.ResourceManager.GetObject("lower_k")), "k"));
            allCharsFound.AddRange(Find(img, MyFunctions.MakeExistingTransparent((Bitmap)Resources.ResourceManager.GetObject("lower_l")), "l"));
            allCharsFound.AddRange(Find(img, MyFunctions.MakeExistingTransparent((Bitmap)Resources.ResourceManager.GetObject("lower_m")), "m"));
            allCharsFound.AddRange(Find(img, MyFunctions.MakeExistingTransparent((Bitmap)Resources.ResourceManager.GetObject("lower_n")), "n"));
            allCharsFound.AddRange(Find(img, MyFunctions.MakeExistingTransparent((Bitmap)Resources.ResourceManager.GetObject("lower_o")), "o"));
            allCharsFound.AddRange(Find(img, MyFunctions.MakeExistingTransparent((Bitmap)Resources.ResourceManager.GetObject("lower_p")), "p"));
            allCharsFound.AddRange(Find(img, MyFunctions.MakeExistingTransparent((Bitmap)Resources.ResourceManager.GetObject("lower_q")), "q"));
            allCharsFound.AddRange(Find(img, MyFunctions.MakeExistingTransparent((Bitmap)Resources.ResourceManager.GetObject("lower_r")), "r"));
            allCharsFound.AddRange(Find(img, MyFunctions.MakeExistingTransparent((Bitmap)Resources.ResourceManager.GetObject("lower_s")), "s"));
            allCharsFound.AddRange(Find(img, MyFunctions.MakeExistingTransparent((Bitmap)Resources.ResourceManager.GetObject("lower_t")), "t"));
            allCharsFound.AddRange(Find(img, MyFunctions.MakeExistingTransparent((Bitmap)Resources.ResourceManager.GetObject("lower_u")), "u"));
            allCharsFound.AddRange(Find(img, MyFunctions.MakeExistingTransparent((Bitmap)Resources.ResourceManager.GetObject("lower_v")), "v"));
            allCharsFound.AddRange(Find(img, MyFunctions.MakeExistingTransparent((Bitmap)Resources.ResourceManager.GetObject("lower_w")), "w"));
            allCharsFound.AddRange(Find(img, MyFunctions.MakeExistingTransparent((Bitmap)Resources.ResourceManager.GetObject("lower_x")), "x"));
            allCharsFound.AddRange(Find(img, MyFunctions.MakeExistingTransparent((Bitmap)Resources.ResourceManager.GetObject("lower_y")), "y"));
            allCharsFound.AddRange(Find(img, MyFunctions.MakeExistingTransparent((Bitmap)Resources.ResourceManager.GetObject("lower_z")), "z"));
            allCharsFound.AddRange(Find(img, MyFunctions.MakeExistingTransparent((Bitmap)Resources.ResourceManager.GetObject("lower_rf")), "rf"));

            allCharsFound.AddRange(Find(img, MyFunctions.MakeExistingTransparent((Bitmap)Resources.ResourceManager.GetObject("underscore")), "_"));
            allCharsFound.AddRange(Find(img, MyFunctions.MakeExistingTransparent((Bitmap)Resources.ResourceManager.GetObject("arond")), "@"));
            allCharsFound.AddRange(Find(img, MyFunctions.MakeExistingTransparent((Bitmap)Resources.ResourceManager.GetObject("_#")), "#"));
            allCharsFound.AddRange(Find(img, MyFunctions.MakeExistingTransparent((Bitmap)Resources.ResourceManager.GetObject("_1")), "1"));
            allCharsFound.AddRange(Find(img, MyFunctions.MakeExistingTransparent((Bitmap)Resources.ResourceManager.GetObject("_2")), "2"));
            allCharsFound.AddRange(Find(img, MyFunctions.MakeExistingTransparent((Bitmap)Resources.ResourceManager.GetObject("_3")), "3"));
            allCharsFound.AddRange(Find(img, MyFunctions.MakeExistingTransparent((Bitmap)Resources.ResourceManager.GetObject("_4")), "4"));
            allCharsFound.AddRange(Find(img, MyFunctions.MakeExistingTransparent((Bitmap)Resources.ResourceManager.GetObject("_5")), "5"));
            allCharsFound.AddRange(Find(img, MyFunctions.MakeExistingTransparent((Bitmap)Resources.ResourceManager.GetObject("_6")), "6"));
            allCharsFound.AddRange(Find(img, MyFunctions.MakeExistingTransparent((Bitmap)Resources.ResourceManager.GetObject("_7")), "7"));
            allCharsFound.AddRange(Find(img, MyFunctions.MakeExistingTransparent((Bitmap)Resources.ResourceManager.GetObject("_8")), "8"));
            allCharsFound.AddRange(Find(img, MyFunctions.MakeExistingTransparent((Bitmap)Resources.ResourceManager.GetObject("_9")), "9"));
            allCharsFound.AddRange(Find(img, MyFunctions.MakeExistingTransparent((Bitmap)Resources.ResourceManager.GetObject("_0")), "0"));

            return allCharsFound;
        }

        private void ScrollAllTheWayUp(IntPtr trHandle, RECT trRect, Bitmap scrollUp, Bitmap scrollDn, bool up)
        {
            var img = MyFunctions.GetPngByHandle(trHandle);
            var scrollState = GetHandleScrollState(img);
            //HideStructureTags(trHandle, GetWordsInHandle(trHandle), trRect);

            int FirstState;
            int SecondState;

            switch (up)
            {
                case true:
                    FirstState = 1;
                    break;
                default:
                    FirstState = 2;
                    break;
            }
            SecondState = 4;

            if (scrollState != FirstState || scrollState != SecondState)
            {
                do
                {
                    if (up == true)
                    {
                        var goHere = Find(img, scrollUp).FirstOrDefault();
                        ClickInWindowAtXY(trHandle, trRect.left + goHere.X, trRect.top + goHere.Y);
                    }
                    else
                    {
                        var goHere = Find(img, scrollDn).FirstOrDefault();
                        ClickInWindowAtXY(trHandle, trRect.left + goHere.X, trRect.top + goHere.Y);
                    }
                    img = MyFunctions.GetPngByHandle(trHandle);

                    scrollState = GetHandleScrollState(img);
                    img.Dispose();
                    System.Threading.Thread.Sleep(10);
                } while (scrollState != FirstState && scrollState != SecondState);
            }
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked == true)
            {
                for (int i = 0; i < checkedListBox1.Items.Count; i++)
                {
                    checkedListBox1.SetItemCheckState(i, CheckState.Checked);
                }
            }
            else
            {
                for (int i = 0; i < checkedListBox1.Items.Count; i++)
                {
                    checkedListBox1.SetItemCheckState(i, CheckState.Unchecked);
                }
            }
            checkedListBox1_SelectedIndexChanged(checkBox1, new System.EventArgs());
        }

        private void HideStructureTags(IntPtr trHandle, List<WordWithLocation> rowData, RECT trRect)
        {
            if (FindClosestMatch(rowData, "Structuretags") == null)
            {
                /*SendKeyHandled(trHandle, "{UP}"); SendKeyHandled(trHandle, "{UP}");*/
                rowData = GetWordsInHandle(trHandle);
            }
            rowData = GetWordsInHandle(trHandle);
            if (FindClosestMatch(rowData, "Structuretags") != null)
            {
                ExpandTreeItem(trHandle, "Structuretags", false, trRect);
            }
        }

        private bool ExpandTreeItem(IntPtr trHandle, string FindThis, bool expand, RECT trRect)
        {
            List<WordWithLocation> rowData = GetWordsInHandle(trHandle);
            WordWithLocation foundElem = FindClosestMatch(rowData, FindThis);
            LogToFile("For " + FindThis + " I found the closest element: " + foundElem.word);
            var getElement = TextHasExpandOrHide(trHandle, foundElem, expand);
            if (getElement != null)
            {
                ClickInWindowAtXY(trHandle, trRect.left + getElement?.X, trRect.top + getElement?.Y);
                return true;
            }
            else
            {
                return false;
            }
        }

        private Point? TextHasExpandOrHide(IntPtr trHandle, WordWithLocation foundElem, bool expand)
        {
            Bitmap find = expand != true ? (Bitmap)Resources.ResourceManager.GetObject("minus") : (Bitmap)Resources.ResourceManager.GetObject("plus");
            var img = MyFunctions.GetPngByHandle(trHandle);
            var foundExpandButtons = Find(img, find);
            if (foundExpandButtons.Where(c => Math.Abs(c.Y - foundElem.y) < 4 && Math.Abs(c.X - foundElem.x) < 150).ToList().Count > 0)
            {
                return foundExpandButtons.FirstOrDefault(c => Math.Abs(c.Y - foundElem.y) < 4 && Math.Abs(c.X - foundElem.x) < 150);
            }
            return null;
        }

        private void ImportTagFile(IntPtr tag, FileInfo f)
        {
            if (isFirstImporting == false)
            {
                SendKeyHandled(tag, "(%)");
                isFirstImporting = false;
            }
            IntPtr importPopup = OpenTagMgmtMenu(tag);
            System.Threading.Thread.Sleep(500);
            //find files in import dialog
            IntPtr duiview = PInvokeLibrary.FindWindowEx(importPopup, IntPtr.Zero, "DUIViewWndClassName", "");
            IntPtr directuihwnd = PInvokeLibrary.FindWindowEx(duiview, IntPtr.Zero, "DirectUIHWND", "");
            List<IntPtr> ctrlNotifySinks = GetAllChildrenWindowHandles(directuihwnd, 5);

            int oldWidth = 0;
            IntPtr CtrlNotifySink = IntPtr.Zero;
            for (int i = 0; i < ctrlNotifySinks.Count; i++)
            {
                IntPtr elem = ctrlNotifySinks[i];
                _ = PInvokeLibrary.GetWindowRect(elem, out RECT rect);
                if (rect.right - rect.left > oldWidth) CtrlNotifySink = elem;
                oldWidth = rect.right - rect.left;
            }

            IntPtr fileListParent = PInvokeLibrary.FindWindowEx(CtrlNotifySink, IntPtr.Zero, "SHELLDLL_DefView", null);
            IntPtr fileList = PInvokeLibrary.FindWindowEx(fileListParent, IntPtr.Zero, "DirectUIHWND", null);

            var filesInDialog = GetWordsInHandle(fileListParent);

            _ = PInvokeLibrary.GetWindowRect(fileList, out RECT fileListRect);

            WordWithLocation tagFile = RobustFileChoice(f, filesInDialog);

            ClickInWindowAtXY(fileList, fileListRect.left + tagFile.x, fileListRect.top + tagFile.y);
            SendKeyHandled(fileList, "{ENTER}");
            System.Threading.Thread.Sleep(1000);

            bool success = false;
            do
            {
                IntPtr confirmationPopup = PInvokeLibrary.FindWindow("#32770", "Import");
                if (confirmationPopup != IntPtr.Zero)
                {
                    SendKeyHandled(confirmationPopup, "{ENTER}");
                    System.Threading.Thread.Sleep(1000);
                    success = true;
                }
                if (confirmationPopup == IntPtr.Zero)
                    System.Threading.Thread.Sleep(1000);
            } while (success == false);
        }

        private static WordWithLocation FindClosestMatch(List<WordWithLocation> rowData, string FindThis)
        {
            int levDist = 100;
            WordWithLocation myElem;
            do
            {
                myElem = rowData.FirstOrDefault(c => c.word == FindThis);
                WordWithLocation chosenFile = new WordWithLocation();
                for (int i = 0; i < rowData.Count; i++)
                {
                    WordWithLocation c = rowData[i];
                    if (LevenshteinDistance.Compute(FindThis, c.word) < levDist)
                    {
                        levDist = LevenshteinDistance.Compute(FindThis, c.word);
                        chosenFile = c;
                    }
                }
                myElem = chosenFile;
            } while (myElem == null);

            if (levDist > 1)
            {
                return null;
            }
            else
                return myElem;
        }

        private static WordWithLocation RobustFileChoice(FileInfo f, List<WordWithLocation> filesInDialog)
        {
            var fileName = Path.GetFileNameWithoutExtension(f.FullName);
            var tagFile = filesInDialog.FirstOrDefault(c => c.word == fileName);
            if (tagFile == null)
            {
                int levDist = 50;
                string chosenFile = "";
                for (int i = 0; i < filesInDialog.Count; i++)
                {
                    WordWithLocation t = filesInDialog[i];
                    if (LevenshteinDistance.Compute(fileName, t.word) < levDist)
                    {
                        levDist = LevenshteinDistance.Compute(fileName, t.word);
                        chosenFile = t.word;
                    }
                }
                tagFile = filesInDialog.FirstOrDefault(c => c.word == chosenFile);
            }
            return tagFile;
        }

        private List<WordWithLocation> GetWordsInHandle(IntPtr fileListParent)
        {
            //find y ranges
            //smallest x for these ranges is the first letter
            //then in these ranges, add all elements to the first letter to get the word

            var sorted = new List<dynamic>();
            var toSort = GetCharactersInHandle(fileListParent).OrderBy(elem => elem.y).ToList();

            var yList = (from MyFunctions.PosLetter c in toSort
                         select c.y).ToList();
            List<int> yDifs = new List<int>();
            for (int i = 1; i < yList.Count; i++)
            {
                yDifs.Add(yList[i] - yList[i - 1]);
            }

            //find y bands and group by band
            var sortedList = new List<List<MyFunctions.PosLetter>>();
            int j = -1;
            for (int i = 0; i < yDifs.Count; i++)
            {
                //add first element and first character
                if (sortedList.Count == 0)
                {
                    sortedList.Add(new List<MyFunctions.PosLetter>());
                    j++;
                    sortedList[j].Add(toSort[i]);
                }
                else //add the rest
                {
                    if (yDifs[i] > 7) //the yband tolerance
                    {
                        sortedList.Add(new List<MyFunctions.PosLetter>());
                        sortedList[j].Add(toSort[i]);
                        j++;
                    }
                    else
                    {
                        sortedList[j].Add(toSort[i]);
                    }
                    if (i == yDifs.Count - 1) //add last element in whole set (smaller number of ydifs)
                    {
                        sortedList[j].Add(toSort[i + 1]);
                    }
                }
            }

            //sort
            for (int i = 0; i < sortedList.Count; i++)
            {
                sortedList[i] = sortedList[i].OrderBy(c => c.x).ToList();
            }

            //group words
            var words = new List<WordWithLocation>();
            int iter = 0;
            for (int i = 0; i < sortedList.Count; i++)
            {
                List<MyFunctions.PosLetter> word = sortedList[i];
                words.Add(new WordWithLocation());
                for (int i1 = 0; i1 < word.Count; i1++)
                {
                    MyFunctions.PosLetter c = word[i1];
                    words[iter].word += c.letter;
                    words[iter].x += c.x;
                    words[iter].y += c.y;
                }
                words[iter].x = words[iter].x / words[iter].word.Length;
                words[iter].y = words[iter].y / words[iter].word.Length;

                iter++;
            }
            return words;
        }

        public static List<IntPtr> GetAllChildrenWindowHandles(IntPtr hParent, int maxCount)
        {
            List<IntPtr> result = new List<IntPtr>();
            int ct = 0;
            IntPtr prevChild = IntPtr.Zero;
            IntPtr currChild = IntPtr.Zero;
            while (true && ct < maxCount)
            {
                currChild = PInvokeLibrary.FindWindowEx(hParent, prevChild, null, null);
                if (currChild == IntPtr.Zero) break;
                result.Add(currChild);
                prevChild = currChild;
                ++ct;
            }
            return result;
        }

        private void checkedListBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            DirectoryInfo dinfo = new DirectoryInfo(textBox1.Text);
            checkedListBox1.Items.Clear();
            if (dinfo.Exists)
            {
                FileInfo[] Files = dinfo.GetFiles("*.txt");

                foreach (var elem in Files)
                {
                    checkedListBox1.Items.Add(elem.Name);
                }
            }
        }

        #region OptimumBitmapFind

        public List<Point> Find(Bitmap haystack, Bitmap needle)
        {
            if (null == haystack || null == needle)
            {
                return null;
            }
            if (haystack.Width < needle.Width || haystack.Height < needle.Height)
            {
                return null;
            }

            var haystackArray = GetPixelArray(haystack);
            var needleArray = GetPixelArray(needle);
            return (from firstLineMatchPoint in FindMatch(haystackArray.Take(haystack.Height - needle.Height), needleArray[0])
                    where IsNeedlePresentAtLocation(haystackArray, needleArray, firstLineMatchPoint, 1)
                    select firstLineMatchPoint).ToList();
        }

        public List<MyFunctions.PosLetter> Find(Bitmap haystack, Bitmap needle, string l)
        {
            if (null == haystack || null == needle)
            {
                return null;
            }
            if (haystack.Width < needle.Width || haystack.Height < needle.Height)
            {
                return null;
            }

            var haystackArray = GetPixelArray(haystack);
            var needleArray = GetPixelArray(needle);
            var list = new List<MyFunctions.PosLetter>();

            foreach (var firstLineMatchPoint in FindMatch(haystackArray.Take(haystack.Height - needle.Height), needleArray[0]))
            {
                if (IsNeedlePresentAtLocation(haystackArray, needleArray, firstLineMatchPoint, 1))
                {
                    list.Add(new MyFunctions.PosLetter()
                    {
                        letter = l,
                        x = firstLineMatchPoint.X,
                        y = firstLineMatchPoint.Y
                    });
                }
            }

            return list;
        }

        private int[][] GetPixelArray(Bitmap bitmap)
        {
            var result = new int[bitmap.Height][];
            var bitmapData = bitmap.LockBits(new Rectangle(0, 0, bitmap.Width, bitmap.Height), ImageLockMode.ReadOnly,
                PixelFormat.Format32bppArgb);

            for (int y = 0; y < bitmap.Height; ++y)
            {
                result[y] = new int[bitmap.Width];
                Marshal.Copy(bitmapData.Scan0 + y * bitmapData.Stride, result[y], 0, result[y].Length);
            }

            bitmap.UnlockBits(bitmapData);

            return result;
        }

        private IEnumerable<Point> FindMatch(IEnumerable<int[]> haystackLines, int[] needleLine)
        {
            var y = 0;
            foreach (var haystackLine in haystackLines)
            {
                for (int x = 0, n = haystackLine.Length - needleLine.Length; x < n; ++x)
                {
                    if (ContainSameElements(haystackLine, x, needleLine, 0, needleLine.Length))
                    {
                        yield return new Point(x, y);
                    }
                }
                y += 1;
            }
        }

        private bool ContainSameElements(int[] first, int firstStart, int[] second, int secondStart, int length)
        {
            for (int i = 0; i < length; ++i)
            {
                if (first[i + firstStart] != second[i + secondStart])
                {
                    return false;
                }
            }
            return true;
        }

        private bool IsNeedlePresentAtLocation(int[][] haystack, int[][] needle, Point point, int alreadyVerified)
        {
            //we already know that "alreadyVerified" lines already match, so skip them
            for (int y = alreadyVerified; y < needle.Length; ++y)
            {
                if (!ContainSameElements(haystack[y + point.Y], point.X, needle[y], 0, needle[y].Length))
                {
                    return false;
                }
            }
            return true;
        }

        #endregion
    }

    class WordWithLocation
    {
        public int x;
        public int y;
        public string word;
    }

    class NameConnection
    {
        public string name;
        public string connection;
    }

    static class LevenshteinDistance
    {
        public static int Compute(string s, string t)
        {
            if (string.IsNullOrEmpty(s))
            {
                if (string.IsNullOrEmpty(t))
                    return 0;
                return t.Length;
            }

            if (string.IsNullOrEmpty(t))
            {
                return s.Length;
            }

            int n = s.Length;
            int m = t.Length;
            int[,] d = new int[n + 1, m + 1];

            // initialize the top and right of the table to 0, 1, 2, ...
            for (int i = 0; i <= n; d[i, 0] = i++) ;
            for (int j = 1; j <= m; d[0, j] = j++) ;

            for (int i = 1; i <= n; i++)
            {
                for (int j = 1; j <= m; j++)
                {
                    int cost = (t[j - 1] == s[i - 1]) ? 0 : 1;
                    int min1 = d[i - 1, j] + 1;
                    int min2 = d[i, j - 1] + 1;
                    int min3 = d[i - 1, j - 1] + cost;
                    d[i, j] = Math.Min(Math.Min(min1, min2), min3);
                }
            }
            return d[n, m];
        }
    }
}
