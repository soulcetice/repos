using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using WindowHandle;
using WindowsUtilities;
using FileDialog;
using Interoperability;
using PInvoke;
using System.IO;
using Microsoft.Win32;
using System.Windows.Forms.VisualStyles;
using System.Drawing.Imaging;
//using CCConfigStudio;

namespace Tag_Importer
{
    public partial class TagImportForm : Form
    {
        public TagImportForm()
        {
            InitializeComponent();
            textBox1_TextChanged(textBox1, new EventArgs());
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
            string configPath = "Tag_Importer.ini";
            var configFile = File.CreateText(configPath);
            configFile.WriteLine(textBox1.Text);
            configFile.WriteLine("");
            configFile.WriteLine("Authored by Muresan Radu-Adrian (MURA02)");
            configFile.Close();
        }

        private void SendKeyHandled(IntPtr windowHandle, string key/*, StreamWriter log*/)
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
                    //log.WriteLine(e.Message);
                    //log.Flush();
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

        private void button1_Click(object sender, EventArgs e)
        {
            KeepConfig();
            ImportTags();
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
            Bitmap scrollUp = (Bitmap)Image.FromFile("characters/scrollUp.png");
            Bitmap scrollDn = (Bitmap)Image.FromFile("characters/scrollDown.png");
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
            //only works for white background for now!!!
            MouseOperations.SetCursorPosition(1, 1); //have to use the found minus/plus coordinates here
            Bitmap img = MyFunctions.GetPngByHandle(handle);
            List<MyFunctions.PosLetter> allCharsFound = new List<MyFunctions.PosLetter>();
            allCharsFound.AddRange(Find(img, MyFunctions.MakeExistingTransparent((Bitmap)Image.FromFile("characters/upper_A.png")), "A"));
            allCharsFound.AddRange(Find(img, MyFunctions.MakeExistingTransparent((Bitmap)Image.FromFile("characters/upper_B.png")), "B"));
            allCharsFound.AddRange(Find(img, MyFunctions.MakeExistingTransparent((Bitmap)Image.FromFile("characters/upper_C.png")), "C"));
            allCharsFound.AddRange(Find(img, MyFunctions.MakeExistingTransparent((Bitmap)Image.FromFile("characters/upper_D.png")), "D"));
            allCharsFound.AddRange(Find(img, MyFunctions.MakeExistingTransparent((Bitmap)Image.FromFile("characters/upper_E.png")), "E"));
            allCharsFound.AddRange(Find(img, MyFunctions.MakeExistingTransparent((Bitmap)Image.FromFile("characters/upper_F.png")), "F"));
            allCharsFound.AddRange(Find(img, MyFunctions.MakeExistingTransparent((Bitmap)Image.FromFile("characters/upper_G.png")), "G"));
            allCharsFound.AddRange(Find(img, MyFunctions.MakeExistingTransparent((Bitmap)Image.FromFile("characters/upper_H.png")), "H"));
            allCharsFound.AddRange(Find(img, MyFunctions.MakeExistingTransparent((Bitmap)Image.FromFile("characters/upper_I.png")), "I"));
            allCharsFound.AddRange(Find(img, MyFunctions.MakeExistingTransparent((Bitmap)Image.FromFile("characters/upper_J.png")), "J"));
            allCharsFound.AddRange(Find(img, MyFunctions.MakeExistingTransparent((Bitmap)Image.FromFile("characters/upper_K.png")), "K"));
            allCharsFound.AddRange(Find(img, MyFunctions.MakeExistingTransparent((Bitmap)Image.FromFile("characters/upper_L.png")), "L"));
            allCharsFound.AddRange(Find(img, MyFunctions.MakeExistingTransparent((Bitmap)Image.FromFile("characters/upper_M.png")), "M"));
            allCharsFound.AddRange(Find(img, MyFunctions.MakeExistingTransparent((Bitmap)Image.FromFile("characters/upper_N.png")), "N"));
            allCharsFound.AddRange(Find(img, MyFunctions.MakeExistingTransparent((Bitmap)Image.FromFile("characters/upper_O.png")), "O"));
            allCharsFound.AddRange(Find(img, MyFunctions.MakeExistingTransparent((Bitmap)Image.FromFile("characters/upper_P.png")), "P"));
            allCharsFound.AddRange(Find(img, MyFunctions.MakeExistingTransparent((Bitmap)Image.FromFile("characters/upper_Q.png")), "Q"));
            allCharsFound.AddRange(Find(img, MyFunctions.MakeExistingTransparent((Bitmap)Image.FromFile("characters/upper_R.png")), "R"));
            allCharsFound.AddRange(Find(img, MyFunctions.MakeExistingTransparent((Bitmap)Image.FromFile("characters/upper_S.png")), "S"));
            allCharsFound.AddRange(Find(img, MyFunctions.MakeExistingTransparent((Bitmap)Image.FromFile("characters/upper_T.png")), "T"));
            allCharsFound.AddRange(Find(img, MyFunctions.MakeExistingTransparent((Bitmap)Image.FromFile("characters/upper_U.png")), "U"));
            allCharsFound.AddRange(Find(img, MyFunctions.MakeExistingTransparent((Bitmap)Image.FromFile("characters/upper_V.png")), "V"));
            allCharsFound.AddRange(Find(img, MyFunctions.MakeExistingTransparent((Bitmap)Image.FromFile("characters/upper_W.png")), "W"));
            allCharsFound.AddRange(Find(img, MyFunctions.MakeExistingTransparent((Bitmap)Image.FromFile("characters/upper_X.png")), "X"));
            allCharsFound.AddRange(Find(img, MyFunctions.MakeExistingTransparent((Bitmap)Image.FromFile("characters/upper_Y.png")), "Y"));
            allCharsFound.AddRange(Find(img, MyFunctions.MakeExistingTransparent((Bitmap)Image.FromFile("characters/upper_Z.png")), "Z"));

            allCharsFound.AddRange(Find(img, MyFunctions.MakeExistingTransparent((Bitmap)Image.FromFile("characters/lower_A.png")), "a"));
            allCharsFound.AddRange(Find(img, MyFunctions.MakeExistingTransparent((Bitmap)Image.FromFile("characters/lower_B.png")), "b"));
            allCharsFound.AddRange(Find(img, MyFunctions.MakeExistingTransparent((Bitmap)Image.FromFile("characters/lower_C.png")), "c"));
            allCharsFound.AddRange(Find(img, MyFunctions.MakeExistingTransparent((Bitmap)Image.FromFile("characters/lower_D.png")), "d"));
            allCharsFound.AddRange(Find(img, MyFunctions.MakeExistingTransparent((Bitmap)Image.FromFile("characters/lower_E.png")), "e"));
            allCharsFound.AddRange(Find(img, MyFunctions.MakeExistingTransparent((Bitmap)Image.FromFile("characters/lower_F.png")), "f"));
            allCharsFound.AddRange(Find(img, MyFunctions.MakeExistingTransparent((Bitmap)Image.FromFile("characters/lower_G.png")), "g"));
            allCharsFound.AddRange(Find(img, MyFunctions.MakeExistingTransparent((Bitmap)Image.FromFile("characters/lower_H.png")), "h"));
            allCharsFound.AddRange(Find(img, MyFunctions.MakeExistingTransparent((Bitmap)Image.FromFile("characters/lower_I.png")), "i"));
            allCharsFound.AddRange(Find(img, MyFunctions.MakeExistingTransparent((Bitmap)Image.FromFile("characters/lower_J.png")), "j"));
            allCharsFound.AddRange(Find(img, MyFunctions.MakeExistingTransparent((Bitmap)Image.FromFile("characters/lower_K.png")), "k"));
            allCharsFound.AddRange(Find(img, MyFunctions.MakeExistingTransparent((Bitmap)Image.FromFile("characters/lower_L.png")), "l"));
            allCharsFound.AddRange(Find(img, MyFunctions.MakeExistingTransparent((Bitmap)Image.FromFile("characters/lower_M.png")), "m"));
            allCharsFound.AddRange(Find(img, MyFunctions.MakeExistingTransparent((Bitmap)Image.FromFile("characters/lower_N.png")), "n"));
            allCharsFound.AddRange(Find(img, MyFunctions.MakeExistingTransparent((Bitmap)Image.FromFile("characters/lower_O.png")), "o"));
            allCharsFound.AddRange(Find(img, MyFunctions.MakeExistingTransparent((Bitmap)Image.FromFile("characters/lower_P.png")), "p"));
            allCharsFound.AddRange(Find(img, MyFunctions.MakeExistingTransparent((Bitmap)Image.FromFile("characters/lower_Q.png")), "q"));
            allCharsFound.AddRange(Find(img, MyFunctions.MakeExistingTransparent((Bitmap)Image.FromFile("characters/lower_R.png")), "r"));
            allCharsFound.AddRange(Find(img, MyFunctions.MakeExistingTransparent((Bitmap)Image.FromFile("characters/lower_S.png")), "s"));
            allCharsFound.AddRange(Find(img, MyFunctions.MakeExistingTransparent((Bitmap)Image.FromFile("characters/lower_T.png")), "t"));
            allCharsFound.AddRange(Find(img, MyFunctions.MakeExistingTransparent((Bitmap)Image.FromFile("characters/lower_U.png")), "u"));
            allCharsFound.AddRange(Find(img, MyFunctions.MakeExistingTransparent((Bitmap)Image.FromFile("characters/lower_V.png")), "v"));
            allCharsFound.AddRange(Find(img, MyFunctions.MakeExistingTransparent((Bitmap)Image.FromFile("characters/lower_W.png")), "w"));
            allCharsFound.AddRange(Find(img, MyFunctions.MakeExistingTransparent((Bitmap)Image.FromFile("characters/lower_X.png")), "x"));
            allCharsFound.AddRange(Find(img, MyFunctions.MakeExistingTransparent((Bitmap)Image.FromFile("characters/lower_Y.png")), "y"));
            allCharsFound.AddRange(Find(img, MyFunctions.MakeExistingTransparent((Bitmap)Image.FromFile("characters/lower_Z.png")), "z"));
            allCharsFound.AddRange(Find(img, MyFunctions.MakeExistingTransparent((Bitmap)Image.FromFile("characters/lower_rf.png")), "rf"));

            allCharsFound.AddRange(Find(img, MyFunctions.MakeExistingTransparent((Bitmap)Image.FromFile("characters/underscore.png")), "_"));
            allCharsFound.AddRange(Find(img, MyFunctions.MakeExistingTransparent((Bitmap)Image.FromFile("characters/arond.png")), "@"));
            allCharsFound.AddRange(Find(img, MyFunctions.MakeExistingTransparent((Bitmap)Image.FromFile("characters/#.png")), "#"));
            allCharsFound.AddRange(Find(img, MyFunctions.MakeExistingTransparent((Bitmap)Image.FromFile("characters/1.png")), "1"));
            allCharsFound.AddRange(Find(img, MyFunctions.MakeExistingTransparent((Bitmap)Image.FromFile("characters/2.png")), "2"));
            allCharsFound.AddRange(Find(img, MyFunctions.MakeExistingTransparent((Bitmap)Image.FromFile("characters/3.png")), "3"));
            allCharsFound.AddRange(Find(img, MyFunctions.MakeExistingTransparent((Bitmap)Image.FromFile("characters/4.png")), "4"));
            allCharsFound.AddRange(Find(img, MyFunctions.MakeExistingTransparent((Bitmap)Image.FromFile("characters/5.png")), "5"));

            return allCharsFound;
        }

        private void ImportTags()
        {
            #region tree view handle capture
            IntPtr tag = PInvokeLibrary.FindWindow("WinCC ConfigurationStudio MainWindow", "Tag Management - WinCC Configuration Studio");
            IntPtr ccAx = PInvokeLibrary.FindWindowEx(tag, IntPtr.Zero, "CCAxControlContainerWindow", null);
            IntPtr navBar = PInvokeLibrary.FindWindowEx(ccAx, IntPtr.Zero, "WinCC NavigationBarControl Window", null);
            IntPtr treeView = PInvokeLibrary.FindWindowEx(navBar, IntPtr.Zero, "WinCC ConfigurationStudio NavigationBarTreeView", null);
            IntPtr trHandle = PInvokeLibrary.FindWindowEx(treeView, IntPtr.Zero, "SysTreeView32", "");
            _ = PInvokeLibrary.GetWindowRect(trHandle, out RECT trRect);
            #endregion

            List<IntPtr> ccAxs = GetAllChildrenWindowHandles(tag, 4);
            var widthsccAx = new List<int>();
            for (int i = 1; i < ccAxs.Count; i++)
            {
                _ = PInvokeLibrary.GetWindowRect(ccAxs[i], out RECT rect);
                int width = rect.right - rect.left;
                widthsccAx.Add(width);
            }
            int no = widthsccAx.FindIndex(a => a == widthsccAx.OrderByDescending(c => c).Skip(1).FirstOrDefault());
            IntPtr ccAxt = ccAxs[no + 1]; //hope this holds together - second largest width of ccax
            _ = PInvokeLibrary.GetWindowRect(ccAxt, out RECT ccAxtRect);
            IntPtr dataGridHandle = PInvokeLibrary.FindWindowEx(ccAxt, IntPtr.Zero, "WinCC DataGridControl Window", null);

            #region lookInDirectory

            DirectoryInfo dinfo = new DirectoryInfo(textBox1.Text);
            FileInfo[] Files = dinfo.GetFiles("*.txt");

            #endregion

            #region image character recognition

            //ExpandOrHideVisibleTree(tag, trHandle, expand: true);

            //find required tag group in treeview, expand and scroll until found
            //if structure tags is not visible even after moving a little bit then it's not in the field of view at all - that's ok
            List<WordWithLocation> rowData = GetWordsInHandle(trHandle);

            #region Expand Only Tag Management
            var img = MyFunctions.GetPngByHandle(trHandle);
            var scrollState = GetHandleScrollState(img);
            //rowData = HideStructureTags(trHandle, rowData, trRect);
            Bitmap scrollUp = (Bitmap)Image.FromFile("characters/scrollUp.png");
            Bitmap scrollDn = (Bitmap)Image.FromFile("characters/scrollDown.png");
            if (scrollState != 1 || scrollState != 4)
            {
                do
                {
                    var goUpHere = Find(img, scrollUp).FirstOrDefault();
                    ClickInWindowAtXY(trHandle, trRect.left + goUpHere.X, trRect.top + goUpHere.Y);
                    img = MyFunctions.GetPngByHandle(trHandle);

                    scrollState = GetHandleScrollState(img);
                    img.Dispose();
                    System.Threading.Thread.Sleep(10);
                } while (scrollState != 1 && scrollState != 4);
            }
            if (FindClosestMatch(rowData, "TagManagement") == null)
            {
                SendKeyHandled(trHandle, "{DOWN}"); SendKeyHandled(trHandle, "{DOWN}");
            }
            ExpandTreeItem(trHandle, "TagManagement", true, trRect);
            #endregion

            //now get structures

            scrollUp.Dispose();
            scrollDn.Dispose();

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
            //checkedListBox1.CheckedItems;
            for (int fileNo = 0; fileNo < Files.Length; fileNo++)
            {
                FileInfo file = Files[fileNo];
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
                }

                ExpandTreeItem(trHandle, grup[1], true, trRect);
                ExpandTreeItem(trHandle, grup[2], true, trRect);

                for (int i = 0; i < nc.Count; i++)
                {
                    DeleteExistingTags(trHandle, trRect, ccAxtRect, dataGridHandle, nc, i);
                }

                ExpandTreeItem(trHandle, grup[2], false, trRect);
                ExpandTreeItem(trHandle, grup[1], false, trRect);

                ImportTagFile(tag, file);


            }

            #endregion

            MessageBox.Show(new Form { TopMost = true }, "Finished importing from specified folder");
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

        private void DeleteExistingTags(IntPtr trHandle, RECT trRect, RECT ccAxtRect, IntPtr dataGridHandle, List<NameConnection> nc, int i)
        {
            List<WordWithLocation> rowData;
            NameConnection connData = nc[i];
            //Console.WriteLine(connData.connection + ", " + connData.name);
            ExpandTreeItem(trHandle, connData.connection, true, trRect);
            rowData = GetWordsInHandle(trHandle);
            var myGroup = FindClosestMatch(rowData, connData.name);
            ClickInWindowAtXY(trHandle, trRect.left + myGroup.x, trRect.top + myGroup.y);
            //now delete the tags
            ClickInWindowAtXY(dataGridHandle, ccAxtRect.left + 100, ccAxtRect.top + 100); //click in data grid
            SendKeyHandled(dataGridHandle, "^(a)");
            SendKeyHandled(dataGridHandle, "{DELETE}");
            ExpandTreeItem(trHandle, connData.connection, false, trRect);
        }

        private List<WordWithLocation> HideStructureTags(IntPtr trHandle, List<WordWithLocation> rowData, RECT trRect)
        {
            if (FindClosestMatch(rowData, "Structuretags") == null)
            {
                SendKeyHandled(trHandle, "{UP}"); SendKeyHandled(trHandle, "{UP}"); rowData = GetWordsInHandle(trHandle);
            }
            if (FindClosestMatch(rowData, "Structuretags") != null)
            {
                rowData = GetWordsInHandle(trHandle);
                ExpandTreeItem(trHandle, "Structuretags", false, trRect);
            }
            return rowData;
        }

        private void ExpandTreeItem(IntPtr trHandle, string FindThis, bool expand, RECT trRect)
        {
            List<WordWithLocation> rowData = GetWordsInHandle(trHandle);
            WordWithLocation foundElem = FindClosestMatch(rowData, FindThis);
            var getElement = TextHasExpandOrHide(trHandle, foundElem, expand);
            if (getElement != null)
            {
                ClickInWindowAtXY(trHandle, trRect.left + getElement?.X, trRect.top + getElement?.Y);
            }
        }

        private Point? TextHasExpandOrHide(IntPtr trHandle, WordWithLocation foundElem, bool expand)
        {
            Bitmap find = expand != true ? (Bitmap)Image.FromFile("characters/minus.png") : (Bitmap)Image.FromFile("characters/plus.png");
            var img = MyFunctions.GetPngByHandle(trHandle);
            var foundExpandButtons = Find(img, find);
            if (foundExpandButtons.Where(c => Math.Abs(c.Y - foundElem.y) < 4 && Math.Abs(c.X - foundElem.x) < 150).ToList().Count > 0)
            {
                return foundExpandButtons.FirstOrDefault(c => Math.Abs(c.Y - foundElem.y) < 4 && Math.Abs(c.X - foundElem.x) < 150);
            }
            return null;
        }

        private bool isFirstImporting;

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

            if (levDist > 5)
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
