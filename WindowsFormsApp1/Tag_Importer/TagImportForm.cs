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
//using CCConfigStudio;

namespace Tag_Importer
{
    public partial class TagImportForm : Form
    {
        public TagImportForm()
        {
            InitializeComponent();
        }

        private IntPtr OpenTagMgmtMenu(IntPtr tagMgmt)
        {
            PInvokeLibrary.SetForegroundWindow(tagMgmt);

            IntPtr importPopup = IntPtr.Zero;
            importPopup = PInvokeLibrary.FindWindow("#32770", "Import");

            if (importPopup == IntPtr.Zero)
            {
                SendKeyHandled(tagMgmt, "(%)");
                SendKeyHandled(tagMgmt, "{RIGHT}");
                SendKeyHandled(tagMgmt, "{ENTER}");
                SendKeyHandled(tagMgmt, "{DOWN}");
                SendKeyHandled(tagMgmt, "{DOWN}");
                SendKeyHandled(tagMgmt, "{DOWN}");
                SendKeyHandled(tagMgmt, "{DOWN}");
                //SendKeyHandled(tagMgmt, "{DOWN}");
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

        private void button1_Click(object sender, EventArgs e)
        {
            //FileDialog.Program.GetFilesInDialog();
            ImportTags();
        }

        private void ClickInWindowAtXY(IntPtr handle, int x, int y)
        {
            PInvokeLibrary.SetForegroundWindow(handle);

            MouseOperations.SetCursorPosition(x, y); //have to use the found minus/plus coordinates here
            MouseOperations.MouseEvent(MouseOperations.MouseEventFlags.LeftDown);
        }

        private void ExpandOrHideVisibleTree(IntPtr tagMgmt, IntPtr trHandle, bool expand)
        {
            Bitmap img;
            Bitmap find = expand != true ? (Bitmap)Image.FromFile("minus.png") : (Bitmap)Image.FromFile("plus.png");
            List<Point> src;
            _ = PInvokeLibrary.SetForegroundWindow(trHandle);
            _ = PInvokeLibrary.GetWindowRect(trHandle, out RECT trRect);

            do //extend or hide all visible
            {
                img = MyFunctions.GetPngByHandle(trHandle);
                src = MyFunctions.FindBitmapsEntry(img, find);
                for (int i = 0; i < src.Count; i++)
                {
                    if (src[i].Y != 6 && src[i].X == 7) break;
                    ClickInWindowAtXY(tagMgmt, trRect.left + src[i].X, trRect.top + src[i].Y); //have to use the found minus/plus coordinates here
                }
            } while (src.Count > 0);

            find.Dispose();
        }

        private List<MyFunctions.PosLetter> GetCharactersInHandle(IntPtr handle)
        {
            Bitmap img = MyFunctions.GetPngByHandle(handle);
            List<MyFunctions.PosLetter> allCharsFound = new List<MyFunctions.PosLetter>();
            allCharsFound.AddRange(MyFunctions.FindLetters(img, (Bitmap)Image.FromFile("characters/upper_A.png"), "A"));
            allCharsFound.AddRange(MyFunctions.FindLetters(img, (Bitmap)Image.FromFile("characters/upper_B.png"), "B"));
            allCharsFound.AddRange(MyFunctions.FindLetters(img, (Bitmap)Image.FromFile("characters/upper_C.png"), "C"));
            allCharsFound.AddRange(MyFunctions.FindLetters(img, (Bitmap)Image.FromFile("characters/upper_D.png"), "D"));
            allCharsFound.AddRange(MyFunctions.FindLetters(img, (Bitmap)Image.FromFile("characters/upper_E.png"), "E"));
            allCharsFound.AddRange(MyFunctions.FindLetters(img, (Bitmap)Image.FromFile("characters/upper_F.png"), "F"));
            allCharsFound.AddRange(MyFunctions.FindLetters(img, (Bitmap)Image.FromFile("characters/upper_G.png"), "G"));
            allCharsFound.AddRange(MyFunctions.FindLetters(img, (Bitmap)Image.FromFile("characters/upper_H.png"), "H"));
            allCharsFound.AddRange(MyFunctions.FindLetters(img, (Bitmap)Image.FromFile("characters/upper_I.png"), "I"));
            allCharsFound.AddRange(MyFunctions.FindLetters(img, (Bitmap)Image.FromFile("characters/upper_J.png"), "J"));
            allCharsFound.AddRange(MyFunctions.FindLetters(img, (Bitmap)Image.FromFile("characters/upper_K.png"), "K"));
            allCharsFound.AddRange(MyFunctions.FindLetters(img, (Bitmap)Image.FromFile("characters/upper_L.png"), "L"));
            allCharsFound.AddRange(MyFunctions.FindLetters(img, (Bitmap)Image.FromFile("characters/upper_M.png"), "M"));
            allCharsFound.AddRange(MyFunctions.FindLetters(img, (Bitmap)Image.FromFile("characters/upper_N.png"), "N"));
            allCharsFound.AddRange(MyFunctions.FindLetters(img, (Bitmap)Image.FromFile("characters/upper_O.png"), "O"));
            allCharsFound.AddRange(MyFunctions.FindLetters(img, (Bitmap)Image.FromFile("characters/upper_P.png"), "P"));
            allCharsFound.AddRange(MyFunctions.FindLetters(img, (Bitmap)Image.FromFile("characters/upper_Q.png"), "Q"));
            allCharsFound.AddRange(MyFunctions.FindLetters(img, (Bitmap)Image.FromFile("characters/upper_R.png"), "R"));
            allCharsFound.AddRange(MyFunctions.FindLetters(img, (Bitmap)Image.FromFile("characters/upper_S.png"), "S"));
            allCharsFound.AddRange(MyFunctions.FindLetters(img, (Bitmap)Image.FromFile("characters/upper_T.png"), "T"));
            allCharsFound.AddRange(MyFunctions.FindLetters(img, (Bitmap)Image.FromFile("characters/upper_U.png"), "U"));
            allCharsFound.AddRange(MyFunctions.FindLetters(img, (Bitmap)Image.FromFile("characters/upper_V.png"), "V"));
            allCharsFound.AddRange(MyFunctions.FindLetters(img, (Bitmap)Image.FromFile("characters/upper_W.png"), "W"));
            allCharsFound.AddRange(MyFunctions.FindLetters(img, (Bitmap)Image.FromFile("characters/upper_X.png"), "X"));
            allCharsFound.AddRange(MyFunctions.FindLetters(img, (Bitmap)Image.FromFile("characters/upper_Y.png"), "Y"));
            allCharsFound.AddRange(MyFunctions.FindLetters(img, (Bitmap)Image.FromFile("characters/upper_Z.png"), "Z"));

            allCharsFound.AddRange(MyFunctions.FindLetters(img, (Bitmap)Image.FromFile("characters/lower_A.png"), "a"));
            allCharsFound.AddRange(MyFunctions.FindLetters(img, (Bitmap)Image.FromFile("characters/lower_B.png"), "b"));
            allCharsFound.AddRange(MyFunctions.FindLetters(img, (Bitmap)Image.FromFile("characters/lower_C.png"), "c"));
            allCharsFound.AddRange(MyFunctions.FindLetters(img, (Bitmap)Image.FromFile("characters/lower_D.png"), "d"));
            allCharsFound.AddRange(MyFunctions.FindLetters(img, (Bitmap)Image.FromFile("characters/lower_E.png"), "e"));
            allCharsFound.AddRange(MyFunctions.FindLetters(img, (Bitmap)Image.FromFile("characters/lower_F.png"), "f"));
            allCharsFound.AddRange(MyFunctions.FindLetters(img, (Bitmap)Image.FromFile("characters/lower_G.png"), "g"));
            allCharsFound.AddRange(MyFunctions.FindLetters(img, (Bitmap)Image.FromFile("characters/lower_H.png"), "h"));
            allCharsFound.AddRange(MyFunctions.FindLetters(img, (Bitmap)Image.FromFile("characters/lower_I.png"), "i"));
            allCharsFound.AddRange(MyFunctions.FindLetters(img, (Bitmap)Image.FromFile("characters/lower_J.png"), "j"));
            allCharsFound.AddRange(MyFunctions.FindLetters(img, (Bitmap)Image.FromFile("characters/lower_K.png"), "k"));
            allCharsFound.AddRange(MyFunctions.FindLetters(img, (Bitmap)Image.FromFile("characters/lower_L.png"), "l"));
            allCharsFound.AddRange(MyFunctions.FindLetters(img, (Bitmap)Image.FromFile("characters/lower_M.png"), "m"));
            allCharsFound.AddRange(MyFunctions.FindLetters(img, (Bitmap)Image.FromFile("characters/lower_N.png"), "n"));
            allCharsFound.AddRange(MyFunctions.FindLetters(img, (Bitmap)Image.FromFile("characters/lower_O.png"), "o"));
            allCharsFound.AddRange(MyFunctions.FindLetters(img, (Bitmap)Image.FromFile("characters/lower_P.png"), "p"));
            allCharsFound.AddRange(MyFunctions.FindLetters(img, (Bitmap)Image.FromFile("characters/lower_Q.png"), "q"));
            allCharsFound.AddRange(MyFunctions.FindLetters(img, (Bitmap)Image.FromFile("characters/lower_R.png"), "r"));
            allCharsFound.AddRange(MyFunctions.FindLetters(img, (Bitmap)Image.FromFile("characters/lower_S.png"), "s"));
            allCharsFound.AddRange(MyFunctions.FindLetters(img, (Bitmap)Image.FromFile("characters/lower_T.png"), "t"));
            allCharsFound.AddRange(MyFunctions.FindLetters(img, (Bitmap)Image.FromFile("characters/lower_U.png"), "u"));
            allCharsFound.AddRange(MyFunctions.FindLetters(img, (Bitmap)Image.FromFile("characters/lower_V.png"), "v"));
            allCharsFound.AddRange(MyFunctions.FindLetters(img, (Bitmap)Image.FromFile("characters/lower_W.png"), "w"));
            allCharsFound.AddRange(MyFunctions.FindLetters(img, (Bitmap)Image.FromFile("characters/lower_X.png"), "x"));
            allCharsFound.AddRange(MyFunctions.FindLetters(img, (Bitmap)Image.FromFile("characters/lower_Y.png"), "y"));
            allCharsFound.AddRange(MyFunctions.FindLetters(img, (Bitmap)Image.FromFile("characters/lower_Z.png"), "z"));
            allCharsFound.AddRange(MyFunctions.FindLetters(img, (Bitmap)Image.FromFile("characters/lower_rf.png"), "rf"));

            allCharsFound.AddRange(MyFunctions.FindLetters(img, (Bitmap)Image.FromFile("characters/underscore.png"), "_"));
            allCharsFound.AddRange(MyFunctions.FindLetters(img, (Bitmap)Image.FromFile("characters/arond.png"), "@"));
            allCharsFound.AddRange(MyFunctions.FindLetters(img, (Bitmap)Image.FromFile("characters/#.png"), "#"));
            allCharsFound.AddRange(MyFunctions.FindLetters(img, (Bitmap)Image.FromFile("characters/1.png"), "1"));
            allCharsFound.AddRange(MyFunctions.FindLetters(img, (Bitmap)Image.FromFile("characters/2.png"), "2"));
            allCharsFound.AddRange(MyFunctions.FindLetters(img, (Bitmap)Image.FromFile("characters/3.png"), "3"));
            allCharsFound.AddRange(MyFunctions.FindLetters(img, (Bitmap)Image.FromFile("characters/4.png"), "4"));
            allCharsFound.AddRange(MyFunctions.FindLetters(img, (Bitmap)Image.FromFile("characters/5.png"), "5"));

            return allCharsFound;
        }

        private List<string> GetTreeText(IntPtr trHandle)
        {
            //all characters should be taken at the same top value x to i can sort easily.

            List<MyFunctions.PosLetter> allCharsFound = GetCharactersInHandle(trHandle);

            Bitmap img = MyFunctions.GetPngByHandle(trHandle);
            int imgHeight = img.Height;
            img.Dispose();

            var orderedChars = allCharsFound.OrderBy(c => c.y).ThenBy(c => c.x).ToList();

            List<List<MyFunctions.PosLetter>> filterRows = new List<List<MyFunctions.PosLetter>>();
            int firsty = 0;
            int inc = -1;
            foreach (MyFunctions.PosLetter c in orderedChars)
            {
                if (firsty == 0)
                {
                    firsty = c.y;
                    filterRows.Add(new List<MyFunctions.PosLetter>());
                    inc++;
                };

                if (c.y >= firsty && c.y < firsty + 15)
                {
                    filterRows[inc].Add(c);
                }
                else
                {
                    firsty = c.y;
                    filterRows.Add(new List<MyFunctions.PosLetter>());
                    inc++;
                    filterRows[inc].Add(c);
                }
            }

            var moreFilter = (from List<MyFunctions.PosLetter> row in filterRows
                              select row.OrderBy(c => c.x).ToList()).ToList();

            //each row is once every 20 px
            List<string> allRows = new List<string>();
            for (int i = 10; i < imgHeight; i += 20)
            {
                int lastx = -1;
                List<MyFunctions.PosLetter> row = new List<MyFunctions.PosLetter>();
                foreach (var c in from item in moreFilter
                                  from c in
                                      from c in item
                                      where c.y > i - 10 && c.y < i + 10 && c.x > lastx
                                      select c
                                  select c)
                {
                    row.Add(c);
                    lastx = c.x;
                }

                List<MyFunctions.PosLetter> orderedRow = row.OrderBy(c => c.x).ToList();
                string myWord = "";
                foreach (MyFunctions.PosLetter c in orderedRow)
                {
                    myWord += c.letter;
                }
                allRows.Add(myWord);
            }
            return allRows;
        }

        private void Testing()
        {
            #region grafexe
            //grafexe.HMIObject t;
            //grafexe.HMIObjects ts;

            //grafexe.Application app;

            //ts = app.ActiveDocument.HMIObjects;

            //foreach (var obj in ts)
            //{

            //}
            #endregion
        }

        private void ImportTags()
        {
            IntPtr tag = PInvokeLibrary.FindWindow("WinCC ConfigurationStudio MainWindow", "Tag Management - WinCC Configuration Studio");
            IntPtr ccAx = PInvokeLibrary.FindWindowEx(tag, IntPtr.Zero, "CCAxControlContainerWindow", null);
            IntPtr navBar = PInvokeLibrary.FindWindowEx(ccAx, IntPtr.Zero, "WinCC NavigationBarControl Window", null);
            IntPtr treeView = PInvokeLibrary.FindWindowEx(navBar, IntPtr.Zero, "WinCC ConfigurationStudio NavigationBarTreeView", null);
            IntPtr trHandle = PInvokeLibrary.FindWindowEx(treeView, IntPtr.Zero, "SysTreeView32", "");

            System.Threading.Thread.Sleep(100);

            #region lookInDirectory

            DirectoryInfo dinfo = new DirectoryInfo(textBox1.Text);
            FileInfo[] Files = dinfo.GetFiles("*.txt");

            #endregion

            #region imageRecognitionAndAction
            //List<string> rowData = GetTreeText(trHandle);
            //int myElem = rowData.IndexOf("iTracking");
            ExpandOrHideVisibleTree(tag, trHandle, expand: true);

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
            //open import dialog

            //do scrolling, same manner

            #endregion

            foreach (var file in Files)
            {
                ImportTagFile(tag, file);
            }
        }

        private void ImportTagFile(IntPtr tag, FileInfo f)
        {
            IntPtr importPopup = OpenTagMgmtMenu(tag);
            System.Threading.Thread.Sleep(500);
            //find files in import dialog
            IntPtr duiview = PInvokeLibrary.FindWindowEx(importPopup, IntPtr.Zero, "DUIViewWndClassName", "");
            IntPtr directuihwnd = PInvokeLibrary.FindWindowEx(duiview, IntPtr.Zero, "DirectUIHWND", "");
            var ctrlNotifySinks = GetAllChildrenWindowHandles(directuihwnd, 5);

            int oldWidth = 0;
            IntPtr CtrlNotifySink = IntPtr.Zero;
            foreach (var elem in ctrlNotifySinks)
            {
                _ = PInvokeLibrary.GetWindowRect(elem, out RECT rect);
                if (rect.right - rect.left > oldWidth) CtrlNotifySink = elem;
                oldWidth = rect.right - rect.left;
            }

            IntPtr fileListParent = PInvokeLibrary.FindWindowEx(CtrlNotifySink, IntPtr.Zero, "SHELLDLL_DefView", null);
            IntPtr fileList = PInvokeLibrary.FindWindowEx(fileListParent, IntPtr.Zero, "DirectUIHWND", null);

            var filesInDialog = GetWordsInHandle(fileListParent);

            _ = PInvokeLibrary.GetWindowRect(fileList, out RECT fileListRect);

            var fileName = Path.GetFileNameWithoutExtension(f.FullName);
            var tagFile = filesInDialog.FirstOrDefault(c => c.word == fileName);

            ClickInWindowAtXY(fileList, fileListRect.left + tagFile.x, fileListRect.top + tagFile.y);
            SendKeyHandled(fileList, "{ENTER}");
            System.Threading.Thread.Sleep(1000);

            IntPtr confirmationPopup = PInvokeLibrary.FindWindow("#32770", "Import");
            if (confirmationPopup != IntPtr.Zero)
            {
                SendKeyHandled(confirmationPopup, "{ENTER}");
                System.Threading.Thread.Sleep(500);
            }

            //IntPtr addressBar = GetChildBySubstring("Address:", importPopup);
            //Console.WriteLine("Done");
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
                    if (yDifs[i] > 10) //the yband tolerance
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
            foreach (List<MyFunctions.PosLetter> word in sortedList)
            {
                words.Add(new WordWithLocation());
                foreach (MyFunctions.PosLetter c in word)
                {
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

        private class WordWithLocation
        {
            public int x;
            public int y;
            public string word;
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
    }
}
