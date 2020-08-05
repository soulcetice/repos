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
                src = Find(img, find);
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

        private void ImportTags()
        {
            #region tree view handle capture
            IntPtr tag = PInvokeLibrary.FindWindow("WinCC ConfigurationStudio MainWindow", "Tag Management - WinCC Configuration Studio");
            IntPtr ccAx = PInvokeLibrary.FindWindowEx(tag, IntPtr.Zero, "CCAxControlContainerWindow", null);
            IntPtr navBar = PInvokeLibrary.FindWindowEx(ccAx, IntPtr.Zero, "WinCC NavigationBarControl Window", null);
            IntPtr treeView = PInvokeLibrary.FindWindowEx(navBar, IntPtr.Zero, "WinCC ConfigurationStudio NavigationBarTreeView", null);
            IntPtr trHandle = PInvokeLibrary.FindWindowEx(treeView, IntPtr.Zero, "SysTreeView32", "");
            #endregion

            #region lookInDirectory

            DirectoryInfo dinfo = new DirectoryInfo(textBox1.Text);
            FileInfo[] Files = dinfo.GetFiles("*.txt");

            foreach (var elem in Files)
            {
                checkedListBox1.Items.Add(elem);
            }

            #endregion

            #region image character recognition

            //find required tag group in treeview, expand and scroll until found
            List<WordWithLocation> rowData = GetWordsInHandle(trHandle);
            var myElem = rowData.FirstOrDefault(c => c.word == "iTracking");
            //ExpandOrHideVisibleTree(tag, trHandle, expand: true);

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
            foreach (var file in Files)
            {
                ImportTagFile(tag, file);
            }

            #endregion

            MessageBox.Show(new Form { TopMost = true }, "Finished importing from specified folder");
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

        private static WordWithLocation RobustFileChoice(FileInfo f, List<WordWithLocation> filesInDialog)
        {
            var fileName = Path.GetFileNameWithoutExtension(f.FullName);
            var tagFile = filesInDialog.FirstOrDefault(c => c.word == fileName);
            if (tagFile == null)
            {
                int levDist = 50;
                string chosenFile = "";
                foreach (var t in filesInDialog)
                {
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
            var list = new List<Point>();

            foreach (var firstLineMatchPoint in FindMatch(haystackArray.Take(haystack.Height - needle.Height), needleArray[0]))
            {
                if (IsNeedlePresentAtLocation(haystackArray, needleArray, firstLineMatchPoint, 1))
                {
                    list.Add(firstLineMatchPoint);
                }
            }

            return list;
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
