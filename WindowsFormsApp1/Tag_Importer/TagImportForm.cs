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
//using CCConfigStudio;
using Tesseract;
using System.Net.PeerToPeer.Collaboration;
using tessnet2;
using System.Runtime.ExceptionServices;

namespace Tag_Importer
{
    public partial class TagImportForm : Form
    {
        public TagImportForm()
        {
            InitializeComponent();
        }

        private void OpenTagMgmtMenu(IntPtr tagMgmt)
        {
            PInvokeLibrary.SetForegroundWindow(tagMgmt);

            SendKeyHandled(tagMgmt, "(%)");
            SendKeyHandled(tagMgmt, "{RIGHT}");
            SendKeyHandled(tagMgmt, "{ENTER}");
            SendKeyHandled(tagMgmt, "{DOWN}");
            SendKeyHandled(tagMgmt, "{DOWN}");
            SendKeyHandled(tagMgmt, "{DOWN}");
            SendKeyHandled(tagMgmt, "{DOWN}");
            //SendKeyHandled(tagMgmt, "{DOWN}");
            SendKeyHandled(tagMgmt, "{ENTER}");
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

        public List<string> ExtractWindowTextByHandle(IntPtr handle)
        {
            var extractedText = new List<string>();
            List<IntPtr> childObjects = new WindowHandleInfo(handle).GetAllChildHandles();
            for (int i = 0; i < childObjects.Count; i++)
            {
                extractedText.Add(GetControlText(childObjects[i]));
            }
            return extractedText;
        }

        public string GetControlText(IntPtr hWnd)
        {

            // Get the size of the string required to hold the window title (including trailing null.) 
            Int32 titleSize = PInvokeLibrary.SendMessage(hWnd, (int)WindowsMessages.WM_GETTEXTLENGTH, 0, IntPtr.Zero).ToInt32();

            // If titleSize is 0, there is no title so return an empty string (or null)
            if (titleSize == 0)
                return String.Empty;

            StringBuilder title = new StringBuilder(titleSize + 1);

            PInvokeLibrary.SendMessage(hWnd, (int)WindowsMessages.WM_GETTEXT, title.Capacity, title);

            return title.ToString();
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

        private void ClickOnExpandOrRevert(IntPtr handle, int x, int y)
        {
            PInvokeLibrary.SetForegroundWindow(handle);

            MouseOperations.SetCursorPosition(x, y); //have to use the found minus/plus coordinates here
            MouseOperations.MouseEvent(MouseOperations.MouseEventFlags.LeftDown);
        }

        private List<string> GetCharacters(Bitmap img)
        {
            //all characters should be taken at the same top value x to i can sort easily.

            Bitmap up_A = (Bitmap)Image.FromFile("characters/upper_A.png");
            Bitmap up_B = (Bitmap)Image.FromFile("characters/upper_B.png");
            Bitmap up_C = (Bitmap)Image.FromFile("characters/upper_C.png");
            Bitmap up_D = (Bitmap)Image.FromFile("characters/upper_D.png");
            Bitmap up_E = (Bitmap)Image.FromFile("characters/upper_E.png");
            Bitmap up_F = (Bitmap)Image.FromFile("characters/upper_F.png");
            Bitmap up_G = (Bitmap)Image.FromFile("characters/upper_G.png");
            Bitmap up_H = (Bitmap)Image.FromFile("characters/upper_H.png");
            Bitmap up_I = (Bitmap)Image.FromFile("characters/upper_I.png");
            Bitmap up_J = (Bitmap)Image.FromFile("characters/upper_J.png");
            Bitmap up_K = (Bitmap)Image.FromFile("characters/upper_K.png");
            Bitmap up_L = (Bitmap)Image.FromFile("characters/upper_L.png");
            Bitmap up_M = (Bitmap)Image.FromFile("characters/upper_M.png");
            Bitmap up_N = (Bitmap)Image.FromFile("characters/upper_N.png");
            Bitmap up_O = (Bitmap)Image.FromFile("characters/upper_O.png");
            Bitmap up_P = (Bitmap)Image.FromFile("characters/upper_P.png");
            Bitmap up_Q = (Bitmap)Image.FromFile("characters/upper_Q.png");
            Bitmap up_R = (Bitmap)Image.FromFile("characters/upper_R.png");
            Bitmap up_S = (Bitmap)Image.FromFile("characters/upper_S.png");
            Bitmap up_T = (Bitmap)Image.FromFile("characters/upper_T.png");
            Bitmap up_U = (Bitmap)Image.FromFile("characters/upper_U.png");
            Bitmap up_V = (Bitmap)Image.FromFile("characters/upper_V.png");
            Bitmap up_W = (Bitmap)Image.FromFile("characters/upper_W.png");
            Bitmap up_X = (Bitmap)Image.FromFile("characters/upper_X.png");
            Bitmap up_Y = (Bitmap)Image.FromFile("characters/upper_Y.png");
            Bitmap up_Z = (Bitmap)Image.FromFile("characters/upper_Z.png");

            Bitmap lo_A = (Bitmap)Image.FromFile("characters/lower_A.png");
            Bitmap lo_B = (Bitmap)Image.FromFile("characters/lower_B.png");
            Bitmap lo_C = (Bitmap)Image.FromFile("characters/lower_C.png");
            Bitmap lo_D = (Bitmap)Image.FromFile("characters/lower_D.png");
            Bitmap lo_E = (Bitmap)Image.FromFile("characters/lower_E.png");
            Bitmap lo_F = (Bitmap)Image.FromFile("characters/lower_F.png");
            Bitmap lo_G = (Bitmap)Image.FromFile("characters/lower_G.png");
            Bitmap lo_H = (Bitmap)Image.FromFile("characters/lower_H.png");
            Bitmap lo_I = (Bitmap)Image.FromFile("characters/lower_I.png");
            Bitmap lo_J = (Bitmap)Image.FromFile("characters/lower_J.png");
            Bitmap lo_K = (Bitmap)Image.FromFile("characters/lower_K.png");
            Bitmap lo_L = (Bitmap)Image.FromFile("characters/lower_L.png");
            Bitmap lo_M = (Bitmap)Image.FromFile("characters/lower_M.png");
            Bitmap lo_N = (Bitmap)Image.FromFile("characters/lower_N.png");
            Bitmap lo_O = (Bitmap)Image.FromFile("characters/lower_O.png");
            Bitmap lo_P = (Bitmap)Image.FromFile("characters/lower_P.png");
            Bitmap lo_Q = (Bitmap)Image.FromFile("characters/lower_Q.png");
            Bitmap lo_R = (Bitmap)Image.FromFile("characters/lower_R.png");
            Bitmap lo_S = (Bitmap)Image.FromFile("characters/lower_S.png");
            Bitmap lo_T = (Bitmap)Image.FromFile("characters/lower_T.png");
            Bitmap lo_U = (Bitmap)Image.FromFile("characters/lower_U.png");
            Bitmap lo_V = (Bitmap)Image.FromFile("characters/lower_V.png");
            Bitmap lo_W = (Bitmap)Image.FromFile("characters/lower_W.png");
            Bitmap lo_X = (Bitmap)Image.FromFile("characters/lower_X.png");
            Bitmap lo_Y = (Bitmap)Image.FromFile("characters/lower_Y.png");
            Bitmap lo_Z = (Bitmap)Image.FromFile("characters/lower_Z.png");
            Bitmap usc = (Bitmap)Image.FromFile("characters/underscore.png");
            Bitmap arn = (Bitmap)Image.FromFile("characters/arond.png");
            Bitmap one = (Bitmap)Image.FromFile("characters/1.png");
            Bitmap two = (Bitmap)Image.FromFile("characters/2.png");
            Bitmap fiv = (Bitmap)Image.FromFile("characters/5.png");

            List<string> alphabet = new List<string>
            {
                "A",
                "B",
                "C",
                "D",
                "E",
                "F",
                "G",
                "H",
                "I",
                "J",
                "K",
                "L",
                "M",
                "N",
                "O",
                "P",
                "Q",
                "R",
                "S",
                "T",
                "U",
                "V",
                "W",
                "X",
                "Y",
                "Z",

                "a",
                "b",
                "c",
                "d",
                "e",
                "f",
                "g",
                "h",
                "i",
                "j",
                "k",
                "l",
                "m",
                "n",
                "o",
                "p",
                "q",
                "r",
                "s",
                "t",
                "u",
                "v",
                "w",
                "x",
                "y",
                "z"
            };

            var allCharsFound = new List<List<MyFunctions.posLetter>>
            {
                MyFunctions.FindLetters(img, up_A, "A"),
                MyFunctions.FindLetters(img, up_B, "B"),
                MyFunctions.FindLetters(img, up_C, "C"),
                MyFunctions.FindLetters(img, up_D, "D"),
                MyFunctions.FindLetters(img, up_E, "E"),
                MyFunctions.FindLetters(img, up_F, "F"),
                MyFunctions.FindLetters(img, up_G, "G"),
                MyFunctions.FindLetters(img, up_H, "H"),
                MyFunctions.FindLetters(img, up_I, "I"),
                MyFunctions.FindLetters(img, up_J, "J"),
                MyFunctions.FindLetters(img, up_K, "K"),
                MyFunctions.FindLetters(img, up_L, "L"),
                MyFunctions.FindLetters(img, up_M, "M"),
                MyFunctions.FindLetters(img, up_N, "N"),
                MyFunctions.FindLetters(img, up_O, "O"),
                MyFunctions.FindLetters(img, up_P, "P"),
                MyFunctions.FindLetters(img, up_Q, "Q"),
                MyFunctions.FindLetters(img, up_R, "R"),
                MyFunctions.FindLetters(img, up_S, "S"),
                MyFunctions.FindLetters(img, up_T, "T"),
                MyFunctions.FindLetters(img, up_U, "U"),
                MyFunctions.FindLetters(img, up_V, "V"),
                MyFunctions.FindLetters(img, up_W, "W"),
                MyFunctions.FindLetters(img, up_X, "X"),
                MyFunctions.FindLetters(img, up_Y, "Y"),
                MyFunctions.FindLetters(img, up_Z, "Z"),

                MyFunctions.FindLetters(img, lo_A,"a"),
                MyFunctions.FindLetters(img, lo_B,"b"),
                MyFunctions.FindLetters(img, lo_C,"c"),
                MyFunctions.FindLetters(img, lo_D,"d"),
                MyFunctions.FindLetters(img, lo_E,"e"),
                MyFunctions.FindLetters(img, lo_F,"f"),
                MyFunctions.FindLetters(img, lo_G,"g"),
                MyFunctions.FindLetters(img, lo_H,"h"),
                MyFunctions.FindLetters(img, lo_I,"i"),
                MyFunctions.FindLetters(img, lo_J,"j"),
                MyFunctions.FindLetters(img, lo_K,"k"),
                MyFunctions.FindLetters(img, lo_L,"l"),
                MyFunctions.FindLetters(img, lo_M,"m"),
                MyFunctions.FindLetters(img, lo_N,"n"),
                MyFunctions.FindLetters(img, lo_O,"o"),
                MyFunctions.FindLetters(img, lo_P,"p"),
                MyFunctions.FindLetters(img, lo_Q,"q"),
                MyFunctions.FindLetters(img, lo_R,"r"),
                MyFunctions.FindLetters(img, lo_S,"s"),
                MyFunctions.FindLetters(img, lo_T,"t"),
                MyFunctions.FindLetters(img, lo_U,"u"),
                MyFunctions.FindLetters(img, lo_V,"v"),
                MyFunctions.FindLetters(img, lo_W,"w"),
                MyFunctions.FindLetters(img, lo_X,"x"),
                MyFunctions.FindLetters(img, lo_Y,"y"),
                MyFunctions.FindLetters(img, lo_Z,"z"),

                MyFunctions.FindLetters(img, usc,"_"),
                MyFunctions.FindLetters(img, arn,"@"),
                MyFunctions.FindLetters(img, one,"1"),
                MyFunctions.FindLetters(img, two,"2"),
                MyFunctions.FindLetters(img, fiv,"5")
            };
            var orderedChars = new List<MyFunctions.posLetter>();
            foreach (var letter in allCharsFound)
            {
                foreach (var item in letter)
                {
                    orderedChars.Add(item);
                }
            }
            orderedChars = orderedChars.OrderBy(c => c.y).ThenBy(c => c.x).ToList();

            var filterRows = new List<List<MyFunctions.posLetter>>();
            var firsty = 0;
            int inc = -1;
            foreach (var c in orderedChars)
            {
                if (firsty == 0)
                {
                    firsty = c.y;
                    filterRows.Add(new List<MyFunctions.posLetter>());
                    inc++;
                };

                if (c.y >= firsty && c.y < firsty + 15)
                {
                    filterRows[inc].Add(c);
                }
                else
                {
                    firsty = c.y;
                    filterRows.Add(new List<MyFunctions.posLetter>());
                    inc++;
                    filterRows[inc].Add(c);
                }
            }

            var moreFilter = new List<List<MyFunctions.posLetter>>();
            foreach (var row in filterRows)
            {
                moreFilter.Add(row.OrderBy(c => c.x).ToList());
            }


            //each row is once every 20 px
            List<string> allRows = new List<string>();
            for (int i = 10; i < img.Height; i += 20)
            {
                int lastx = -1;
                var row = new List<MyFunctions.posLetter>();
                foreach (var item in moreFilter)
                {
                    foreach (var c in item)
                    {
                        if (c.y > i - 10 && c.y < i + 10 && c.x > lastx)
                        {
                            row.Add(c);
                            lastx = c.x;
                        }
                    }
                }
                var orderedRow = row.OrderBy(c => c.x).ToList();
                string myWord = "";
                foreach (var c in orderedRow)
                {
                    myWord += c.letter;
                }
                allRows.Add(myWord);
            }
            return allRows;
        }

        private void ImportTags()
        {
            IntPtr tag = PInvokeLibrary.FindWindow("WinCC ConfigurationStudio MainWindow", "Tag Management - WinCC Configuration Studio");
            IntPtr ccAx = PInvokeLibrary.FindWindowEx(tag, IntPtr.Zero, "CCAxControlContainerWindow", null);
            IntPtr navBar = PInvokeLibrary.FindWindowEx(ccAx, IntPtr.Zero, "WinCC NavigationBarControl Window", null);
            IntPtr treeView = PInvokeLibrary.FindWindowEx(navBar, IntPtr.Zero, "WinCC ConfigurationStudio NavigationBarTreeView", null);
            IntPtr trHandle = PInvokeLibrary.FindWindowEx(treeView, IntPtr.Zero, "SysTreeView32", "");

            var trRect = new RECT();
            PInvokeLibrary.SetForegroundWindow(trHandle);
            PInvokeLibrary.GetWindowRect(trHandle, out trRect);
            System.Threading.Thread.Sleep(500);

            #region imageRecognitionAndAction
            //defaults preloading
            Bitmap minus = (Bitmap)Image.FromFile("minus.png");
            Bitmap plus = (Bitmap)Image.FromFile("plus.png");
            Bitmap img;
            List<Point> searchMinus;
            List<Point> searchPlus;
            List<Point> src;

            img = MyFunctions.GetPngByHandle(trHandle);

            var rowData = GetCharacters(img);

            //ClickOnExpandOrRevert(tag, trRect.left + src[0].X, trRect.top + src[0].Y); //have to use the found minus/plus coordinates here

            do //extend all visible
            {
                img = MyFunctions.GetPngByHandle(trHandle);
                searchMinus = MyFunctions.FindBitmapsEntry(img, minus);
                src = searchMinus;
                for (int i = 0; i < src.Count; i++)
                {
                    ClickOnExpandOrRevert(tag, trRect.left + src[i].X, trRect.top + src[i].Y); //have to use the found minus/plus coordinates here
                } //always click from lower side to upper side for plus
            } while (src.Count > 0);

            do //hide all visible
            {
                img = MyFunctions.GetPngByHandle(trHandle);
                searchPlus = MyFunctions.FindBitmapsEntry(img, plus);
                src = searchPlus;
                for (int i = 0; i < src.Count; i++)
                {
                    ClickOnExpandOrRevert(tag, trRect.left + src[i].X, trRect.top + src[i].Y); //have to use the found minus/plus coordinates here
                } //always click from lower side to upper side for plus
            } while (src.Count > 0);

            #region grafexe
            //grafexe.HMIObject t;
            //grafexe.HMIObjects ts;

            //grafexe.Application app;

            //ts = app.ActiveDocument.HMIObjects;

            //foreach (var obj in ts)
            //{

            //}
            #endregion

            #region tessnet2
            //var ocr = new tessnet2.Tesseract();
            //ocr.Init(@"\\vmware-host\Shared Folders\C\Users\MURA02\source\repos\WindowsFormsApp1\packages\NuGet.Tessnet2.1.1.1\content\Content\tessdata", "eng", false);
            //var result = ocr.DoOCR(img, Rectangle.Empty);
            //foreach (tessnet2.Word word in result)
            //{
            //    //Console.WriteLine(word.Text);
            //}
            #endregion

            #region tesseract
            TesseractEngine tess = new TesseractEngine(@"\\vmware-host\Shared Folders\C\Users\MURA02\source\repos\WindowsFormsApp1\packages\NuGet.Tessnet2.1.1.1\content\Content\tessdata", "eng");
            #endregion

            //do scrolling, same manner

            minus.Dispose();
            plus.Dispose();
            img.Dispose();
            #endregion

            OpenTagMgmtMenu(tag);

            IntPtr importPopup = PInvokeLibrary.FindWindow("#32770", "Import");
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

            SendKeyHandled(importPopup, "{TAB}");
            SendKeyHandled(importPopup, "{TAB}");
            SendKeyHandled(importPopup, "{TAB}");
            SendKeyHandled(importPopup, "{TAB}");
            SendKeyHandled(importPopup, "{TAB}");

            //var len = SendMessage(importPopup, (uint)WindowsMessages.WM_GETTEXTLENGTH, 0, null);
            //var sb = new StringBuilder(len + 1);
            //SendMessage(importPopup, (uint)WindowsMessages.WM_GETTEXT, sb.Capacity, sb);

            IntPtr fileListParent = PInvokeLibrary.FindWindowEx(importPopup, IntPtr.Zero, "DUIViewWndClassName", null);
            IntPtr fileList = PInvokeLibrary.FindWindowEx(fileListParent, IntPtr.Zero, "DirectUIHWND", null);


            IntPtr addressBar = GetChildBySubstring("Address:", importPopup);

            Console.WriteLine("Done");
        }
    }
}
