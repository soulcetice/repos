using PInvoke;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.ConstrainedExecution;
using System.Runtime.InteropServices;
using System.Security;
using System.Text;
using System.Threading.Tasks;
using System.Drawing;
using System.ComponentModel;
using System.Drawing.Imaging;
using System.Xml;

namespace Interoperability
{
    public class MouseOperations
    {
        [Flags]
        public enum MouseEventFlags
        {
            LeftDown = 0x00000002,
            LeftUp = 0x00000004,
            MiddleDown = 0x00000020,
            MiddleUp = 0x00000040,
            Move = 0x00000001,
            Absolute = 0x00008000,
            RightDown = 0x00000008,
            RightUp = 0x00000010
        }

        [DllImport("user32.dll", EntryPoint = "SetCursorPos")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool SetCursorPos(int x, int y);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool GetCursorPos(out MousePoint lpMousePoint);

        [DllImport("user32.dll")]
        private static extern void mouse_event(int dwFlags, int dx, int dy, int dwData, int dwExtraInfo);

        public static void SetCursorPosition(int x, int y)
        {
            SetCursorPos(x, y);
        }

        public static void SetCursorPosition(MousePoint point)
        {
            SetCursorPos(point.X, point.Y);
        }

        public static MousePoint GetCursorPosition()
        {
            MousePoint currentMousePoint;
            var gotPoint = GetCursorPos(out currentMousePoint);
            if (!gotPoint) { currentMousePoint = new MousePoint(0, 0); }
            return currentMousePoint;
        }

        public static void MouseEvent(MouseEventFlags value)
        {
            MousePoint position = GetCursorPosition();

            mouse_event
                ((int)value,
                 position.X,
                 position.Y,
                 0,
                 0)
                ;
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct MousePoint
        {
            public int X;
            public int Y;

            public MousePoint(int x, int y)
            {
                X = x;
                Y = y;
            }
        }
    }

    public class PInvokeLibrary
    {
        public enum ShowWindowCommands
        {
            /// <summary>
            /// Hides the window and activates another window.
            /// </summary>
            Hide = 0,
            /// <summary>
            /// Activates and displays a window. If the window is minimized or
            /// maximized, the system restores it to its original size and position.
            /// An application should specify this flag when displaying the window
            /// for the first time.
            /// </summary>
            Normal = 1,
            /// <summary>
            /// Activates the window and displays it as a minimized window.
            /// </summary>
            ShowMinimized = 2,
            /// <summary>
            /// Maximizes the specified window.
            /// </summary>
            Maximize = 3, // is this the right value?
            /// <summary>
            /// Activates the window and displays it as a maximized window.
            /// </summary>      
            ShowMaximized = 3,
            /// <summary>
            /// Displays a window in its most recent size and position. This value
            /// is similar to <see cref="Win32.ShowWindowCommand.Normal"/>, except
            /// the window is not activated.
            /// </summary>
            ShowNoActivate = 4,
            /// <summary>
            /// Activates the window and displays it in its current size and position.
            /// </summary>
            Show = 5,
            /// <summary>
            /// Minimizes the specified window and activates the next top-level
            /// window in the Z order.
            /// </summary>
            Minimize = 6,
            /// <summary>
            /// Displays the window as a minimized window. This value is similar to
            /// <see cref="Win32.ShowWindowCommand.ShowMinimized"/>, except the
            /// window is not activated.
            /// </summary>
            ShowMinNoActive = 7,
            /// <summary>
            /// Displays the window in its current size and position. This value is
            /// similar to <see cref="Win32.ShowWindowCommand.Show"/>, except the
            /// window is not activated.
            /// </summary>
            ShowNA = 8,
            /// <summary>
            /// Activates and displays the window. If the window is minimized or
            /// maximized, the system restores it to its original size and position.
            /// An application should specify this flag when restoring a minimized window.
            /// </summary>
            Restore = 9,
            /// <summary>
            /// Sets the show state based on the SW_* value specified in the
            /// STARTUPINFO structure passed to the CreateProcess function by the
            /// program that started the application.
            /// </summary>
            ShowDefault = 10,
            /// <summary>
            ///  <b>Windows 2000/XP:</b> Minimizes a window, even if the thread
            /// that owns the window is not responding. This flag should only be
            /// used when minimizing windows from a different thread.
            /// </summary>
            ForceMinimize = 11
        }
        /// <summary>
        /// Contains information about the placement of a window on the screen.
        /// </summary>
        [Serializable]
        [StructLayout(LayoutKind.Sequential)]
        public struct WINDOWPLACEMENT
        {
            /// <summary>
            /// The length of the structure, in bytes. Before calling the GetWindowPlacement or SetWindowPlacement functions, set this member to sizeof(WINDOWPLACEMENT).
            /// <para>
            /// GetWindowPlacement and SetWindowPlacement fail if this member is not set correctly.
            /// </para>
            /// </summary>
            public int Length;

            /// <summary>
            /// Specifies flags that control the position of the minimized window and the method by which the window is restored.
            /// </summary>
            public int Flags;

            /// <summary>
            /// The current show state of the window.
            /// </summary>
            public ShowWindowCommands ShowCmd;

            /// <summary>
            /// The coordinates of the window's upper-left corner when the window is minimized.
            /// </summary>
            public POINT MinPosition;

            /// <summary>
            /// The coordinates of the window's upper-left corner when the window is maximized.
            /// </summary>
            public POINT MaxPosition;

            /// <summary>
            /// The window's coordinates when the window is in the restored position.
            /// </summary>
            public RECT NormalPosition;

            /// <summary>
            /// Gets the default (empty) value.
            /// </summary>
            public static WINDOWPLACEMENT Default
            {
                get
                {
                    WINDOWPLACEMENT result = new WINDOWPLACEMENT();
                    result.Length = Marshal.SizeOf(result);
                    return result;
                }
            }
        }

        //moved from ncm_downloader
        #region ImportDlls

        //only using these dlls to change the rdp window size
        [DllImport("user32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool GetWindowPlacement(IntPtr hWnd, ref WINDOWPLACEMENT lpwndpl);

        [DllImport("user32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool SetWindowPlacement(IntPtr hWnd, [In] ref WINDOWPLACEMENT lpwndpl);

        // Get a handle to an application window.
        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        public static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

        [DllImport("user32.dll", SetLastError = true)]
        public static extern bool GetWindowRect(IntPtr hwnd, out RECT lpRect);

        // Define the SetWindowPos API function.
        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int X, int Y, int cx, int cy, SetWindowPosFlags uFlags);

        //get objects in window ?
        [DllImport("user32.dll")]
        public static extern IntPtr FindWindowEx(IntPtr handleParent, IntPtr handleChild, string className, string WindowName);

        // Activate an application window.
        [DllImport("user32.dll")]
        public static extern bool SetForegroundWindow(IntPtr hWnd);

        internal delegate int WindowEnumProc(IntPtr hwnd, IntPtr lparam);

        // Sends message to ptr
        [DllImport("user32.dll")]
        public static extern IntPtr SendMessage(IntPtr hWnd, int Msg, int wParam, IntPtr lParam);
        [System.Runtime.InteropServices.DllImport("user32.dll", EntryPoint = "SendMessage", CharSet = System.Runtime.InteropServices.CharSet.Auto)]
        public static extern bool SendMessage(IntPtr hWnd, uint Msg, int wParam, StringBuilder lParam);
        [System.Runtime.InteropServices.DllImport("user32.dll", SetLastError = true)]
        public static extern IntPtr SendMessage(int hWnd, int Msg, int wparam, int lparam);

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        public static extern IntPtr SendMessage(HandleRef hWnd, uint Msg, int wParam, StringBuilder lParam);


        [DllImport("user32.dll")]
        public static extern IntPtr PostMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);

        [DllImport("user32.dll", SetLastError = true)]
        public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

        [DllImport("kernel32.dll", SetLastError = true)]
        public static extern IntPtr OpenProcess(ProcessAccessFlags processAccess, bool bInheritHandle, int processId);
        public static IntPtr OpenProcess(Process proc, ProcessAccessFlags flags) { return OpenProcess(flags, false, proc.Id); }

        [DllImport("kernel32.dll", SetLastError = true)]
        [ReliabilityContract(Consistency.WillNotCorruptState, Cer.Success)]
        [SuppressUnmanagedCodeSecurity]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool CloseHandle(IntPtr hObject);

        [DllImport("user32.dll")]
        public static extern bool IsWindowUnicode(IntPtr hwnd);

        [DllImport("kernel32.dll", SetLastError = true, ExactSpelling = true)]
        public static extern bool VirtualFreeEx(IntPtr hProcess, IntPtr lpAddress, int dwSize, AllocationType dwFreeType);

        [DllImport("kernel32.dll", SetLastError = true)]
        public static extern bool ReadProcessMemory(
            IntPtr hProcess,
            IntPtr lpBaseAddress,
            [Out] byte[] lpBuffer,
            int dwSize,
            out IntPtr lpNumberOfBytesRead);

        [DllImport("kernel32.dll", SetLastError = true)]
        public static extern bool ReadProcessMemory(
            IntPtr hProcess,
            IntPtr lpBaseAddress,
            [Out, MarshalAs(UnmanagedType.AsAny)] object lpBuffer,
            int dwSize,
            out IntPtr lpNumberOfBytesRead);

        [DllImport("kernel32.dll", SetLastError = true)]
        public static extern bool ReadProcessMemory(
            IntPtr hProcess,
            IntPtr lpBaseAddress,
            IntPtr lpBuffer,
            int dwSize,
            out IntPtr lpNumberOfBytesRead);


        [DllImport("kernel32.dll", SetLastError = true, ExactSpelling = true)]
        public static extern IntPtr VirtualAllocEx(IntPtr hProcess, IntPtr lpAddress, uint dwSize, AllocationType flAllocationType, MemoryProtection flProtect);

        [DllImport("kernel32.dll", SetLastError = true)]
        public static extern bool WriteProcessMemory(IntPtr hProcess, IntPtr lpBaseAddress, [MarshalAs(UnmanagedType.AsAny)] object lpBuffer, int dwSize, out IntPtr lpNumberOfBytesWritten);



        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        public static extern IntPtr GetDlgItem(IntPtr hwnd, int childID);

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        public static extern int SendMessage(HandleRef hwnd, int wMsg, int wParam, String s);

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        public static extern String SendMessage(HandleRef hwnd, uint WM_GETTEXT);

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        public static extern int SendMessage(int hWnd, int msg, int wParam, IntPtr lParam);
        // to get file size import
        [DllImport("kernel32.dll")]
        public static extern bool GetFileSizeEx(IntPtr hFile, out long lpFileSize);



        [DllImport("kernel32.dll", CharSet = CharSet.Auto)]
        static extern IntPtr FindFirstFile(string lpFileName, out WIN32_FIND_DATA lpFindFileData);

        [DllImport("kernel32.dll")]
        static extern IntPtr FindFirstFile(IntPtr lpfilename, ref WIN32_FIND_DATA findfiledata);

        [DllImport("kernel32.dll")]
        static extern IntPtr FindClose(IntPtr pff);

        [DllImport("user32.dll", SetLastError = true)]
        public static extern IntPtr GetWindow(IntPtr hWnd, GetWindowType uCmd);

        #endregion


        public enum GetWindowType : uint
        {
            /// <summary>
            /// The retrieved handle identifies the window of the same type that is highest in the Z order.
            /// <para/>
            /// If the specified window is a topmost window, the handle identifies a topmost window.
            /// If the specified window is a top-level window, the handle identifies a top-level window.
            /// If the specified window is a child window, the handle identifies a sibling window.
            /// </summary>
            GW_HWNDFIRST = 0,
            /// <summary>
            /// The retrieved handle identifies the window of the same type that is lowest in the Z order.
            /// <para />
            /// If the specified window is a topmost window, the handle identifies a topmost window.
            /// If the specified window is a top-level window, the handle identifies a top-level window.
            /// If the specified window is a child window, the handle identifies a sibling window.
            /// </summary>
            GW_HWNDLAST = 1,
            /// <summary>
            /// The retrieved handle identifies the window below the specified window in the Z order.
            /// <para />
            /// If the specified window is a topmost window, the handle identifies a topmost window.
            /// If the specified window is a top-level window, the handle identifies a top-level window.
            /// If the specified window is a child window, the handle identifies a sibling window.
            /// </summary>
            GW_HWNDNEXT = 2,
            /// <summary>
            /// The retrieved handle identifies the window above the specified window in the Z order.
            /// <para />
            /// If the specified window is a topmost window, the handle identifies a topmost window.
            /// If the specified window is a top-level window, the handle identifies a top-level window.
            /// If the specified window is a child window, the handle identifies a sibling window.
            /// </summary>
            GW_HWNDPREV = 3,
            /// <summary>
            /// The retrieved handle identifies the specified window's owner window, if any.
            /// </summary>
            GW_OWNER = 4,
            /// <summary>
            /// The retrieved handle identifies the child window at the top of the Z order,
            /// if the specified window is a parent window; otherwise, the retrieved handle is NULL.
            /// The function examines only child windows of the specified window. It does not examine descendant windows.
            /// </summary>
            GW_CHILD = 5,
            /// <summary>
            /// The retrieved handle identifies the enabled popup window owned by the specified window (the
            /// search uses the first such window found using GW_HWNDNEXT); otherwise, if there are no enabled
            /// popup windows, the retrieved handle is that of the specified window.
            /// </summary>
            GW_ENABLEDPOPUP = 6
        }

        public struct WIN32_FIND_DATA
        {
            public int dwFileAttributes;
            public System.Runtime.InteropServices.ComTypes.FILETIME ftCreationTime;
            public System.Runtime.InteropServices.ComTypes.FILETIME ftLastAccessTime;
            public System.Runtime.InteropServices.ComTypes.FILETIME ftLastWriteTime;
            public int nFileSizeHigh;
            public int nFileSizeLow;
            public int dwReserved0;
            public int dwReserved1;
            public string cFileName; //mite need marshalling, TCHAR size = MAX_PATH???
            public string cAlternateFileName; //mite need marshalling, TCHAR size = 14
        }
        public struct WIN32_FIND_DATA1
        {
            public int dwFileAttributes;
            public System.Runtime.InteropServices.ComTypes.FILETIME ftCreationTime;
            public System.Runtime.InteropServices.ComTypes.FILETIME ftLastAccessTime;
            public System.Runtime.InteropServices.ComTypes.FILETIME ftLastWriteTime;
            public int nFileSizeHigh;
            public int nFileSizeLow;
            public int dwReserved0;
            public int dwReserved1;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 260)]
            public string cFileName;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 14)]
            public string cAlternateFileName;
        }

        #region Flags
        [Flags]

        public enum SetWindowPosFlags : uint
        {
            SynchronousWindowPosition = 0x4000,
            DeferErase = 0x2000,
            DrawFrame = 0x0020,
            FrameChanged = 0x0020,
            HideWindow = 0x0080,
            DoNotActivate = 0x0010,
            DoNotCopyBits = 0x0100,
            IgnoreMove = 0x0002,
            DoNotChangeOwnerZOrder = 0x0200,
            DoNotRedraw = 0x0008,
            DoNotReposition = 0x0200,
            DoNotSendChangingEvent = 0x0400,
            IgnoreResize = 0x0001,
            IgnoreZOrder = 0x0004,
            ShowWindow = 0x0040,
        }
        public enum ProcessAccessFlags : uint
        {
            All = 0x001F0FFF,
            Terminate = 0x00000001,
            CreateThread = 0x00000002,
            VirtualMemoryOperation = 0x00000008,
            VirtualMemoryRead = 0x00000010,
            VirtualMemoryWrite = 0x00000020,
            DuplicateHandle = 0x00000040,
            CreateProcess = 0x000000080,
            SetQuota = 0x00000100,
            SetInformation = 0x00000200,
            QueryInformation = 0x00000400,
            QueryLimitedInformation = 0x00001000,
            Synchronize = 0x00100000
        }
        public enum AllocationType
        {
            Commit = 0x1000,
            Reserve = 0x2000,
            Decommit = 0x4000,
            Release = 0x8000,
            Reset = 0x80000,
            Physical = 0x400000,
            TopDown = 0x100000,
            WriteWatch = 0x200000,
            LargePages = 0x20000000
        }
        public enum MemoryProtection
        {
            Execute = 0x10,
            ExecuteRead = 0x20,
            ExecuteReadWrite = 0x40,
            ExecuteWriteCopy = 0x80,
            NoAccess = 0x01,
            ReadOnly = 0x02,
            ReadWrite = 0x04,
            WriteCopy = 0x08,
            GuardModifierflag = 0x100,
            NoCacheModifierflag = 0x200,
            WriteCombineModifierflag = 0x400
        }
        public struct TVITEM
        {
            public uint mask;
            public IntPtr hItem;
            public uint state;
            public uint stateMask;
            public IntPtr pszText;
            public int cchTextMax;
            public int iImage;
            public int iSelectedImage;
            public int cChildren;
            public IntPtr lParam;
        }
        public enum TreeViewMsg
        {
            BN_CLICKED = 0xf5,
            TV_CHECKED = 0x2000,
            TV_FIRST = 0x1100,
            TVGN_ROOT = 0x0,
            TVGN_NEXT = 0x1,
            TVGN_CHILD = 0x4,
            TVGN_FIRSTVISIBLE = 0x5,
            TVGN_NEXTVISIBLE = 0x6,
            TVGN_CARET = 0x9,
            TVM_SELECTITEM = (TV_FIRST + 11),
            TVM_GETNEXTITEM = (TV_FIRST + 10),
            TVM_GETITEM = (TV_FIRST + 12),
            TVIF_TEXT = 0x1
        }
        public enum FreeType
        {
            Decommit = 0x4000,
            Release = 0x8000,
        }
        #endregion
    }

    public class MyFunctions
    {
        public static PInvokeLibrary.WINDOWPLACEMENT GetPlacementByHandle(IntPtr handle)
        {
            // Prepare the WINDOWPLACEMENT structure.
            var placement = new PInvokeLibrary.WINDOWPLACEMENT();
            placement.Length = Marshal.SizeOf(placement);
            // Get the window's current placement.
            PInvokeLibrary.GetWindowPlacement(handle, ref placement);

            return placement;
        }

        public static Rectangle RECTToRectangle(RECT rect)
        {
            var rectangle = new Rectangle()
            {
                X = rect.left,
                Y = rect.top,
                Height = rect.bottom - rect.top,
                Width = rect.right - rect.left,
                //Location = new Point(rect.left, rect.top),
                //Size = new Size(rect.right - rect.left, rect.bottom - rect.top)
            };
            return rectangle;
        }

        public static Bitmap GetPngByHandle(IntPtr handle)
        {
            var rect = new RECT();

            if (!PInvokeLibrary.SetForegroundWindow(handle))
                throw new Win32Exception(Marshal.GetLastWin32Error());

            if (!PInvokeLibrary.GetWindowRect(handle, out rect))
                throw new Win32Exception(Marshal.GetLastWin32Error());

            System.Threading.Thread.Sleep(500);

            Rectangle windowSize = RECTToRectangle(rect);
            Bitmap target = new Bitmap(windowSize.Width, windowSize.Height);
            using (Graphics g = Graphics.FromImage(target))
            {
                g.CopyFromScreen(windowSize.X, windowSize.Y, 0, 0, new Size(windowSize.Width, windowSize.Height));
            }

            target.MakeTransparent();

            //target = target.GetPixel();

            target.Save("output.png", System.Drawing.Imaging.ImageFormat.Png);            

            return target;
        }

        public static Bitmap MakeExistingTransparent(Bitmap img)
        {
            img.MakeTransparent();
            return img;
        }

        public static Bitmap ConvertPng(Bitmap minus, System.Drawing.Imaging.PixelFormat format)
        {
            Bitmap minusClone = new Bitmap(minus.Width, minus.Height,
                format);
            using (Graphics gr = Graphics.FromImage(minusClone))
            {
                gr.DrawImage(minus, new Rectangle(0, 0, minusClone.Width, minusClone.Height));
            }

            return minusClone;
        }

        private static List<Point> FindBitmapsEntry(Bitmap sourceBitmap, Bitmap searchingBitmap)
        {
            #region Arguments check
            if (sourceBitmap.PixelFormat != searchingBitmap.PixelFormat)
                searchingBitmap = ConvertPng(searchingBitmap, sourceBitmap.PixelFormat);

            if (sourceBitmap == null || searchingBitmap == null)
                throw new ArgumentNullException();

            if (sourceBitmap.PixelFormat != searchingBitmap.PixelFormat)
                throw new ArgumentException("Pixel formats aren't equal");

            if (sourceBitmap.Width < searchingBitmap.Width || sourceBitmap.Height < searchingBitmap.Height)
                throw new ArgumentException("Size of searchingBitmap bigger then sourceBitmap");

            #endregion

            var pixelFormatSize = Image.GetPixelFormatSize(sourceBitmap.PixelFormat) / 8;


            // Copy sourceBitmap to byte array
            var sourceBitmapData = sourceBitmap.LockBits(new Rectangle(0, 0, sourceBitmap.Width, sourceBitmap.Height),
                ImageLockMode.ReadOnly, sourceBitmap.PixelFormat);
            var sourceBitmapBytesLength = sourceBitmapData.Stride * sourceBitmap.Height;
            var sourceBytes = new byte[sourceBitmapBytesLength];
            Marshal.Copy(sourceBitmapData.Scan0, sourceBytes, 0, sourceBitmapBytesLength);
            sourceBitmap.UnlockBits(sourceBitmapData);

            // Copy searchingBitmap to byte array
            var searchingBitmapData =
                searchingBitmap.LockBits(new Rectangle(0, 0, searchingBitmap.Width, searchingBitmap.Height),
                    ImageLockMode.ReadOnly, searchingBitmap.PixelFormat);
            var searchingBitmapBytesLength = searchingBitmapData.Stride * searchingBitmap.Height;
            var searchingBytes = new byte[searchingBitmapBytesLength];
            Marshal.Copy(searchingBitmapData.Scan0, searchingBytes, 0, searchingBitmapBytesLength);
            searchingBitmap.UnlockBits(searchingBitmapData);

            var pointsList = new List<Point>();

            // Serching entries
            // minimazing searching zone
            // sourceBitmap.Height - searchingBitmap.Height + 1
            for (var mainY = 0; mainY < sourceBitmap.Height - searchingBitmap.Height + 1; mainY++)
            {
                var sourceY = mainY * sourceBitmapData.Stride;

                for (var mainX = 0; mainX < sourceBitmap.Width - searchingBitmap.Width + 1; mainX++)
                {// mainY & mainX - pixel coordinates of sourceBitmap
                 // sourceY + sourceX = pointer in array sourceBitmap bytes
                    var sourceX = mainX * pixelFormatSize;

                    var isEqual = true;
                    for (var c = 0; c < pixelFormatSize; c++)
                    {// through the bytes in pixel
                        if (sourceBytes[sourceX + sourceY + c] == searchingBytes[c])
                            continue;
                        isEqual = false;
                        break;
                    }

                    if (!isEqual) continue;

                    var isStop = false;

                    // find fist equalation and now we go deeper) 
                    for (var secY = 0; secY < searchingBitmap.Height; secY++)
                    {
                        var searchY = secY * searchingBitmapData.Stride;

                        var sourceSecY = (mainY + secY) * sourceBitmapData.Stride;

                        for (var secX = 0; secX < searchingBitmap.Width; secX++)
                        {// secX & secY - coordinates of searchingBitmap
                         // searchX + searchY = pointer in array searchingBitmap bytes

                            var searchX = secX * pixelFormatSize;

                            var sourceSecX = (mainX + secX) * pixelFormatSize;

                            for (var c = 0; c < pixelFormatSize; c++)
                            {// through the bytes in pixel
                                if (sourceBytes[sourceSecX + sourceSecY + c] == searchingBytes[searchX + searchY + c]) continue;

                                // not equal - abort iteration
                                isStop = true;
                                break;
                            }

                            if (isStop) break;
                        }

                        if (isStop) break;
                    }

                    if (!isStop)
                    {// searching bitmap is found!!
                        pointsList.Add(new Point(mainX, mainY));
                    }
                }
            }
            return pointsList;
        }

        private static List<PosLetter> FindLetters(Bitmap sourceBitmap, Bitmap searchingBitmap, [Optional] string myLetter)
        {
            #region Arguments check
            if (sourceBitmap.PixelFormat != searchingBitmap.PixelFormat)
                searchingBitmap = ConvertPng(searchingBitmap, sourceBitmap.PixelFormat);

            if (sourceBitmap == null || searchingBitmap == null)
                throw new ArgumentNullException();

            if (sourceBitmap.PixelFormat != searchingBitmap.PixelFormat)
                throw new ArgumentException("Pixel formats aren't equal");

            if (sourceBitmap.Width < searchingBitmap.Width || sourceBitmap.Height < searchingBitmap.Height)
                throw new ArgumentException("Size of searchingBitmap bigger then sourceBitmap");
            #endregion

            var pixelFormatSize = Image.GetPixelFormatSize(sourceBitmap.PixelFormat) / 8;


            // Copy sourceBitmap to byte array
            var sourceBitmapData = sourceBitmap.LockBits(new Rectangle(0, 0, sourceBitmap.Width, sourceBitmap.Height),
                ImageLockMode.ReadOnly, sourceBitmap.PixelFormat);
            var sourceBitmapBytesLength = sourceBitmapData.Stride * sourceBitmap.Height;
            var sourceBytes = new byte[sourceBitmapBytesLength];
            Marshal.Copy(sourceBitmapData.Scan0, sourceBytes, 0, sourceBitmapBytesLength);
            sourceBitmap.UnlockBits(sourceBitmapData);

            // Copy searchingBitmap to byte array
            var searchingBitmapData =
                searchingBitmap.LockBits(new Rectangle(0, 0, searchingBitmap.Width, searchingBitmap.Height),
                    ImageLockMode.ReadOnly, searchingBitmap.PixelFormat);
            var searchingBitmapBytesLength = searchingBitmapData.Stride * searchingBitmap.Height;
            var searchingBytes = new byte[searchingBitmapBytesLength];
            Marshal.Copy(searchingBitmapData.Scan0, searchingBytes, 0, searchingBitmapBytesLength);
            searchingBitmap.UnlockBits(searchingBitmapData);

            var pointsList = new List<PosLetter>();

            // Serching entries
            // minimazing searching zone
            // sourceBitmap.Height - searchingBitmap.Height + 1
            for (var mainY = 0; mainY < sourceBitmap.Height - searchingBitmap.Height + 1; mainY++)
            {
                var sourceY = mainY * sourceBitmapData.Stride;

                for (var mainX = 0; mainX < sourceBitmap.Width - searchingBitmap.Width + 1; mainX++)
                {// mainY & mainX - pixel coordinates of sourceBitmap
                 // sourceY + sourceX = pointer in array sourceBitmap bytes
                    var sourceX = mainX * pixelFormatSize;

                    var isEqual = true;
                    for (var c = 0; c < pixelFormatSize; c++)
                    {// through the bytes in pixel
                        if (sourceBytes[sourceX + sourceY + c] == searchingBytes[c])
                            continue;
                        isEqual = false;
                        break;
                    }

                    if (!isEqual) continue;

                    var isStop = false;

                    // find fist equalation and now we go deeper) 
                    for (var secY = 0; secY < searchingBitmap.Height; secY++)
                    {
                        var searchY = secY * searchingBitmapData.Stride;

                        var sourceSecY = (mainY + secY) * sourceBitmapData.Stride;

                        for (var secX = 0; secX < searchingBitmap.Width; secX++)
                        {// secX & secY - coordinates of searchingBitmap
                         // searchX + searchY = pointer in array searchingBitmap bytes

                            var searchX = secX * pixelFormatSize;

                            var sourceSecX = (mainX + secX) * pixelFormatSize;

                            for (var c = 0; c < pixelFormatSize; c++)
                            {// through the bytes in pixel
                                int tol = 90;
                                if (sourceBytes[sourceSecX + sourceSecY + c] > searchingBytes[searchX + searchY + c] - tol && sourceBytes[sourceSecX + sourceSecY + c] < searchingBytes[searchX + searchY + c] + tol)
                                {
                                    continue;
                                }
                                else
                                {
                                    // not equal - abort iteration
                                    isStop = true;
                                    break;
                                }
                            }

                            if (isStop) break;
                        }

                        if (isStop) break;
                    }

                    if (!isStop)
                    {// searching bitmap is found!!
                        pointsList.Add(new PosLetter() { x = mainX, y = mainY, letter = myLetter });
                    }
                }
            }
            return pointsList;
        }

        public class PosLetter
        {
            public int x;
            public int y;
            public string letter;
        }
    }
}
