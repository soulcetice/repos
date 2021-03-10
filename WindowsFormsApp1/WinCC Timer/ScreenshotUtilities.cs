using CCHMIRUNTIME;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using CommonInterops;
using grafexe;

namespace RuntimeUtilities
{
    public class RuntimeFunctions
    {
        public static string currentPage = "";
        public static string currentActiveScreen = "";
        public static List<string> navigatedPages = new List<string>();

        public static void GetRuntimeScreens(out IHMIScreens screens, out IHMIScreen activeScreen)
        {
            CCHMIRUNTIME.HMIRuntime rt = new CCHMIRUNTIME.HMIRuntime();
            CCHMIRTGRAPHICS.HMIRTGraphics graphics = new CCHMIRTGRAPHICS.HMIRTGraphics();

            screens = rt.Screens;
            activeScreen = rt.ActiveScreen;

            currentActiveScreen = activeScreen.ObjectName;
            currentPage = currentActiveScreen;
        }

        public static List<IHMIScreen> GetCurrentRuntimeFilelist()
        {
            IHMIScreens screens;
            IHMIScreen activeScreen;
            GetRuntimeScreens(out screens, out activeScreen);

            var foundScreens = new List<IHMIScreen>();

            var rt = new CCHMIRUNTIME.HMIRuntime();

            var screensList = screens.Cast<IHMIScreen>();
            List<CCHMIRUNTIME.IHMIDataSet> datasets = screensList.
                Where(c => c.Parent != null).
                Where(c => c.Parent.ObjectName.StartsWith("@DataClass")).
                Select(c => c.DataSet).
                ToList();

            foreach (CCHMIRUNTIME.IHMIScreen s in screens)
            {
                foundScreens.Add(s);
            }

            return foundScreens;
        }

        public void GetRuntimeInfo(string type, out List<ObjectData> dataList, string nameContains = "")
        {
            var objects = new List<grafexe.HMIObject>();
            dataList = new List<ObjectData>();

            GetScreenInfo(out List<IHMIScreen> filteredScreens, out List<WindowData> windowData);

            dataList = GetObjectsInfo(type, nameContains, filteredScreens, windowData);
        }

        public string GetMainScreen()
        {
            var list = new List<string>();
            GetRuntimeScreens(out IHMIScreens screens, out IHMIScreen activeScreen);
            foreach (IHMIScreen s in screens)
            {
                list.Add(s.ObjectName);
            }
            var currentNormal = list?.FirstOrDefault(c => c.ToUpper().Contains("_n_".ToUpper()));
            var currentWide = list?.FirstOrDefault(c => c.ToUpper().Contains("_w_".ToUpper()));

            var mainScreen = currentNormal != null ? currentNormal : currentWide;
            return mainScreen;
        }

        public grafexe.HMIObject RetrieveObjectProperties(string pdlName, string objName, grafexe.Application g, grafexe.HMIOpenDocumentType openType = grafexe.HMIOpenDocumentType.hmiOpenDocumentTypeInvisible)
        {
            string seldocfullname = Path.Combine(g.ApplicationDataPath, pdlName + ".pdl");

            Console.WriteLine("Opening... " + seldocfullname);

            grafexe.HMIObject go = null;

            grafexe.Document seldoc = g.Documents.Open(seldocfullname, openType);
            grafexe.HMIObjects selos = seldoc.HMIObjects;

            go = selos.Find(ObjectName: objName).Count > 0 ? selos.Find(ObjectName: objName)[1] : null;

            if (go != null)
            {
                if (go.Visible.value == false && go.Visible.DynamicStateType == HMIDynamicStateType.hmiDynamicStateTypeNoDynamic)
                    go = null;
            }

            if (go != null)
            {
                Console.WriteLine(go.ObjectName.value);
            }

            return go;
        }

        public List<ObjectData> GetObjectsInfo(string type, string nameContains, List<IHMIScreen> filteredScreens, List<WindowData> windowData)
        {
            var g = new grafexe.Application();
            var dataList = new List<ObjectData>();

            foreach (var s in filteredScreens)
            {
                var mainScreen = GetMainScreen();

                var screenData = windowData.FirstOrDefault(c => c.Name == s.Parent.ObjectName);
                var left = screenData.ScreenLeft;
                var top = screenData.ScreenTop;

                var screenItems = s.ScreenItems.Cast<IHMIScreenItem>();
                var query = screenItems.Where(o => o.ObjectName.Contains(nameContains) && o.Type == type).ToList();

                foreach (IHMIScreenItem o in query) //screenitems is rt objects
                {
                    var obj = RetrieveObjectProperties(s.ObjectName, o.ObjectName, g);

                    if (obj != null)
                    {
                        Console.WriteLine(obj.ObjectName.value + "," + obj.Left.value + "," + obj.Top.value);

                        Console.WriteLine(o.Parent.ObjectName + "," + obj.ObjectName.value + "," + obj.Left.value + "," + obj.Top.value, "\\Screen.log", false);
                        Console.WriteLine(o.Parent.ObjectName + "," + obj.ObjectName.value + "," +
                            (left + obj.Left.value) + "," +
                            (top + obj.Top.value) + "," +
                            (left + obj.Left.value + obj.Width.value) + "," +
                            (top + obj.Top.value + obj.Height.value), "\\Screen.log", false);

                        dataList.Add(new ObjectData()
                        {
                            RealLeft = left + obj.Left.value,
                            RealRight = left + obj.Left.value + obj.Width.value,
                            RealTop = top + obj.Top.value,
                            RealBottom = top + obj.Top.value + obj.Height.value,

                            OffsetLeft = left,
                            OffsetTop = top,
                            ObjectName = obj.ObjectName.value,
                            Page = s.ObjectName,
                            CallsPdl = ""
                        });

                        var events = obj.Events.Cast<HMIEvent>();
                        HMIEvent evs = events.FirstOrDefault(e => e.EventName == "OnLButtonUp" && e.Actions.Count > 0);

                        if (evs == null)
                            break;

                        HMIActions acts = evs.Actions;

                        string pdl = "";
                        foreach (HMIScriptInfo act in acts)
                        {
                            if (pdl == "")
                            {
                                var vb = act.SourceCode;

                                var callingRow = vb.Split("\r\n".ToCharArray()).FirstOrDefault(c => c.Contains("pdlName = "));
                                pdl = callingRow?.Split("\"".ToCharArray())?.ElementAtOrDefault(1);

                                var embeddedRow = vb.Split("\r\n".ToCharArray()).FirstOrDefault(c => c.Contains(".PictureName = "));
                                if (embeddedRow != null)
                                {
                                    pdl = embeddedRow?.Split("\"".ToCharArray())?.ElementAtOrDefault(1);
                                }
                                if (pdl != "")
                                {
                                    break;
                                }
                            }
                        }

                        pdl = pdl != null ? pdl.Replace(".pdl", "") : "";

                        dataList.FirstOrDefault(c => c.ObjectName == o.ObjectName).CallsPdl = pdl != null ? pdl : "";

                        Console.WriteLine(obj.ObjectName.value + " calls " + dataList.Last().CallsPdl, "\\PdlCalls.log", false);
                    }
                }
            }

            dataList = dataList.Where(c => c.CallsPdl != "" && c.CallsPdl != null && !navigatedPages.Contains(c.CallsPdl)).ToList();
            return dataList;
        }

        public void GetScreenInfo(out List<IHMIScreen> relevantScreens, out List<WindowData> windowData)
        {
            relevantScreens = new List<IHMIScreen>();

            GetRuntimeScreens(out IHMIScreens screens, out IHMIScreen screen);

            List<IHMIScreen> screensList = screens.Cast<IHMIScreen>().ToList();

            IEnumerable<IHMIScreenItem> parents = screensList.Where(c => c.Parent != null).Select(c => c.Parent);
            IEnumerable<IHMIScreenItem> menus = screensList.Where(c => c.Parent != null && c.Parent.ObjectName.Contains("menu")).Select(c => c.Parent);
            IEnumerable<string> accessPaths = screensList.Where(c => c.Parent != null).Select(c => c.AccessPath);

            string mainscreen = GetMainScreen();

            //relevantScreens = screensList.Where(c => c.Parent != null && (c.Parent.ObjectName.StartsWith("@DataClass") || c.Parent.ObjectName == "@ProcAreaLight"));
            foreach (IHMIScreen s in screensList)
            {
                if (s.Parent != null)
                {
                    if (s.Parent.ObjectName == "@DataClass1" || s.Parent.ObjectName == "@DataClass2")
                    {
                        if (s.Parent.Parent.ObjectName == "@frame_s_content")
                        {
                            relevantScreens.Add(s);
                        }
                    }
                    else if (s.Parent.ObjectName == "@DataClass3" && s.Parent.Parent.ObjectName == mainscreen)
                    {
                        relevantScreens.Add(s);
                    }

                    if (s.Parent.ObjectName == "@ProcAreaLight")
                    {
                        relevantScreens.Add(s);
                    }
                }
            }

            var result = from c in relevantScreens
                         join c2 in relevantScreens on c.Parent.ObjectName equals c2.Parent.ObjectName
                         where c != c2
                         select c;
            if (result.Count() > 0)
            {
                Console.WriteLine("Found some duplicates at c.parent.objectname boss");
            }

            var screenNames = relevantScreens.Select(c => c.ObjectName).ToList();

            var screenWindowsQuery = relevantScreens.Select(c => c.Parent).ToList();
            windowData = GetWindowData(screenWindowsQuery);
        }

        public static List<WindowData> GetWindowData(List<IHMIScreenItem> screenWindowsQuery)
        {
            List<WindowData> windowData = new List<WindowData>();
            foreach (IHMIScreenItem s in screenWindowsQuery)
            {
                int recLeft = 0;
                int recTop = 0;
                GetPositionRec(ref recLeft, ref recTop, s);

                windowData.Add(new WindowData()
                {
                    Width = s.Width,
                    Top = s.Top,
                    Height = s.Height,
                    Left = s.Left,
                    Name = s.ObjectName,
                    Visible = s.Visible,
                    Type = s.Type,
                    Layer = s.Layer,
                    Parent = s.Parent.ObjectName,
                    ScreenLeft = recLeft,
                    ScreenTop = recTop,
                    Enabled = s.Enabled
                });
            }

            return windowData;
        }

        public static void GetPositionRec(ref int recLeft, ref int recTop, IHMIScreenItem screenWindow)
        {
            if (screenWindow != null)
            {
                recLeft += screenWindow.Left;
                recTop += screenWindow.Top;
                if (screenWindow.Parent.Parent != null)
                {
                    GetPositionRec(ref recLeft, ref recTop, screenWindow.Parent.Parent);
                }
            }
        }

    }
}
