using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;

namespace NCM_Downloader.Services
{
    public class ProcessManager
    {
        private static readonly List<string> _knownWinCCProcesses = new List<string>
        {
            "apdiag", "asosheartbeatx", "AutoStartRT", "BildAnw", 
            "CAArchiveManagerBackupCopyX", "CADeleteWinCCMessageQueueing",
            "CAReportMergerX", "CC_DPSRV", "CC_MShh", "ccaeprovider",
            "CCAgent.exe", "CCAlgIAlarmDataCollector", "CCAlgIAlarmDataProxyServer",
            "CcAlgRtServer", "CCArchiveConnMon", "CCArchiveManager",
            "CCAuthorInformation", "CCAuditProviderSrv.exe", "CCAuditTrailServer.exe",
            "CCAuditViewer", "CCAuditDCV", "CCClientAPIs", "CCCloudConnect",
            "CCConfigServerMemory", "CCConfigStudio", "CCConfigStudio_x64",
            "CCCSigRTServer", "CCDBUtils.exe", "CCDeltaLoader",
            "CCDiagnosisView", "CCDiagRTServer", "CCDMClientHelper",
            "CCDmIVarProxyServer", "CCDmRtChannelHost", "CCDmRuntimePersistence",
            "CCEClient_x64", "CCEClient", "CCEmergencyWatchRTServer",
            "CCEServer_x64", "CCEServer", "CCLBMRTServer", "CCLicenseService",
            "CCMcpAutServer", "CCMigrator", "CCNSInfo2Provider", "CCOnlCmp",
            "CCOPCConfigPerm", "CCPackageMgr", "CCPerfMon", "CcPJob",
            "CCProfileServer", "CCProgressDlg", "CCProjectDuplicator",
            "CCProjectMgr", "CCPtmRTServer", "CCRedCodi", "CCRedundancyAgent",
            "CCRemoteService.exe", "CCRtsLoader", "CCRtsLoader_x64",
            "CCRunDTSPackage", "CCRunRedCodiCS", "CCScriptEditor",
            "CCSplash", "CCSsmRTServer", "CCSyncAgent", "CCSystemDiagnosticsHost",
            "CCSysTray", "CCSysTrayProxy", "CCTextDistributor",
            "CCTextServer", "CCTlgArchiveClient", "CCTlgServer",
            "CCTMConfiguration", "CCTMTimeSync", "CCTMTimeSyncServer",
            "CCTMTimeSyncServerV5", "CCTMTimeSyncV5", "CCTTRTServer",
            "CCTxtProxy", "CCUAEditor", "CCUAIUABasicProxyServer",
            "CCUAIUABasicStubServer", "CCUASetupDCOM", "CCUCSurrogate",
            "CCUsrAcv", "CCWinCCMTEditor", "CCWinCCStart", "CCWriteArchiveServer",
            "CCXREF.Presentation.Editor", "ChannelWrapperCS", "DdeServ",
            "DiagServ", "Grafexe", "gsccs", "gscrt", "HMRT", "HornCS",
            "imserverx", "LBMS", "MessageCorrectorx", "MSCS", "OPCTags",
            "OpcUaServerWinCC.exe", "osltmhandlerx", "osstatemachinex",
            "PassCS", "PassDBRT", "PassReg", "PDLCSApiTest", "PdlRT",
            "PdlServ", "Print1", "PrintIt", "Projecteditor", "ProtCS",
            "PrtScr", "PTMCS", "RebootX", "Rd1CS", "RedundancyControl",
            "RedundancyState", "RptRT", "s7asysvx", "s7aversx",
            "s7dosTraceControlPanelx", "s7oiehsx", "s7otbxsx", "s7sninsx",
            "S7TraceServiceX.exe", "s7tgtopx", "s7ubTstx", "sfcrt",
            "SAMExportToolx", "script", "SCSDialogX", "SCSDistServiceX.exe",
            "SCSFsX", "SCSMonitor", "SCSMX.exe", "SDiagRT", "setupdcom",
            "Siemens.Automation.FrameApplicationLoader",
            "Siemens.Automation.ObjectFrame.FileStorage.Server",
            "Siemens.Automation.Portal", "sn_cp1612", "ssc_visual",
            "SSMSettings", "TagAnw", "textbib", "TextLibrary", "TlgCS",
            "TouchInputPC", "TrendOnl", "trv", "TTEdit", "WebNavigatorRT",
            "WinCCChnDiag", "WinCCExplorer", "WinCCWebLicense", "LBMCS",
            "CCOnScreenKeyboard", "CCCAPHServer", "CalendarAccessProvider.exe",
            "WinCC_Calendar_Server", "WinCC_Calendar_Viewer",
            "PlantIntelligenceService.exe", "PMNSInfo2Provider", "CCSESRTSrv",
            "GfxRTS", "CCKeyboardHook", "CCProjectMgr.exe"
        };
        
        public IEnumerable<string> KnownWinCCProcesses => _knownWinCCProcesses;

        public void CloseConflictingProcesses(string exceptName = null)
        {
            var running = Process.GetProcesses();
            foreach (var p in running)
            {
                if (p.ProcessName.Contains("NCM_Downloader") 
                    && (string.IsNullOrEmpty(exceptName) || p.ProcessName != exceptName))
                {
                    try
                    {
                        if (p.Id != Process.GetCurrentProcess().Id)
                        {
                            p.Kill();
                        }
                    }
                    catch { /* Best effort */ }
                }
            }
        }
    }
}
