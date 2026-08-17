using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;

namespace AudioBt
{
    [ComImport]
    [Guid("BCDE0395-E52F-467C-8E3D-C4579291692E")]
    class MMDeviceEnumeratorCom { }

    [ComImport]
    [Guid("870AF99C-171D-4F9E-AF0D-E63DF40C2BC9")]
    class PolicyConfigClient { }

    [ComImport]
    [Guid("294935CE-F637-4E7C-A41B-AB255460B862")]
    class PolicyConfigVistaClient { }

    [UnmanagedFunctionPointer(CallingConvention.StdCall)]
    delegate int EnumAudioEndpointsProc(IntPtr thisPtr, int dataFlow, int dwStateMask, out IntPtr ppDevices);

    [UnmanagedFunctionPointer(CallingConvention.StdCall)]
    delegate int GetDefaultAudioEndpointProc(IntPtr thisPtr, int dataFlow, int role, out IntPtr ppEndpoint);

    [UnmanagedFunctionPointer(CallingConvention.StdCall)]
    delegate int CollectionGetCountProc(IntPtr thisPtr, out uint pcDevices);

    [UnmanagedFunctionPointer(CallingConvention.StdCall)]
    delegate int CollectionItemProc(IntPtr thisPtr, uint nDevice, out IntPtr ppDevice);

    [UnmanagedFunctionPointer(CallingConvention.StdCall)]
    delegate int DeviceOpenPropertyStoreProc(IntPtr thisPtr, int stgmAccess, out IntPtr ppProperties);

    [UnmanagedFunctionPointer(CallingConvention.StdCall)]
    delegate int DeviceGetIdProc(IntPtr thisPtr, out IntPtr ppstrId);

    [UnmanagedFunctionPointer(CallingConvention.StdCall)]
    delegate int DeviceGetStateProc(IntPtr thisPtr, out int pdwState);

    [UnmanagedFunctionPointer(CallingConvention.StdCall)]
    delegate int PropStoreGetValueProc(IntPtr thisPtr, ref PROPERTYKEY key, out PROPVARIANT pv);

    [UnmanagedFunctionPointer(CallingConvention.StdCall)]
    delegate int PolicySetDefaultProc(IntPtr thisPtr, [MarshalAs(UnmanagedType.LPWStr)] string pszDeviceName, uint role);

    [UnmanagedFunctionPointer(CallingConvention.StdCall)]
    delegate int PolicySetVisibleProc(IntPtr thisPtr, [MarshalAs(UnmanagedType.LPWStr)] string pszDeviceName, int bVisible);

    [StructLayout(LayoutKind.Sequential)]
    struct PROPERTYKEY
    {
        public Guid fmtid;
        public uint pid;
    }

    [StructLayout(LayoutKind.Explicit)]
    struct PROPVARIANT
    {
        [FieldOffset(0)] public ushort vt;
        [FieldOffset(8)] public IntPtr pointerValue;
    }

    [StructLayout(LayoutKind.Sequential)]
    struct SYSTEMTIME
    {
        public ushort wYear, wMonth, wDayOfWeek, wDay, wHour, wMinute, wSecond, wMilliseconds;
    }

    [StructLayout(LayoutKind.Sequential)]
    struct BLUETOOTH_FIND_RADIO_PARAMS
    {
        public uint dwSize;
    }

    [StructLayout(LayoutKind.Sequential)]
    struct BLUETOOTH_DEVICE_SEARCH_PARAMS
    {
        public uint dwSize;
        [MarshalAs(UnmanagedType.Bool)] public bool fReturnAuthenticated;
        [MarshalAs(UnmanagedType.Bool)] public bool fReturnRemembered;
        [MarshalAs(UnmanagedType.Bool)] public bool fReturnUnknown;
        [MarshalAs(UnmanagedType.Bool)] public bool fReturnConnected;
        [MarshalAs(UnmanagedType.Bool)] public bool fIssueInquiry;
        public byte cTimeoutMultiplier;
        public IntPtr hRadio;
    }

    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    struct BLUETOOTH_DEVICE_INFO
    {
        public uint dwSize;
        public ulong Address;
        public uint ulClassOfDevice;
        [MarshalAs(UnmanagedType.Bool)] public bool fConnected;
        [MarshalAs(UnmanagedType.Bool)] public bool fRemembered;
        [MarshalAs(UnmanagedType.Bool)] public bool fAuthenticated;
        public SYSTEMTIME stLastSeen;
        public SYSTEMTIME stLastUsed;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 248)]
        public string szName;
    }

    class EndpointInfo
    {
        public string Id;
        public string Name;
        public int Flow; // 0 render, 1 capture
        public int State;
        public bool IsDefault;
        public bool IsBluetooth;
        public string BtAddrHex;
    }

    class BtDeviceInfo
    {
        public ulong Address;
        public string AddrHex;
        public string Name;
        public bool Connected;
        public uint ClassOfDevice;
    }

    public static class Helper
    {
        const int eRender = 0;
        const int eCapture = 1;
        const int DEVICE_STATE_ACTIVE = 0x1;
        const int DEVICE_STATE_DISABLED = 0x2;
        const int DEVICE_STATE_NOTPRESENT = 0x4;
        const int DEVICE_STATE_UNPLUGGED = 0x8;
        const int DEVICE_STATEMASK_ALL = 0xF;
        const int STGM_READ = 0;
        const uint eConsole = 0;
        const uint eMultimedia = 1;
        const uint eCommunications = 2;
        const uint BLUETOOTH_SERVICE_DISABLE = 0;
        const uint BLUETOOTH_SERVICE_ENABLE = 1;

        static readonly Guid PKEY_Device_FriendlyNameFmtid = new Guid("a45c254e-df1c-4efd-8020-67d146a850e0");
        static readonly Guid PKEY_DeviceInterface_FriendlyNameFmtid = new Guid("026e516e-b814-414b-83cd-856d6fef4822");
        static readonly Guid GuidA2dpSink = new Guid("0000110B-0000-1000-8000-00805F9B34FB");
        static readonly Guid GuidHandsfree = new Guid("0000111E-0000-1000-8000-00805F9B34FB");
        static readonly Guid GuidHeadset = new Guid("00001108-0000-1000-8000-00805F9B34FB");
        static string logPath = "";
        static readonly object logLock = new object();
        static string lastDroppedBt = "";

        public static void SetLogPath(string path)
        {
            logPath = path ?? "";
        }

        public static void LogLine(string msg)
        {
            Log(msg);
        }

        public static void LogSnapshot(string label)
        {
            Log("--- snapshot " + label + " ---");
            try
            {
                List<BtDeviceInfo> bts = CollectBluetoothAudio();
                Log("classic BT audio count=" + bts.Count);
                foreach (BtDeviceInfo bt in bts)
                    Log("  bt name=" + bt.Name + " addr=" + bt.AddrHex + " classicConnected=" + bt.Connected + " class=0x" + bt.ClassOfDevice.ToString("X"));
                List<EndpointInfo> endpoints = CollectEndpoints();
                Log("endpoints count=" + endpoints.Count);
                foreach (EndpointInfo ep in endpoints)
                {
                    Log("  ep flow=" + (ep.Flow == eRender ? "Out" : "In")
                        + " state=0x" + ep.State.ToString("X")
                        + " active=" + IsActiveEndpoint(ep)
                        + " def=" + ep.IsDefault
                        + " bt=" + ep.IsBluetooth
                        + " addr=" + ep.BtAddrHex
                        + " name=" + ep.Name
                        + " id=" + ep.Id);
                }
            }
            catch (Exception ex)
            {
                Log("snapshot error: " + ex.Message);
            }
        }

        static string DroppedSuffix()
        {
            if (string.IsNullOrEmpty(lastDroppedBt))
                return "";
            return " (dropped " + lastDroppedBt + ")";
        }

        static void Log(string msg)
        {
            if (string.IsNullOrEmpty(logPath))
                return;
            try
            {
                string line = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff") + "  " + (msg ?? "") + Environment.NewLine;
                string dir = Path.GetDirectoryName(logPath);
                lock (logLock)
                {
                    if (!string.IsNullOrEmpty(dir) && !Directory.Exists(dir))
                        Directory.CreateDirectory(dir);
                    File.AppendAllText(logPath, line, new UTF8Encoding(false));
                }
            }
            catch
            {
            }
        }

        [DllImport("ole32.dll")]
        static extern int PropVariantClear(ref PROPVARIANT pvar);

        [DllImport("ole32.dll")]
        static extern void CoTaskMemFree(IntPtr pv);

        [DllImport("ole32.dll")]
        static extern int CoCreateInstance(ref Guid rclsid, IntPtr pUnkOuter, uint dwClsContext, ref Guid riid, out IntPtr ppv);

        static IntPtr VtblSlot(IntPtr com, int slot)
        {
            IntPtr vtbl = Marshal.ReadIntPtr(com);
            return Marshal.ReadIntPtr(vtbl, slot * IntPtr.Size);
        }

        static T ComFn<T>(IntPtr com, int slot) where T : class
        {
            return (T)(object)Marshal.GetDelegateForFunctionPointer(VtblSlot(com, slot), typeof(T));
        }

        static IntPtr CreateCom(Type comClass, Guid iid)
        {
            object obj = Activator.CreateInstance(comClass);
            IntPtr unk = Marshal.GetIUnknownForObject(obj);
            IntPtr p;
            int hr = Marshal.QueryInterface(unk, ref iid, out p);
            Marshal.Release(unk);
            if (hr != 0 || p == IntPtr.Zero)
                throw new COMException("QueryInterface " + iid, hr);
            return p;
        }

        [DllImport("BluetoothAPIs.dll", SetLastError = true)]
        static extern IntPtr BluetoothFindFirstRadio(ref BLUETOOTH_FIND_RADIO_PARAMS pbtfrp, out IntPtr phRadio);

        [DllImport("BluetoothAPIs.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        static extern bool BluetoothFindNextRadio(IntPtr hFind, out IntPtr phRadio);

        [DllImport("BluetoothAPIs.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        static extern bool BluetoothFindRadioClose(IntPtr hFind);

        [DllImport("BluetoothAPIs.dll", SetLastError = true)]
        static extern IntPtr BluetoothFindFirstDevice(ref BLUETOOTH_DEVICE_SEARCH_PARAMS pbtsp, ref BLUETOOTH_DEVICE_INFO pbtdi);

        [DllImport("BluetoothAPIs.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        static extern bool BluetoothFindNextDevice(IntPtr hFind, ref BLUETOOTH_DEVICE_INFO pbtdi);

        [DllImport("BluetoothAPIs.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        static extern bool BluetoothFindDeviceClose(IntPtr hFind);

        [DllImport("BluetoothAPIs.dll", SetLastError = true)]
        static extern uint BluetoothSetServiceState(IntPtr hRadio, ref BLUETOOTH_DEVICE_INFO pbtdi, ref Guid guidService, uint dwServiceFlags);

        [DllImport("BluetoothAPIs.dll", SetLastError = true)]
        static extern uint BluetoothGetDeviceInfo(IntPtr hRadio, ref BLUETOOTH_DEVICE_INFO pbtdi);

        [DllImport("kernel32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        static extern bool CloseHandle(IntPtr hObject);

        [DllImport("kernel32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        static extern bool DeviceIoControl(IntPtr hDevice, uint dwIoControlCode, ref ulong lpInBuffer, uint nInBufferSize,
            IntPtr lpOutBuffer, uint nOutBufferSize, out uint lpBytesReturned, IntPtr lpOverlapped);

        [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
        static extern IntPtr CreateFile(string lpFileName, uint dwDesiredAccess, uint dwShareMode, IntPtr lpSecurityAttributes,
            uint dwCreationDisposition, uint dwFlagsAndAttributes, IntPtr hTemplateFile);

        [DllImport("setupapi.dll", CharSet = CharSet.Unicode, SetLastError = true, EntryPoint = "SetupDiGetClassDevsW")]
        static extern IntPtr SetupDiGetClassDevsGuid(ref Guid ClassGuid, IntPtr Enumerator, IntPtr hwndParent, uint Flags);

        [DllImport("setupapi.dll", CharSet = CharSet.Unicode, SetLastError = true, EntryPoint = "SetupDiGetClassDevsW")]
        static extern IntPtr SetupDiGetClassDevsEnum(IntPtr ClassGuid, string Enumerator, IntPtr hwndParent, uint Flags);

        [DllImport("setupapi.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        static extern bool SetupDiEnumDeviceInterfaces(IntPtr DeviceInfoSet, IntPtr DeviceInfoData, ref Guid InterfaceClassGuid,
            uint MemberIndex, ref SP_DEVICE_INTERFACE_DATA DeviceInterfaceData);

        [DllImport("setupapi.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        static extern bool SetupDiGetDeviceInterfaceDetail(IntPtr DeviceInfoSet, ref SP_DEVICE_INTERFACE_DATA DeviceInterfaceData,
            IntPtr DeviceInterfaceDetailData, uint DeviceInterfaceDetailDataSize, out uint RequiredSize, IntPtr DeviceInfoData);

        [DllImport("setupapi.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        static extern bool SetupDiDestroyDeviceInfoList(IntPtr DeviceInfoSet);

        [DllImport("setupapi.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        static extern bool SetupDiEnumDeviceInfo(IntPtr DeviceInfoSet, uint MemberIndex, ref SP_DEVINFO_DATA DeviceInfoData);

        [DllImport("setupapi.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        static extern bool SetupDiGetDeviceInstanceId(IntPtr DeviceInfoSet, ref SP_DEVINFO_DATA DeviceInfoData,
            StringBuilder DeviceInstanceId, int DeviceInstanceIdSize, out int RequiredSize);

        [DllImport("cfgmgr32.dll", CharSet = CharSet.Unicode, EntryPoint = "CM_Locate_DevNodeW")]
        static extern int CM_Locate_DevNode(out uint pdnDevInst, string pDeviceID, uint ulFlags);

        [DllImport("cfgmgr32.dll")]
        static extern int CM_Enable_DevNode(uint dnDevInst, uint ulFlags);

        [DllImport("cfgmgr32.dll")]
        static extern int CM_Disable_DevNode(uint dnDevInst, uint ulFlags);

        const uint IOCTL_BTH_DISCONNECT_DEVICE = 0x0041000C;
        const uint GENERIC_READ = 0x80000000;
        const uint GENERIC_WRITE = 0x40000000;
        const uint FILE_SHARE_READ = 0x1;
        const uint FILE_SHARE_WRITE = 0x2;
        const uint OPEN_EXISTING = 3;
        const uint DIGCF_PRESENT = 0x2;
        const uint DIGCF_ALLCLASSES = 0x4;
        const uint DIGCF_DEVICEINTERFACE = 0x10;
        static readonly Guid GuidBthPortDeviceInterface = new Guid("0850302A-B344-4FDA-9BE9-90576B8D46F0");

        [StructLayout(LayoutKind.Sequential)]
        struct SP_DEVICE_INTERFACE_DATA
        {
            public uint cbSize;
            public Guid InterfaceClassGuid;
            public uint Flags;
            public IntPtr Reserved;
        }

        [StructLayout(LayoutKind.Sequential)]
        struct SP_DEVINFO_DATA
        {
            public uint cbSize;
            public Guid ClassGuid;
            public uint DevInst;
            public IntPtr Reserved;
        }

        public static string ListTsv()
        {
            try
            {
                List<EndpointInfo> endpoints = CollectEndpoints();
                List<BtDeviceInfo> btDevs = CollectBluetoothAudio();
                int[] activeByFlow = CountActiveByFlow(endpoints);
                StringBuilder sb = new StringBuilder();
                sb.AppendLine("id\tkind\tname\tstate\tisDefault\tcanConnect\tiso");

                Dictionary<string, bool> seenBt = new Dictionary<string, bool>(StringComparer.OrdinalIgnoreCase);
                foreach (BtDeviceInfo bt in btDevs)
                {
                    if (seenBt.ContainsKey(bt.AddrHex))
                        continue;
                    seenBt[bt.AddrHex] = true;
                    string state = (bt.Connected || HasActiveEndpoint(endpoints, bt, -1)) ? "Connected" : "Disconnected";
                    bool isDef = false;
                    foreach (EndpointInfo ep in endpoints)
                    {
                        if (ep.IsDefault && EndpointMatchesBt(ep, bt))
                        {
                            isDef = true;
                            break;
                        }
                    }
                    bool isoOut = BtFlowIsolated(bt, endpoints, activeByFlow, eRender);
                    bool isoIn = BtFlowIsolated(bt, endpoints, activeByFlow, eCapture);
                    string iso = IsolatedFlowsCode(isoOut, isoIn);
                    if (isoOut && isoIn)
                        state = "Connected · Isolated";
                    else if (isoOut)
                        state = "Connected · Isolated out";
                    else if (isoIn)
                        state = "Connected · Isolated in";
                    else if (isDef)
                        state = "Connected · Default";
                    else if (bt.Connected || HasActiveEndpoint(endpoints, bt, -1))
                    {
                        foreach (EndpointInfo ep in endpoints)
                        {
                            if (EndpointMatchesBt(ep, bt) && IsActiveEndpoint(ep))
                            {
                                state = "Connected";
                                break;
                            }
                        }
                    }
                    sb.AppendLine(Tsv("bt:" + bt.AddrHex, "BT", bt.Name, state, isDef ? "1" : "0", "1", iso));
                }

                foreach (EndpointInfo ep in endpoints)
                {
                    if ((ep.State & DEVICE_STATE_NOTPRESENT) != 0 && !ep.IsBluetooth)
                        continue;
                    string kind = ep.Flow == eRender ? "Out" : "In";
                    bool isolated = EndpointIsIsolated(ep, activeByFlow);
                    string state = StateLabel(ep, isolated);
                    string iso = isolated ? kind : "";
                    sb.AppendLine(Tsv(ep.Id, kind, ep.Name, state, ep.IsDefault ? "1" : "0", ep.IsBluetooth ? "1" : "0", iso));
                }
                return sb.ToString();
            }
            catch (Exception ex)
            {
                return "ERR\t" + Cell(ex.Message);
            }
        }

        public static string SetDefault(string id)
        {
            try
            {
                EndpointInfo ep = ResolveEndpoint(id, true);
                if (ep == null)
                    return "ERR\tNo audio endpoint found for that device";
                string err;
                if (!TrySetDefaultTarget(ref ep, ResolveBt(id), out err))
                    return "ERR\t" + Cell(err);
                return "OK\tDefault set: " + Cell(ep.Name);
            }
            catch (Exception ex)
            {
                return "ERR\t" + Cell(ex.Message);
            }
        }

        public static string SetVisible(string id, bool visible)
        {
            try
            {
                EndpointInfo ep = ResolveEndpoint(id, false);
                if (ep == null)
                    return "ERR\tNo audio endpoint found for that device";
                int hr = PolicySetVisible(ep.Id, visible);
                if (hr != 0)
                    return "ERR\t" + (visible ? "Enable" : "Disable") + " failed " + FormatHr(hr);
                return "OK\t" + (visible ? "Enabled: " : "Disabled: ") + Cell(ep.Name);
            }
            catch (Exception ex)
            {
                return "ERR\t" + Cell(ex.Message);
            }
        }

        public static string Isolate(string id)
        {
            try
            {
                Log("Isolate id=" + id);
                BtDeviceInfo bt = ResolveBt(id);
                Log("Isolate resolveBt name=" + (bt == null ? "<null>" : bt.Name) + " addr=" + (bt == null ? "" : bt.AddrHex) + " classicConnected=" + (bt != null && bt.Connected));
                lastDroppedBt = "";
                bool didConnect = false;
                string connectMethod = "";
                int requestedFlow = -1;
                if (id != null && !id.StartsWith("bt:", StringComparison.OrdinalIgnoreCase))
                {
                    EndpointInfo selectedBefore = FindEndpointById(CollectEndpoints(), id);
                    if (selectedBefore != null)
                        requestedFlow = selectedBefore.Flow;
                }
                if (id != null && id.StartsWith("bt:", StringComparison.OrdinalIgnoreCase) && bt == null)
                    return "ERR\tBluetooth device not found";
                if (bt != null)
                {
                    string dropped;
                    DisconnectOtherBtAudio(bt, out dropped);
                }
                if (bt != null && !bt.Connected)
                {
                    string connectErr;
                    Log("Isolate needs connect");
                    if (!ConnectBtWithFallback(bt, out connectMethod, out connectErr))
                    {
                        Log("Isolate connect failed: " + connectErr);
                        return "ERR\t" + Cell(connectErr);
                    }
                    didConnect = true;
                    Log("Isolate connected via " + connectMethod);
                }

                List<EndpointInfo> endpoints = CollectEndpoints();
                List<EndpointInfo> targets = CollectIsolateTargets(id, bt, endpoints, requestedFlow);
                if (targets.Count == 0)
                {
                    System.Threading.Thread.Sleep(500);
                    endpoints = CollectEndpoints();
                    targets = CollectIsolateTargets(id, bt, endpoints, requestedFlow);
                }
                if (targets.Count == 0)
                    return "ERR\tNo audio endpoint found to isolate";

                bool fromOutputEndpoint = requestedFlow == eRender;

                HashSet<int> isolatedFlows = new HashSet<int>();
                List<string> failed = new List<string>();
                StringBuilder names = new StringBuilder();
                for (int i = 0; i < targets.Count; i++)
                {
                    EndpointInfo t = targets[i];
                    if (fromOutputEndpoint && t.Flow != eRender)
                        continue;
                    string setErr;
                    if (!TrySetDefaultTarget(ref t, bt, out setErr))
                    {
                        failed.Add(setErr);
                        continue;
                    }
                    targets[i] = t;
                    isolatedFlows.Add(t.Flow);
                    if (names.Length > 0)
                        names.Append(", ");
                    names.Append(t.Name);
                }
                if (isolatedFlows.Count == 0)
                {
                    string msg = failed.Count == 0 ? "SetDefault failed" : string.Join("; ", failed);
                    return "ERR\t" + Cell(msg);
                }

                endpoints = CollectEndpoints();
                foreach (EndpointInfo ep in endpoints)
                {
                    if (!isolatedFlows.Contains(ep.Flow))
                        continue;
                    if ((ep.State & DEVICE_STATE_NOTPRESENT) != 0)
                        continue;
                    bool isTarget = false;
                    foreach (EndpointInfo t in targets)
                    {
                        if (!isolatedFlows.Contains(t.Flow))
                            continue;
                        if (string.Equals(t.Id, ep.Id, StringComparison.OrdinalIgnoreCase)
                            || (bt != null && EndpointMatchesBt(ep, bt) && ep.Flow == t.Flow))
                        {
                            isTarget = true;
                            break;
                        }
                    }
                    if (isTarget)
                        continue;
                    if ((ep.State & DEVICE_STATE_ACTIVE) != 0)
                        PolicySetVisible(ep.Id, false);
                }
                string prefix = didConnect
                    ? ("Connected via " + connectMethod + DroppedSuffix() + " and isolated: ")
                    : ("Isolated" + DroppedSuffix() + ": ");
                return "OK\t" + prefix + Cell(names.ToString());
            }
            catch (Exception ex)
            {
                return "ERR\t" + Cell(ex.Message);
            }
        }

        public static string Connect(string id)
        {
            try
            {
                Log("Connect id=" + id);
                BtDeviceInfo bt = ResolveBt(id);
                if (bt == null)
                {
                    Log("Connect resolveBt=null");
                    return "ERR\tNot a paired Bluetooth audio device";
                }
                Log("Connect resolveBt name=" + bt.Name + " addr=" + bt.AddrHex + " classicConnected=" + bt.Connected);
                lastDroppedBt = "";
                string method;
                string err;
                if (!ConnectBtWithFallback(bt, out method, out err))
                {
                    Log("Connect failed: " + err);
                    return "ERR\t" + Cell(err);
                }
                Log("Connect ok via " + method);
                return "OK\tConnected via " + method + DroppedSuffix() + ": " + Cell(bt.Name);
            }
            catch (Exception ex)
            {
                return "ERR\t" + Cell(ex.Message);
            }
        }

        public static string ConfirmConnected(string id, string method)
        {
            try
            {
                BtDeviceInfo bt = ResolveBt(id);
                if (bt == null)
                    return "ERR\tNot a paired Bluetooth audio device";
                string label = string.IsNullOrEmpty(method) ? "fallback" : method;
                Log("ConfirmConnected method=" + label + " name=" + bt.Name);
                if (!WaitForBtAudio(bt, 10000))
                {
                    Log("ConfirmConnected wait failed");
                    return "ERR\tConnected via " + label + " but no audio endpoint appeared: " + Cell(bt.Name);
                }
                Log("ConfirmConnected wait ok");
                List<EndpointInfo> endpoints = CollectEndpoints();
                foreach (EndpointInfo ep in endpoints)
                {
                    if (EndpointMatchesBt(ep, bt))
                        EnsureVisible(ep.Id, true);
                }
                return "OK\tConnected via " + label + DroppedSuffix() + ": " + Cell(bt.Name);
            }
            catch (Exception ex)
            {
                return "ERR\t" + Cell(ex.Message);
            }
        }

        public static string BluetoothAddressHex(string id)
        {
            try
            {
                BtDeviceInfo bt = ResolveBt(id);
                return bt == null ? "" : bt.AddrHex;
            }
            catch
            {
                return "";
            }
        }

        public static string Disconnect(string id)
        {
            try
            {
                BtDeviceInfo bt = ResolveBt(id);
                if (bt == null)
                    return "ERR\tNot a paired Bluetooth audio device";
                string ioctlErr;
                bool ioctlOk = DisconnectRadio(bt, out ioctlErr);
                string svcErr;
                bool svcOk = SetBtServices(bt, false, out svcErr);
                if (!ioctlOk && !svcOk)
                {
                    string msg = ioctlErr;
                    if (!string.IsNullOrEmpty(svcErr))
                        msg = string.IsNullOrEmpty(msg) ? svcErr : msg + "; " + svcErr;
                    if (string.IsNullOrEmpty(msg))
                        msg = "Windows refused the Bluetooth disconnect";
                    return "ERR\t" + Cell(msg);
                }
                System.Threading.Thread.Sleep(500);
                return "OK\tDisconnected: " + Cell(bt.Name);
            }
            catch (Exception ex)
            {
                return "ERR\t" + Cell(ex.Message);
            }
        }

        static bool IsCountableActive(EndpointInfo ep)
        {
            if (ep == null)
                return false;
            if ((ep.State & DEVICE_STATE_DISABLED) != 0)
                return false;
            if ((ep.State & DEVICE_STATE_NOTPRESENT) != 0)
                return false;
            return (ep.State & DEVICE_STATE_ACTIVE) != 0;
        }

        static int[] CountActiveByFlow(List<EndpointInfo> endpoints)
        {
            int[] counts = new int[2];
            foreach (EndpointInfo ep in endpoints)
            {
                if (IsCountableActive(ep) && (ep.Flow == eRender || ep.Flow == eCapture))
                    counts[ep.Flow]++;
            }
            return counts;
        }

        static bool EndpointIsIsolated(EndpointInfo ep, int[] activeByFlow)
        {
            if (ep == null || !ep.IsDefault || activeByFlow == null)
                return false;
            if ((ep.State & DEVICE_STATE_DISABLED) != 0)
                return false;
            if ((ep.State & DEVICE_STATE_NOTPRESENT) != 0)
                return false;
            if (ep.Flow != eRender && ep.Flow != eCapture)
                return false;
            return activeByFlow[ep.Flow] == 1;
        }

        static bool BtFlowIsolated(BtDeviceInfo bt, List<EndpointInfo> endpoints, int[] activeByFlow, int flow)
        {
            if (bt == null)
                return false;
            bool any = false;
            foreach (EndpointInfo ep in endpoints)
            {
                if (!EndpointMatchesBt(ep, bt) || ep.Flow != flow)
                    continue;
                if ((ep.State & DEVICE_STATE_NOTPRESENT) != 0)
                    continue;
                if ((ep.State & DEVICE_STATE_DISABLED) != 0)
                    continue;
                any = true;
                if (!EndpointIsIsolated(ep, activeByFlow))
                    return false;
            }
            return any;
        }

        static string IsolatedFlowsCode(bool isoOut, bool isoIn)
        {
            if (isoOut && isoIn)
                return "InOut";
            if (isoOut)
                return "Out";
            if (isoIn)
                return "In";
            return "";
        }

        static string StateLabel(EndpointInfo ep, bool isolated)
        {
            if ((ep.State & DEVICE_STATE_DISABLED) != 0)
                return "Disabled";
            if (isolated)
                return "Isolated";
            if (ep.IsDefault)
                return "Default";
            if ((ep.State & DEVICE_STATE_UNPLUGGED) != 0)
                return ep.IsBluetooth ? "Disconnected" : "Unplugged";
            if ((ep.State & DEVICE_STATE_NOTPRESENT) != 0)
                return "Not present";
            return "Enabled";
        }

        static string Tsv(string id, string kind, string name, string state, string isDefault, string canConnect, string iso)
        {
            return Cell(id) + "\t" + Cell(kind) + "\t" + Cell(name) + "\t" + Cell(state) + "\t" + isDefault + "\t" + canConnect + "\t" + Cell(iso);
        }

        static string Cell(string s)
        {
            if (string.IsNullOrEmpty(s))
                return "";
            return s.Replace("\t", " ").Replace("\r", " ").Replace("\n", " ");
        }

        static IntPtr CreateEnumerator()
        {
            return CreateCom(typeof(MMDeviceEnumeratorCom), new Guid("A95664D2-9614-4F35-A746-DE8DB63617E6"));
        }

        static List<EndpointInfo> CollectEndpoints()
        {
            List<EndpointInfo> list = new List<EndpointInfo>();
            IntPtr en = CreateEnumerator();
            try
            {
                string defRender = GetDefaultId(en, eRender);
                string defCapture = GetDefaultId(en, eCapture);
                CollectFlow(en, eRender, defRender, list);
                CollectFlow(en, eCapture, defCapture, list);
            }
            finally
            {
                Marshal.Release(en);
            }
            LinkBluetoothEndpoints(list);
            return list;
        }

        static void CollectFlow(IntPtr en, int flow, string defaultId, List<EndpointInfo> list)
        {
            IntPtr col;
            int hr = ComFn<EnumAudioEndpointsProc>(en, 3)(en, flow, DEVICE_STATEMASK_ALL, out col);
            if (hr != 0 || col == IntPtr.Zero)
                return;
            try
            {
                uint count;
                if (ComFn<CollectionGetCountProc>(col, 3)(col, out count) != 0)
                    return;
                for (uint i = 0; i < count; i++)
                {
                    IntPtr dev;
                    if (ComFn<CollectionItemProc>(col, 4)(col, i, out dev) != 0 || dev == IntPtr.Zero)
                        continue;
                    try
                    {
                        string id = DeviceId(dev);
                        if (string.IsNullOrEmpty(id))
                            continue;
                        int state;
                        if (ComFn<DeviceGetStateProc>(dev, 6)(dev, out state) != 0)
                            state = 0;
                        string name = GetFriendlyName(dev);
                        if (string.IsNullOrEmpty(name))
                            name = id;
                        EndpointInfo ep = new EndpointInfo();
                        ep.Id = id;
                        ep.Name = name;
                        ep.Flow = flow;
                        ep.State = state;
                        ep.IsDefault = !string.IsNullOrEmpty(defaultId) && string.Equals(id, defaultId, StringComparison.OrdinalIgnoreCase);
                        ep.IsBluetooth = IsBluetoothEndpoint(id);
                        ep.BtAddrHex = ExtractBtAddrFromEndpointId(id);
                        list.Add(ep);
                    }
                    finally
                    {
                        Marshal.Release(dev);
                    }
                }
            }
            finally
            {
                Marshal.Release(col);
            }
        }

        static string GetDefaultId(IntPtr en, int flow)
        {
            string id = TryDefault(en, flow, (int)eMultimedia);
            if (string.IsNullOrEmpty(id))
                id = TryDefault(en, flow, (int)eConsole);
            if (string.IsNullOrEmpty(id))
                id = TryDefault(en, flow, (int)eCommunications);
            return id;
        }

        static string TryDefault(IntPtr en, int flow, int role)
        {
            IntPtr dev;
            if (ComFn<GetDefaultAudioEndpointProc>(en, 4)(en, flow, role, out dev) != 0 || dev == IntPtr.Zero)
                return null;
            try
            {
                return DeviceId(dev);
            }
            finally
            {
                Marshal.Release(dev);
            }
        }

        static string DeviceId(IntPtr dev)
        {
            IntPtr p;
            if (ComFn<DeviceGetIdProc>(dev, 5)(dev, out p) != 0 || p == IntPtr.Zero)
                return null;
            try
            {
                return Marshal.PtrToStringUni(p);
            }
            finally
            {
                CoTaskMemFree(p);
            }
        }

        static string GetFriendlyName(IntPtr dev)
        {
            IntPtr store;
            if (ComFn<DeviceOpenPropertyStoreProc>(dev, 4)(dev, STGM_READ, out store) != 0 || store == IntPtr.Zero)
                return "";
            try
            {
                string name = ReadPropString(store, PKEY_Device_FriendlyNameFmtid, 14);
                if (string.IsNullOrEmpty(name))
                    name = ReadPropString(store, PKEY_DeviceInterface_FriendlyNameFmtid, 2);
                return name;
            }
            finally
            {
                Marshal.Release(store);
            }
        }

        static string ReadPropString(IntPtr store, Guid fmtid, uint pid)
        {
            PROPERTYKEY key = new PROPERTYKEY();
            key.fmtid = fmtid;
            key.pid = pid;
            PROPVARIANT pv;
            if (ComFn<PropStoreGetValueProc>(store, 5)(store, ref key, out pv) != 0)
                return "";
            try
            {
                if (pv.vt == 31 || pv.vt == 8)
                    return Marshal.PtrToStringUni(pv.pointerValue) ?? "";
                return "";
            }
            finally
            {
                PropVariantClear(ref pv);
            }
        }

        static void LinkBluetoothEndpoints(List<EndpointInfo> endpoints)
        {
            List<BtDeviceInfo> btDevs = CollectBluetoothAudio();
            foreach (EndpointInfo ep in endpoints)
            {
                if (ep.IsBluetooth && !string.IsNullOrEmpty(ep.BtAddrHex))
                    continue;
                foreach (BtDeviceInfo bt in btDevs)
                {
                    if (EndpointMatchesBt(ep, bt))
                    {
                        ep.IsBluetooth = true;
                        if (string.IsNullOrEmpty(ep.BtAddrHex))
                            ep.BtAddrHex = bt.AddrHex;
                        break;
                    }
                }
            }
        }

        static bool IsBluetoothEndpoint(string id)
        {
            if (string.IsNullOrEmpty(id))
                return false;
            string u = id.ToUpperInvariant();
            return u.Contains("BTHENUM") || u.Contains("BTHHFENUM") || u.Contains("BTHLE") || u.Contains("BTHAUDIO");
        }

        static string ExtractBtAddrFromEndpointId(string id)
        {
            if (string.IsNullOrEmpty(id))
                return "";
            // Common: ...&00_11_22_33_44_55 or ...#00_11_22_33_44_55
            System.Text.RegularExpressions.Match m = System.Text.RegularExpressions.Regex.Match(
                id, @"([0-9A-Fa-f]{2}[_:][0-9A-Fa-f]{2}[_:][0-9A-Fa-f]{2}[_:][0-9A-Fa-f]{2}[_:][0-9A-Fa-f]{2}[_:][0-9A-Fa-f]{2})");
            if (!m.Success)
                return "";
            return m.Groups[1].Value.Replace(":", "").Replace("_", "").ToUpperInvariant();
        }

        static EndpointInfo FindEndpointById(List<EndpointInfo> endpoints, string id)
        {
            if (string.IsNullOrEmpty(id))
                return null;
            foreach (EndpointInfo ep in endpoints)
            {
                if (string.Equals(ep.Id, id, StringComparison.OrdinalIgnoreCase))
                    return ep;
            }
            return null;
        }

        static List<EndpointInfo> CollectIsolateTargets(string id, BtDeviceInfo bt, List<EndpointInfo> endpoints, int requestedFlow)
        {
            List<EndpointInfo> targets = new List<EndpointInfo>();
            if (id != null && id.StartsWith("bt:", StringComparison.OrdinalIgnoreCase))
            {
                EndpointInfo render = BestBtEndpoint(endpoints, bt, eRender);
                EndpointInfo capture = BestBtEndpoint(endpoints, bt, eCapture);
                if (render != null)
                    targets.Add(render);
                if (capture != null)
                    targets.Add(capture);
                return targets;
            }
            int flow = requestedFlow >= 0 ? requestedFlow : eRender;
            if (bt != null)
            {
                EndpointInfo best = BestBtEndpoint(endpoints, bt, flow);
                if (best != null)
                {
                    targets.Add(best);
                    return targets;
                }
            }
            EndpointInfo one = FindEndpointById(endpoints, id);
            if (one != null && (one.State & DEVICE_STATE_NOTPRESENT) == 0)
                targets.Add(one);
            return targets;
        }

        static EndpointInfo BestBtEndpoint(List<EndpointInfo> endpoints, BtDeviceInfo bt, int flow)
        {
            if (endpoints == null || bt == null)
                return null;
            EndpointInfo best = null;
            int bestScore = -1;
            foreach (EndpointInfo ep in endpoints)
            {
                if (!EndpointMatchesBt(ep, bt) || ep.Flow != flow)
                    continue;
                if ((ep.State & DEVICE_STATE_NOTPRESENT) != 0)
                    continue;
                int score = 0;
                if ((ep.State & DEVICE_STATE_ACTIVE) != 0)
                    score += 8;
                if ((ep.State & DEVICE_STATE_DISABLED) == 0)
                    score += 4;
                if ((ep.State & DEVICE_STATE_UNPLUGGED) == 0)
                    score += 2;
                if (ep.IsDefault)
                    score += 1;
                string n = ep.Name ?? "";
                if (flow == eRender)
                {
                    if (n.IndexOf("Hands-Free", StringComparison.OrdinalIgnoreCase) >= 0
                        || n.IndexOf("Handsfree", StringComparison.OrdinalIgnoreCase) >= 0
                        || n.IndexOf("Headset", StringComparison.OrdinalIgnoreCase) >= 0)
                        score -= 3;
                    if (n.IndexOf("Stereo", StringComparison.OrdinalIgnoreCase) >= 0
                        || n.IndexOf("Headphones", StringComparison.OrdinalIgnoreCase) >= 0
                        || n.IndexOf("Speaker", StringComparison.OrdinalIgnoreCase) >= 0)
                        score += 3;
                }
                if (score > bestScore)
                {
                    bestScore = score;
                    best = ep;
                }
            }
            return best;
        }

        static bool TrySetDefaultTarget(ref EndpointInfo target, BtDeviceInfo bt, out string err)
        {
            err = "";
            if (target == null)
            {
                err = "No audio endpoint found to set default";
                return false;
            }
            int flow = target.Flow;
            string originalId = target.Id;
            DateTime deadline = DateTime.UtcNow.AddMilliseconds(12000);
            int lastHr = -1;
            int lastState = 0;
            while (true)
            {
                List<EndpointInfo> endpoints = CollectEndpoints();
                EndpointInfo ep = FindEndpointById(endpoints, originalId);
                if (bt != null)
                {
                    EndpointInfo byBt = BestBtEndpoint(endpoints, bt, flow);
                    if (byBt != null)
                        ep = byBt;
                }
                if (ep == null)
                    ep = FindEndpointById(endpoints, target.Id);
                if (ep != null && (ep.State & DEVICE_STATE_NOTPRESENT) == 0)
                {
                    lastState = ep.State;
                    if ((ep.State & DEVICE_STATE_DISABLED) != 0)
                        EnsureVisible(ep.Id, true);
                    bool ready = (ep.State & DEVICE_STATE_ACTIVE) != 0
                        && (ep.State & DEVICE_STATE_UNPLUGGED) == 0
                        && (ep.State & DEVICE_STATE_DISABLED) == 0;
                    if (ready)
                    {
                        Log("SetDefault try name=" + ep.Name + " flow=" + (ep.Flow == eRender ? "Out" : "In")
                            + " state=0x" + ep.State.ToString("X") + " id=" + ep.Id);
                        lastHr = PolicySetDefault(ep.Id, ep.Flow);
                        EndpointInfo check = FindEndpointById(CollectEndpoints(), ep.Id);
                        Log("SetDefault hr=0x" + unchecked((uint)lastHr).ToString("X8") + " nowDefault=" + (check != null && check.IsDefault));
                        if (lastHr == 0 || (check != null && check.IsDefault))
                        {
                            target = check ?? ep;
                            return true;
                        }
                    }
                    originalId = ep.Id;
                }
                if (DateTime.UtcNow >= deadline)
                {
                    string who = ep != null ? ep.Name : target.Name;
                    if ((lastState & DEVICE_STATE_UNPLUGGED) != 0 || (lastState & DEVICE_STATE_ACTIVE) == 0)
                    {
                        err = (string.IsNullOrEmpty(who) ? "Device" : who) + " is not an active output yet";
                        return false;
                    }
                    err = "SetDefault failed " + FormatHr(lastHr) + (string.IsNullOrEmpty(who) ? "" : ": " + who);
                    return false;
                }
                System.Threading.Thread.Sleep(350);
            }
        }

        static string FormatHr(int hr)
        {
            uint u = unchecked((uint)hr);
            string name = "";
            switch (u)
            {
                case 0x80070002: name = "not found"; break;
                case 0x80070005: name = "access denied"; break;
                case 0x80070015: name = "device not ready"; break;
                case 0x8007001F: name = "device not ready"; break;
                case 0x80070057: name = "invalid argument"; break;
                case 0x80070490: name = "element not found"; break;
                case 0x80004005: name = "unspecified failure"; break;
                case 0x88890004: name = "device invalidated"; break;
            }
            string s = "HRESULT 0x" + u.ToString("X8");
            if (name.Length > 0)
                s += " (" + name + ")";
            return s;
        }

        static EndpointInfo ResolveEndpoint(string id, bool preferRender)
        {
            List<EndpointInfo> endpoints = CollectEndpoints();
            EndpointInfo direct = FindEndpointById(endpoints, id);
            if (direct != null)
                return direct;
            BtDeviceInfo bt = ResolveBt(id);
            if (bt == null)
                return null;
            EndpointInfo render = null;
            EndpointInfo capture = null;
            foreach (EndpointInfo ep in endpoints)
            {
                if (!EndpointMatchesBt(ep, bt))
                    continue;
                if (ep.Flow == eRender && render == null)
                    render = ep;
                if (ep.Flow == eCapture && capture == null)
                    capture = ep;
            }
            if (preferRender)
            {
                EndpointInfo bestRender = BestBtEndpoint(endpoints, bt, eRender);
                if (bestRender != null)
                    return bestRender;
                return render ?? capture;
            }
            EndpointInfo bestCapture = BestBtEndpoint(endpoints, bt, eCapture);
            if (bestCapture != null)
                return bestCapture;
            return capture ?? render;
        }

        static BtDeviceInfo ResolveBt(string id)
        {
            if (string.IsNullOrEmpty(id))
                return null;
            if (id.StartsWith("bt:", StringComparison.OrdinalIgnoreCase))
                return FindBt(id.Substring(3));
            List<EndpointInfo> endpoints = CollectEndpoints();
            EndpointInfo ep = FindEndpointById(endpoints, id);
            if (ep == null || !ep.IsBluetooth)
                return null;
            if (!string.IsNullOrEmpty(ep.BtAddrHex))
            {
                BtDeviceInfo byAddr = FindBt(ep.BtAddrHex);
                if (byAddr != null)
                    return byAddr;
            }
            foreach (BtDeviceInfo bt in CollectBluetoothAudio())
            {
                if (EndpointMatchesBt(ep, bt))
                    return bt;
            }
            return null;
        }

        static BtDeviceInfo FindBt(string addrHex)
        {
            if (string.IsNullOrEmpty(addrHex))
                return null;
            string norm = addrHex.Replace(":", "").Replace("_", "").Replace("-", "").ToUpperInvariant();
            foreach (BtDeviceInfo bt in CollectBluetoothAudio())
            {
                if (bt.AddrHex == norm)
                    return bt;
            }
            return null;
        }

        static bool EndpointMatchesBt(EndpointInfo ep, BtDeviceInfo bt)
        {
            if (ep == null || bt == null)
                return false;
            if (!string.IsNullOrEmpty(ep.BtAddrHex) && ep.BtAddrHex == bt.AddrHex)
                return true;
            if (!string.IsNullOrEmpty(ep.Name) && !string.IsNullOrEmpty(bt.Name)
                && ep.Name.IndexOf(bt.Name, StringComparison.OrdinalIgnoreCase) >= 0)
                return true;
            return false;
        }

        static bool IsActiveEndpoint(EndpointInfo ep)
        {
            if (ep == null)
                return false;
            if ((ep.State & DEVICE_STATE_ACTIVE) == 0)
                return false;
            if ((ep.State & DEVICE_STATE_DISABLED) != 0)
                return false;
            if ((ep.State & DEVICE_STATE_UNPLUGGED) != 0)
                return false;
            if ((ep.State & DEVICE_STATE_NOTPRESENT) != 0)
                return false;
            return true;
        }

        static bool HasActiveEndpoint(List<EndpointInfo> endpoints, BtDeviceInfo bt, int flow)
        {
            if (endpoints == null || bt == null)
                return false;
            foreach (EndpointInfo ep in endpoints)
            {
                if (!EndpointMatchesBt(ep, bt))
                    continue;
                if (flow >= 0 && ep.Flow != flow)
                    continue;
                if (IsActiveEndpoint(ep))
                    return true;
            }
            return false;
        }

        static void EnsureVisible(string deviceId, bool visible)
        {
            PolicySetVisible(deviceId, visible);
        }

        static bool WaitForBtAudio(BtDeviceInfo bt, int timeoutMs)
        {
            DateTime deadline = DateTime.UtcNow.AddMilliseconds(timeoutMs);
            while (true)
            {
                List<EndpointInfo> endpoints = CollectEndpoints();
                if (HasActiveEndpoint(endpoints, bt, eRender))
                {
                    Log("WaitForBtAudio active render=true");
                    return true;
                }
                foreach (EndpointInfo ep in endpoints)
                {
                    if (!EndpointMatchesBt(ep, bt) || ep.Flow != eRender)
                        continue;
                    if ((ep.State & DEVICE_STATE_DISABLED) != 0)
                        EnsureVisible(ep.Id, true);
                }
                if (DateTime.UtcNow >= deadline)
                {
                    bool ok = HasActiveEndpoint(CollectEndpoints(), bt, eRender);
                    Log("WaitForBtAudio timeout activeRender=" + ok);
                    return ok;
                }
                System.Threading.Thread.Sleep(350);
            }
        }

        static IntPtr CreatePolicyConfig()
        {
            Guid[] iids = new Guid[] {
                new Guid("F8679F50-850A-41CF-9C72-430F290290C8"),
                new Guid("568B9108-44BF-40B4-9006-86AFE5B5A620"),
                new Guid("CA286FC3-91FD-42C3-8E9B-CAAFA66242E3"),
                new Guid("6BE54BE8-A068-4875-A49D-0C2966473B11")
            };
            Exception last = null;
            foreach (Guid iid in iids)
            {
                try
                {
                    return CreateCom(typeof(PolicyConfigClient), iid);
                }
                catch (Exception ex)
                {
                    last = ex;
                }
            }
            throw last ?? new COMException("IPolicyConfig unavailable");
        }

        static int PolicySetDefaultSlot(Guid iid)
        {
            if (iid == new Guid("568B9108-44BF-40B4-9006-86AFE5B5A620"))
                return 12;
            return 13;
        }

        static bool IsCurrentDefaultId(string deviceId, int flow)
        {
            if (string.IsNullOrEmpty(deviceId))
                return false;
            IntPtr en = CreateEnumerator();
            try
            {
                string def = GetDefaultId(en, flow);
                return string.Equals(def, deviceId, StringComparison.OrdinalIgnoreCase);
            }
            finally
            {
                Marshal.Release(en);
            }
        }

        static int PolicySetDefaultOn(IntPtr cfg, int slot, string deviceId, int flow, out int lastFail)
        {
            lastFail = -1;
            uint[] roles = flow == eCapture
                ? new uint[] { eCommunications, eMultimedia, eConsole }
                : new uint[] { eMultimedia, eConsole, eCommunications };
            PolicySetDefaultProc fn = ComFn<PolicySetDefaultProc>(cfg, slot);
            bool anyOk = false;
            foreach (uint role in roles)
            {
                int hr = fn(cfg, deviceId, role);
                if (hr == 0)
                    anyOk = true;
                else
                    lastFail = hr;
            }
            if (anyOk || IsCurrentDefaultId(deviceId, flow))
                return 0;
            return lastFail;
        }

        static int PolicySetDefault(string deviceId, int flow)
        {
            Guid[] clsids = new Guid[] {
                new Guid("870AF99C-171D-4F9E-AF0D-E63DF40C2BC9"),
                new Guid("294935CE-F637-4E7C-A41B-AB255460B862")
            };
            Guid[] iids = new Guid[] {
                new Guid("F8679F50-850A-41CF-9C72-430F290290C8"),
                new Guid("568B9108-44BF-40B4-9006-86AFE5B5A620"),
                new Guid("CA286FC3-91FD-42C3-8E9B-CAAFA66242E3"),
                new Guid("6BE54BE8-A068-4875-A49D-0C2966473B11")
            };
            int lastFail = -1;
            foreach (Guid clsid in clsids)
            {
                foreach (Guid iid in iids)
                {
                    Guid c = clsid;
                    Guid i = iid;
                    IntPtr cfg;
                    if (CoCreateInstance(ref c, IntPtr.Zero, 23, ref i, out cfg) != 0 || cfg == IntPtr.Zero)
                        continue;
                    try
                    {
                        int hr = PolicySetDefaultOn(cfg, PolicySetDefaultSlot(i), deviceId, flow, out lastFail);
                        Log("PolicySetDefault iid=" + i.ToString() + " slot=" + PolicySetDefaultSlot(i)
                            + " hr=0x" + unchecked((uint)hr).ToString("X8") + " lastFail=0x" + unchecked((uint)lastFail).ToString("X8"));
                        if (hr == 0)
                            return 0;
                    }
                    catch
                    {
                    }
                    finally
                    {
                        Marshal.Release(cfg);
                    }
                }
            }
            return lastFail;
        }

        static int PolicySetVisible(string deviceId, bool visible)
        {
            IntPtr cfg = CreatePolicyConfig();
            try
            {
                return ComFn<PolicySetVisibleProc>(cfg, 14)(cfg, deviceId, visible ? 1 : 0);
            }
            finally
            {
                Marshal.Release(cfg);
            }
        }

        static bool IsAudioClass(uint ulClassOfDevice)
        {
            uint major = (ulClassOfDevice >> 8) & 0x1F;
            return major == 0x04;
        }

        static string AddrHex(ulong address)
        {
            byte[] b = BitConverter.GetBytes(address);
            return string.Format("{0:X2}{1:X2}{2:X2}{3:X2}{4:X2}{5:X2}", b[5], b[4], b[3], b[2], b[1], b[0]);
        }

        static List<BtDeviceInfo> CollectBluetoothAudio()
        {
            List<BtDeviceInfo> list = new List<BtDeviceInfo>();
            Dictionary<string, bool> seen = new Dictionary<string, bool>(StringComparer.OrdinalIgnoreCase);
            List<IntPtr> radios = OpenRadioHandles();
            try
            {
                foreach (IntPtr hRadio in radios)
                    CollectBtOnRadio(hRadio, list, seen);
            }
            catch
            {
            }
            finally
            {
                CloseRadioHandles(radios);
            }
            return list;
        }

        static void CollectBtOnRadio(IntPtr hRadio, List<BtDeviceInfo> list, Dictionary<string, bool> seen)
        {
            BLUETOOTH_DEVICE_SEARCH_PARAMS sp = new BLUETOOTH_DEVICE_SEARCH_PARAMS();
            sp.dwSize = (uint)Marshal.SizeOf(typeof(BLUETOOTH_DEVICE_SEARCH_PARAMS));
            sp.fReturnAuthenticated = true;
            sp.fReturnRemembered = true;
            sp.fReturnUnknown = false;
            sp.fReturnConnected = true;
            sp.fIssueInquiry = false;
            sp.cTimeoutMultiplier = 0;
            sp.hRadio = hRadio;

            BLUETOOTH_DEVICE_INFO info = NewDeviceInfo();
            IntPtr hFind = BluetoothFindFirstDevice(ref sp, ref info);
            if (hFind == IntPtr.Zero)
                return;
            try
            {
                do
                {
                    AddBtIfAudio(info, list, seen);
                    info = NewDeviceInfo();
                }
                while (BluetoothFindNextDevice(hFind, ref info));
            }
            finally
            {
                BluetoothFindDeviceClose(hFind);
            }
        }

        static void AddBtIfAudio(BLUETOOTH_DEVICE_INFO info, List<BtDeviceInfo> list, Dictionary<string, bool> seen)
        {
            if (!info.fAuthenticated && !info.fRemembered)
                return;
            bool audio = IsAudioClass(info.ulClassOfDevice);
            string name = info.szName ?? "";
            if (!audio)
            {
                string n = name.ToLowerInvariant();
                audio = n.Contains("headphone") || n.Contains("headset") || n.Contains("earbud")
                    || n.Contains("speaker") || n.Contains("airpod") || n.Contains("audio")
                    || n.Contains("buds") || n.Contains("fone") || n.Contains("ouvido");
            }
            if (!audio)
                return;
            string hex = AddrHex(info.Address);
            if (seen.ContainsKey(hex))
                return;
            seen[hex] = true;
            BtDeviceInfo bt = new BtDeviceInfo();
            bt.Address = info.Address;
            bt.AddrHex = hex;
            bt.Name = string.IsNullOrEmpty(name) ? hex : name;
            bt.Connected = info.fConnected;
            bt.ClassOfDevice = info.ulClassOfDevice;
            list.Add(bt);
        }

        static BLUETOOTH_DEVICE_INFO NewDeviceInfo()
        {
            BLUETOOTH_DEVICE_INFO info = new BLUETOOTH_DEVICE_INFO();
            info.dwSize = (uint)Marshal.SizeOf(typeof(BLUETOOTH_DEVICE_INFO));
            return info;
        }

        static bool ForEachRadio(Func<IntPtr, bool> fn, out bool openedAny)
        {
            openedAny = false;
            bool anyFn = false;
            List<IntPtr> radios = OpenRadioHandles();
            try
            {
                foreach (IntPtr hRadio in radios)
                {
                    openedAny = true;
                    if (fn(hRadio))
                        anyFn = true;
                }
            }
            finally
            {
                CloseRadioHandles(radios);
            }
            return anyFn;
        }

        static List<IntPtr> OpenRadioHandles()
        {
            List<IntPtr> list = new List<IntPtr>();
            try
            {
                BLUETOOTH_FIND_RADIO_PARAMS rp = new BLUETOOTH_FIND_RADIO_PARAMS();
                rp.dwSize = (uint)Marshal.SizeOf(typeof(BLUETOOTH_FIND_RADIO_PARAMS));
                IntPtr hRadio;
                IntPtr hFindRadio = BluetoothFindFirstRadio(ref rp, out hRadio);
                if (hFindRadio != IntPtr.Zero)
                {
                    try
                    {
                        do
                        {
                            if (IsValidHandle(hRadio))
                                list.Add(hRadio);
                        }
                        while (BluetoothFindNextRadio(hFindRadio, out hRadio));
                    }
                    finally
                    {
                        BluetoothFindRadioClose(hFindRadio);
                    }
                }
                else
                    Log("BluetoothFindFirstRadio=0 lastWin32=" + Marshal.GetLastWin32Error());

                int nFind = list.Count;
                TryAddRadioFile(@"\\.\BthHci", list);
                int afterHci = list.Count;
                AddSetupApiRadios(list);
                Log("OpenRadioHandles findFirst=" + nFind + " plusBthHci=" + (afterHci - nFind)
                    + " plusSetupApi=" + (list.Count - afterHci) + " total=" + list.Count);
                return list;
            }
            catch
            {
                CloseRadioHandles(list);
                throw;
            }
        }

        static bool IsValidHandle(IntPtr h)
        {
            return h != IntPtr.Zero && h != new IntPtr(-1);
        }

        static void TryAddRadioFile(string path, List<IntPtr> list)
        {
            IntPtr h = CreateFile(path, GENERIC_READ | GENERIC_WRITE, FILE_SHARE_READ | FILE_SHARE_WRITE,
                IntPtr.Zero, OPEN_EXISTING, 0, IntPtr.Zero);
            if (IsValidHandle(h))
                list.Add(h);
        }

        static void AddSetupApiRadios(List<IntPtr> list)
        {
            Guid iface = GuidBthPortDeviceInterface;
            IntPtr devs = SetupDiGetClassDevsGuid(ref iface, IntPtr.Zero, IntPtr.Zero, DIGCF_PRESENT | DIGCF_DEVICEINTERFACE);
            if (!IsValidHandle(devs))
                return;
            try
            {
                for (uint i = 0; ; i++)
                {
                    SP_DEVICE_INTERFACE_DATA data = new SP_DEVICE_INTERFACE_DATA();
                    data.cbSize = (uint)Marshal.SizeOf(typeof(SP_DEVICE_INTERFACE_DATA));
                    if (!SetupDiEnumDeviceInterfaces(devs, IntPtr.Zero, ref iface, i, ref data))
                        break;
                    uint needed;
                    SetupDiGetDeviceInterfaceDetail(devs, ref data, IntPtr.Zero, 0, out needed, IntPtr.Zero);
                    if (needed == 0)
                        continue;
                    IntPtr detail = Marshal.AllocHGlobal((int)needed);
                    try
                    {
                        Marshal.WriteInt32(detail, IntPtr.Size == 8 ? 8 : 6);
                        if (!SetupDiGetDeviceInterfaceDetail(devs, ref data, detail, needed, out needed, IntPtr.Zero))
                            continue;
                        int pathOffset = IntPtr.Size == 8 ? 8 : 4;
                        string path = Marshal.PtrToStringUni(new IntPtr(detail.ToInt64() + pathOffset));
                        if (!string.IsNullOrEmpty(path))
                            TryAddRadioFile(path, list);
                    }
                    finally
                    {
                        Marshal.FreeHGlobal(detail);
                    }
                }
            }
            finally
            {
                SetupDiDestroyDeviceInfoList(devs);
            }
        }

        static void CloseRadioHandles(List<IntPtr> list)
        {
            foreach (IntPtr h in list)
            {
                if (IsValidHandle(h))
                    CloseHandle(h);
            }
            list.Clear();
        }

        static bool DisconnectRadio(BtDeviceInfo bt, out string err)
        {
            err = "";
            int lastWinErr = 0;
            bool ok = false;
            bool openedAny = false;
            ForEachRadio(delegate (IntPtr hRadio)
            {
                ulong addr = bt.Address;
                uint bytes;
                if (DeviceIoControl(hRadio, IOCTL_BTH_DISCONNECT_DEVICE, ref addr, 8, IntPtr.Zero, 0, out bytes, IntPtr.Zero))
                {
                    ok = true;
                    return true;
                }
                lastWinErr = Marshal.GetLastWin32Error();
                return false;
            }, out openedAny);
            if (ok)
                return true;
            if (!openedAny && lastWinErr == 0)
                err = "No Bluetooth radio found";
            else
                err = "Disconnect IOCTL " + lastWinErr;
            return false;
        }

        static bool OtherBtHoldsAudio(BtDeviceInfo other, List<EndpointInfo> endpoints)
        {
            if (other == null)
                return false;
            if (other.Connected)
                return true;
            return HasActiveEndpoint(endpoints, other, eRender);
        }

        static void DisconnectOtherBtAudio(BtDeviceInfo keep, out string dropped)
        {
            dropped = "";
            if (keep == null || string.IsNullOrEmpty(keep.AddrHex))
                return;
            List<BtDeviceInfo> peers = new List<BtDeviceInfo>();
            List<EndpointInfo> endpoints = CollectEndpoints();
            foreach (BtDeviceInfo other in CollectBluetoothAudio())
            {
                if (string.Equals(other.AddrHex, keep.AddrHex, StringComparison.OrdinalIgnoreCase))
                    continue;
                if (!OtherBtHoldsAudio(other, endpoints))
                    continue;
                peers.Add(other);
            }
            if (peers.Count == 0)
            {
                Log("DisconnectOtherBtAudio: no other connected BT audio");
                return;
            }
            List<string> names = new List<string>();
            foreach (BtDeviceInfo other in peers)
            {
                Log("DisconnectOtherBtAudio dropping " + other.Name + " " + other.AddrHex);
                string ioctlErr;
                string svcErr;
                bool ioctlOk = DisconnectRadio(other, out ioctlErr);
                bool svcOk = SetBtServices(other, false, out svcErr);
                Log("  ioctlOk=" + ioctlOk + " ioctlErr=" + ioctlErr + " svcOk=" + svcOk + " svcErr=" + svcErr);
                names.Add(string.IsNullOrEmpty(other.Name) ? other.AddrHex : other.Name);
            }
            dropped = string.Join(", ", names.ToArray());
            lastDroppedBt = dropped;
            DateTime deadline = DateTime.UtcNow.AddMilliseconds(3000);
            while (DateTime.UtcNow < deadline)
            {
                bool anyHeld = false;
                endpoints = CollectEndpoints();
                foreach (BtDeviceInfo other in CollectBluetoothAudio())
                {
                    if (string.Equals(other.AddrHex, keep.AddrHex, StringComparison.OrdinalIgnoreCase))
                        continue;
                    if (OtherBtHoldsAudio(other, endpoints))
                    {
                        anyHeld = true;
                        break;
                    }
                }
                if (!anyHeld)
                    break;
                System.Threading.Thread.Sleep(250);
            }
            Log("DisconnectOtherBtAudio dropped=" + dropped);
        }

        static bool ConnectBtWithFallback(BtDeviceInfo bt, out string method, out string err)
        {
            method = "";
            err = "";
            List<string> errors = new List<string>();
            Log("ConnectBtWithFallback name=" + bt.Name + " addr=" + bt.AddrHex);
            string dropped;
            DisconnectOtherBtAudio(bt, out dropped);

            string svcErr;
            if (SetBtServices(bt, true, out svcErr))
            {
                method = "BluetoothSetServiceState";
                Log("SetBtServices ok, waiting for render");
                if (FinishConnect(bt))
                    return true;
                errors.Add("BluetoothSetServiceState: no audio endpoint appeared");
                Log("SetBtServices: no active render");
            }
            else if (!string.IsNullOrEmpty(svcErr))
            {
                errors.Add(svcErr);
                Log("SetBtServices fail: " + svcErr);
            }

            string pnpErr;
            if (EnablePnpBtAudio(bt, out pnpErr))
            {
                method = "PnP";
                Log("PnP bounce ok, waiting for render");
                if (FinishConnect(bt))
                    return true;
                errors.Add("PnP: no audio endpoint appeared");
                Log("PnP: no active render");
            }
            else if (!string.IsNullOrEmpty(pnpErr))
            {
                errors.Add(pnpErr);
                Log("PnP fail: " + pnpErr);
            }

            method = "";
            err = errors.Count == 0 ? "Bluetooth connect failed" : string.Join("; ", errors);
            Log("ConnectBtWithFallback done fail: " + err);
            return false;
        }

        static bool FinishConnect(BtDeviceInfo bt)
        {
            if (!WaitForBtAudio(bt, 10000))
                return false;
            foreach (EndpointInfo ep in CollectEndpoints())
            {
                if (EndpointMatchesBt(ep, bt) && ep.Flow == eRender)
                    EnsureVisible(ep.Id, true);
            }
            return HasActiveEndpoint(CollectEndpoints(), bt, eRender);
        }

        static bool IsBtAudioPnpNode(string instanceId)
        {
            if (string.IsNullOrEmpty(instanceId))
                return false;
            string u = instanceId.ToUpperInvariant();
            return u.Contains("0000110B") || u.Contains("0000110A")
                || u.Contains("0000111E") || u.Contains("00001108");
        }

        static int BtAudioPnpPriority(string instanceId)
        {
            string u = instanceId.ToUpperInvariant();
            if (u.Contains("0000110B") || u.Contains("0000110A"))
                return 2;
            if (u.Contains("0000111E") || u.Contains("00001108"))
                return 1;
            return 0;
        }

        static bool EnablePnpBtAudio(BtDeviceInfo bt, out string err)
        {
            err = "";
            if (bt == null || string.IsNullOrEmpty(bt.AddrHex))
            {
                err = "No Bluetooth address for PnP connect";
                return false;
            }
            string hex = bt.AddrHex.ToUpperInvariant();
            string colon = FormatAddrColon(hex);
            IntPtr devs = SetupDiGetClassDevsEnum(IntPtr.Zero, "BTHENUM", IntPtr.Zero, DIGCF_PRESENT | DIGCF_ALLCLASSES);
            if (!IsValidHandle(devs))
            {
                err = "PnP BTHENUM enumerate failed " + Marshal.GetLastWin32Error();
                return false;
            }
            List<string> nodes = new List<string>();
            int lastCr = 0;
            try
            {
                for (uint i = 0; ; i++)
                {
                    SP_DEVINFO_DATA info = new SP_DEVINFO_DATA();
                    info.cbSize = (uint)Marshal.SizeOf(typeof(SP_DEVINFO_DATA));
                    if (!SetupDiEnumDeviceInfo(devs, i, ref info))
                        break;
                    StringBuilder sb = new StringBuilder(1024);
                    int needed;
                    if (!SetupDiGetDeviceInstanceId(devs, ref info, sb, sb.Capacity, out needed))
                        continue;
                    string instanceId = sb.ToString();
                    if (string.IsNullOrEmpty(instanceId))
                        continue;
                    string compact = instanceId.ToUpperInvariant().Replace(":", "").Replace("-", "").Replace("_", "");
                    if (compact.IndexOf(hex, StringComparison.Ordinal) < 0
                        && instanceId.IndexOf(colon, StringComparison.OrdinalIgnoreCase) < 0)
                        continue;
                    if (!IsBtAudioPnpNode(instanceId))
                        continue;
                    nodes.Add(instanceId);
                }
            }
            finally
            {
                SetupDiDestroyDeviceInfoList(devs);
            }
            nodes.Sort(delegate (string a, string b)
            {
                return BtAudioPnpPriority(b).CompareTo(BtAudioPnpPriority(a));
            });
            int bounced = 0;
            Log("PnP audio nodes=" + nodes.Count);
            foreach (string instanceId in nodes)
            {
                uint devInst;
                int cr = CM_Locate_DevNode(out devInst, instanceId, 0);
                if (cr != 0)
                {
                    lastCr = cr;
                    Log("PnP locate fail cr=" + cr + " " + instanceId);
                    continue;
                }
                int dis = CM_Disable_DevNode(devInst, 0);
                System.Threading.Thread.Sleep(400);
                cr = CM_Enable_DevNode(devInst, 0);
                Log("PnP bounce disable=" + dis + " enable=" + cr + " " + instanceId);
                if (cr == 0)
                    bounced++;
                else
                    lastCr = cr;
            }
            Log("PnP bounced=" + bounced);
            if (bounced > 0)
                return true;
            if (nodes.Count == 0)
                err = "No BTHENUM audio node for " + hex;
            else
                err = "PnP CM_Enable_DevNode " + lastCr;
            return false;
        }

        static string FormatAddrColon(string hex)
        {
            if (string.IsNullOrEmpty(hex) || hex.Length < 12)
                return hex;
            return string.Format("{0}:{1}:{2}:{3}:{4}:{5}",
                hex.Substring(0, 2), hex.Substring(2, 2), hex.Substring(4, 2),
                hex.Substring(6, 2), hex.Substring(8, 2), hex.Substring(10, 2));
        }

        static bool SetBtServices(BtDeviceInfo bt, bool enable, out string err)
        {
            err = "";
            uint flags = enable ? BLUETOOTH_SERVICE_ENABLE : BLUETOOTH_SERVICE_DISABLE;
            Guid[] services = new Guid[] { GuidA2dpSink, GuidHandsfree, GuidHeadset };
            int ok = 0;
            uint lastCode = 0;
            bool openedAny = false;
            ForEachRadio(delegate (IntPtr hRadio)
            {
                BLUETOOTH_DEVICE_INFO info = NewDeviceInfo();
                info.Address = bt.Address;
                BluetoothGetDeviceInfo(hRadio, ref info);
                info.Address = bt.Address;
                int localOk = 0;
                foreach (Guid svc in services)
                {
                    Guid g = svc;
                    uint code = BluetoothSetServiceState(hRadio, ref info, ref g, flags);
                    lastCode = code;
                    Log("BluetoothSetServiceState enable=" + enable + " svc=" + g.ToString() + " code=" + code);
                    if (code == 0)
                    {
                        ok++;
                        localOk++;
                    }
                }
                return localOk > 0;
            }, out openedAny);
            Log("SetBtServices enable=" + enable + " openedAny=" + openedAny + " ok=" + ok + " lastCode=" + lastCode);
            if (ok > 0)
                return true;
            if (!openedAny)
            {
                err = "No Bluetooth radio found";
                return false;
            }
            if (lastCode == 0 || lastCode == 1168)
            {
                err = enable ? "No audio service to connect" : "No audio service to disconnect";
                return false;
            }
            err = "BluetoothSetServiceState " + lastCode;
            return false;
        }
    }
}
