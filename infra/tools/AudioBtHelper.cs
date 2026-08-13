using System;
using System.Collections.Generic;
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

        [DllImport("ole32.dll")]
        static extern int PropVariantClear(ref PROPVARIANT pvar);

        [DllImport("ole32.dll")]
        static extern void CoTaskMemFree(IntPtr pv);

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

        const uint IOCTL_BTH_DISCONNECT_DEVICE = 0x0041000C;
        const uint GENERIC_READ = 0x80000000;
        const uint GENERIC_WRITE = 0x40000000;
        const uint FILE_SHARE_READ = 0x1;
        const uint FILE_SHARE_WRITE = 0x2;
        const uint OPEN_EXISTING = 3;

        public static string ListTsv()
        {
            try
            {
                List<EndpointInfo> endpoints = CollectEndpoints();
                List<BtDeviceInfo> btDevs = CollectBluetoothAudio();
                StringBuilder sb = new StringBuilder();
                sb.AppendLine("id\tkind\tname\tstate\tisDefault\tcanConnect");

                Dictionary<string, bool> seenBt = new Dictionary<string, bool>(StringComparer.OrdinalIgnoreCase);
                foreach (BtDeviceInfo bt in btDevs)
                {
                    if (seenBt.ContainsKey(bt.AddrHex))
                        continue;
                    seenBt[bt.AddrHex] = true;
                    string state = bt.Connected ? "Connected" : "Disconnected";
                    bool isDef = false;
                    foreach (EndpointInfo ep in endpoints)
                    {
                        if (ep.IsDefault && EndpointMatchesBt(ep, bt))
                        {
                            isDef = true;
                            state = "Connected · Default";
                            break;
                        }
                    }
                    if (bt.Connected && !isDef)
                    {
                        foreach (EndpointInfo ep in endpoints)
                        {
                            if (EndpointMatchesBt(ep, bt) && (ep.State & DEVICE_STATE_ACTIVE) != 0)
                            {
                                state = "Connected";
                                break;
                            }
                        }
                    }
                    sb.AppendLine(Tsv("bt:" + bt.AddrHex, "BT", bt.Name, state, isDef ? "1" : "0", "1"));
                }

                foreach (EndpointInfo ep in endpoints)
                {
                    if ((ep.State & DEVICE_STATE_NOTPRESENT) != 0 && !ep.IsBluetooth)
                        continue;
                    string kind = ep.Flow == eRender ? "Out" : "In";
                    string state = StateLabel(ep);
                    sb.AppendLine(Tsv(ep.Id, kind, ep.Name, state, ep.IsDefault ? "1" : "0", ep.IsBluetooth ? "1" : "0"));
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
                EnsureVisible(ep.Id, true);
                int hr = PolicySetDefault(ep.Id, ep.Flow);
                if (hr != 0)
                    return "ERR\tSetDefault failed HRESULT 0x" + hr.ToString("X8");
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
                    return "ERR\t" + (visible ? "Enable" : "Disable") + " failed HRESULT 0x" + hr.ToString("X8");
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
                List<EndpointInfo> endpoints = CollectEndpoints();
                List<EndpointInfo> targets = new List<EndpointInfo>();
                if (id != null && id.StartsWith("bt:", StringComparison.OrdinalIgnoreCase))
                {
                    BtDeviceInfo bt = FindBt(id.Substring(3));
                    if (bt == null)
                        return "ERR\tBluetooth device not found";
                    foreach (EndpointInfo ep in endpoints)
                    {
                        if (EndpointMatchesBt(ep, bt) && (ep.State & DEVICE_STATE_NOTPRESENT) == 0)
                            targets.Add(ep);
                    }
                }
                else
                {
                    EndpointInfo one = FindEndpointById(endpoints, id);
                    if (one != null)
                        targets.Add(one);
                }
                if (targets.Count == 0)
                    return "ERR\tNo audio endpoint found to isolate";

                HashSet<int> flows = new HashSet<int>();
                foreach (EndpointInfo t in targets)
                    flows.Add(t.Flow);

                StringBuilder names = new StringBuilder();
                foreach (EndpointInfo t in targets)
                {
                    EnsureVisible(t.Id, true);
                    int hr = PolicySetDefault(t.Id, t.Flow);
                    if (hr != 0)
                        return "ERR\tSetDefault failed HRESULT 0x" + hr.ToString("X8");
                    if (names.Length > 0)
                        names.Append(", ");
                    names.Append(t.Name);
                }

                foreach (EndpointInfo ep in endpoints)
                {
                    if (!flows.Contains(ep.Flow))
                        continue;
                    if ((ep.State & DEVICE_STATE_NOTPRESENT) != 0)
                        continue;
                    bool isTarget = false;
                    foreach (EndpointInfo t in targets)
                    {
                        if (string.Equals(t.Id, ep.Id, StringComparison.OrdinalIgnoreCase))
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
                return "OK\tIsolated: " + Cell(names.ToString());
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
                BtDeviceInfo bt = ResolveBt(id);
                if (bt == null)
                    return "ERR\tNot a paired Bluetooth audio device";
                string err;
                if (!SetBtServices(bt, true, out err))
                    return "ERR\t" + Cell(err);
                List<EndpointInfo> endpoints = CollectEndpoints();
                foreach (EndpointInfo ep in endpoints)
                {
                    if (EndpointMatchesBt(ep, bt))
                        EnsureVisible(ep.Id, true);
                }
                return "OK\tConnect requested: " + Cell(bt.Name);
            }
            catch (Exception ex)
            {
                return "ERR\t" + Cell(ex.Message);
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

        static string StateLabel(EndpointInfo ep)
        {
            if ((ep.State & DEVICE_STATE_DISABLED) != 0)
                return "Disabled";
            if (ep.IsDefault)
                return "Default";
            if ((ep.State & DEVICE_STATE_UNPLUGGED) != 0)
                return ep.IsBluetooth ? "Disconnected" : "Unplugged";
            if ((ep.State & DEVICE_STATE_NOTPRESENT) != 0)
                return "Not present";
            return "Enabled";
        }

        static string Tsv(string id, string kind, string name, string state, string isDefault, string canConnect)
        {
            return Cell(id) + "\t" + Cell(kind) + "\t" + Cell(name) + "\t" + Cell(state) + "\t" + isDefault + "\t" + canConnect;
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
                return render ?? capture;
            return render ?? capture;
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

        static void EnsureVisible(string deviceId, bool visible)
        {
            PolicySetVisible(deviceId, visible);
        }

        static IntPtr CreatePolicyConfig()
        {
            Guid[] iids = new Guid[] {
                new Guid("F8679F50-850A-41CF-9C72-430F290290C8"),
                new Guid("568B9108-44BF-40B4-9006-86AFE5B5A620"),
                new Guid("CA286FC3-91FD-42C3-8E9B-CAAFA66242E3")
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

        static int PolicySetDefault(string deviceId, int flow)
        {
            IntPtr cfg = CreatePolicyConfig();
            try
            {
                int last = -1;
                uint[] roles = flow == eCapture
                    ? new uint[] { eCommunications, eMultimedia, eConsole }
                    : new uint[] { eMultimedia, eConsole, eCommunications };
                PolicySetDefaultProc fn = ComFn<PolicySetDefaultProc>(cfg, 13);
                foreach (uint role in roles)
                    last = fn(cfg, deviceId, role);
                return last;
            }
            finally
            {
                Marshal.Release(cfg);
            }
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
            try
            {
                BLUETOOTH_FIND_RADIO_PARAMS rp = new BLUETOOTH_FIND_RADIO_PARAMS();
                rp.dwSize = (uint)Marshal.SizeOf(typeof(BLUETOOTH_FIND_RADIO_PARAMS));
                IntPtr hRadio;
                IntPtr hFindRadio = BluetoothFindFirstRadio(ref rp, out hRadio);
                if (hFindRadio == IntPtr.Zero)
                    return list;
                try
                {
                    do
                    {
                        CollectBtOnRadio(hRadio, list, seen);
                        CloseHandle(hRadio);
                    }
                    while (BluetoothFindNextRadio(hFindRadio, out hRadio));
                }
                finally
                {
                    BluetoothFindRadioClose(hFindRadio);
                }
            }
            catch
            {
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

        static bool ForEachRadio(Func<IntPtr, bool> fn)
        {
            BLUETOOTH_FIND_RADIO_PARAMS rp = new BLUETOOTH_FIND_RADIO_PARAMS();
            rp.dwSize = (uint)Marshal.SizeOf(typeof(BLUETOOTH_FIND_RADIO_PARAMS));
            IntPtr hRadio;
            IntPtr hFindRadio = BluetoothFindFirstRadio(ref rp, out hRadio);
            if (hFindRadio == IntPtr.Zero)
                return false;
            bool any = false;
            try
            {
                do
                {
                    if (fn(hRadio))
                        any = true;
                    CloseHandle(hRadio);
                }
                while (BluetoothFindNextRadio(hFindRadio, out hRadio));
            }
            finally
            {
                BluetoothFindRadioClose(hFindRadio);
            }
            return any;
        }

        static bool DisconnectRadio(BtDeviceInfo bt, out string err)
        {
            err = "";
            int lastWinErr = 0;
            bool ok = false;
            bool foundRadio = ForEachRadio(delegate (IntPtr hRadio)
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
            });
            if (ok)
                return true;

            IntPtr hci = CreateFile(@"\\.\BthHci", GENERIC_READ | GENERIC_WRITE, FILE_SHARE_READ | FILE_SHARE_WRITE,
                IntPtr.Zero, OPEN_EXISTING, 0, IntPtr.Zero);
            if (hci != IntPtr.Zero && hci != new IntPtr(-1))
            {
                try
                {
                    ulong addr = bt.Address;
                    uint bytes;
                    if (DeviceIoControl(hci, IOCTL_BTH_DISCONNECT_DEVICE, ref addr, 8, IntPtr.Zero, 0, out bytes, IntPtr.Zero))
                        return true;
                    lastWinErr = Marshal.GetLastWin32Error();
                }
                finally
                {
                    CloseHandle(hci);
                }
            }

            if (!foundRadio && lastWinErr == 0)
                err = "No Bluetooth radio found";
            else
                err = "Disconnect IOCTL " + lastWinErr;
            return false;
        }

        static bool SetBtServices(BtDeviceInfo bt, bool enable, out string err)
        {
            err = "";
            uint flags = enable ? BLUETOOTH_SERVICE_ENABLE : BLUETOOTH_SERVICE_DISABLE;
            Guid[] services = new Guid[] { GuidA2dpSink, GuidHandsfree, GuidHeadset };
            int ok = 0;
            uint lastCode = 0;
            bool foundRadio = ForEachRadio(delegate (IntPtr hRadio)
            {
                BLUETOOTH_DEVICE_INFO info = NewDeviceInfo();
                info.Address = bt.Address;
                BluetoothGetDeviceInfo(hRadio, ref info);
                info.Address = bt.Address;
                foreach (Guid svc in services)
                {
                    Guid g = svc;
                    uint code = BluetoothSetServiceState(hRadio, ref info, ref g, flags);
                    lastCode = code;
                    if (code == 0)
                        ok++;
                }
                return ok > 0;
            });
            if (!foundRadio)
            {
                err = "No Bluetooth radio found";
                return false;
            }
            if (ok > 0)
                return true;
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
