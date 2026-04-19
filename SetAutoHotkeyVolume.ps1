param(
    [int]$Level = 70
)

$source = @"
using System;
using System.Runtime.InteropServices;
using System.Diagnostics;

namespace AudioVol {
    [ComImport]
    [Guid("BCDE0395-E52F-467C-8E3D-C4579291692E")]
    class MMDeviceEnumerator {}

    [Guid("A95664D2-9614-4F35-A746-DE8DB63617E6"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    interface IMMDeviceEnumerator {
        void EnumAudioEndpoints(int dataFlow, int stateMask, out IntPtr devices);
        void GetDefaultAudioEndpoint(int dataFlow, int role, out IMMDevice ppEndpoint);
    }

    [Guid("D666063F-1587-4E43-81F1-B948E807363F"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    interface IMMDevice {
        void Activate(ref Guid iid, int dwClsCtx, IntPtr pActivationParams, [MarshalAs(UnmanagedType.IUnknown)] out object ppInterface);
    }

    [Guid("77AA99A0-1BD6-484F-8BC7-2C654C9A9B6F"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    interface IAudioSessionManager2 {
        void GetAudioSessionControl(ref Guid AudioSessionGuid, int StreamFlags, out IntPtr SessionControl);
        void GetSimpleAudioVolume(ref Guid AudioSessionGuid, int StreamFlags, out IntPtr SimpleAudioVolume);
        void GetSessionEnumerator(out IAudioSessionEnumerator SessionEnum);
        void RegisterSessionNotification(IntPtr SessionNotification);
        void UnregisterSessionNotification(IntPtr SessionNotification);
        void RegisterDuckNotification(IntPtr sessionID, IntPtr duckNotification);
        void UnregisterDuckNotification(IntPtr duckNotification);
    }

    [Guid("E2F5BB11-0570-40CA-ACDD-3AA01277DEE8"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    interface IAudioSessionEnumerator {
        void GetCount(out int SessionCount);
        void GetSession(int SessionCount, out IAudioSessionControl Session);
    }

    [Guid("F4B1A599-7266-4319-A8CA-E70ACB11E8CD"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    interface IAudioSessionControl {
        void GetState(out int pRetVal);
        void GetDisplayName(out IntPtr pRetVal);
        void SetDisplayName(string Value, Guid EventContext);
        void GetIconPath(out IntPtr pRetVal);
        void SetIconPath(string Value, Guid EventContext);
        void GetGroupingParam(out Guid pRetVal);
        void SetGroupingParam(Guid Override, Guid EventContext);
        void RegisterAudioSessionNotification(IntPtr NewNotifications);
        void UnregisterAudioSessionNotification(IntPtr NewNotifications);
    }

    [Guid("BFB7FF88-7239-4FC9-8FA2-07C950BE9C6D"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    interface IAudioSessionControl2 {
        void GetState(out int pRetVal);
        void GetDisplayName(out IntPtr pRetVal);
        void SetDisplayName(string Value, Guid EventContext);
        void GetIconPath(out IntPtr pRetVal);
        void SetIconPath(string Value, Guid EventContext);
        void GetGroupingParam(out Guid pRetVal);
        void SetGroupingParam(Guid Override, Guid EventContext);
        void RegisterAudioSessionNotification(IntPtr NewNotifications);
        void UnregisterAudioSessionNotification(IntPtr NewNotifications);
        void GetSessionIdentifier(out IntPtr pRetVal);
        void GetSessionInstanceIdentifier(out IntPtr pRetVal);
        void GetProcessId(out int pRetVal);
        void IsSystemSoundsSession();
        void SetDuckingPreference(bool optOut);
    }

    [Guid("87CE5498-68D6-44E5-9215-6DA47EF883D8"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    interface ISimpleAudioVolume {
        void SetMasterVolume(float fLevel, ref Guid EventContext);
        void GetMasterVolume(out float pfLevel);
        void SetMute(bool bMute, ref Guid EventContext);
        void GetMute(out bool pbMute);
    }

    public class VolumeSetter {
        public static void SetAutoHotkeyVolume(int percent) {
            try {
                float level = Math.Max(0f, Math.Min(100f, (float)percent)) / 100.0f;
                IMMDeviceEnumerator enumerator = (IMMDeviceEnumerator)new MMDeviceEnumerator();
                IMMDevice device;
                enumerator.GetDefaultAudioEndpoint(0, 1, out device);
                
                Guid iid = typeof(IAudioSessionManager2).GUID;
                object managerObj;
                device.Activate(ref iid, 23, IntPtr.Zero, out managerObj);
                IAudioSessionManager2 manager = (IAudioSessionManager2)managerObj;
                
                IAudioSessionEnumerator sessionEnumerator;
                manager.GetSessionEnumerator(out sessionEnumerator);
                
                int sessionCount;
                sessionEnumerator.GetCount(out sessionCount);
                
                for (int i = 0; i < sessionCount; i++) {
                    IAudioSessionControl control;
                    sessionEnumerator.GetSession(i, out control);
                    
                    IAudioSessionControl2 control2 = control as IAudioSessionControl2;
                    if (control2 != null) {
                        int pid;
                        control2.GetProcessId(out pid);
                        if (pid > 0) {
                            try {
                                Process proc = Process.GetProcessById(pid);
                                if (proc.ProcessName.ToLower().Contains("autohotkey")) {
                                    ISimpleAudioVolume simpleVol = control as ISimpleAudioVolume;
                                    if (simpleVol != null) {
                                        Guid ctx = Guid.Empty;
                                        simpleVol.SetMasterVolume(level, ref ctx);
                                    }
                                }
                            } catch {}
                        }
                    }
                }
            } catch (Exception ex) {
                Console.WriteLine("Error: " + ex.Message);
            }
        }
    }
}
"@
Add-Type -TypeDefinition $source
[AudioVol.VolumeSetter]::SetAutoHotkeyVolume($Level)
