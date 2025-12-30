param([int]$Level = 100)

$scalar = [float]([Math]::Max(0, [Math]::Min(100, $Level)) / 100)

$code = @"
using System;
using System.Runtime.InteropServices;

[ComImport, Guid("BCDE0395-E52F-467C-8E3D-C4579291692E")]
public class MMDeviceEnumerator {}

[Guid("A95664D2-9614-4F35-A746-DE8DB63617E6")]
[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
public interface IMMDeviceEnumerator {
  int NotImpl1();
  int GetDefaultAudioEndpoint(int dataFlow, int role, out IMMDevice ppDevice);
}

[Guid("D666063F-1587-4E43-81F1-B948E807363F")]
[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
public interface IMMDevice {
  int Activate(ref Guid iid, int dwClsCtx, IntPtr pActivationParams, out IAudioEndpointVolume ppInterface);
  int OpenPropertyStore(int stgmAccess, out IntPtr ppProperties);
  int GetId([MarshalAs(UnmanagedType.LPWStr)] out string ppstrId);
  int GetState(out int pdwState);
}

[Guid("5CDF2C82-841E-4546-9722-0CF74078229A")]
[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
public interface IAudioEndpointVolume {
  int RegisterControlChangeNotify(IntPtr pNotify);
  int UnregisterControlChangeNotify(IntPtr pNotify);
  int GetChannelCount(out uint pnChannelCount);
  int SetMasterVolumeLevel(float fLevelDB, Guid pguidEventContext);
  int SetMasterVolumeLevelScalar(float fLevel, Guid pguidEventContext);
  int GetMasterVolumeLevel(out float pfLevelDB);
  int GetMasterVolumeLevelScalar(out float pfLevel);
  int SetChannelVolumeLevel(uint nChannel, float fLevelDB, Guid pguidEventContext);
  int SetChannelVolumeLevelScalar(uint nChannel, float fLevel, Guid pguidEventContext);
  int GetChannelVolumeLevel(uint nChannel, out float pfLevelDB);
  int GetChannelVolumeLevelScalar(uint nChannel, out float pfLevel);
  int SetMute(bool bMute, Guid pguidEventContext);
  int GetMute(out bool pbMute);
  int GetVolumeStepInfo(out uint pnStep, out uint pnStepCount);
  int VolumeStepUp(Guid pguidEventContext);
  int VolumeStepDown(Guid pguidEventContext);
  int QueryHardwareSupport(out uint pdwHardwareSupportMask);
  int GetVolumeRange(out float pflVolumeMindB, out float pflVolumeMaxdB, out float pflVolumeIncrementdB);
}

public static class MicVolNative {
  // dataFlow eCapture=1, roles: eCommunications=2, eMultimedia=1, eConsole=0
  public static float SetDefaultCaptureVolumeScalar(float scalar) {
    var enumerator = (IMMDeviceEnumerator)Activator.CreateInstance(typeof(MMDeviceEnumerator));
    IMMDevice dev;
    int hr = enumerator.GetDefaultAudioEndpoint(1, 2, out dev);
    if (hr != 0) hr = enumerator.GetDefaultAudioEndpoint(1, 1, out dev);
    if (hr != 0) hr = enumerator.GetDefaultAudioEndpoint(1, 0, out dev);
    if (hr != 0) Marshal.ThrowExceptionForHR(hr);

    Guid iid = typeof(IAudioEndpointVolume).GUID;
    IAudioEndpointVolume vol;
    Marshal.ThrowExceptionForHR(dev.Activate(ref iid, 23, IntPtr.Zero, out vol)); // CLSCTX_ALL = 23

    Guid g = Guid.Empty;
    Marshal.ThrowExceptionForHR(vol.SetMasterVolumeLevelScalar(scalar, g));
    float cur;
    Marshal.ThrowExceptionForHR(vol.GetMasterVolumeLevelScalar(out cur));
    return cur;
  }
}
"@

Add-Type -TypeDefinition $code

$result = [MicVolNative]::SetDefaultCaptureVolumeScalar($scalar)
Write-Output ("Mic volume scalar: {0:P0}" -f $result)
