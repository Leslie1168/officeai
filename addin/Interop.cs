using System;
using System.Runtime.InteropServices;

namespace OfficeAi
{
    public enum ext_ConnectMode
    {
        ext_cm_AfterStartup = 0,
        ext_cm_Startup = 1,
        ext_cm_External = 2,
        ext_cm_CommandLine = 3
    }

    public enum ext_DisconnectMode
    {
        ext_dm_HostShutdown = 0,
        ext_dm_UserClosed = 1
    }

    /// <summary>
    /// 自行声明 IDTExtensibility2（本机没有 Extensibility.dll PIA，且这套写法已在
    /// PDF 工具箱加载项里跑通）。Office 侧的其余接口统一用 GAC 里的 Office PIA。
    /// </summary>
    [Guid("B65AD801-ABAF-11D0-BB8B-00A0C90F2744")]
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface IDTExtensibility2
    {
        [DispId(1)]
        void OnConnection([In, MarshalAs(UnmanagedType.IDispatch)] object Application,
                          [In] ext_ConnectMode ConnectMode,
                          [In, MarshalAs(UnmanagedType.IDispatch)] object AddInInst,
                          [In, MarshalAs(UnmanagedType.SafeArray)] ref Array custom);

        [DispId(2)]
        void OnDisconnection([In] ext_DisconnectMode RemoveMode,
                             [In, MarshalAs(UnmanagedType.SafeArray)] ref Array custom);

        [DispId(3)]
        void OnAddInsUpdate([In, MarshalAs(UnmanagedType.SafeArray)] ref Array custom);

        [DispId(4)]
        void OnStartupComplete([In, MarshalAs(UnmanagedType.SafeArray)] ref Array custom);

        [DispId(5)]
        void OnBeginShutdown([In, MarshalAs(UnmanagedType.SafeArray)] ref Array custom);
    }
}
