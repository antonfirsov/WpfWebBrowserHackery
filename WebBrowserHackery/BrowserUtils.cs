using System;
using System.Diagnostics;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Windows;
using SHDocVw;
using WebBrowser = System.Windows.Controls.WebBrowser;

namespace WebBrowserHackery
{
    public static class BrowserUtils
    {
        /// <summary>
        /// Gets an interop web browser.
        /// </summary>
        /// <param name="browser"></param>
        /// <returns></returns>
        public static SHDocVw.WebBrowser GetInteropWebBrowser(this WebBrowser browser)
        {
            SHDocVw.IWebBrowser2 browser2 = browser.GetInteropWebBrowser2();
            SHDocVw.WebBrowser wb = (SHDocVw.WebBrowser)browser2;

            return wb;
        }

        public static IServiceProvider GetServiceProvider(this WebBrowser browser)
        {
            return (IServiceProvider)browser.Document;
        }

        public static SHDocVw.IWebBrowser2 GetInteropWebBrowser2(this WebBrowser browser)
        {
            Guid serviceGuid = new Guid("0002DF05-0000-0000-C000-000000000046");
            Guid iid = typeof(SHDocVw.IWebBrowser2).GUID;
            IServiceProvider serviceProvider = browser.GetServiceProvider();
            SHDocVw.IWebBrowser2 browser2 = (SHDocVw.IWebBrowser2)serviceProvider.QueryService(ref serviceGuid, ref iid);
            return browser2;
        }

        public static void SetSilent(WebBrowser browser, bool silent)
        {
            if (browser == null)
                throw new ArgumentNullException(nameof(browser));

            SHDocVw.WebBrowser browser2 = browser.GetInteropWebBrowser();
            if (browser2 != null)
                browser2.Silent = silent;
        }

        public static void Hack(this WebBrowser browser)
        {
            IWebBrowser2 b = browser.GetInteropWebBrowser2();

            DWebBrowserEvents2_Event evt = (DWebBrowserEvents2_Event) b;
            evt.DownloadBegin += Evt_DownloadBegin;
        }

        private static void Evt_DownloadBegin()
        {
            Debug.Print("Download begin!");
        }
    }

    /// <summary>
    /// Provides the COM interface for the IServiceProvider.
    /// </summary>
    [ComImport, Guid("6D5140C1-7436-11CE-8034-00AA006009FA"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IServiceProvider
    {
        /// <summary>
        /// Queries the service.
        /// </summary>
        /// <param name="serviceGuid"></param>
        /// <param name="riid"></param>
        /// <returns></returns>
        [return: MarshalAs(UnmanagedType.IUnknown)]
        object QueryService(ref Guid serviceGuid, ref Guid riid);
    }

    [ComImport,
     GuidAttribute("6d5140c1-7436-11ce-8034-00aa006009fa"),
     InterfaceTypeAttribute(ComInterfaceType.InterfaceIsIUnknown),
     ComVisible(false)]
    public interface IServiceProviderNative
    {
        [return: MarshalAs(UnmanagedType.I4)]
        [PreserveSig]
        int QueryService(ref Guid guidService, ref Guid riid, out IntPtr ppvObject);
    }

    [ComImport]
    [Guid("00000118-0000-0000-C000-000000000046")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IOleClientSite
    {
        void SaveObject();
        void GetMoniker(uint dwAssign, uint dwWhichMoniker, ref object ppmk);
        void GetContainer(ref object ppContainer);
        void ShowObject();
        void OnShowWindow(bool fShow);
        void RequestNewObjectLayout();
    }

    [ComImport]
    [Guid("00000112-0000-0000-C000-000000000046")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IOleObject
    {
        void SetClientSite(IOleClientSite pClientSite);
        void GetClientSite(ref IOleClientSite ppClientSite);
        void SetHostNames(object szContainerApp, object szContainerObj);
        void Close(uint dwSaveOption);
        void SetMoniker(uint dwWhichMoniker, object pmk);
        void GetMoniker(uint dwAssign, uint dwWhichMoniker, object ppmk);
        void InitFromData(object pDataObject, bool fCreation, uint dwReserved);
        void GetClipboardData(uint dwReserved, ref object ppDataObject);
        void DoVerb(uint iVerb, uint lpmsg, object pActiveSite, uint lindex, uint hwndParent, uint lprcPosRect);
        void EnumVerbs(ref object ppEnumOleVerb);
        void Update();
        void IsUpToDate();
        void GetUserClassID(uint pClsid);
        void GetUserType(uint dwFormOfType, uint pszUserType);
        void SetExtent(uint dwDrawAspect, uint psizel);
        void GetExtent(uint dwDrawAspect, uint psizel);
        void Advise(object pAdvSink, uint pdwConnection);
        void Unadvise(uint dwConnection);
        void EnumAdvise(ref object ppenumAdvise);
        void GetMiscStatus(uint dwAspect, uint pdwStatus);
        void SetColorScheme(object pLogpal);
    };

    [ComVisible(false), ComImport]
    [Guid("988934A4-064B-11D3-BB80-00104B35E7F9")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IDownloadManager
    {
        [return: MarshalAs(UnmanagedType.I4)]
        [PreserveSig]
        int Download(
            [In, MarshalAs(UnmanagedType.Interface)] IMoniker pmk,
            [In, MarshalAs(UnmanagedType.Interface)] IBindCtx pbc,
            [In, MarshalAs(UnmanagedType.U4)] UInt32 dwBindVerb,
            [In] int grfBINDF,
            [In] IntPtr pBindInfo,
            [In, MarshalAs(UnmanagedType.LPWStr)] string pszHeaders,
            [In, MarshalAs(UnmanagedType.LPWStr)] string pszRedir,
            [In, MarshalAs(UnmanagedType.U4)] uint uiCP);
    }

    [System.Runtime.InteropServices.ComVisible(true)]
    [System.Runtime.InteropServices.Guid("bdb9c34c-d0ca-448e-b497-8de62e709744")]
    public class DownloadManagerImplementation : IDownloadManager
    {

        /// <summary>
        /// Return S_OK (0) so that IE will stop to download the file itself. 
        /// Else the default download user interface is used.
        /// </summary>
        public int Download(IMoniker pmk, IBindCtx pbc, uint dwBindVerb, int grfBINDF,
            IntPtr pBindInfo, string pszHeaders, string pszRedir, uint uiCP)
        {
            // Get the display name of the pointer to an IMoniker interface that specifies
            // the object to be downloaded.
            string name = string.Empty;
            pmk.GetDisplayName(pbc, null, out name);

            if (!string.IsNullOrEmpty(name))
            {
                Uri url = null;
                bool result = Uri.TryCreate(name, UriKind.Absolute, out url);

                if (result)
                {
                    //Implement your custom download manager here
                    //Example:
                    //WebDownloadForm manager = new WebDownloadForm();
                    //manager.FileToDownload = url.AbsoluteUri;
                    //manager.Show();
                    MessageBox.Show("Download URL is: " + url);
                    return 0; //Return S_OK
                }
            }
            return 1; //unspecified error occured.
        }

    }
}