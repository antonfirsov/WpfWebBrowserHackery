using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace WebBrowserHackery
{
    /// <summary>
    /// Interaction logic for BrowserHolder.xaml
    /// </summary>
    public partial class BrowserHolder : UserControl, IOleClientSite, IServiceProviderNative
    {
        public BrowserHolder()
        {
            InitializeComponent();

            this.Loaded += BrowserHolder_Loaded;
        }

        private IServiceProvider _oldServiceProvider;

        private void BrowserHolder_Loaded(object sender, RoutedEventArgs e)
        {
            Browser.Navigate("about:blank");
            BrowserUtils.SetSilent(Browser, true);

            _oldServiceProvider = Browser.GetServiceProvider();

            IOleObject wut = (IOleObject) Browser.GetInteropWebBrowser2();
            wut.SetClientSite(this);

            Browser.Navigating += Browser_Navigating;

            if (!DesignerProperties.GetIsInDesignMode(this))
            {
                Browser.Navigate("https://program.momentum.hu/");
            }
        }

        private void Browser_Navigating(object sender, NavigatingCancelEventArgs e)
        {

            Debug.Print("Navigating to: "+e.Uri);
        }

        void IOleClientSite.SaveObject()
        {
        }

        void IOleClientSite.GetMoniker(uint dwAssign, uint dwWhichMoniker, ref object ppmk)
        {
        }

        void IOleClientSite.GetContainer(ref object ppContainer)
        {
            ppContainer = this;
        }

        void IOleClientSite.ShowObject()
        {
        }

        void IOleClientSite.OnShowWindow(bool fShow)
        {
        }

        void IOleClientSite.RequestNewObjectLayout()
        {
        }

        private DownloadManagerImplementation _downloadManager = new DownloadManagerImplementation();

        public int QueryService(ref Guid guidService, ref Guid riid, out IntPtr ppvObject)
        {
            Guid SID_SDownloadManager = new Guid("988934A4-064B-11D3-BB80-00104B35E7F9");
            Guid IID_IDownloadManager = new Guid("988934A4-064B-11D3-BB80-00104B35E7F9");

            if ((guidService == IID_IDownloadManager && riid == IID_IDownloadManager))
            {
                ppvObject = Marshal.GetComInterfaceForObject(_downloadManager, typeof(IDownloadManager));
                return 0; //S_OK
            }

            ppvObject = new IntPtr();
            return INET_E_DEFAULT_ACTION;
        }

        public const int INET_E_DEFAULT_ACTION = unchecked((int)0x800C0011);
        public const int S_OK = unchecked((int)0x00000000);

    }
}
