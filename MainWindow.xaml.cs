using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Windows.Threading;
using System.IO;
using System.Reflection;

namespace WebPageRotator
{
    public partial class MainWindow : Window
    {
        private struct AddressAndTimeToShow
        {
            public string URL;
            public int TimeToShow;
        }

        private const int DefaultTimeToShowWebpage = 10;
        private string AppTitle = "Cycling web pages";

        DispatcherTimer _dispatcherTimer;
        List<AddressAndTimeToShow> _webPagesToShow = new List<AddressAndTimeToShow>();
        int _currentPageCounter = 0;

        public MainWindow()
        {
            InitializeComponent();

            try
            {
                TextReader fileReader = File.OpenText(@"WebPagesToShow.txt");
                string configLine = fileReader.ReadLine();
                while (!string.IsNullOrEmpty(configLine))
                {
                    var settings = configLine.Split(' ');
                    AddressAndTimeToShow addressAndTimeToShow = new AddressAndTimeToShow();

                    if (settings.Length == 2)
                    {
                        addressAndTimeToShow.URL = settings[0];
                        int time = DefaultTimeToShowWebpage;
                        if (int.TryParse(settings[1], out time))
                        {
                            addressAndTimeToShow.TimeToShow = time;
                        }
                    }
                    else if (settings.Length == 1)
                    {
                        addressAndTimeToShow.URL = settings[0];
                        addressAndTimeToShow.TimeToShow = DefaultTimeToShowWebpage;
                    }

                    _webPagesToShow.Add(addressAndTimeToShow);
                    configLine = fileReader.ReadLine();
                }
                if (_webPagesToShow.Count > 0)
                {
                    _dispatcherTimer = new DispatcherTimer();
                    _dispatcherTimer.Tick += new EventHandler(dispatcherTimer_Tick);
                    _dispatcherTimer.Interval = new TimeSpan(0, 0, 5);
                    _dispatcherTimer.Start();
                    dispatcherTimer_Tick(null, null);
                }
            }
            catch (Exception ex)
            {
                ReportError("WebPagesToShow.txt", "ex.Message");
            }
        }
        
        private void dispatcherTimer_Tick(object sender, EventArgs e)
        {
            string page = string.Empty;

            try
            {
                if (_currentPageCounter >= _webPagesToShow.Count)
                {
                    _currentPageCounter = 0;
                }
                page = _webPagesToShow[_currentPageCounter].URL;

                if (page.Contains('\\') && !File.Exists(page))
                {
                    throw new Exception(string.Format("The page {0} could not be found", page));
                }

                _webBrowser.Navigate(page);

                int interval = _webPagesToShow[_currentPageCounter].TimeToShow;
                _dispatcherTimer.Interval = new TimeSpan(0, 0, interval);
                Title = string.Format("{0} (displaying page {1} for {2} seconds)", AppTitle, page, interval);

                CommandManager.InvalidateRequerySuggested();
            }
            catch (Exception ex)
            {
                ReportError(page, ex.Message);
            }

            _currentPageCounter++;
        }

        private void ReportError(string page, string errorMessage)
        {
            try
            {
                string message = string.Format("Unable to navigate to page: {0}. Exception: {1}", page, errorMessage);
                _webBrowser.NavigateToString(message);
                CommandManager.InvalidateRequerySuggested();
            }
            catch { }
        }

        public void HideScriptErrors(WebBrowser wb, bool Hide)
        {
            FieldInfo fiComWebBrowser = typeof(WebBrowser).GetField("_axIWebBrowser2", BindingFlags.Instance | BindingFlags.NonPublic);
            if (fiComWebBrowser == null) return;
            object objComWebBrowser = fiComWebBrowser.GetValue(wb);
            if (objComWebBrowser == null) return;
            objComWebBrowser.GetType().InvokeMember("Silent", BindingFlags.SetProperty, null, objComWebBrowser, new object[] { Hide });
        }

        private void _webBrowser_Navigated(object sender, NavigationEventArgs e)
        {
            HideScriptErrors((WebBrowser)sender, true);
        }
    }
}
