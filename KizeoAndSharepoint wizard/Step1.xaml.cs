using KizeoAndSharepoint_wizard.Models;
using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core;
using Microsoft.SharePoint;
using Microsoft.Win32;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using OfficeDevPnP.Core;
using System.IO;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Reflection;
using System.Security;
using System.Security.Principal;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Media;
using System.Text;
using System.Web.Script.Serialization;
using AuthenticationManager = OfficeDevPnP.Core.AuthenticationManager;


namespace KizeoAndSharepoint_wizard
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class Step1 : Window
    {
        public Step1()
        {
            InitializeComponent();

            Config config = new Config();

            config.KizeoConfig = new KizeoConfig();
            config.SharepointConfig = new SharepointConfig();
            config.FormsToSpLists = new ObservableCollection<FormToSpList>();
            config.FormsToSpLibraries = new ObservableCollection<FormToSpLibrary>();
            config.SpListsToExtLists = new ObservableCollection<SpListToExtList>();
            config.PeriodicExports = new ObservableCollection<PeriodicExport>();

            DataContext = config;

            string path = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "Kizeo");
            string filePath = Path.Combine(path, "sharepoint_kf_connector_config.json");

            if (System.IO.File.Exists(filePath))
            {
                using (var sr = new StreamReader(filePath))
                {
                    string jsonText = sr.ReadToEnd();
                    var tmpContext = JsonConvert.DeserializeObject<Config>(jsonText);

                    if (tmpContext == null)
                    {
                        sr.Close();
                        createFile(filePath, true);
                    }
                    else
                    {
                        DataContext = tmpContext;
                    }
                }
            }
            else
            {
                createFile(filePath);
            }
        }

        private void createFile(string filePath, bool exist = false)
        {
            string jsonText = JsonConvert.SerializeObject((Config)DataContext, Formatting.Indented);
            if (exist)
            {
                using (var sw = new StreamWriter(filePath, false))
                {
                    sw.Write(jsonText);
                }
            }
            else
            {
                using (FileStream fs = System.IO.File.Create(filePath))
                {
                    byte[] info = new UTF8Encoding(true).GetBytes(jsonText);
                    fs.Write(info, 0, info.Length);
                }
            }
        }

        // Test that TextBoxes aren't empty
        private bool TestTextBox()
        {
            return (!string.IsNullOrEmpty(txtKfUrl.Text) && !string.IsNullOrEmpty(txtToken.Text) && !string.IsNullOrEmpty(sp_client_id.Text) && !string.IsNullOrEmpty(sp_client_secret.Text) && !string.IsNullOrEmpty(sp_domain.Text));
        }

        // Initialize textbox frames
        // 1 - txtKFUrl  /  2 - txtUrlSharepoint / 3 - All
        private void InitializeTextBoxFrame(int numberBox)
        {
            // MessageBox.Show(txtUrlSharepoint.BorderBrush.ToString());
            SolidColorBrush mySolidColorBrush = new SolidColorBrush();
            mySolidColorBrush = (SolidColorBrush)(new BrushConverter().ConvertFrom("#ffabadb3"));
            if (numberBox == 1) txtKfUrl.Background = Brushes.Transparent;
            if (numberBox == 2) sp_domain.Background = Brushes.Transparent;
            if (numberBox == 3)
            {
                txtKfUrl.BorderBrush = mySolidColorBrush;
                sp_domain.BorderBrush = mySolidColorBrush;
                txtKfUrl.Background = Brushes.Transparent;
                sp_domain.Background = Brushes.Transparent;
            }
        }

        // Test URL Kizeo Forms
        private async void testUrlKizeoForms(object sender, System.EventArgs e)
        {

            try
            {
                var HttpClient = new HttpClient();
                HttpClient.BaseAddress = new Uri(txtKfUrl.Text);
                HttpClient.DefaultRequestHeaders.Accept.Clear();
                var testToken = await HttpClient.GetAsync(txtKfUrl.Text);
                txtKfUrl.Background = Brushes.Transparent;
                CheckButton.IsEnabled = true;
                NextButton.IsEnabled = true;
            }
            catch (Exception)
            {
                txtKfUrl.Background = Brushes.Red;
                CheckButton.IsEnabled = false;
                NextButton.IsEnabled = false;
            }
        }


        private async void ButtonSuivant_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (!TestTextBox()) return;
                if (await TestKfApi() && TestSharePointConnection())
                {
                    ((Config)DataContext).SharepointConfig.Context = new AuthenticationManager().GetAppOnlyAuthenticatedContext(sp_domain.Text, sp_client_id.Text, sp_client_secret.Text);
                    Hide();
                    var step2 = new Step2();
                    step2.DataContext = DataContext;
                    step2.PreviousWindow = this;
                    step2.Show();
                }
                else
                {
                    MessageBox.Show("Please check Sharepoint and Kizeo Forms connections before");
                }
            }
            catch (Exception ee)
            {
                MessageBox.Show(ee.ToString());
            }

        }

        private async void ButtonTestConnection_Click(object sender, RoutedEventArgs e)
        {

            InitializeTextBoxFrame(3);
            if (TestTextBox())
            {
                var sharepointResult = TestSharePointConnection();
                var kfResult = await TestKfApi();
                var message = "Kizeo Forms : " + ((kfResult) ? "Successful" : "Failed") + "\n";
                message = message + "Sharepoint : " + ((sharepointResult) ? "Successful" : "Failed") + "\n";
                MessageBox.Show(message);
            }
            else
            {
                MessageBox.Show("Please fill all fields before.");
            }
        }


        private void ButtonAnnuler_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void MenuItemImporter_Click(object sender, RoutedEventArgs e)
        {
            var fileBrowser = new OpenFileDialog { Filter = "Json File (.json)|*.json| All Files (*.*)|*.*", FilterIndex = 1 };

            if (fileBrowser.ShowDialog() ?? false)
            {

                using (var sr = new StreamReader(fileBrowser.FileName))
                {

                    string jsonText = sr.ReadToEnd();
                    DataContext = JsonConvert.DeserializeObject<Config>(jsonText);
                }

            }
        }

        private bool TrySharePointConnection(string url, string clientId, string clientSecret)
        {
            try
            {
                using (var Context = new AuthenticationManager().GetAppOnlyAuthenticatedContext(url, clientId, clientSecret))
                {
                    var web = Context.Web;
                    try
                    {
                        Context.Load(web);
                        Context.ExecuteQuery();
                        return true;
                    }
                    catch (Exception)
                    {
                        return false;
                    }
                };
            }
            catch (Exception e)
            {
                return false;
            }

        }

        public bool TestSharePointConnection()
        {
            if (!sp_domain.Text.StartsWith("https://"))
            {
                MessageBox.Show("SharePoint's url must start with https://");
                return false;
            }

            var connected = TrySharePointConnection(sp_domain.Text, sp_client_id.Text, sp_client_secret.Text);
            if (!connected)
            {
                MessageBox.Show("Invalid credentials");
            }
            return connected;
        }


        public async Task<bool> TestKfApi()
        {
            try
            {
                var HttpClient = new HttpClient();
                HttpClient.BaseAddress = new Uri(txtKfUrl.Text);
                HttpClient.DefaultRequestHeaders.Accept.Clear();
                HttpClient.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
                HttpClient.DefaultRequestHeaders.TryAddWithoutValidation("Authorization", txtToken.Text);
                ((Config)DataContext).KizeoConfig.HttpClient = HttpClient;
                var testToken = await HttpClient.GetAsync($"{txtKfUrl.Text}/rest/v3/testapi/sharepoint");

                return testToken.IsSuccessStatusCode;
            }
            catch (Exception e)
            {
                MessageBox.Show("Couldn't connect to Kizeo Forms. Check both url and token.");
            }
            finally
            {
                txtKfUrl.Focus();
                txtKfUrl.SelectionStart = 0;
                txtKfUrl.SelectionLength = txtKfUrl.Text.Length;
            }
            return false;
        }

        private void MenuItem_Click(object sender, RoutedEventArgs e)
        {
            MessageBox.Show("Version " + GetType().Assembly.GetName().Version.ToString());
        }
    }
}
