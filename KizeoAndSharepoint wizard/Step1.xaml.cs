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
                    DataContext = JsonConvert.DeserializeObject<Config>(jsonText);
                }
            }



        }

        // Test that TextBoxes aren't empty
        private bool TestTextBox()
        {
            return (!string.IsNullOrEmpty(txtKfUrl.Text) && !string.IsNullOrEmpty(txtToken.Text) && !string.IsNullOrEmpty(sp_client_id.Text) && !string.IsNullOrEmpty(sp_client_secret.Text) && !string.IsNullOrEmpty(sp_tenant_id.Text) && !string.IsNullOrEmpty(sp_domain.Text));
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

            var HttpClient = new HttpClient();
            HttpClient.BaseAddress = new Uri(txtKfUrl.Text);
            HttpClient.DefaultRequestHeaders.Accept.Clear();

            try
            {
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

            if (!TestTextBox()) return;
            if (await TestKfApi() && TestSharePointConnection())
            {
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

        private string ExtractDomainNameFromURL(string Url)
        {
            return System.Text.RegularExpressions.Regex.Replace(
                Url,
                @"^([a-zA-Z]+:\/\/)?([^\/]+)\/.*?$",
                "$2"
            );
        }


        private string TrySharePointConnection()
        {
          
            string access_url = $"https://accounts.accesscontrol.windows.net/{sp_tenant_id.Text }/tokens/OAuth/2";
            const string resource_id = "00000003-0000-0ff1-ce00-000000000000";
            try
            {
                var request = (HttpWebRequest)WebRequest.Create(access_url);

                var postData = "grant_type=client_credentials";
                postData += $"&client_id={ sp_client_id.Text}@{sp_tenant_id.Text}";
                postData += "&client_secret=" + sp_client_secret.Text;
                postData += $"&resource={resource_id}/{ExtractDomainNameFromURL(sp_domain.Text)}@{sp_tenant_id.Text}";

                byte[] data = Encoding.UTF8.GetBytes(postData);

                request.Method = "POST";
                request.ContentType = "application/x-www-form-urlencoded";
                request.ContentLength = data.Length;

                using (var stream = request.GetRequestStream())
                {
                    stream.Write(data, 0, data.Length);

                }
                var response = (HttpWebResponse)request.GetResponse();

                var responseString = new StreamReader(response.GetResponseStream()).ReadToEnd();


                dynamic json = JsonConvert.DeserializeObject(responseString);
                return json.access_token;
            }
            catch (System.Net.WebException e)
            {
                MessageBox.Show("Une ou plusieurs informations de SharePoint sont fausses");
            }
            catch (Exception)
            {
                return "undefined";
            }
            return "undefined";
        }

        private ClientContext GetClientContextWithAccessToken(string targetUrl, string accessToken)
        {
            Uri targetUri = new Uri(targetUrl);

            ClientContext clientContext = new ClientContext(targetUrl);


            clientContext.AuthenticationMode = ClientAuthenticationMode.Anonymous;
            clientContext.FormDigestHandlingEnabled = false;
            clientContext.ExecutingWebRequest +=
                delegate (object oSender, WebRequestEventArgs webRequestEventArgs)
                {
                    webRequestEventArgs.WebRequestExecutor.RequestHeaders["Authorization"] =
                        "Bearer " + accessToken;
                };

            return clientContext;
        }
        private bool IsTokenCorrect(string token)
        {
            try
            {
                HttpWebRequest endpointRequest = (HttpWebRequest)HttpWebRequest.Create($"{sp_domain.Text}/_api/web/lists");
                endpointRequest.Method = "GET";
                endpointRequest.Accept = "application/json;odata=verbose";
                endpointRequest.Headers.Add("Authorization",
                  "Bearer " + token);
                HttpWebResponse endpointResponse =
                  (HttpWebResponse)endpointRequest.GetResponse();
                string json;
                using (var sr = new StreamReader(endpointResponse.GetResponseStream()))
                {
                    json = sr.ReadToEnd();
                }
                dynamic data = JsonConvert.DeserializeObject(json);

                return true;
            }
            catch (Exception e)
            {
                MessageBox.Show(e.ToString());

                return false;
            }

        }

        public bool TestSharePointConnection()
        {
            if (!sp_domain.Text.StartsWith("https://"))
            {
                MessageBox.Show("L'url du domaine doit commencer par https://");
                return false;

            }

            var token = TrySharePointConnection();
            if (!token.Equals("undefined"))
            {
                if (IsTokenCorrect(token))
                {
                    try
                    {
                        ClientContext Context = GetClientContextWithAccessToken(sp_domain.Text, token);
                        ((Config)(DataContext)).SharepointConfig.Context = Context;
                        return true;
                    }
                    catch (Exception)
                    {
                        return false;
                    }

                }

            }
            return false;
        }


        public async Task<bool> TestKfApi()
        {
            try
            {
                var HttpClient = new HttpClient();
                HttpClient.BaseAddress = new Uri(txtKfUrl.Text);
                HttpClient.DefaultRequestHeaders.Accept.Clear();
                HttpClient.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
                HttpClient.DefaultRequestHeaders.Add("Authorization", txtToken.Text);

                ((Config)(DataContext)).KizeoConfig.HttpClient = HttpClient;
                var testToken = await HttpClient.GetAsync($"{txtKfUrl.Text}/rest/v3/testapi/sharepoint");

                return testToken.IsSuccessStatusCode;
            }
            catch (HttpRequestException)
            {
                txtKfUrl.BorderBrush = Brushes.Red;
                MessageBox.Show("URL Kizeo Forms doesn't exist.");
            }
            catch (System.UriFormatException)
            {
                txtKfUrl.BorderBrush = Brushes.Red;
                MessageBox.Show("Please enter a valid URL Kizeo Forms.\n(example : https://www.adresse_serveur.ext)");
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
