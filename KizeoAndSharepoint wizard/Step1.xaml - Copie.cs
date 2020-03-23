using KizeoAndSharepoint_wizard.Models;
using Microsoft.SharePoint.Client;
using Microsoft.Win32;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Reflection;
using System.Security;
using System.Security.Principal;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Media;


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

            config.KizeoConfig = new KizeoConfig ();
            config.SharepointConfig = new SharepointConfig ();
            config.FormsToSpLists = new ObservableCollection<FormToSpList>();
            config.FormsToSpLibraries = new ObservableCollection<FormToSpLibrary>();
            config.SpListsToExtLists = new ObservableCollection<SpListToExtList>();
            config.PeriodicExports = new ObservableCollection<PeriodicExport>();
            
            DataContext = config;

            string path = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "Kizeo");
            string filePath = Path.Combine(path, "sharepoint_kf_connector_config.json");

            if (System.IO.File.Exists(filePath)) {
                using (var sr = new StreamReader(filePath))
                {
                    string jsonText = sr.ReadToEnd();
                    DataContext = JsonConvert.DeserializeObject<Config>(jsonText);
                }
            }



        }

        private void MenuItemSave_Click(object sender, RoutedEventArgs e)
        {
            // MessageBox.Show("Vous avez appuyé sur Save");
            // MessageBox.Show("Sauvegarde réussite");
        }

        // ↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓

        // Test that TextBoxes aren't empty - Patrick & Mehdi
        private bool TestTextBox()
        {
            return (!string.IsNullOrEmpty(txtKfUrl.Text) && !string.IsNullOrEmpty(txtToken.Text) && !string.IsNullOrEmpty(txtPwd.Text) && !string.IsNullOrEmpty(txtLogin.Text) && !string.IsNullOrEmpty(txtUrlSharepoint.Text));
        }

        // Initialize textbox frames - Patrick
        // 1 - txtKFUrl  /  2 - txtUrlSharepoint / 3 - All
        private void InitializeTextBoxFrame(int numberBox)
        {
            // MessageBox.Show(txtUrlSharepoint.BorderBrush.ToString());
            SolidColorBrush mySolidColorBrush = new SolidColorBrush();
            mySolidColorBrush = (SolidColorBrush)(new BrushConverter().ConvertFrom("#ffabadb3"));
            if (numberBox == 1) txtKfUrl.Background = Brushes.Transparent;
            if (numberBox == 2) txtUrlSharepoint.Background = Brushes.Transparent;
            if (numberBox == 3)
            {
                txtKfUrl.BorderBrush = mySolidColorBrush;
                txtUrlSharepoint.BorderBrush = mySolidColorBrush;
                txtKfUrl.Background = Brushes.Transparent;
                txtUrlSharepoint.Background = Brushes.Transparent;
            }
        }

        // Test URL
        private async Task<bool> testURL(string adressUrl)
        {
            var HttpClient = new HttpClient();
            HttpClient.BaseAddress = new Uri(adressUrl);
            HttpClient.DefaultRequestHeaders.Accept.Clear();
            MessageBox.Show(adressUrl);
            
            try
            {
                var testToken = await HttpClient.GetAsync(adressUrl);
                MessageBox.Show(adressUrl + " : Vrai");
                return true;
            }catch(Exception e)
            {
                MessageBox.Show(adressUrl + " : Faux");
                return false;
            }
        }


        // Test URL Kizeo Forms
        private async void testUrlKizeoForms()
        {
            if (testURL(txtKfUrl.Text).Result)
            {
                txtKfUrl.Background = Brushes.Transparent;
                if(txtUrlSharepoint.Background == Brushes.Transparent)
                {
                    CheckButton.IsEnabled = true;
                    NextButton.IsEnabled = true;
                }
                else
                {
                    txtKfUrl.Background = Brushes.Red;
                    CheckButton.IsEnabled = false;
                    NextButton.IsEnabled = false;
                }
            }
        }


            // test URL SharePoint on lost Focus - Patrick
            private void txtUrlSharepoint_LostFocus(object sender, System.EventArgs e)
        {
            if (testURL(txtUrlSharepoint.Text).Result)
            {
                txtUrlSharepoint.Background = Brushes.Transparent;
                if (txtKfUrl.Background == Brushes.Transparent)
                {
                    CheckButton.IsEnabled = true;
                    NextButton.IsEnabled = true;
                }
                else
                {
                    txtUrlSharepoint.Background = Brushes.Red;
                    CheckButton.IsEnabled = false;
                    NextButton.IsEnabled = false;
                }
            }
        }

        // Test URL Kizeo Forms on lost Focus - Patrick
        private void txtKfUrl_Leave(object sender, System.EventArgs e)
        {
            testUrlKizeoForms();
            // txtKfUrl.Background = Brushes.Red;
        } 

        // ↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑

        private async void ButtonSuivant_Click(object sender, RoutedEventArgs e)
        {
            if (!TestTextBox()) return;
            if (await TestKfApi() && TestSharePointConnection())
            {
                ((Config)DataContext).SharepointConfig.Context = new ClientContext(txtUrlSharepoint.Text);
                ((Config)DataContext).SharepointConfig.Context.Credentials = new SharePointOnlineCredentials(txtLogin.Text, ConvertToSecureString(txtPwd.Text));

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
                var message = "Kizeo Forms : " + ((kfResult == true) ? "Successful" : "Failed") + "\n";
                message = message + "Sharepoint : " + ((sharepointResult == true) ? "Successful" : "Failed") + "\n";
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

                // MessageBox.Show("File opened");
            }
        }



        public bool TestSharePointConnection()
        {
            try
            {
            ClientContext Context = new ClientContext(txtUrlSharepoint.Text);
            Context.Credentials = new SharePointOnlineCredentials(txtLogin.Text, ConvertToSecureString(txtPwd.Text));

            ((Config)(DataContext)).SharepointConfig.Context = Context;

            var web = Context.Web;
                Context.Load(web);
                Context.ExecuteQuery();
                return true;
            }
            catch (System.UriFormatException)
            {
                txtUrlSharepoint.BorderBrush = Brushes.Red;
                MessageBox.Show("Please enter a valid URL SharePoint.\n(example : https://www.adresse_serveur.ext)");
            }
            catch (Exception)
            {
                txtUrlSharepoint.BorderBrush = Brushes.Red;
                MessageBox.Show("Can't access serveur SharePoint.");
            }
            finally
            {
                txtUrlSharepoint.Focus();
                txtUrlSharepoint.SelectionStart = 0;
                txtUrlSharepoint.SelectionLength = txtUrlSharepoint.Text.Length;
            }
            return false;

        }


        public static SecureString ConvertToSecureString(string password)
        {
            var securePassword = new SecureString();

            if (password == null)
                throw new ArgumentNullException("password");

            foreach (char c in password)
                securePassword.AppendChar(c);

            securePassword.MakeReadOnly();

            return securePassword;

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

        private void TxtUrlSharepoint_LostFocus(object sender, RoutedEventArgs e)
        {

        }
    }
}
