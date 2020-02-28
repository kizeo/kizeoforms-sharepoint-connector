using KizeoAndSharepoint_wizard.Models;
using Microsoft.Win32;
using Newtonsoft.Json;
using System.Collections.ObjectModel;
using System;
using System.IO;
using System.Windows;

namespace KizeoAndSharepoint_wizard
{
    /// <summary>
    /// Interaction logic for Step5.xaml
    /// </summary>
    public partial class Step5 : Window
    {
        public Step4 PreviousWindow { get; internal set; }
        public ObservableCollection<PeriodicExportsChoices> PeriodicChoices { get; set; }

        public Step5()
        {
            InitializeComponent();
            PeriodicChoices = new ObservableCollection<PeriodicExportsChoices> {
            new PeriodicExportsChoices { Id=0,Name="Disabled"},
            new PeriodicExportsChoices { Id=1,Name="Daily"},
            new PeriodicExportsChoices { Id=2,Name="Weekly"},
            new PeriodicExportsChoices { Id=3,Name="D & W"}
            };
        }
        private void ButtonAnnuler_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }
        private void Button_Ajouter_Click(object sender, RoutedEventArgs e)
        {
            var step5AddOrUpdate = new Step5AddorUpdate();
            step5AddOrUpdate.PeriodicExports = ((Config)DataContext).PeriodicExports;
            step5AddOrUpdate.Context = ((Config)DataContext).SharepointConfig.Context;
            step5AddOrUpdate.OpnedForNewItem = true;
            var item = new PeriodicExport();
            item.PeriodicChoices = PeriodicChoices;
            step5AddOrUpdate.DataContext = item;
            step5AddOrUpdate.Show();
        }


        private void Window_Activated(object sender, EventArgs e)
        {
            RefreshList();
        }

        private void RefreshList()
        {
            var itemSource = lvPeriodicExports.ItemsSource;
            lvPeriodicExports.ItemsSource = null;
            lvPeriodicExports.ItemsSource = itemSource;
        }

        private void ButtonSaveACopy_Click(object sender, RoutedEventArgs e)
        {
            if (((Config)DataContext).PeriodicExports == null)
            {
                ((Config)DataContext).PeriodicExports=  new ObservableCollection<PeriodicExport>();
            }
            var x = (Config)DataContext;
            x.SharepointConfig.Context = null;

            string jsonText = JsonConvert.SerializeObject((Config)DataContext, Formatting.Indented);

            var fileBrowser = new SaveFileDialog { Filter = "Json File (.json)|*.json| All Files (*.*)|*.*", FilterIndex = 1 };

            if (fileBrowser.ShowDialog() ?? false)
            {

                using (var sw = new StreamWriter(fileBrowser.FileName, false))
                {
                    sw.Write(jsonText);
                }

                MessageBox.Show("Copy saved");
            }
        }

        private void ButtonTerminer_Click(object sender, RoutedEventArgs e)
        {

            if (((Config)DataContext).PeriodicExports == null)
            {
                ((Config)DataContext).PeriodicExports = new ObservableCollection<PeriodicExport>();
            }
            var x = (Config)DataContext;
            x.SharepointConfig.Context = null;

            string path = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "Kizeo");
            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
            }

            string jsonText = JsonConvert.SerializeObject((Config)DataContext, Formatting.Indented);

            using (var sw = new StreamWriter(Path.Combine(path, "sharepoint_kf_connector_config.json"), false))
            {
                sw.Write(jsonText);
            }

            MessageBox.Show("File saved to : " + Path.Combine(path, "sharepoint_kf_connector_config.json"));
            this.Close();


            var fileBrowser = new SaveFileDialog { Filter = "Json File (.json)|*.json| All Files (*.*)|*.*", FilterIndex = 1 };
        }

        private void ButtonPrevious_Click(object sender, RoutedEventArgs e)
        {
            PreviousWindow.Show();
            this.Hide();
        }

        private void ButtonUpdate_Click(object sender, RoutedEventArgs e)
        {
            if (lvPeriodicExports.SelectedItem != null)
            {
                var step5AddOrUpdate = new Step5AddorUpdate();
                var item = (PeriodicExport)lvPeriodicExports.SelectedItem;
                item.PeriodicChoices = PeriodicChoices;
                step5AddOrUpdate.PeriodicExports = ((Config)DataContext).PeriodicExports;
                step5AddOrUpdate.Context = ((Config)DataContext).SharepointConfig.Context;
                step5AddOrUpdate.DataContext = item;
                step5AddOrUpdate.OpnedForNewItem = false;
                step5AddOrUpdate.InitComboboxes();
                step5AddOrUpdate.Show();

            }



        }


        private void ButtonDelete_Click(object sender, RoutedEventArgs e)
        {
            PeriodicExport item = (PeriodicExport)lvPeriodicExports.SelectedItem;
            if (lvPeriodicExports.SelectedItem != null && MessageBox.Show("Are you sure ?", "Question", MessageBoxButton.YesNo, MessageBoxImage.Warning) == MessageBoxResult.Yes)
            {
                ((Config)DataContext).PeriodicExports.Remove(item);
            }

        }



    }
}
