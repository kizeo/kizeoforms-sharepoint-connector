using KizeoAndSharepoint_wizard.Models;
using Microsoft.Win32;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;

namespace KizeoAndSharepoint_wizard
{
    /// <summary>
    /// Interaction logic for Step4.xaml
    /// </summary>
    public partial class Step4 : Window
    {
        public Step4()
        {
            InitializeComponent();
        }

        public Step3 PreviousWindow { get; internal set; }

        private void Window_Activated(object sender, EventArgs e)
        {
            RefreshList();
        }
        
        private void RefreshList() {
            var itemSource = lvSpListsToExLists.ItemsSource;
            lvSpListsToExLists.ItemsSource = null;
            lvSpListsToExLists.ItemsSource = itemSource;
        }

        private void ButtonSuivant_Click(object sender, RoutedEventArgs e)
        {
            var step5 = new Step5();
            step5.DataContext = DataContext;
            step5.Show();
            step5.PreviousWindow = this;
            this.Hide();
        }

        private void Button_Ajouter_Click(object sender, RoutedEventArgs e)
        {
            var step4AddOrUpdate = new Step4AddOrUpdate();
            step4AddOrUpdate.DataContext = new SpListToExtList();
            step4AddOrUpdate.SpListToExtLists = ((Config)DataContext).SpListsToExtLists;
            step4AddOrUpdate.OpnedForNewItem = true;
            step4AddOrUpdate.Context = ((Config)DataContext).SharepointConfig.Context;
            step4AddOrUpdate.Show();
            step4AddOrUpdate.FillCbBox();
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

                MessageBox.Show("Config file imported");
            }
        }

        private void MenuItemSave_Click(object sender, RoutedEventArgs e)
        {

        }

        private void ButtonPrevious_Click(object sender, RoutedEventArgs e)
        {
            PreviousWindow.Show();
            this.Hide();
        }

        private void ButtonUpdate_Click(object sender, RoutedEventArgs e)
        {
            if (lvSpListsToExLists.SelectedItem != null)
            {
                var step4AddOrUpdate = new Step4AddOrUpdate();
                var item = (SpListToExtList)lvSpListsToExLists.SelectedItem;
                step4AddOrUpdate.SpListToExtLists = ((Config)DataContext).SpListsToExtLists;
                step4AddOrUpdate.DataContext = item;
                step4AddOrUpdate.OpnedForNewItem = false;
                step4AddOrUpdate.Context = ((Config)DataContext).SharepointConfig.Context;
                step4AddOrUpdate.Show();
                step4AddOrUpdate.FillCbBox();
            }
        }

        private void ButtonDelete_Click(object sender, RoutedEventArgs e)
        {
            SpListToExtList item = (SpListToExtList)lvSpListsToExLists.SelectedItem;
            if (item != null && MessageBox.Show("Are you sure ?", "Question", MessageBoxButton.YesNo, MessageBoxImage.Warning) == MessageBoxResult.Yes)
            {
                ((Config)DataContext).SpListsToExtLists.Remove(item);
            }

        }
    }
}
