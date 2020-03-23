using KizeoAndSharepoint_wizard.Models;
using System;
using System.Collections.Generic;
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
using System.Collections.ObjectModel;
using Microsoft.Win32;
using System.IO;
using Newtonsoft.Json;

namespace KizeoAndSharepoint_wizard
{
    /// <summary>
    /// Interaction logic for Step2.xaml
    /// </summary>
    public partial class Step2 : Window
    {
        public Window PreviousWindow { get; set; }

        public Step2()
        {
            InitializeComponent();
        }

        private void Window_Activated(object sender, EventArgs e)
        {
            RefreshList();
        }
        
        private void RefreshList() {
            var itemSource = lvFormsToSpLists.ItemsSource;
            lvFormsToSpLists.ItemsSource = null;
            lvFormsToSpLists.ItemsSource = itemSource;
        }

        private void MenuItemSave_Click(object sender, RoutedEventArgs e)
        {
            MessageBox.Show("Vous avez appuyé sur Save");
            MessageBox.Show("Sauvegarde réussite");
        }

        private void ButtonSuivant_Click(object sender, RoutedEventArgs e)
        {
            this.Hide();
            var step3 = new Step3();
            step3.DataContext = DataContext;
            step3.PreviousWindow = this;
            step3.Show();
        }


        private void Button_Ajouter_Click(object sender, RoutedEventArgs e)
        {
            var step2AddOrUpdate = new Step2AddOrUpdate();
            step2AddOrUpdate.DataContext = new FormToSpList { DataMapping = new ObservableCollection<DataMapping>() };
            step2AddOrUpdate.Context = ((Config)DataContext).SharepointConfig.Context;
            step2AddOrUpdate.FormTospLists = ((Config)DataContext).FormsToSpLists;
            step2AddOrUpdate.OpnedForNewItem = true;
            step2AddOrUpdate.Show();
            step2AddOrUpdate.FillCbBox();
        }

        private void ButtonPrevious_Click(object sender, RoutedEventArgs e)
        {
            PreviousWindow.Show();
            this.Hide();
        }

        private void ButtonUpdate_Click(object sender, RoutedEventArgs e)
        {
            if (lvFormsToSpLists.SelectedItem != null)
            {
                FormToSpList item = (FormToSpList)lvFormsToSpLists.SelectedItem;
                item.DataMapping = item.DataMapping ?? new ObservableCollection<DataMapping>();
                var step2AddOrUpdate = new Step2AddOrUpdate();
                step2AddOrUpdate.DataContext = item;
                step2AddOrUpdate.OpnedForNewItem = false;
                step2AddOrUpdate.Context = ((Config)DataContext).SharepointConfig.Context;
                step2AddOrUpdate.Show();
                step2AddOrUpdate.FillCbBox();

            }

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

        private void ButtonDelete_Click(object sender, RoutedEventArgs e)
        {
            FormToSpList item = (FormToSpList)lvFormsToSpLists.SelectedItem;
            if ((FormToSpList)lvFormsToSpLists.SelectedItem != null && MessageBox.Show("Are you sure ? ", "Question", MessageBoxButton.YesNo, MessageBoxImage.Warning) == MessageBoxResult.Yes)
            {
                ((Config)DataContext).FormsToSpLists.Remove(item);
            }

        }
    }
}
