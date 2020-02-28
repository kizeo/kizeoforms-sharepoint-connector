using KizeoAndSharepoint_wizard.Models;
using Microsoft.Win32;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
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
    /// Interaction logic for Step3.xaml
    /// </summary>
    public partial class Step3 : Window
    {
        public Step2 PreviousWindow { get; internal set; }

        public Step3()
        {
            InitializeComponent();
            
        }
        private void ButtonAnnuler_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }


        private void Window_Activated(object sender, EventArgs e)
        {
            RefreshList();
        }
        
        private void RefreshList() {
            var itemSource = lvFormsToSplibraries.ItemsSource;
            lvFormsToSplibraries.ItemsSource = null;
            lvFormsToSplibraries.ItemsSource = itemSource;
        }

 
        private void ButtonSuivant_Click(object sender, RoutedEventArgs e)
        {
            var step4 = new Step4();
            if (((Config)DataContext).FormsToSpLibraries == null)
            {
                ((Config)DataContext).FormsToSpLibraries = new ObservableCollection<FormToSpLibrary>();
            }
            step4.DataContext = DataContext;
            step4.Show();
            step4.PreviousWindow = this;
            this.Hide();
        }


        private void Button_Ajouter_Click(object sender, RoutedEventArgs e)
        {
            var step3AddOrUpdate = new Step3AddOrUpdate();
            step3AddOrUpdate.DataContext = new FormToSpLibrary { Exports = new ObservableCollection<Export>() , MetaData = new ObservableCollection<DataMapping>() };
            step3AddOrUpdate.OpnedForNewItem = true;
            step3AddOrUpdate.FormToSpLibraries = ((Config)DataContext).FormsToSpLibraries;
            step3AddOrUpdate.Context = ((Config)DataContext).SharepointConfig.Context;
            step3AddOrUpdate.Show();
            
        }

        private void ButtonPrevious_Click(object sender, RoutedEventArgs e)
        {
            PreviousWindow.Show();
            this.Hide();
        }

        private void ButtonUpdate_Click(object sender, RoutedEventArgs e)
        {
            if ((FormToSpLibrary)lvFormsToSplibraries.SelectedItem != null)
            {
                FormToSpLibrary item = (FormToSpLibrary)lvFormsToSplibraries.SelectedItem;
                item.Exports = item.Exports ?? new ObservableCollection<Export>();
                item.MetaData = item.MetaData ?? new ObservableCollection<DataMapping>();
                var step3AddOrUpdate = new Step3AddOrUpdate();
                step3AddOrUpdate.DataContext = item;
                step3AddOrUpdate.OpnedForNewItem = false;
                step3AddOrUpdate.Context = ((Config)DataContext).SharepointConfig.Context;
                step3AddOrUpdate.Show();
            }

        }


        private void ButtonDelete_Click(object sender, RoutedEventArgs e)
        {
            FormToSpLibrary item = (FormToSpLibrary)lvFormsToSplibraries.SelectedItem;
            if ((FormToSpLibrary)lvFormsToSplibraries.SelectedItem !=null && MessageBox.Show("Are you sure ? ", "Question", MessageBoxButton.YesNo, MessageBoxImage.Warning) == MessageBoxResult.Yes)
            {
                ((Config)DataContext).FormsToSpLibraries.Remove(item);
            }

        }
    }
}
