using KizeoAndSharepoint_wizard.Models;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
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
    /// Interaction logic for Step3AddMapping.xaml
    /// </summary>
    public partial class Step3AddMetaData : Window
    {
        public ObservableCollection<DataMapping> Metadata { get; set; }

        public Step3AddMetaData()
        {
            InitializeComponent();

        }

        private void ButtonAjouter_Click(object sender, RoutedEventArgs e)
        {
            var item = (FormToSpLibrary)DataContext;
            item.MetaData.Add(new DataMapping { KfColumnSelector = TxtKfColumnId.Text, SpColumnId = (string)cbSpColumnId.SelectedValue, SpecialType = cbSpecialType.Text });

        }

        private void ButtonUpdate_Click(object sender, RoutedEventArgs e)
        {
            if (lvMapping.SelectedItem != null)
            {
                var item = (DataMapping)lvMapping.SelectedItem;

                item.KfColumnSelector = TxtKfColumnId.Text;
                item.SpColumnId = (string)cbSpColumnId.SelectedValue;
                item.SpecialType = cbSpecialType.Text;
                lvMapping.ItemsSource = null;
                lvMapping.Items.Clear();
                lvMapping.ItemsSource = ((FormToSpLibrary)DataContext).MetaData;
            }
        }
        

    private void ButtonValider_Click(object sender, RoutedEventArgs e)
        {

            this.Hide();
        }

        private void lvMapping_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (lvMapping.SelectedItem != null)
            {
                var item = (DataMapping)lvMapping.SelectedItem;
                TxtKfColumnId.Text = item.KfColumnSelector;
                cbSpColumnId.SelectedValue = item.SpColumnId;
                cbSpecialType.Text = item.SpecialType;
            }
        }

        private void ButtonDelete_Click(object sender, RoutedEventArgs e)
        {

            if (lvMapping.SelectedItem != null)
            {
                var item = (DataMapping)lvMapping.SelectedItem;
                var item2 = ((FormToSpLibrary)DataContext);
                item2.MetaData.Remove(item);

                lvMapping.ItemsSource = null;
                lvMapping.Items.Clear();
                lvMapping.ItemsSource = item2.MetaData;

            }
        }

        
    }
}
