using KizeoAndSharepoint_wizard.Models;
using Microsoft.SharePoint.Client;
using System;
using System.Collections.ObjectModel;
using System.Drawing;
using System.Windows;
using System.Windows.Controls;

namespace KizeoAndSharepoint_wizard
{
    /// <summary>
    /// Interaction logic for Step5AddorUpdate.xaml
    /// </summary>
    public partial class Step5AddorUpdate : Window
    {
        public ObservableCollection<PeriodicExport> PeriodicExports { get; set; }
        public bool OpnedForNewItem { get; set; }
        public ClientContext Context;


        public Step5AddorUpdate()
        {
            InitializeComponent();

        }

        public void InitComboboxes()
        {

            cbExcelList.SelectedValue = ((PeriodicExport)DataContext).ExcelListPeriod;
            cbExcelListCustom.SelectedValue = ((PeriodicExport)DataContext).ExcelListCustomPeriod;
            cbCsv.SelectedValue = ((PeriodicExport)DataContext).CsvPeriod;
            cbCsvCustom.SelectedValue = ((PeriodicExport)DataContext).CsvCustomPeriod;

        }
        private void sp_library_id_changed(object sender, TextChangedEventArgs e)
        {
        
        }

        private bool LibraryExists()
        {
            const string list_mask = "00000000-0000-0000-0000-000000000000";
            if (
             !string.IsNullOrEmpty(lib_sp.Text)
             && (
                 !lib_sp.Text.Equals(list_mask)
             ) && lib_sp.Text.Length == 36
         )
            {
                try
                {

                    var spList = Context.Web.Lists.GetById(new Guid(lib_sp.Text));

                    Context.Load(spList);
                    CamlQuery query = CamlQuery.CreateAllItemsQuery();
                    Microsoft.SharePoint.Client.ListItemCollection items = spList.GetItems(query);
                    Context.ExecuteQuery();
                    lib_sp.BorderBrush = System.Windows.Media.Brushes.ForestGreen;
                    return true;

                }
                catch (Exception e) { lib_sp.BorderBrush = System.Windows.Media.Brushes.Red;
                    MessageBox.Show("SharePoint's list couldn't be loaded.\nPlease check out that:\n- The client is associated to the SharePoint url you entered.\n- The url is well formated");
                }
            }
            return false;
        }
        private void ButtonValider_Click(object sender, RoutedEventArgs e)
        {
            if (LibraryExists())
            {
                if (OpnedForNewItem)
                {
                    PeriodicExports.Add((PeriodicExport)DataContext);
                }

                this.Hide();
            }
      
        }

        private void Cb_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            var leSendeur = (ComboBox)sender;

            switch (leSendeur.Name)
            {
                case "cbExcelList":
                    txtExcelList.IsEnabled = ((int)leSendeur.SelectedValue != 0);
                    ((PeriodicExport)DataContext).ExcelListPeriod = (int)leSendeur.SelectedValue;
                    break;
                case "cbExcelListCustom":
                    txtExcelListCustom.IsEnabled = ((int)leSendeur.SelectedValue != 0);
                    ((PeriodicExport)DataContext).ExcelListCustomPeriod = (int)leSendeur.SelectedValue;
                    break;
                case "cbCsv":
                    txtCsv.IsEnabled = ((int)leSendeur.SelectedValue != 0);
                    ((PeriodicExport)DataContext).CsvPeriod = (int)leSendeur.SelectedValue;
                    break;
                case "cbCsvCustom":
                    txtCsvCustom.IsEnabled = ((int)leSendeur.SelectedValue != 0);
                    ((PeriodicExport)DataContext).CsvCustomPeriod = (int)leSendeur.SelectedValue;
                    break;
                default:
                    break;
            }

        }

    }

}
