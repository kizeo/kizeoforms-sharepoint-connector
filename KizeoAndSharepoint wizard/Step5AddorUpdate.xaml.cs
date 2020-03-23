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

        private void ChangeElementsState(bool state)
        {
            if (cbExcelList.Text.Equals("Disable")) ;
            cbExcelList.IsEnabled = state;
            txtExcelList.IsEnabled = (int)((ComboBox)cbExcelList).SelectedValue == 0 ? false : state;
            cbExcelListCustom.IsEnabled = state;
            txtExcelListCustom.IsEnabled = (int)((ComboBox)cbExcelListCustom).SelectedValue == 0 ? false : state;
            cbCsv.IsEnabled = state;
            txtCsv.IsEnabled = (int)((ComboBox)cbCsv).SelectedValue == 0 ? false : state;
            cbCsvCustom.IsEnabled = state;
            txtCsvCustom.IsEnabled = (int)((ComboBox)cbCsvCustom).SelectedValue == 0 ? false : state;
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
          /*  if (LibraryExists())
                ChangeElementsState(true);
            else
                ChangeElementsState(false);
*/
        }

        private bool LibraryExists()
        {
            //b69b051e-e649-4841-b39d-610f9d5d2cd0
            if (
             !string.IsNullOrEmpty(lib_sp.Text)
             && (
                 !lib_sp.Text.Equals("00000000-0000-0000-0000-000000000000")
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
                catch (Exception e) { lib_sp.BorderBrush = System.Windows.Media.Brushes.Red; Console.WriteLine(e.ToString()); MessageBox.Show("Impossible de charger la librairie SharePoint.\nVeuillez vérifier que:\n- Le client utilisé est bien associé l'url SharePoint renseigné\n- L'url est correct\n- L'ID de la librairie est correct");
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

        private void TextBox_DataContextChanged(object sender, DependencyPropertyChangedEventArgs e)
        {

        }
    }

}
