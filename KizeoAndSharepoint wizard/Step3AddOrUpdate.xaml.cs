using KizeoAndSharepoint_wizard.Models;
using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Runtime.Remoting.Contexts;
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
    /// Interaction logic for Step33.xaml
    /// </summary>
    public partial class Step3AddOrUpdate : Window
    {
        public ObservableCollection<FormToSpLibrary> FormToSpLibraries { get; set; }
        public bool OpnedForNewItem { get; internal set; }
        public ClientContext Context;

        public Step3AddOrUpdate()
        {
            InitializeComponent();
            spmeta.Visibility = Visibility.Hidden;


        }

        private void ButtonMetadata_Click(object sender, RoutedEventArgs e)
        {
            var metadata = new Step3AddMetaData();
            metadata.DataContext = ((FormToSpLibrary)DataContext);
            metadata.Show();
            try
            {
                /*      var spList = Context.Web.Lists.GetById(((FormToSpLibrary)DataContext).SpLibraryId);*/
                var spList = Context.Web.Lists.GetById(new Guid(spLib.Text));
                Context.Load(spList);
                CamlQuery query = CamlQuery.CreateAllItemsQuery();
                Microsoft.SharePoint.Client.ListItemCollection items = spList.GetItems(query);

                Context.Load(spList.Fields);
                Context.ExecuteQuery();
                var x = spList.Fields.Where(f => f.Hidden == false).Where(f => f.ReadOnlyField == false).Where(f => f.StaticName == "Title" || f.CanBeDeleted == true).ToList();
                metadata.cbSpColumnId.ItemsSource = x;

            }
            catch (Exception ee)
            {
                MessageBox.Show("SharePoint's list couldn't be loaded.\nPlease check out that:\n- The client is associated to the SharePoint url you entered.\n- The url is well formated");
            }

        }

        private void ButtonAjouterExport_Click(object sender, RoutedEventArgs e)
        {
            if (PreventInvalidCharacters()) return;
            var formToSpLibrary = (FormToSpLibrary)DataContext;
            if (formToSpLibrary.Exports == null)
            {
                formToSpLibrary.Exports = new ObservableCollection<Export>();

            }
            formToSpLibrary.Exports.Add(new Export { Id = txtIdExport.Text, ToInitialType = cbToInitial.IsChecked ?? false, InitialTypePath = txtInitialTypePath.Text, ToPdf = cbToPdf.IsChecked ?? false, PdfPath = txtPdfPath.Text });
        }

        private bool PreventInvalidCharacters()
        {
            if (txtInitialTypePath.Text.Contains("/") || txtInitialTypePath.Text.Contains("\\"))
            {
                txtInitialTypePath.BorderBrush = Brushes.Red;
                MessageBox.Show("Incorrect path for inital type path. Illegal characters: /, \\");
                return true;
            }
            return false;
        }

        private void ButtonValider_Click(object sender, RoutedEventArgs e)
        {
            const string list_mask = "00000000-0000-0000-0000-000000000000";

            if (
          !string.IsNullOrEmpty(spLib.Text)
          && (
              !spLib.Text.Equals(list_mask)
          ) && spLib.Text.Length == 36
      )
            {
                try
                {
                    var spList = Context.Web.Lists.GetById(new Guid(spLib.Text));
                    Context.Load(spList);
                    CamlQuery query = CamlQuery.CreateAllItemsQuery();
                    Microsoft.SharePoint.Client.ListItemCollection items = spList.GetItems(query);
                    Context.ExecuteQuery();
                    spLib.BorderBrush = Brushes.ForestGreen;
                    ok_btn.IsEnabled = true;
                    if (OpnedForNewItem)
                    {
                        FormToSpLibraries.Add((FormToSpLibrary)DataContext);
                    }

                    this.Hide();
                }
                catch (Exception)
                {
                    spLib.BorderBrush = Brushes.Red;
                    MessageBox.Show("Impossible de charger la librairie SharePoint.\nVeuillez vérifier que:\n- Le client utilisé est bien associé l'url SharePoint renseigné\n- L'url est correct\n- L'ID de la librairie est correct");

                }
            }
            else
            {
                spLib.BorderBrush = Brushes.Red;
                MessageBox.Show("Impossible de charger la librairie SharePoint.\nVeuillez vérifier que:\n- Le client utilisé est bien associé l'url SharePoint renseigné\n- L'url est correct\n- L'ID de la librairie est correct");

            }

        }

        private void ListView_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if ((Export)lvExports.SelectedItem != null)
            {
                txtInitialTypePath.BorderBrush = Brushes.Gray;
                var export = (Export)lvExports.SelectedItem;
                txtIdExport.Text = export.Id;
                cbToInitial.IsChecked = export.ToInitialType;
                txtInitialTypePath.Text = export.InitialTypePath;
                cbToPdf.IsChecked = export.ToPdf;
                txtPdfPath.Text = export.PdfPath;
            }
        }

        private void ButtonDeleteExport_Click(object sender, RoutedEventArgs e)
        {
            if (MessageBox.Show("Are you sure ? ", "Question", MessageBoxButton.YesNo, MessageBoxImage.Warning) == MessageBoxResult.Yes && lvExports.SelectedItem != null)
            {
                Export item = (Export)lvExports.SelectedItem;
                ((FormToSpLibrary)DataContext).Exports.Remove(item);
            }
        }

        private void ButtonUpdateExport_Click(object sender, RoutedEventArgs e)
        {
            if ((Export)lvExports.SelectedItem != null)
            {
                if (PreventInvalidCharacters()) return;
                Export item = (Export)lvExports.SelectedItem;
                item.Id = txtIdExport.Text;
                item.InitialTypePath = txtInitialTypePath.Text;
                item.PdfPath = txtPdfPath.Text;

                item.ToInitialType = cbToInitial.IsChecked ?? false;
                item.ToPdf = cbToPdf.IsChecked ?? false;

                lvExports.ItemsSource = null;
                lvExports.Items.Clear();
                lvExports.ItemsSource = ((FormToSpLibrary)DataContext).Exports;
            }


        }



    }
}
