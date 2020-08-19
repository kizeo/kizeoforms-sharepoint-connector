using KizeoAndSharepoint_wizard.Models;
using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows;
using System.Windows.Controls;

namespace KizeoAndSharepoint_wizard
{
    /// <summary>
    /// Interaction logic for Step2AddOrUpdate.xaml
    /// </summary>
    public partial class Step2AddOrUpdate : Window
    {
        public ObservableCollection<FormToSpList> FormTospLists { get; set; }
        public bool OpnedForNewItem { get; internal set; }
        public ClientContext Context;

        public Step2AddOrUpdate()
        {
            InitializeComponent();
            
        }

        private void ButtonAjouter_Click(object sender, RoutedEventArgs e)
        {
            var formToSpList = (FormToSpList)DataContext;
            formToSpList.DataMapping.Add(new DataMapping { KfColumnSelector = TxtKfColumnId.Text, SpColumnId = (string)cbSpColumnId.SelectedValue, SpecialType = cbSpecialType.Text });

        }

        private void ButtonValider_Click(object sender, RoutedEventArgs e)
        {
            if (OpnedForNewItem)
            {
                FormTospLists.Add((FormToSpList)DataContext);
            }

            this.Hide();
        }

        private void LvMapping_SelectionChanged(object sender, SelectionChangedEventArgs e)
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
            if (lvMapping.SelectedItem != null && MessageBox.Show("Are you sure ? ", "Delete", MessageBoxButton.YesNo, MessageBoxImage.Warning) == MessageBoxResult.Yes)
            {
                var item = (DataMapping)lvMapping.SelectedItem;
                ((FormToSpList)DataContext).DataMapping.Remove(item);
            }
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
                lvMapping.ItemsSource = ((FormToSpList)DataContext).DataMapping;
            }
        }

       
        private void txtListId_LostFocus(object sender, RoutedEventArgs e)
        {
            FillCbBox();
        }



        public void FillCbBox()
        {
            const string list_mask = "00000000-0000-0000-0000-000000000000";
            if (
                !string.IsNullOrEmpty(txtListId.Text) 
                && (
                    !txtListId.Text.Equals(list_mask)
                )
            )
            {
                try
                {
                    Guid listId = new Guid(txtListId.Text);
                    var spList = Context.Web.Lists.GetById(listId);
                    Context.Load(spList);
                    CamlQuery query = CamlQuery.CreateAllItemsQuery();
                    ListItemCollection items = spList.GetItems(query);

                    Context.Load(spList.Fields);
                    Context.ExecuteQuery();
                    var x = spList.Fields.Where(f => f.Hidden == false).Where(f => f.ReadOnlyField == false).Where(f => f.StaticName == "Title" || f.CanBeDeleted == true).ToList();
                    cbSpColumnId.ItemsSource = x;
                }
                catch (Exception)
                {
                    MessageBox.Show("SharePoint's list couldn't be loaded.\nPlease check out that:\n- The client is associated to the SharePoint url you entered.\n- The url is well formated");
                 
                    cbSpColumnId.ItemsSource = null;
                    cbSpColumnId.Items.Clear();

                }

            }
        }
    }
}
