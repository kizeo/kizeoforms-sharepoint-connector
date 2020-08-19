﻿using KizeoAndSharepoint_wizard.Models;
using Microsoft.SharePoint.Client;
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
    /// Interaction logic for Step4.xaml
    /// </summary>
    public partial class Step4AddOrUpdate : Window
    {
        public ObservableCollection<SpListToExtList> SpListToExtLists { get; set; }
        public bool OpnedForNewItem { get; set; }
        public ClientContext Context;

        public Step4AddOrUpdate()
        {
            InitializeComponent();
            
        }

        private void ButtonValider_Click(object sender, RoutedEventArgs e)
        {
            if (OpnedForNewItem)
            {
                SpListToExtLists.Add((SpListToExtList)DataContext);
            }
            this.Hide();
        }

        private void cbSpColumnsId_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            ((SpListToExtList)DataContext).DataSchema = txtDataSchema.Text + "$$" + cbSpColumnsId.SelectedValue + "$$";
        }

       
        private void txtSpListId_LostFocus(object sender, RoutedEventArgs e)
        {
            FillCbBox();
        }

        public void FillCbBox()
        {
            const string list_mask = "00000000-0000-0000-0000-000000000000";
            if (
                !string.IsNullOrEmpty(txtSpListId.Text)
                && (
                    !txtSpListId.Text.Equals(list_mask)
                )
            )
            {
                try
                {
                    Guid listId = new Guid(txtSpListId.Text);
                    var spList = Context.Web.Lists.GetById(listId);
                    Context.Load(spList);
                    CamlQuery query = CamlQuery.CreateAllItemsQuery();
                    Microsoft.SharePoint.Client.ListItemCollection items = spList.GetItems(query);

                    Context.Load(spList.Fields);
                    Context.ExecuteQuery();
                    var x = spList.Fields.Where(f => f.Hidden == false).Where(f => f.ReadOnlyField == false).Where(f => f.StaticName == "Title" || f.CanBeDeleted == true).ToList();
                    cbSpColumnsId.ItemsSource = x;
                }
                catch (Exception )
                {
                    MessageBox.Show("SharePoint's list couldn't be loaded.\nPlease check out that:\n- The client is associated to the SharePoint url you entered.\n- The url is well formated");

                    cbSpColumnsId.ItemsSource = null;
                    cbSpColumnsId.Items.Clear();
                }

            }
        }
    }
}
