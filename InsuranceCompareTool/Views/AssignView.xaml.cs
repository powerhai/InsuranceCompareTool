﻿using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using InsuranceCompareTool.Domain;
using InsuranceCompareTool.Models;
using InsuranceCompareTool.ShareCommon;
using InsuranceCompareTool.ShareCommon.ValueConverter;
using Xceed.Wpf.Toolkit;
namespace InsuranceCompareTool.Views
{
    /// <summary>
    /// Interaction logic for AssignView
    /// </summary>
    public partial class AssignView 
    {
        private readonly List<string> mSystemColumns = BillTableColumns.Columns.Where(a => a.IsSystemColumn == true).Select(a => a.Name).ToList();

        private readonly List<string> mDateColumns = BillTableColumns.Columns.Where(a => a.Type == typeof(DateTime)).Select(a => a.Name).ToList();
        private readonly List<string> mIntColumns =  BillTableColumns.Columns.Where(a => a.Type == typeof(int)).Select(a=>a.Name).ToList();
        private readonly List<string> mNumberColumns = BillTableColumns.Columns.Where(a => a.Type == typeof(double)).Select(a => a.Name).ToList();
        public AssignView()
        {
            InitializeComponent();

            var logger = new VisualLogger
            {
                TextBox = T_Log
            };
            log4net.Config.BasicConfigurator.Configure(logger);
            
        }
        private void DataGrid_OnAutoGeneratingColumn(object sender, DataGridAutoGeneratingColumnEventArgs e)
        {
            var col = e.Column as DataGridTextColumn;

            if(col != null)
            {
                if(mSystemColumns.Contains(e.PropertyName))
                {
                    e.Cancel = true;
                }
                Style s = new Style();
                s.Setters.Add(new Setter() { Property = TextBlock.TextWrappingProperty, Value = TextWrapping.Wrap });
                s.Setters.Add(new Setter() { Property = TextBlock.HeightProperty, Value = (double.NaN)});
                if(mIntColumns.Contains(e.PropertyName))
                {
                    s.Setters.Add(new Setter(){ Property = TextBlock.HorizontalAlignmentProperty, Value =  HorizontalAlignment.Center});
                }

                if(mNumberColumns.Contains(e.PropertyName))
                {
                    s.Setters.Add(new Setter() { Property = TextBlock.HorizontalAlignmentProperty, Value = HorizontalAlignment.Right });

                }

                col.ElementStyle = s;
                DataGridTextColumn dc = e.Column as DataGridTextColumn;
                System.Windows.Data.Binding binding = dc.Binding as Binding;
                if(mDateColumns.Contains(binding.Path.Path))
                {
                    binding.StringFormat = "d";
                }


                //col.CellStyle.Setters.Add(new  Setter(){ Property = TextBlock.TextWrappingProperty, Value = TextWrapping.Wrap });
                //col.CellStyle.Setters.Add(new  Setter(){ Property = TextBlock.HeightProperty, Value =90}); 
            }
        }
        private void DataGrid_OnAutoGeneratedColumns(object sender, EventArgs e)
        {
          
            
            var dataGrid = sender as DataGrid;
            for(var i = 0; i < dataGrid.Columns.Count; i++)
            {
                var col = dataGrid.Columns[i];
                Binding b = new Binding();
                b.Source = this.DataContext;
                b.Path = new PropertyPath($"Columns[{i}].IsVisible");
                b.Converter = new Boolean2VisibleConverter();
                BindingOperations.SetBinding(col, DataGridColumn.VisibilityProperty, b);
            }
        }
        private void ButtonBase_OnClick(object sender, RoutedEventArgs e)
        {
            T_Log.Document.Blocks.Clear();
        }
    }
}
