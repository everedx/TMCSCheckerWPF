using System;
using System.Collections.Generic;
using System.Data;
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

namespace TMCSCheckerWPF
{
    /// <summary>
    /// Interaction logic for ExpandedGridWindow.xaml
    /// </summary>
    public partial class ExpandedGridWindow : Window
    {
        public ExpandedGridWindow()
        {
            InitializeComponent();
        }

        public ExpandedGridWindow(DataGrid dg)
        {
            InitializeComponent();
          
            foreach (DataGridColumn col in dg.Columns)
            {
                DataGridTextColumn column = new DataGridTextColumn();
                column.Header = col.Header;
                column.Width = col.Width;
                column.Binding = ((DataGridTextColumn)col).Binding;
                dgConnections.Columns.Add(column);
            }
            dgConnections.ItemsSource = dg.ItemsSource;

        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void Window_MouseDown_1(object sender, MouseButtonEventArgs e)
        {
            if (e.ChangedButton == MouseButton.Left)
                this.DragMove();
        }
    }
}
