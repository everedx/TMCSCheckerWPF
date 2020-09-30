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

namespace TMCSCheckerWPF
{
    /// <summary>
    /// Interaction logic for CustomModalWindow.xaml
    /// </summary>
    public partial class CustomModalWindow : Window
    {
        public CustomModalWindow()
        {
            InitializeComponent();

        }
        public CustomModalWindow(string title, string content)
        {
            InitializeComponent();
            titleLabel.Content = title;
            contentLabel.Text = content;
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
