using System;
using System.Collections.Generic;
using System.ComponentModel.Design;
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
    /// Interaction logic for InputModalWindow.xaml
    /// </summary>
    public partial class InputModalWindow : Window
    {
        public string ReturnValue
        {
            get { return inputBox.Text; }
        }

        public InputModalWindow()
        {
            InitializeComponent();
        }
        
        public InputModalWindow(string title)
        {
            InitializeComponent();
            titleLabel.Content = title;
            inputBox.Text = "";
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
