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
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace Wpf_LLU1
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            /*StackPanel stackPanel = new StackPanel();
            this.Content = stackPanel;

            // Create the TextBlock 
            TextBlock textBlock = new TextBlock();
            textBlock.Text = "Welcome to XAML Tutorial";
            textBlock.Height = 20;
            textBlock.Width = 200;
            textBlock.Margin = new Thickness(5);
            stackPanel.Children.Add(textBlock);
            
            // Create the Button 
            Button button = new Button();
            button.Content = "OK";
            button.Height = 20;
            button.Width = 50;
            button.Margin = new Thickness(20);
            stackPanel.Children.Add(button);
            */
        }

        private void btButton1_Click(object sender, RoutedEventArgs e)
        {
            ExcelFunctions ExcelFunctions = new ExcelFunctions();
            ExcelFunctions.CreateExcel(@"c:\temp\mynames.xlsx");

        }

        private void btButton2_Click(object sender, RoutedEventArgs e)
        {
            WordFunctions wordFunctions = new WordFunctions();
            wordFunctions.CreateWord();
        }
    }
}
