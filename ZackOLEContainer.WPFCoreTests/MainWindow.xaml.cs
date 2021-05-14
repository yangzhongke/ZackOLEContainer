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

namespace ZackOLEContainer.WPFCoreTests
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();

          
        }

        private void Window_Loaded(System.Object sender, System.Windows.RoutedEventArgs e)
        {
            this.oleContainer.OpenFile(@"E:\主同步盘\我的坚果云\UoZ\SE101-玩着学编程\Part3课件\5-搞“对象”.pptx");
        }
    }
}
