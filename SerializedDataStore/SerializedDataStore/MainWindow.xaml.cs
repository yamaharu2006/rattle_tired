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

namespace SerializedDataStore
{
    /// <summary>
    /// MainWindow.xaml の相互作用ロジック
    /// </summary>
    public partial class MainWindow : Window
    {
        AddressBook addressBook = new AddressBook();
        BackupManager backupManager = new BackupManager();

        public MainWindow()
        {
            InitializeComponent();

            backupManager.Load(ref addressBook);

            CustomerListView.ItemsSource = addressBook.AddressList;
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            addressBook.Add(new PersonalData(TextBoxName.Text, TextBoxAddress.Text));
        }

        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            backupManager.Save(addressBook);
        }
    }
}
