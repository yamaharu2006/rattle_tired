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

namespace ReservationPostingRocketChat
{
    /// <summary>
    /// MainWindow.xaml の相互作用ロジック
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            TextBoxPostingTime.Text = DateTime.Now.ToString();
            ListViewReservationPosting.ItemsSource = TimerManager.TimerEventList;
        }

        void ButtonPost_Click(object sender, RoutedEventArgs e)
        {
            if(CheckEnablePost())
            {
                TimerManager.StartTimer();
                TimerManager.AddTimerEvent(new PostingEvent(TextBoxRoomName.Text, TextBoxPostingTime.Text, TextBoxPostingContext.Text));
            }
        }

        bool CheckEnablePost()
        {
            if(PasswordBoxAccessToken.Password == "" )
            {

            }
            if (PasswordBoxUserId.Password == "")
            {

            }
            return true;
        }

    }
}
