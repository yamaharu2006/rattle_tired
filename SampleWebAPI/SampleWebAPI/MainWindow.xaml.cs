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
using System.Net.Http;
using System.Windows.Threading;

namespace SampleWebAPI
{
    /// <summary>
    /// MainWindow.xaml の相互作用ロジック
    /// </summary>
    public partial class MainWindow : Window
    {
        // HttpClientは、使用ごとではなく、アプリケーションごとに1回インスタンス化されることを目的としています。備考を参照してください。
        static readonly HttpClient client = new HttpClient();

        private DispatcherTimer timer = new DispatcherTimer();
        private DateTime postTime;

        public MainWindow()
        {
            InitializeComponent();

            // 精度は求めていないので1分ごとにタイマーを発行
            TimerTextBox.Text = DateTime.Now.ToString();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            ResultTextbox.Text = "";

            timer.Interval = TimeSpan.FromMilliseconds(1000);
            timer.Tick += new EventHandler(CallBack);
            timer.Start();
        }

        // [@note] https://qiita.com/rawr/items/f78a3830d894042f891b#ボディを送るcontent-typeを指定する
        private async Task GetZipcode()
        {
#if false
            // RequestHeaderを設定したい場合
            client.DefaultRequestHeaders.Add("zipcode", "0010000
#endif

#if true
            // BodyDataを設定したい場合
            var zipcode = ZipcodeTextBox.Text;
            var parameters = new Dictionary<string, string>()
            {
                { "zipcode", zipcode }
            };
            var content = new FormUrlEncodedContent(parameters);
            
            var response = await client.PostAsync($"https://zipcloud.ibsnet.co.jp/api/search", content);

#endif

            // EnsureSuccessStatusCode : HTTPレスポンスが失敗した場合は例外を投げる
            response.EnsureSuccessStatusCode();

            string responseBody = await response.Content.ReadAsStringAsync();

            Console.WriteLine(responseBody);
            ResultTextbox.Text = responseBody;
        }

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            postTime = DateTime.Parse(TimerTextBox.Text);
        }

        private void CallBack(object sender, EventArgs e)
        {
            if(postTime < DateTime.Now)
            {
                GetZipcode();
                timer.Stop();
            }
        }
    }
}
