using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Threading;
using System.Collections.Concurrent;
using System.Collections.ObjectModel;
using System.Windows;

namespace ReservationPostingRocketChat
{
    static class TimerManager
    {
        private static ObservableCollection<PostingEvent> timerEventList = new ObservableCollection<PostingEvent>();
        public static ObservableCollection<PostingEvent> TimerEventList
        {
            get { return timerEventList; }
            set { timerEventList = TimerEventList; }
        }

        // UIスレッドのDispatcherを取得
        static Dispatcher dispatcher = Application.Current.Dispatcher;
        static DispatcherTimer timer;

        const double interval = 10000; // ケチなので10秒おきにイベント発行

        public static void StartTimer()
        {
            timer = new DispatcherTimer();
            timer.Interval = TimeSpan.FromMilliseconds(10000);
            timer.Tick += TickTimer;
            timer.Start();
        }

        public static void StopTimer()
        {
            if(timer.IsEnabled)
            {
                timer.Stop();
            }
        }

        // 過度なアクセスを防止するため、1回のTickで1回のイベントしか実行させない
        static void TickTimer(object sender, EventArgs e)
        {
            if (TimerEventList.Count() > 0)
            {
                StopTimer();
            }

            var nextEvent =  TimerEventList
                .OrderBy(x => x.PostingTime)
                .FirstOrDefault();

            if (nextEvent != null && nextEvent.PostingTime < DateTime.Now)
            {
                nextEvent.Exec();
                // 失敗しても成功してもdelete
                dispatcher.Invoke(() => {
                    TimerEventList.Remove(nextEvent);
                    UpdateTextBlockTabReserved();
                });
            }
        }

        public static void AddTimerEvent(PostingEvent timerEvent)
        {
            // 必ずUIスレッドからアクセスするように排他制御
            dispatcher.Invoke(() => {
                TimerEventList.Add(timerEvent);
                UpdateTextBlockTabReserved();
            });
        }

        // ちなみにUI層とデータ層を分けることはWPF界隈で悪手と言われている
        static void UpdateTextBlockTabReserved()
        {
            var mainWindow = (MainWindow) App.Current.MainWindow;

            if (TimerEventList.Count() > 0)
            {
                mainWindow.TextBlockTabReserved.Text = "予約済み(" + TimerEventList.Count() + ")";
            }
            else
            {
                mainWindow.TextBlockTabReserved.Text = "予約済み";
            }
        }
    }
}
