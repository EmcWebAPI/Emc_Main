using System.Diagnostics;

namespace EmcReportWebApi.Utils
{
    /// <summary>
    /// 计时器
    /// </summary>
    public class TimerUtil
    {
        private readonly Stopwatch _stopwatch;
        /// <summary>
        /// 构造函数计时
        /// </summary>
        /// <param name="stopwatch"></param>
        public TimerUtil(Stopwatch stopwatch)
        {
            _stopwatch = stopwatch;
            StartTimer();

        }

        /// <summary>
        /// 开始计时
        /// </summary>
        private void StartTimer()
        {
            _stopwatch?.Start();
        }

        /// <summary>
        /// 停止计时器
        /// </summary>
        /// <returns></returns>
        public double StopTimer()
        {
            if (_stopwatch == null)
                return 0d;
            _stopwatch?.Stop();
            return (double)_stopwatch.ElapsedMilliseconds / 1000;
        }
    }
}