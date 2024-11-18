using Microsoft.Office.Interop.PowerPoint;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Timers;

namespace PowerPointTimer.Model
{
    class TimerData : IDisposable
    {
        private static readonly TimeSpan OneSecond = TimeSpan.FromSeconds(1);

        private readonly Timer _timer;
        private readonly string _duration;
        private readonly Shape _timerShape;
        private readonly Action _onTimeout;

        public TimerData(Shape timerShape, Action onTimeout = null)
        {
            _timer = new Timer(OneSecond.TotalMilliseconds);
            _timer.Elapsed += TimerTick;
            _timer.Start();
            _timerShape = timerShape;
            _onTimeout = onTimeout;
            _duration = GetShapeText();
        }

        private void TimerTick(object _, ElapsedEventArgs __)
        {
            string timeText = GetShapeText();
            if (timeText.StartsWith("@") && TimeSpan.TryParseExact(timeText.Substring(1), "h\\:mm",
                CultureInfo.InvariantCulture, out TimeSpan targetTime))
            {
                DateTime now = DateTime.Now;
                DateTime todayTarget = now.Date + targetTime;
                if (todayTarget <= now)
                    todayTarget = todayTarget.AddDays(1);
                timeText = (todayTarget - now).ToString("h\\:mm\\:ss");
            }
            if (TimeSpan.TryParseExact(timeText, new string[] { "h\\:mm\\:ss", "m\\:ss" },
                CultureInfo.InvariantCulture, out TimeSpan time))
            {
                if (time.TotalSeconds == 0)
                {
                    _onTimeout?.Invoke();
                    return;
                }
                time = time - OneSecond;
                SetShapeText(time.TotalHours >= 1 ? time.ToString("h\\:mm\\:ss") : time.ToString("m\\:ss"));
            } else
            {
                SetShapeText($"Invalid format: {_duration}");
            }
        }

        private void SetShapeText(string text)
        {
            _timerShape.TextFrame.TextRange.Text = text;
        }

        private string GetShapeText()
        {
            if (_timerShape.HasTextFrame == Microsoft.Office.Core.MsoTriState.msoTrue)
            {
                return _timerShape.TextFrame2.TextRange.Text;
            }

            return "No text found.";
        }

        public void Dispose()
        {
            SetShapeText(_duration);
            _timer.Dispose();
        }
    }
}
