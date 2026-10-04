using System;
using System.Diagnostics;
using System.Globalization;

namespace SlidePace
{
    public enum TimerMode { None, CountUp, CountDown, Clock }

    public interface ITimeSource
    {
        double MonotonicSeconds { get; }
        DateTime LocalTime { get; }
    }

    public sealed class SystemTimeSource : ITimeSource
    {
        public double MonotonicSeconds { get { return (double)Stopwatch.GetTimestamp() / Stopwatch.Frequency; } }
        public DateTime LocalTime { get { return DateTime.Now; } }
    }

    public sealed class TimerSnapshot
    {
        public TimerMode Mode { get; internal set; }
        public string Text { get; internal set; }
        public string Status { get; internal set; }
        public bool IsRunning { get; internal set; }
        public bool HasStarted { get; internal set; }
        public bool IsOvertime { get; internal set; }
        public double ElapsedSeconds { get; internal set; }
    }

    public sealed class TimerEngine
    {
        private sealed class Counter
        {
            public double Accumulated;
            public double Anchor;
            public bool Running;
            public bool Started;
            public double Elapsed(double now) { return Accumulated + (Running ? Math.Max(0, now - Anchor) : 0); }
            public void Start(double now)
            {
                if (Running) return;
                Anchor = now;
                Running = true;
                Started = true;
            }
            public void Pause(double now)
            {
                Accumulated = Elapsed(now);
                Running = false;
            }
            public void Reset() { Accumulated = 0; Running = false; Started = false; }
        }

        private readonly ITimeSource time;
        private readonly Counter up = new Counter();
        private readonly Counter down = new Counter();
        public TimerMode Mode { get; private set; }
        public bool IsSlideShowActive { get; private set; }
        public int CountdownSeconds { get; private set; }

        public TimerEngine(ITimeSource source)
        {
            if (source == null) throw new ArgumentNullException("source");
            time = source;
            CountdownSeconds = 600;
        }

        private Counter Current
        {
            get { return Mode == TimerMode.CountUp ? up : Mode == TimerMode.CountDown ? down : null; }
        }

        public void SelectMode(TimerMode mode)
        {
            if (!Enum.IsDefined(typeof(TimerMode), mode)) throw new ArgumentOutOfRangeException("mode");
            if (mode == Mode) return;
            Pause();
            Mode = mode;
            if (IsSlideShowActive) Start();
        }

        public void BeginSlideShow()
        {
            if (IsSlideShowActive) return;
            IsSlideShowActive = true;
            Start();
        }

        public void EndSlideShow()
        {
            double now = time.MonotonicSeconds;
            up.Pause(now);
            down.Pause(now);
            IsSlideShowActive = false;
        }

        public void Start()
        {
            Counter counter = Current;
            if (IsSlideShowActive && counter != null) counter.Start(time.MonotonicSeconds);
        }

        public void Pause()
        {
            Counter counter = Current;
            if (counter != null) counter.Pause(time.MonotonicSeconds);
        }

        public void Reset()
        {
            Counter counter = Current;
            if (counter != null) counter.Reset();
        }

        public void SetCountdownSeconds(int seconds)
        {
            if (seconds < 1 || seconds > 86399) throw new ArgumentOutOfRangeException("seconds", "时长须在 00:00:01 至 23:59:59 之间。");
            if (seconds == CountdownSeconds) return;
            if (down.Running) throw new InvalidOperationException("请先暂停倒计时，再修改时长。");
            CountdownSeconds = seconds;
            down.Reset();
        }

        public TimerSnapshot Snapshot()
        {
            var result = new TimerSnapshot { Mode = Mode, Text = "", Status = "未选择模式" };
            if (Mode == TimerMode.None) return result;
            if (Mode == TimerMode.Clock)
            {
                result.Text = time.LocalTime.ToString("HH:mm:ss", CultureInfo.InvariantCulture);
                result.Status = "系统时间";
                return result;
            }
            Counter counter = Current;
            double elapsed = counter.Elapsed(time.MonotonicSeconds);
            result.IsRunning = counter.Running;
            result.HasStarted = counter.Started;
            result.ElapsedSeconds = elapsed;
            result.IsOvertime = Mode == TimerMode.CountDown && elapsed >= CountdownSeconds;
            if (Mode == TimerMode.CountUp) result.Text = FormatSeconds((long)Math.Floor(elapsed));
            else if (!result.IsOvertime) result.Text = FormatSeconds((long)Math.Ceiling(CountdownSeconds - elapsed));
            else
            {
                long overtime = (long)Math.Floor(elapsed - CountdownSeconds);
                result.Text = (overtime > 0 ? "-" : "") + FormatSeconds(overtime);
            }
            result.Status = !counter.Started ? "未开始" : counter.Running ? "运行中" : "已暂停";
            if (result.IsOvertime) result.Status = counter.Running ? "超时" : "超时 · 已暂停";
            return result;
        }

        public static string FormatSeconds(long seconds)
        {
            seconds = Math.Max(0, seconds);
            return string.Format(CultureInfo.InvariantCulture, "{0:00}:{1:00}:{2:00}", seconds / 3600, seconds / 60 % 60, seconds % 60);
        }

        public static bool TryParseDuration(string text, out int seconds)
        {
            seconds = 0;
            if (string.IsNullOrWhiteSpace(text)) return false;
            string[] parts = text.Trim().Split(':');
            int hours, minutes, secs;
            if (parts.Length != 3 || !int.TryParse(parts[0], NumberStyles.None, CultureInfo.InvariantCulture, out hours) ||
                !int.TryParse(parts[1], NumberStyles.None, CultureInfo.InvariantCulture, out minutes) ||
                !int.TryParse(parts[2], NumberStyles.None, CultureInfo.InvariantCulture, out secs) ||
                hours < 0 || hours > 23 || minutes < 0 || minutes > 59 || secs < 0 || secs > 59) return false;
            seconds = hours * 3600 + minutes * 60 + secs;
            return seconds > 0;
        }

        public static string ModeName(TimerMode mode)
        {
            return mode == TimerMode.CountUp ? "顺计时" : mode == TimerMode.CountDown ? "倒计时" : mode == TimerMode.Clock ? "系统时间" : "不显示";
        }
    }
}
