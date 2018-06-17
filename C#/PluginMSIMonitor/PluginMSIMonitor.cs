using Rainmeter;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Management;
using System.Runtime.InteropServices;
using System.Threading.Tasks;

// Overview: This is a blank canvas on which to build your plugin.

// Note: GetString, ExecuteBang and an unnamed function for use as a section variable
// have been commented out. If you need GetString, ExecuteBang, and/or section variables 
// and you have read what they are used for from the SDK docs, uncomment the function(s)
// and/or add a function name to use for the section variable function(s). 
// Otherwise leave them commented out (or get rid of them)!

namespace PluginMSIMonitor
{
    internal static class Monitor
    {
        public static int CpuTemp { get; private set; }
        public static int GpuTemp { get; private set; }
        public static int CpuFan { get; private set; }
        public static int GpuFan { get; private set; }

        private static readonly ManagementObjectSearcher WmiCpuTemp;
        private static readonly ManagementObjectSearcher WmiGpuTemp;
        private static readonly ManagementObjectSearcher WmiFan;

        private static readonly object PollLock;

        // Key is observer instance key, Value is update rate in milliseconds
        private static readonly Dictionary<int, int> Observers;
        private static int _observerCount;

        private static bool _pollRunning;
        private static Task _pollTask;
        private static int _pollInterval;

        static Monitor()
        {
            PollLock = new object();
            Observers = new Dictionary<int, int>();

            WmiCpuTemp = new ManagementObjectSearcher("root\\WMI", "SELECT CPU FROM MSI_CPU");
            WmiGpuTemp = new ManagementObjectSearcher("root\\WMI", "SELECT VGA FROM MSI_VGA");
            WmiFan = new ManagementObjectSearcher("root\\WMI", "SELECT AP FROM MSI_AP");
        }

        public static int AddObserver(int updateRate)
        {
            lock (PollLock)
            {
                _observerCount++;
                Observers.Add(_observerCount, updateRate);
                _pollInterval = Observers.Select(m => m.Value).Min();
                _pollInterval = _pollInterval >= 100 ? _pollInterval : 100;

                if (!_pollRunning)
                {
                    _pollRunning = true;
                    _pollTask = Task.Run(() => PollLoop());
                }

                return _observerCount;
            }
        }

        public static void RemoveObserver(int observerKey)
        {
            lock (PollLock)
            {
                Observers.Remove(observerKey);

                if (Observers.Count == 0)
                {
                    _pollRunning = false;
                    _pollTask.Dispose();
                    _pollTask = null;
                }
                else
                {
                    _pollInterval = Observers.Select(m => m.Value).Min();
                    _pollInterval = _pollInterval >= 100 ? _pollInterval : 100;
                }
            }
        }

        public static void ChangeUpdateRate(int observerKey, int updateRate)
        {
            lock (PollLock)
            {
                Observers[observerKey] = updateRate;
                _pollInterval = Observers.Select(m => m.Value).Min();
                _pollInterval = _pollInterval >= 100 ? _pollInterval : 100;
            }
        }

        private static async void PollLoop()
        {
            while (_pollRunning)
            {
                var cpuTempItem = WmiCpuTemp.Get().GetEnumerator();
                cpuTempItem.MoveNext();
                CpuTemp = Convert.ToInt16(cpuTempItem.Current["CPU"].ToString());

                var gpuTempItem = WmiGpuTemp.Get().GetEnumerator();
                gpuTempItem.MoveNext();
                GpuTemp = Convert.ToInt16(gpuTempItem.Current["VGA"].ToString());

                var fanSpeedItem = WmiFan.Get().GetEnumerator();
                fanSpeedItem.MoveNext();

                fanSpeedItem.MoveNext();
                var cpuFanLow8 = Convert.ToInt16(fanSpeedItem.Current["AP"].ToString());
                fanSpeedItem.MoveNext();
                var cpuFanHigh8 = Convert.ToInt16(fanSpeedItem.Current["AP"].ToString());
                if (cpuFanLow8 != 0 || cpuFanHigh8 != 0)
                {
                    CpuFan = (int)(60000000.0 / ((double)(((cpuFanHigh8 << 8) + cpuFanLow8) * 2) * 62.5));
                }
                else
                {
                    CpuFan = 0;
                }

                fanSpeedItem.MoveNext();
                var gpuFanLow8 = Convert.ToInt16(fanSpeedItem.Current["AP"].ToString());
                fanSpeedItem.MoveNext();
                var gpuFanHigh8 = Convert.ToInt16(fanSpeedItem.Current["AP"].ToString());
                if (gpuFanLow8 != 0 || gpuFanHigh8 != 0)
                {
                    GpuFan = (int)(60000000.0 / ((double)(((gpuFanHigh8 << 8) + gpuFanLow8) * 2) * 62.5));
                }
                else
                {
                    GpuFan = 0;
                }

                await Task.Delay(_pollInterval);
            }
        }
    }

    internal class Measure : IDisposable
    {
        public static implicit operator Measure(IntPtr data)
        {
            return (Measure)GCHandle.FromIntPtr(data).Target;
        }

        public enum MeasureType
        {
            CpuTemp,
            GpuTemp,
            CpuFan,
            GpuFan
        }

        public Func<int> GetLatestValue { get; private set; }

        private readonly int _observerKey;

        public Measure()
        {
            _observerKey = Monitor.AddObserver(1000);
        }

        public void Reload(MeasureType type, int updateRate)
        {
            Monitor.ChangeUpdateRate(_observerKey, updateRate);

            switch (type)
            {
                case MeasureType.CpuTemp:
                    GetLatestValue = () => Monitor.CpuTemp;
                    break;
                case MeasureType.GpuTemp:
                    GetLatestValue = () => Monitor.GpuTemp;
                    break;
                case MeasureType.CpuFan:
                    GetLatestValue = () => Monitor.CpuFan;
                    break;
                case MeasureType.GpuFan:
                    GetLatestValue = () => Monitor.GpuFan;
                    break;
                default:
                    throw new ArgumentOutOfRangeException(nameof(type), type, null);
            }
        }

        private void DisposeOp()
        {
            Monitor.RemoveObserver(_observerKey);
        }

        public void Dispose()
        {
            DisposeOp();
            GC.SuppressFinalize(this);
        }

        ~Measure()
        {
            DisposeOp();
        }
    }

    public class Plugin
    {
        [DllExport]
        public static void Initialize(ref IntPtr data, IntPtr rm)
        {
            data = GCHandle.ToIntPtr(GCHandle.Alloc(new Measure()));
        }

        [DllExport]
        public static void Finalize(IntPtr data)
        {
            var measure = (Measure)data;
            measure.Dispose();

            GCHandle.FromIntPtr(data).Free();
        }

        [DllExport]
        public static void Reload(IntPtr data, IntPtr rm, ref double maxValue)
        {
            var api = (API)rm;
            if (Enum.TryParse(
                api.ReadString("MeasureType", string.Empty),
                true,
                out Measure.MeasureType type))
            {
                var measure = (Measure)data;
                measure.Reload(type, api.ReadInt("PollRate", 0));
            }
            else
            {
                api.Log(API.LogType.Error, "MeasureType invalid");
            }
        }

        [DllExport]
        public static double Update(IntPtr data)
        {
            var measure = (Measure)data;
            return measure.GetLatestValue();
        }

        //[DllExport]
        //public static IntPtr GetString(IntPtr data)
        //{
        //    Measure measure = (Measure)data;
        //    if (measure.buffer != IntPtr.Zero)
        //    {
        //        Marshal.FreeHGlobal(measure.buffer);
        //        measure.buffer = IntPtr.Zero;
        //    }
        //
        //    measure.buffer = Marshal.StringToHGlobalUni("");
        //
        //    return measure.buffer;
        //}

        //[DllExport]
        //public static void ExecuteBang(IntPtr data, [MarshalAs(UnmanagedType.LPWStr)]String args)
        //{
        //    Measure measure = (Measure)data;
        //}

        //[DllExport]
        //public static IntPtr (IntPtr data, int argc,
        //    [MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.LPWStr, SizeParamIndex = 1)] string[] argv)
        //{
        //    Measure measure = (Measure)data;
        //    if (measure.buffer != IntPtr.Zero)
        //    {
        //        Marshal.FreeHGlobal(measure.buffer);
        //        measure.buffer = IntPtr.Zero;
        //    }
        //
        //    measure.buffer = Marshal.StringToHGlobalUni("");
        //
        //    return measure.buffer;
        //}
    }
}