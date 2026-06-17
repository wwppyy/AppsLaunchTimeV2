namespace AppsLaunchTime
{
    using System.Diagnostics;
    using System.Runtime.InteropServices;

    public partial class Form1 : Form
    {
        private const string ExcelFilePath = @"c:\test\sample.xlsx";

        // avicap32 / VFW interop for activating the camera without rendering.
        private const uint WS_CHILD = 0x40000000;
        private const int WM_CAP_START = 0x0400;
        private const int WM_CAP_DRIVER_CONNECT = WM_CAP_START + 10;
        private const int WM_CAP_DRIVER_DISCONNECT = WM_CAP_START + 11;

        [DllImport("avicap32.dll", CharSet = CharSet.Unicode)]
        private static extern IntPtr capCreateCaptureWindow(
            string lpszWindowName,
            uint dwStyle,
            int x, int y, int nWidth, int nHeight,
            IntPtr hwndParent, int nID);

        [DllImport("user32.dll")]
        private static extern IntPtr SendMessage(IntPtr hWnd, int msg, IntPtr wParam, IntPtr lParam);

        [DllImport("user32.dll")]
        private static extern bool DestroyWindow(IntPtr hWnd);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

        private IntPtr _captureWindow = IntPtr.Zero;

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private async void button1_Click(object sender, EventArgs e)
        {
            button1.Enabled = false;
            numericUpDownIterations.Enabled = false;
            textBox1.Text = string.Empty;

            try
            {
                if (!File.Exists(ExcelFilePath))
                {
                    textBox1.Text = $"File not found: {ExcelFilePath}";
                    return;
                }

                Type? excelType = Type.GetTypeFromProgID("Excel.Application");
                if (excelType == null)
                {
                    textBox1.Text = "Microsoft Excel is not installed.";
                    return;
                }

                int iterations = (int)numericUpDownIterations.Value;
                var results = new List<(double FileOpen, double Render, double Total)>();

                AppendLine($"Running {iterations} iteration(s) against:");
                AppendLine(ExcelFilePath);
                AppendLine(new string('=', 50));

                for (int iter = 1; iter <= iterations; iter++)
                {
                    AppendLine($"Iteration {iter}/{iterations} ...");
                    var result = await RunSingleMeasurementAsync(excelType);
                    if (result == null)
                    {
                        AppendLine("  (failed)");
                        continue;
                    }

                    var (fileOpen, render, total) = result.Value;
                    results.Add((fileOpen, render, total));

                    AppendLine($"  File Open : {fileOpen,7:N3} s");
                    AppendLine($"  Rendering : {render,7:N3} s");
                    AppendLine($"  Total     : {total,7:N3} s");
                    AppendLine(new string('-', 50));
                }

                if (results.Count > 0)
                {
                    double avgOpen = results.Average(r => r.FileOpen);
                    double avgRender = results.Average(r => r.Render);
                    double avgTotal = results.Average(r => r.Total);

                    AppendLine("Averages:");
                    AppendLine($"  File Open : {avgOpen,7:N3} s");
                    AppendLine($"  Rendering : {avgRender,7:N3} s");
                    AppendLine($"  Total     : {avgTotal,7:N3} s");
                }
            }
            catch (Exception ex)
            {
                AppendLine($"Error: {ex.Message}");
            }
            finally
            {
                button1.Enabled = true;
                numericUpDownIterations.Enabled = true;
            }
        }

        private void AppendLine(string text)
        {
            textBox1.AppendText(text + "\r\n");
        }

        private async Task<(double FileOpen, double Render, double Total)?> RunSingleMeasurementAsync(Type excelType)
        {
            object? excelApp = null;
            object? workbooks = null;
            object? workbook = null;

            try
            {
                // Measure: start Excel, open workbook (loads content), and run any auto-open macros.
                var stopwatch = Stopwatch.StartNew();

                excelApp = Activator.CreateInstance(excelType);
                if (excelApp == null)
                {
                    return null;
                }

                excelType.InvokeMember("Visible", System.Reflection.BindingFlags.SetProperty, null, excelApp, new object[] { true });
                excelType.InvokeMember("DisplayAlerts", System.Reflection.BindingFlags.SetProperty, null, excelApp, new object[] { false });
                excelType.InvokeMember("AutomationSecurity", System.Reflection.BindingFlags.SetProperty, null, excelApp, new object[] { 1 /* msoAutomationSecurityLow - allow macros */ });

                workbooks = excelType.InvokeMember("Workbooks", System.Reflection.BindingFlags.GetProperty, null, excelApp, null);

                workbook = workbooks!.GetType().InvokeMember(
                    "Open",
                    System.Reflection.BindingFlags.InvokeMethod,
                    null,
                    workbooks,
                    new object[] { ExcelFilePath });

                double fileOpenSeconds = stopwatch.Elapsed.TotalSeconds;

                try
                {
                    excelType.InvokeMember("CalculateUntilAsyncQueriesDone", System.Reflection.BindingFlags.InvokeMethod, null, excelApp, null);
                }
                catch
                {
                }

                Process? excelProcess = null;
                try
                {
                    object hwndObj = excelType.InvokeMember("Hwnd", System.Reflection.BindingFlags.GetProperty, null, excelApp, null)!;
                    IntPtr excelHwnd = new IntPtr(Convert.ToInt64(hwndObj));
                    if (excelHwnd != IntPtr.Zero)
                    {
                        GetWindowThreadProcessId(excelHwnd, out uint pid);
                        if (pid != 0)
                        {
                            excelProcess = Process.GetProcessById((int)pid);
                        }
                    }
                }
                catch
                {
                }

                var renderStopwatch = Stopwatch.StartNew();
                int stableCount = 0;
                const int maxWaitCycles = 200;
                const int sleepIntervalMs = 100;
                const int requiredStableCycles = 5;

                if (excelProcess != null)
                {
                    for (int i = 0; i < maxWaitCycles; i++)
                    {
                        double cpuBefore = excelProcess.TotalProcessorTime.TotalMilliseconds;
                        await Task.Delay(sleepIntervalMs);
                        excelProcess.Refresh();
                        double cpuAfter = excelProcess.TotalProcessorTime.TotalMilliseconds;
                        double cpuDelta = cpuAfter - cpuBefore;

                        if (cpuDelta < 2)
                        {
                            stableCount++;
                            if (stableCount >= requiredStableCycles) break;
                        }
                        else
                        {
                            stableCount = 0;
                        }
                    }
                }

                renderStopwatch.Stop();
                stopwatch.Stop();

                double bufferSeconds = (stableCount * sleepIntervalMs) / 1000.0;
                double renderSeconds = Math.Max(0, renderStopwatch.Elapsed.TotalSeconds - bufferSeconds);
                double totalSeconds = Math.Max(0, stopwatch.Elapsed.TotalSeconds - bufferSeconds);

                return (fileOpenSeconds, renderSeconds, totalSeconds);
            }
            finally
            {
                try
                {
                    if (workbook != null)
                    {
                        workbook.GetType().InvokeMember("Close", System.Reflection.BindingFlags.InvokeMethod, null, workbook, new object[] { false });
                        Marshal.FinalReleaseComObject(workbook);
                    }
                    if (workbooks != null)
                    {
                        Marshal.FinalReleaseComObject(workbooks);
                    }
                    if (excelApp != null)
                    {
                        excelType.InvokeMember("Quit", System.Reflection.BindingFlags.InvokeMethod, null, excelApp, null);
                        Marshal.FinalReleaseComObject(excelApp);
                    }
                }
                catch
                {
                }

                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (_captureWindow == IntPtr.Zero)
            {
                // Create a hidden capture window (no WS_VISIBLE) and connect to camera driver 0.
                // Without WM_CAP_SET_PREVIEW the camera is powered on but no frames are rendered.
                IntPtr hwnd = capCreateCaptureWindow(
                    "CaptureWindow",
                    WS_CHILD,
                    0, 0, 0, 0,
                    this.Handle,
                    0);

                if (hwnd == IntPtr.Zero)
                {
                    textBox1.Text = "Failed to create capture window.";
                    return;
                }

                IntPtr connected = SendMessage(hwnd, WM_CAP_DRIVER_CONNECT, IntPtr.Zero, IntPtr.Zero);
                if (connected == IntPtr.Zero)
                {
                    DestroyWindow(hwnd);
                    textBox1.Text = "Failed to connect to camera.";
                    return;
                }

                _captureWindow = hwnd;
                button2.Text = "Stop Camera";
                textBox1.Text = "Camera activated (no rendering).";
            }
            else
            {
                SendMessage(_captureWindow, WM_CAP_DRIVER_DISCONNECT, IntPtr.Zero, IntPtr.Zero);
                DestroyWindow(_captureWindow);
                _captureWindow = IntPtr.Zero;
                button2.Text = "Camera";
                textBox1.Text = "Camera stopped.";
            }
        }

        protected override void OnFormClosed(FormClosedEventArgs e)
        {
            if (_captureWindow != IntPtr.Zero)
            {
                SendMessage(_captureWindow, WM_CAP_DRIVER_DISCONNECT, IntPtr.Zero, IntPtr.Zero);
                DestroyWindow(_captureWindow);
                _captureWindow = IntPtr.Zero;
            }
            base.OnFormClosed(e);
        }
    }
}
