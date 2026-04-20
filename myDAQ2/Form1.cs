using NationalInstruments.DAQmx;
using System;
using System.ComponentModel;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using System.Windows.Forms.DataVisualization.Charting;

namespace myDAQ2
{
    public partial class Form1 : Form
    {
        // UI Controls
        private IContainer components = null;
        private ComboBox cmbDevice;
        private CheckBox chkSaveData;
        private NumericUpDown numSampleRate;
        private NumericUpDown numSamples;
        private NumericUpDown numLedAmp;
        private NumericUpDown numLedFreq;
        private NumericUpDown numLedOffset;
        private NumericUpDown numTimerInterval;
        private Button btnStart;
        private Button btnStop;
        private Label lblAled;
        private Label lblAledRange;
        private Chart chartDrive;
        private Chart chartDetected;
        private Chart chartMultiplied;
        private Chart chartLockIn;
        private Timer timer1;

        // DAQ Tasks and Readers/Writers
        private Task outTask;
        private Task inTask;
        private AnalogSingleChannelWriter writerAO0;
        private AnalogMultiChannelReader readerAI;

        // Measurement State Variables
        private StreamWriter fileWriter;
        private int measurementTick = 0;

        public Form1()
        {
            InitializeComponent();
            LoadConnectedDevices();

            btnStart.Click += BtnStart_Click;
            btnStop.Click += BtnStop_Click;
            timer1.Tick += Timer1_Tick;
        }

        private void InitializeComponent()
        {
            this.components = new Container();

            // Initialize Controls
            this.cmbDevice = new ComboBox();
            this.chkSaveData = new CheckBox();
            this.numSampleRate = new NumericUpDown();
            this.numSamples = new NumericUpDown();
            this.numLedAmp = new NumericUpDown();
            this.numLedFreq = new NumericUpDown();
            this.numLedOffset = new NumericUpDown();
            this.numTimerInterval = new NumericUpDown();
            this.btnStart = new Button();
            this.btnStop = new Button();
            this.lblAled = new Label();
            this.lblAledRange = new Label();
            this.chartDrive = new Chart();
            this.chartDetected = new Chart();
            this.chartMultiplied = new Chart();
            this.chartLockIn = new Chart();
            this.timer1 = new Timer(this.components);

            Label lblInput1 = new Label() { Text = "Sample Rate (Hz)", Location = new Point(20, 20), AutoSize = true };
            Label lblInput2 = new Label() { Text = "Num Samples", Location = new Point(20, 60), AutoSize = true };
            Label lblInput3 = new Label() { Text = "LED Amp (V)", Location = new Point(20, 100), AutoSize = true };
            Label lblInput4 = new Label() { Text = "LED Freq (Hz)", Location = new Point(20, 140), AutoSize = true };
            Label lblInput5 = new Label() { Text = "LED Offset (V)", Location = new Point(20, 180), AutoSize = true };
            Label lblInput6 = new Label() { Text = "Interval (ms)", Location = new Point(20, 220), AutoSize = true };
            Label lblDevice = new Label() { Text = "DAQ Device", Location = new Point(20, 260), AutoSize = true };

            ((ISupportInitialize)(this.numSampleRate)).BeginInit();
            ((ISupportInitialize)(this.numSamples)).BeginInit();
            ((ISupportInitialize)(this.numLedAmp)).BeginInit();
            ((ISupportInitialize)(this.numLedFreq)).BeginInit();
            ((ISupportInitialize)(this.numLedOffset)).BeginInit();
            ((ISupportInitialize)(this.numTimerInterval)).BeginInit();
            ((ISupportInitialize)(this.chartDrive)).BeginInit();
            ((ISupportInitialize)(this.chartDetected)).BeginInit();
            ((ISupportInitialize)(this.chartMultiplied)).BeginInit();
            ((ISupportInitialize)(this.chartLockIn)).BeginInit();
            this.SuspendLayout();

            // NumericUpDown Configurations
            this.numSampleRate.Location = new Point(130, 18);
            this.numSampleRate.Maximum = 200000;
            this.numSampleRate.Value = 10000;

            this.numSamples.Location = new Point(130, 58);
            this.numSamples.Maximum = 100000;
            this.numSamples.Value = 1000;

            this.numLedAmp.DecimalPlaces = 2;
            this.numLedAmp.Location = new Point(130, 98);
            this.numLedAmp.Value = 2.0M;

            this.numLedFreq.Location = new Point(130, 138);
            this.numLedFreq.Maximum = 10000;
            this.numLedFreq.Value = 100;

            this.numLedOffset.DecimalPlaces = 2;
            this.numLedOffset.Location = new Point(130, 178);
            this.numLedOffset.Value = 2.5M; // Keep LED forward biased

            this.numTimerInterval.Location = new Point(130, 218);
            this.numTimerInterval.Maximum = 2000;
            this.numTimerInterval.Value = 100;

            // Device Combo and Checkbox
            this.cmbDevice.Location = new Point(130, 258);
            this.cmbDevice.DropDownStyle = ComboBoxStyle.DropDownList;
            this.cmbDevice.Size = new Size(120, 24);

            this.chkSaveData.Location = new Point(20, 300);
            this.chkSaveData.Text = "Save Lock-In Data to CSV";
            this.chkSaveData.AutoSize = true;
            this.chkSaveData.Checked = false;

            // Buttons & Labels
            this.btnStart.Location = new Point(20, 330);
            this.btnStart.Size = new Size(100, 40);
            this.btnStart.Text = "Start";

            this.btnStop.Location = new Point(130, 330);
            this.btnStop.Size = new Size(100, 40);
            this.btnStop.Text = "Stop";

            this.lblAled.Location = new Point(20, 390);
            this.lblAled.Text = "A_LED = 0.000 V";
            this.lblAled.Font = new Font("Microsoft Sans Serif", 12F, FontStyle.Bold);
            this.lblAled.AutoSize = true;

            this.lblAledRange.Location = new Point(20, 425);
            this.lblAledRange.Text = "A_LED (L-B) = 0.000 V";
            this.lblAledRange.Font = new Font("Microsoft Sans Serif", 12F, FontStyle.Bold);
            this.lblAledRange.AutoSize = true;

            // Setup Charts in a 2x2 Grid on the right side
            SetupChart(this.chartDrive, "Drive Signal (AI0)", SeriesChartType.Line, 280, 20, "Samples", "Voltage (V)");
            SetupChart(this.chartDetected, "Detected Signal (AI1)", SeriesChartType.Line, 700, 20, "Samples", "Voltage (V)");
            SetupChart(this.chartMultiplied, "Multiplied Signal (AI0 x AI1)", SeriesChartType.Line, 280, 350, "Samples", "Product (V^2)");
            SetupChart(this.chartLockIn, "Lock-In Signal (A_LED) vs Time", SeriesChartType.Point, 700, 350, "Time (Ticks)", "A_LED (V)");

            // Form Configurations
            this.ClientSize = new Size(1120, 700);
            this.Controls.AddRange(new Control[] { this.cmbDevice, this.chkSaveData, this.numSampleRate, this.numSamples,
                                                   this.numLedAmp, this.numLedFreq, this.numLedOffset, this.numTimerInterval,
                                                   this.btnStart, this.btnStop, this.lblAled, this.lblAledRange,
                                                   this.chartDrive, this.chartDetected, this.chartMultiplied, this.chartLockIn,
                                                   lblDevice, lblInput1, lblInput2, lblInput3, lblInput4, lblInput5, lblInput6 });
            this.Name = "Form1";
            this.Text = "Lock-In Photodiode Measurement";

            ((ISupportInitialize)(this.numSampleRate)).EndInit();
            ((ISupportInitialize)(this.numSamples)).EndInit();
            ((ISupportInitialize)(this.numLedAmp)).EndInit();
            ((ISupportInitialize)(this.numLedFreq)).EndInit();
            ((ISupportInitialize)(this.numLedOffset)).EndInit();
            ((ISupportInitialize)(this.numTimerInterval)).EndInit();
            ((ISupportInitialize)(this.chartDrive)).EndInit();
            ((ISupportInitialize)(this.chartDetected)).EndInit();
            ((ISupportInitialize)(this.chartMultiplied)).EndInit();
            ((ISupportInitialize)(this.chartLockIn)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();
        }

        private void SetupChart(Chart chart, string title, SeriesChartType type, int x, int y, string xTitle, string yTitle)
        {
            ChartArea area = new ChartArea("ChartArea1");
            area.AxisX.Title = xTitle;
            area.AxisY.Title = yTitle;
            area.AxisX.LabelStyle.Format = "0";
            area.AxisY.LabelStyle.Format = "0.00";

            Series series = new Series("Series1") { ChartArea = "ChartArea1", ChartType = type };
            chart.ChartAreas.Add(area);
            chart.Series.Add(series);
            chart.Titles.Add(title);
            chart.Location = new Point(x, y);
            chart.Size = new Size(400, 300);
        }

        private void LoadConnectedDevices()
        {
            try
            {
                string[] devices = DaqSystem.Local.Devices;
                if (devices.Length > 0)
                {
                    cmbDevice.Items.AddRange(devices);
                    cmbDevice.SelectedIndex = 0;
                }
                else
                {
                    cmbDevice.Items.Add("No devices found");
                    cmbDevice.SelectedIndex = 0;
                    btnStart.Enabled = false;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Could not load DAQ devices: " + ex.Message);
            }
        }

        private void InitializeTasksAndWaveform()
        {
            try
            {
                string dev = cmbDevice.SelectedItem.ToString();
                double rate = (double)numSampleRate.Value;
                int samples = (int)numSamples.Value;

                // 1. Setup Analog Output Task (Continuous)
                outTask = new Task();
                outTask.AOChannels.CreateVoltageChannel($"{dev}/ao0", "", -10, 10, AOVoltageUnits.Volts);
                outTask.Timing.ConfigureSampleClock("", rate, SampleClockActiveEdge.Rising, SampleQuantityMode.ContinuousSamples, samples);

                writerAO0 = new AnalogSingleChannelWriter(outTask.Stream);

                // Generate Sine Wave Buffer for the LED
                double freq = (double)numLedFreq.Value;
                double amp = (double)numLedAmp.Value;
                double offset = (double)numLedOffset.Value;
                double[] waveform = new double[samples];

                for (int i = 0; i < samples; i++)
                {
                    waveform[i] = offset + amp * Math.Sin(2 * Math.PI * freq * (i / rate));
                }

                // Write buffer to card and start continuous generation
                writerAO0.WriteMultiSample(false, waveform);
                outTask.Start();

                // 2. Setup Analog Input Task (Finite)
                inTask = new Task();
                inTask.AIChannels.CreateVoltageChannel($"{dev}/ai0", "", AITerminalConfiguration.Differential, -10.0, 10.0, AIVoltageUnits.Volts);
                inTask.AIChannels.CreateVoltageChannel($"{dev}/ai1", "", AITerminalConfiguration.Differential, -10.0, 10.0, AIVoltageUnits.Volts);
                inTask.Timing.ConfigureSampleClock("", rate, SampleClockActiveEdge.Rising, SampleQuantityMode.FiniteSamples, samples);

                readerAI = new AnalogMultiChannelReader(inTask.Stream);
            }
            catch (DaqException ex)
            {
                MessageBox.Show("DAQ Initialization Error: " + ex.Message);
            }
        }

        private void BtnStart_Click(object sender, EventArgs e)
        {
            if (cmbDevice.SelectedItem.ToString() == "No devices found") return;

            if (chkSaveData.Checked)
            {
                using (SaveFileDialog sfd = new SaveFileDialog())
                {
                    sfd.Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*";
                    sfd.Title = "Save Lock-In Data";
                    sfd.FileName = "lockin_data.csv";

                    if (sfd.ShowDialog() == DialogResult.OK)
                    {
                        fileWriter = new StreamWriter(sfd.FileName);
                        fileWriter.WriteLine("Tick,S_mean,A_LED");
                    }
                    else return;
                }
            }
            else fileWriter = null;

            chartDrive.Series[0].Points.Clear();
            chartDetected.Series[0].Points.Clear();
            chartMultiplied.Series[0].Points.Clear();
            chartLockIn.Series[0].Points.Clear();
            measurementTick = 0;

            InitializeTasksAndWaveform();

            timer1.Interval = (int)numTimerInterval.Value;
            timer1.Start();
        }

        private void Timer1_Tick(object sender, EventArgs e)
        {
            try
            {
                int samples = (int)numSamples.Value;
                double a_g = (double)numLedAmp.Value;

                // Execute finite read
                double[,] data = readerAI.ReadMultiSample(samples);
                double sumMult = 0;
                double minDetectSig = double.MaxValue;
                double maxDetectSig = double.MinValue;

                chartDrive.Series[0].Points.Clear();
                chartDetected.Series[0].Points.Clear();
                chartMultiplied.Series[0].Points.Clear();

                // Process arrays and calculate multiplication
                for (int i = 0; i < samples; i++)
                {
                    double driveSig = data[0, i]; // AI0
                    double detectSig = data[1, i]; // AI1
                    double multSig = driveSig * detectSig;

                    // Plot raw scopes (optional: downsample for UI performance if sample size is huge)
                    chartDrive.Series[0].Points.AddY(driveSig);
                    chartDetected.Series[0].Points.AddY(detectSig);
                    chartMultiplied.Series[0].Points.AddY(multSig);

                    sumMult += multSig;
                    if (detectSig < minDetectSig) minDetectSig = detectSig;
                    if (detectSig > maxDetectSig) maxDetectSig = detectSig;
                }

                // Lock-In Calculation
                double s_mean = sumMult / samples;
                double a_led = (2.0 * s_mean) / a_g;
                double a_led_range = (maxDetectSig - minDetectSig);

                lblAled.Text = $"A_LED = {a_led:F4} V";
                lblAledRange.Text = $"A_LED (L-B) = {a_led_range:F4} V";

                chartLockIn.Series[0].Points.AddXY(measurementTick, a_led);

                if (fileWriter != null)
                {
                    fileWriter.WriteLine($"{measurementTick},{s_mean:F6},{a_led:F6}");
                }

                measurementTick++;
            }
            catch (DaqException ex)
            {
                timer1.Stop();
                MessageBox.Show("DAQ Execution Error: " + ex.Message);
            }
        }

        private void BtnStop_Click(object sender, EventArgs e)
        {
            timer1.Stop();

            if (fileWriter != null)
            {
                fileWriter.Close();
                fileWriter.Dispose();
            }

            // Zero out LED before disposing
            try { writerAO0?.WriteSingleSample(true, 0); } catch { }

            DisposeDAQ();
        }

        private void DisposeDAQ()
        {
            outTask?.Dispose();
            inTask?.Dispose();
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            BtnStop_Click(this, EventArgs.Empty);
            base.OnFormClosing(e);
        }
    }
}
