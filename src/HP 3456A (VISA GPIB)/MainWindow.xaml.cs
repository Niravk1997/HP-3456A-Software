using System;
using System.Collections.Generic;
using System.Collections.Concurrent;
using System.IO;
using System.IO.Ports;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Timers;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Speech.Synthesis;
using System.Windows.Threading;
using System.Globalization;
using System.Reflection;
using Ivi.Visa;
using Ivi.Visa.FormattedIO;

namespace HP_3456A
{
    public static class GPIB_Address_Info
    {
        public static bool isConnected = false;

        //HP 3456A GPIB Device Info
        public static string GPIB_Address;

        public static string folder_Directory;
    }

    public partial class MainWindow : Window
    {
        //Reference to the graph window
        DateTime_Graph_Window HP3456A_DateTime_Graph_Window;

        //Reference to the graph window
        Graphing_Window HP3456A_Graph_Window;

        //Reference to the N graph Window
        N_Sample_Graph_Window HP3456A_N_Graph_Window;

        //Reference to Measurement Table
        Measurement_Data_Table HP3456A_Table;
        string Current_Measurement_Unit = "VDC";

        //HP34401A GPIB connection
        public IMessageBasedSession session;
        public MessageBasedFormattedIO formattedIO;

        public int GPIB_Lock = 0;

        //Which Measurement is currently selected
        int Measurement_Selected = 0;
        int Selected_Measurement_type = 0;
        int NPLC_indicator = 4;
        int Filter_indicator = 0;
        int AutoZero_indicator = 1;
        //VDC = 0
        //VAC = 1
        //2Ohm = 2
        //4Ohm = 3
        //ACDCV = 4
        //DCV_DCV = 5
        //ACV_DCV = 6
        //ACV_DCV_DCV = 7
        //OCTwoOhms = 8
        //OCFourOhms = 9

        //All Serial Write Commands are stored in this queue
        BlockingCollection<string> SerialWriteQueue = new BlockingCollection<string>();

        //Clear Logs after this count
        int Auto_Clear_Output_Log_Count = 20;

        //Lets the function know the queue has data
        bool isUserSendCommand = false;
        //if default Ndigit is not selected
        //remove 3 digits from measurement
        bool NDigit3 = false;
        //remove 2 digit from measurement
        bool NDigit4 = false;
        //if user only wants to read data, for fast sampling
        bool isSamplingOnly = false;
        bool isUpdateSpeed_Changed = false;

        //User decides whether to save data to text file or not
        //to save output log or not
        bool saveOutputLog = false;
        //to save measurements or not
        bool saveMeasurements = false;
        //to add data to table
        bool save_to_Table = false;
        //to add data to graphs
        bool save_to_Graph = false;
        bool Save_to_N_Graph = false;
        bool Save_to_DateTime_Graph = false;

        //Data is stored in these queues, waiting for it to be written to text files
        BlockingCollection<string> save_data_DCV = new BlockingCollection<string>();
        BlockingCollection<string> save_data_ACV = new BlockingCollection<string>();
        BlockingCollection<string> save_data_ACDCV = new BlockingCollection<string>();
        BlockingCollection<string> save_data_2Ohm = new BlockingCollection<string>();
        BlockingCollection<string> save_data_4Ohm = new BlockingCollection<string>();
        BlockingCollection<string> save_data_DCV_DCV = new BlockingCollection<string>();
        BlockingCollection<string> save_data_ACV_DCV = new BlockingCollection<string>();
        BlockingCollection<string> save_data_ACV_DCV_DCV = new BlockingCollection<string>();
        BlockingCollection<string> save_data_OCTwoOhms = new BlockingCollection<string>();
        BlockingCollection<string> save_data_OCFourOhms = new BlockingCollection<string>();

        //Options for Speech Synthesizer
        SpeechSynthesizer Voice = new SpeechSynthesizer();
        int Speech_Value_Precision = 1;
        int isSpeechActive = 0;
        int isSpeechContinuous = 0;
        int isSpeechMIN = 0;
        int isSpeechMAX = 0;
        double Speech_Continuous_Voice_Value = 0;
        double Speech_min_value = 0;
        double Speech_max_value = 0;

        //Default border color for when a switch is selected or not
        SolidColorBrush Selected = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#00CE30"));
        SolidColorBrush Deselected = new SolidColorBrush((Color)ColorConverter.ConvertFromString("White"));

        //Options for Measurement Data sampling speed
        double UpdateSpeed = 1000;

        //COM Select Window
        GPIB_Select_Window GPIB_Select;

        //Timer for getting data from multimeter at specified update speed.
        private System.Timers.Timer Speech_MIN_Max;
        private System.Timers.Timer Speech_Measurement_Interval;
        private System.Timers.Timer DataTimer;
        private DispatcherTimer runtime_Timer;
        private DispatcherTimer Process_Data;
        private System.Timers.Timer saveMeasurements_Timer;

        //Allow data timer to get data from multimeter or not
        bool DataSampling = false;

        //Data is stored here for display
        BlockingCollection<string> measurements = new BlockingCollection<string>();
        int Total_Samples = 0;
        int Invalid_Samples = 0;

        //Display Measurement as
        bool Original_Display = true;

        //Calculate Runtime from this
        DateTime StartDateTime;

        //Min, Max, Avg values
        //Program will compare input values to these value
        //and update these values
        decimal min = 0;
        decimal max = 0;
        decimal avg = 0;
        int AVG_Calculate = 1;
        int avg_count = 0;
        int avg_factor = 1000;
        int avg_resolution = 5;
        int resetMinMaxAvg = 1;

        public MainWindow()
        {
            InitializeComponent();
            if (Thread.CurrentThread.CurrentCulture.Name != "en-US")
            {
                Thread.CurrentThread.CurrentCulture = CultureInfo.CreateSpecificCulture("en-US");
                Thread.CurrentThread.CurrentUICulture = CultureInfo.CreateSpecificCulture("en-US");
                insert_Log("Culture set to english-US, decimal numbers will use dot as the seperator.", 0);
                insert_Log("Write decimal values with a dot as a seperator, not a comma.", 2);
            }
            Create_GetDataTimer();
            General_Timer();
            SetupSpeechSythesis();
            Check_Speech_MIN_MAX_Timer();
            Continuous_Voice_Measurement();
            Save_measurements_to_files_Timer();
            Load_Main_Window_Settings();
            insert_Log("Click the Config Menu then click Connect.", 5);
            insert_Log("VISA Compatible GPIB Adapter required.", 5);
            insert_Log("NI-VISA 20.0 and Keysight IO Libraries Suite 2021.", 5);
            insert_Log("If using Keysight 82357B, then Set NI-VISA as primary VISA.", 5);
            insert_Log("This software only works with HP 3456A.", 5);
        }

        private void Save_measurements_to_files_Timer()
        {
            saveMeasurements_Timer = new System.Timers.Timer();
            saveMeasurements_Timer.Interval = 60000; //Default is 1 minute;
            saveMeasurements_Timer.AutoReset = false;
            saveMeasurements_Timer.Enabled = false;
            saveMeasurements_Timer.Elapsed += Save_MeasurementData_to_files;
        }

        private void Save_MeasurementData_to_files(Object source, ElapsedEventArgs e)
        {
            string Date = DateTime.UtcNow.ToString("yyyy-MM-dd");
            int VDC_Count = save_data_DCV.Count;
            int ACV_Count = save_data_ACV.Count;
            int TwoOhm_Count = save_data_2Ohm.Count;
            int FourOhm_Count = save_data_4Ohm.Count;
            int ACDCV_Count = save_data_ACDCV.Count;
            int DCV_DCV_Count = save_data_DCV_DCV.Count;
            int ACV_DCV_Count = save_data_ACV_DCV.Count;
            int ACV_DCV_DCV_Count = save_data_ACV_DCV_DCV.Count;
            int OCTwoOhms_Count = save_data_OCTwoOhms.Count;
            int OCFourOhms_Count = save_data_OCFourOhms.Count;

            if (VDC_Count > 0)
            {
                try
                {
                    using (TextWriter datatotxt = new StreamWriter(GPIB_Address_Info.folder_Directory + @"\" + "DCV" + @"\" + Date + "_" + GPIB_Address_Info.GPIB_Address.Replace(":", "") + "_DCV.txt", true))
                    {
                        for (int i = 0; i < VDC_Count; i++)
                        {
                            datatotxt.WriteLine(save_data_DCV.Take());
                        }
                    }
                }
                catch (Exception)
                {
                    insert_Log("Cannot save DCV measurements to text file.", 1);
                }
            }

            if (ACV_Count > 0)
            {
                try
                {
                    using (TextWriter datatotxt = new StreamWriter(GPIB_Address_Info.folder_Directory + @"\" + "ACV" + @"\" + Date + "_" + GPIB_Address_Info.GPIB_Address.Replace(":", "") + "_ACV.txt", true))
                    {
                        for (int i = 0; i < ACV_Count; i++)
                        {
                            datatotxt.WriteLine(save_data_ACV.Take());
                        }
                    }
                }
                catch (Exception)
                {
                    insert_Log("Cannot save ACV measurements to text file.", 1);
                }
            }

            if (TwoOhm_Count > 0)
            {
                try
                {
                    using (TextWriter datatotxt = new StreamWriter(GPIB_Address_Info.folder_Directory + @"\" + "2WireOhms" + @"\" + Date + "_" + GPIB_Address_Info.GPIB_Address.Replace(":", "") + "_2WireOhms.txt", true))
                    {
                        for (int i = 0; i < TwoOhm_Count; i++)
                        {
                            datatotxt.WriteLine(save_data_2Ohm.Take());
                        }
                    }
                }
                catch (Exception)
                {
                    insert_Log("Cannot save 2 Wire Ohms measurements to text file.", 1);
                }
            }

            if (FourOhm_Count > 0)
            {
                try
                {
                    using (TextWriter datatotxt = new StreamWriter(GPIB_Address_Info.folder_Directory + @"\" + "4WireOhms" + @"\" + Date + "_" + GPIB_Address_Info.GPIB_Address.Replace(":", "") + "_4WireOhms.txt", true))
                    {
                        for (int i = 0; i < FourOhm_Count; i++)
                        {
                            datatotxt.WriteLine(save_data_4Ohm.Take());
                        }
                    }
                }
                catch (Exception)
                {
                    insert_Log("Cannot save 4 Wire Ohms measurements to text file.", 1);
                }
            }

            if (ACDCV_Count > 0)
            {
                try
                {
                    using (TextWriter datatotxt = new StreamWriter(GPIB_Address_Info.folder_Directory + @"\" + "ACDCV" + @"\" + Date + "_" + GPIB_Address_Info.GPIB_Address.Replace(":", "") + "_ACDCV.txt", true))
                    {
                        for (int i = 0; i < ACDCV_Count; i++)
                        {
                            datatotxt.WriteLine(save_data_ACDCV.Take());
                        }
                    }
                }
                catch (Exception)
                {
                    insert_Log("Cannot save ACV + DCV measurements to text file.", 1);
                }
            }

            if (DCV_DCV_Count > 0)
            {
                try
                {
                    using (TextWriter datatotxt = new StreamWriter(GPIB_Address_Info.folder_Directory + @"\" + "DCV_DCV" + @"\" + Date + "_" + GPIB_Address_Info.GPIB_Address.Replace(":", "") + "_DCV_DCV.txt", true))
                    {
                        for (int i = 0; i < DCV_DCV_Count; i++)
                        {
                            datatotxt.WriteLine(save_data_DCV_DCV.Take());
                        }
                    }
                }
                catch (Exception)
                {
                    insert_Log("Cannot save DCV/DCV measurements to text file.", 1);
                }
            }

            if (ACV_DCV_Count > 0)
            {
                try
                {
                    using (TextWriter datatotxt = new StreamWriter(GPIB_Address_Info.folder_Directory + @"\" + "ACV_DCV" + @"\" + Date + "_" + GPIB_Address_Info.GPIB_Address.Replace(":", "") + "_ACV_DCV.txt", true))
                    {
                        for (int i = 0; i < ACV_DCV_Count; i++)
                        {
                            datatotxt.WriteLine(save_data_ACV_DCV.Take());
                        }
                    }
                }
                catch (Exception)
                {
                    insert_Log("Cannot save ACV/DCV measurements to text file.", 1);
                }
            }

            if (ACV_DCV_DCV_Count > 0)
            {
                try
                {
                    using (TextWriter datatotxt = new StreamWriter(GPIB_Address_Info.folder_Directory + @"\" + "ACV_DCV_DCV" + @"\" + Date + "_" + GPIB_Address_Info.GPIB_Address.Replace(":", "") + "_ACV_DCV_DCV.txt", true))
                    {
                        for (int i = 0; i < ACV_DCV_DCV_Count; i++)
                        {
                            datatotxt.WriteLine(save_data_ACV_DCV_DCV.Take());
                        }
                    }
                }
                catch (Exception)
                {
                    insert_Log("Cannot save ACV + DCV/DCV measurements to text file.", 1);
                }
            }

            if (OCTwoOhms_Count > 0)
            {
                try
                {
                    using (TextWriter datatotxt = new StreamWriter(GPIB_Address_Info.folder_Directory + @"\" + "2Wire_OC_Ohms" + @"\" + Date + "_" + GPIB_Address_Info.GPIB_Address.Replace(":", "") + "_2Wire_OC_Ohms.txt", true))
                    {
                        for (int i = 0; i < OCTwoOhms_Count; i++)
                        {
                            datatotxt.WriteLine(save_data_OCTwoOhms.Take());
                        }
                    }
                }
                catch (Exception)
                {
                    insert_Log("Cannot save OC 2 Wire Ohms measurements to text file.", 1);
                }
            }

            if (OCFourOhms_Count > 0)
            {
                try
                {
                    using (TextWriter datatotxt = new StreamWriter(GPIB_Address_Info.folder_Directory + @"\" + "4Wire_OC_Ohms" + @"\" + Date + "_" + GPIB_Address_Info.GPIB_Address.Replace(":", "") + "_4Wire_OC_Ohms.txt", true))
                    {
                        for (int i = 0; i < OCFourOhms_Count; i++)
                        {
                            datatotxt.WriteLine(save_data_OCFourOhms.Take());
                        }
                    }
                }
                catch (Exception)
                {
                    insert_Log("Cannot save OC 4 Wire Ohms measurements to text file.", 1);
                }
            }

            saveMeasurements_Timer.Enabled = true;
            if (saveMeasurements == false)
            {
                while (save_data_DCV.TryTake(out _)) { }
                while (save_data_ACV.TryTake(out _)) { }
                while (save_data_2Ohm.TryTake(out _)) { }
                while (save_data_4Ohm.TryTake(out _)) { }
                while (save_data_ACDCV.TryTake(out _)) { }
                while (save_data_DCV_DCV.TryTake(out _)) { }
                while (save_data_ACV_DCV.TryTake(out _)) { }
                while (save_data_ACV_DCV_DCV.TryTake(out _)) { }
                while (save_data_OCTwoOhms.TryTake(out _)) { }
                while (save_data_OCFourOhms.TryTake(out _)) { }
                saveMeasurements_Timer.Enabled = false;
                saveMeasurements_Timer.Stop();
                insert_Log("Save Measurements Queues Cleared.", 0);
            }
        }

        private void SetupSpeechSythesis()
        {
            Voice.Volume = 100;
            Voice.SelectVoiceByHints(VoiceGender.Male);
            Voice.Rate = 1;
        }

        private void General_Timer()
        {
            runtime_Timer = new DispatcherTimer();
            runtime_Timer.Interval = TimeSpan.FromSeconds(1);
            runtime_Timer.Tick += runtime_Update;
            runtime_Timer.Start();
        }

        public void GPIB_COM_Selected()
        {
            if (GPIB_Address_Info.isConnected == true)
            {
                Connect.IsEnabled = false;
                unlockControls();
                GPIB_Connect();
                this.Title = "HP 3456A " + GPIB_Address_Info.GPIB_Address;
                DataSampling = true;
                saveOutputLog = true;
                saveMeasurements = true;
                Stop_Sampling.IsEnabled = true;
                DataTimer.Enabled = true;
                StartDateTime = DateTime.Now;
                Data_process();
                saveMeasurements_Timer.Enabled = true;
                Sampling_Only.IsEnabled = true;
                Local_Exit.IsEnabled = true;
                DataLogger.IsEnabled = true;
            }
        }

        private void runtime_Update(object sender, EventArgs e)
        {
            InvalidSamples_Total.Content = Invalid_Samples.ToString();
            Samples_Total.Content = Total_Samples.ToString();
            if (DataSampling == true)
            {
                Runtime_Timer.Content = GetTimeSpan();
            }
        }

        private void Continuous_Voice_Measurement()
        {
            Speech_Measurement_Interval = new System.Timers.Timer();
            Speech_Measurement_Interval.Interval = 60000; //Default is 1 minute;
            Speech_Measurement_Interval.AutoReset = false;
            Speech_Measurement_Interval.Enabled = false;
            Speech_Measurement_Interval.Elapsed += Check_Continuous_Voice_Measurement;
        }

        private void Check_Continuous_Voice_Measurement(Object source, ElapsedEventArgs e)
        {
            if (isSpeechActive == 1)
            {
                try
                {
                    if (isSpeechContinuous == 1)
                    {
                        if (Speech_Continuous_Voice_Value > 999999999)
                        {
                            Voice.Speak("Overload" + " " + MeasurementType_String());
                        }
                        else
                        {
                            Voice.Speak((decimal)Math.Round(Speech_Continuous_Voice_Value, Speech_Value_Precision) + " " + MeasurementType_String());
                        }
                        Speech_Measurement_Interval.Enabled = true;
                    }
                }
                catch (Exception)
                {
                    insert_Log("Speech Synthesizer Continuous Voice measurement feature failed.", 1);
                    insert_Log("Don't worry. Trying again.", 2);
                    Speech_Measurement_Interval.Enabled = true;
                }
            }
            else
            {
                Interlocked.Exchange(ref isSpeechContinuous, 0);
            }
        }

        private void Check_Speech_MIN_MAX_Timer()
        {
            Speech_MIN_Max = new System.Timers.Timer();
            Speech_MIN_Max.Interval = 1000;
            Speech_MIN_Max.AutoReset = false;
            Speech_MIN_Max.Enabled = false;
            Speech_MIN_Max.Elapsed += Check_Speech_MIN_MAX;
        }

        private void Check_Speech_MIN_MAX(Object source, ElapsedEventArgs e)
        {
            if (isSpeechActive == 1)
            {
                try
                {
                    if (isSpeechMAX == 1)
                    {
                        if (Speech_max_value <= (double)max)
                        {
                            Voice.Speak("Warning, maximum value of " + (decimal)Math.Round(Speech_max_value, Speech_Value_Precision) + " " + MeasurementType_String() + " reached.");
                            if ((double)max > 999999999)
                            {
                                Voice.Speak("maximum value is " + "overload" + " " + MeasurementType_String());
                            }
                            else
                            {
                                Voice.Speak("maximum value is " + (decimal)Math.Round(max, Speech_Value_Precision) + " " + MeasurementType_String());
                            }
                        }
                    }
                    if (isSpeechMIN == 1)
                    {
                        if (Speech_min_value >= (double)min)
                        {
                            Voice.Speak("Warning, minimum value of " + (decimal)Math.Round(Speech_min_value, Speech_Value_Precision) + " " + MeasurementType_String() + " reached.");
                            Voice.Speak("minimum value is " + (decimal)Math.Round(min, Speech_Value_Precision) + " " + MeasurementType_String());
                        }
                    }
                    Speech_MIN_Max.Enabled = true;
                }
                catch (Exception)
                {
                    insert_Log("Speech Synthesizer MIN and MAX feature failed.", 1);
                    insert_Log("Don't worry. Trying again.", 2);
                    Speech_MIN_Max.Enabled = true;
                }
            }
            if (isSpeechMAX == 0 & isSpeechMIN == 0)
            {
                Speech_MIN_Max.Enabled = false;
                Speech_MIN_Max.Stop();
            }
        }

        private string MeasurementType_String()
        {
            switch (Selected_Measurement_type)
            {
                case 0:
                    return "volts DC";
                case 1:
                    return "volts AC";
                case 2:
                    return "ohms";
                case 3:
                    return "ohms";
                case 4:
                    return "volts AC";
                case 5:
                    return "volts DC";
                case 6:
                    return "volts AC";
                case 7:
                    return "volts AC";
                case 8:
                    return "ohms";
                case 9:
                    return "ohms";
                default:
                    return "value";
            }
        }

        private (string, string) MeasurementUnit_String()
        {
            switch (Selected_Measurement_type)
            {
                case 0:
                    return ("VDC", "DCV Voltage");
                case 1:
                    return ("VAC", "ACV Voltage");
                case 2:
                    return ("Ω", "Ω 2Wire Ohms");
                case 3:
                    return ("Ω", "Ω 4Wire Ohms");
                case 4:
                    return ("VAC", "ACV + DCV Voltage");
                case 5:
                    return ("VDC", "DCV/DCV Voltage");
                case 6:
                    return ("VAC", "ACV/DCV Voltage");
                case 7:
                    return ("VAC", "ACV + DCV/DCV Voltage");
                case 8:
                    return ("Ω", "OC Ω 2Wire Ohms");
                case 9:
                    return ("Ω", "OC Ω 4Wire Ohms");
                default:
                    return ("Unk", "Unknown");
            }
        }

        private string GetTimeSpan()
        {
            TimeSpan span = (DateTime.Now - StartDateTime);
            return (String.Format("{0:00}:{1:00}:{2:00}", span.Hours, span.Minutes, span.Seconds));
        }

        private void unlockControls()
        {
            Application.Current.Dispatcher.Invoke(DispatcherPriority.Background, new ThreadStart(delegate
            {
                Measurements.IsEnabled = true;
                Range.IsEnabled = true;
                Resolution_Config.IsEnabled = true;
                Trigger_Config.IsEnabled = true;
                UpdateSpeed_Box.IsEnabled = true;
            }));
        }

        private void lockControls()
        {
            Application.Current.Dispatcher.Invoke(DispatcherPriority.Background, new ThreadStart(delegate
            {
                Measurements.IsEnabled = false;
                Range.IsEnabled = false;
                Resolution_Config.IsEnabled = false;
                Trigger_Config.IsEnabled = false;
                UpdateSpeed_Box.IsEnabled = false;
            }));
        }

        private void Speedup_Interval()
        {
            if (UpdateSpeed > 2000)
            {
                DataTimer.Interval = 0.01;
            }
        }

        private void Restore_Interval()
        {
            DataTimer.Interval = UpdateSpeed;
        }

        private void GPIB_Connect()
        {
            session = GlobalResourceManager.Open(GPIB_Address_Info.GPIB_Address, (AccessModes)GPIB_Lock, 10000) as IMessageBasedSession;
            session.TimeoutMilliseconds = 20000;
            session.TerminationCharacterEnabled = true;
            formattedIO = new MessageBasedFormattedIO(session);
            formattedIO.ReadBufferSize = 8192;
            formattedIO.WriteBufferSize = 8192;
        }

        private void GPIB_Reconnect()
        {
            try
            {
                insert_Log("Trying to reestablish GPIB Connection, please Wait.", 2);
                formattedIO = null;
                if (GPIB_Lock == 1)
                {
                    session.UnlockResource();
                }
                session.Dispose();
                Thread.Sleep(10000);
                session = GlobalResourceManager.Open(GPIB_Address_Info.GPIB_Address, (AccessModes)GPIB_Lock, 10000) as IMessageBasedSession;
                session.TimeoutMilliseconds = 20000;
                session.TerminationCharacterEnabled = true;
                formattedIO = new MessageBasedFormattedIO(session);
                formattedIO.ReadBufferSize = 8192;
                formattedIO.WriteBufferSize = 8192;
                insert_Log("GPIB Reconnect successful.", 0);
            }
            catch (Exception Ex)
            {
                insert_Log(Ex.Message, 1);
                insert_Log("GPIB Reconnect failed.", 1);
            }
        }

        private string Query(string Command)
        {
            formattedIO.WriteLine(Command);
            return formattedIO.ReadLine();
        }

        private string Read()
        {
            return formattedIO.ReadLine();
        }

        private void Write(string Command)
        {
            formattedIO.WriteLine(Command);
        }

        private void Create_GetDataTimer()
        {
            DataTimer = new System.Timers.Timer();
            DataTimer.Interval = 1000;
            DataTimer.Elapsed += HP3456ACommunicateEvent;
            DataTimer.AutoReset = false;
        }

        private void Data_process()
        {
            Process_Data = new DispatcherTimer();
            Process_Data.Interval = TimeSpan.FromSeconds(0);
            Process_Data.Tick += DataProcessor;
            Process_Data.Start();
        }

        private void DataProcessor(object sender, EventArgs e)
        {
            while (measurements.Count > 0)
            {
                try
                {
                    string measurement = measurements.Take();
                    decimal value = decimal.Parse(measurement, NumberStyles.Float);
                    DisplayData(measurement, value);
                    Display_MIN_MAX_AVG(value);
                    setContinuousVoiceMeasurement(value);
                }
                catch (Exception Ex)
                {
                    if (Show_Display_Error.IsChecked == true)
                    {
                        insert_Log(Ex.ToString(), 2);
                        insert_Log("Sample display process failed. Trying again.", 2);
                    }
                }
            }
            Process_Data.Stop();
        }

        private void setContinuousVoiceMeasurement(decimal value)
        {
            if (isSpeechContinuous == 1)
            {
                Interlocked.Exchange(ref Speech_Continuous_Voice_Value, (double)value);
            }
        }

        private void DisplayData(string measurement, decimal value)
        {
            if (Original_Display == true)
            {
                Measurement_Value.Content = measurement.Substring(0, measurement.IndexOf("E"));
                string unit = measurement.Substring(measurement.Length - 2);
                switch (unit)
                {
                    case "-9":
                        Measurement_Value.Content = "OVLD";
                        if (Measurement_Selected == 2 || Measurement_Selected == 3 || Measurement_Selected == 8 || Measurement_Selected == 9)
                        {
                            Measurement_Scale.Content = "M";
                        }
                        else
                        {
                            Measurement_Scale.Content = "";
                        }
                        break;
                    case "-6":
                        Measurement_Scale.Content = "μ";
                        break;
                    case "-3":
                        Measurement_Scale.Content = "m";
                        break;
                    case "+3":
                        Measurement_Scale.Content = "K";
                        break;
                    case "+6":
                        Measurement_Scale.Content = "M";
                        break;
                    case "+9":
                        Measurement_Value.Content = "OVLD";
                        if (Measurement_Selected == 2 || Measurement_Selected == 3 || Measurement_Selected == 8 || Measurement_Selected == 9)
                        {
                            Measurement_Scale.Content = "M";
                        }
                        else
                        {
                            Measurement_Scale.Content = "";
                        }
                        break;
                    default:
                        Measurement_Scale.Content = "";
                        break;
                }
            }
            else
            {
                if (value > -1 & value < 1)
                {
                    Measurement_Value.Content = (double)(value * 1000);
                    Measurement_Scale.Content = "m";
                }
                else if (value > 999999999)
                {
                    Measurement_Value.Content = "OVLD";
                    if (Measurement_Selected == 2 || Measurement_Selected == 3 || Measurement_Selected == 8 || Measurement_Selected == 9)
                    {
                        Measurement_Scale.Content = "M";
                    }
                    else
                    {
                        Measurement_Scale.Content = "";
                    }
                }
                else if (value > 999999)
                {
                    Measurement_Value.Content = (double)value / 1000000;
                    Measurement_Scale.Content = "M";
                }
                else if (value > 999)
                {
                    Measurement_Value.Content = (double)value / 1000;
                    Measurement_Scale.Content = "K";
                }
                else
                {
                    Measurement_Value.Content = value;
                    Measurement_Scale.Content = "";
                }
            }
        }

        private void Display_MIN_MAX_AVG(decimal measurement)
        {
            if (resetMinMaxAvg == 1)
            {
                min = measurement;
                max = measurement;
                avg = 0;
                avg_count = 0;
                insert_Log("Reset MIN, MAX, AVG values.", 0);
                updateMIN(measurement);
                updateMAX(measurement);
                Interlocked.Exchange(ref resetMinMaxAvg, 0);
            }
            if (measurement < min)
            {
                updateMIN(measurement);
            }
            if (measurement > max)
            {
                updateMAX(measurement);
            }
            if (AVG_Calculate == 1)
            {
                updateAVG(measurement);
            }
        }

        private void Reset_Click_MIN_MAX_AVG(object sender, MouseButtonEventArgs e)
        {
            Interlocked.Exchange(ref resetMinMaxAvg, 1);
            insert_Log("Reset MIN, MAX, AVG command has been send.", 4);
        }

        private void updateAVG(decimal measurement)
        {
            avg_count += 1;
            avg = avg + (measurement - avg) / Math.Min(avg_count, avg_factor);
            if (avg == 0)
            {
                AVG_Value.Content = (decimal)Math.Round(avg, avg_resolution);
                AVG_Scale.Content = "";
            }
            else if (avg < 1 & avg > -1)
            {
                AVG_Value.Content = (decimal)Math.Round((double)(avg * 1000), avg_resolution);
                AVG_Scale.Content = "m";
            }
            else if (avg < -99999999999 || avg > 99999999999)
            {
                AVG_Value.Content = "OVLD";
                if (Measurement_Selected == 2 || Measurement_Selected == 3 || Measurement_Selected == 8 || Measurement_Selected == 9)
                {
                    AVG_Scale.Content = "M";
                }
                else
                {
                    AVG_Scale.Content = "";
                }
            }
            else if (avg < -999999 || avg > 999999)
            {
                AVG_Value.Content = (decimal)Math.Round((double)avg / 1000000, avg_resolution);
                AVG_Scale.Content = "M";
            }
            else if (avg < -999 || avg > 999)
            {
                AVG_Value.Content = (decimal)Math.Round((double)avg / 1000, avg_resolution);
                AVG_Scale.Content = "K";
            }
            else
            {
                AVG_Value.Content = (decimal)Math.Round(avg, avg_resolution);
                AVG_Scale.Content = "";
            }
        }

        private void updateMIN(decimal measurement)
        {
            min = measurement;
            if (min == 0)
            {
                MIN_Value.Content = (decimal)(min);
                MIN_Scale.Content = "";
            }
            else if (min < 1 & min > -1)
            {
                MIN_Value.Content = (decimal)((double)(min * 1000));
                MIN_Scale.Content = "m";
            }
            else if (min < -99999999999 || min > 99999999999)
            {
                MIN_Value.Content = "OVLD";
                if (Measurement_Selected == 2 || Measurement_Selected == 3 || Measurement_Selected == 8 || Measurement_Selected == 9)
                {
                    MIN_Scale.Content = "M";
                }
                else
                {
                    MIN_Scale.Content = "";
                }
            }
            else if (min < -999999 || min > 999999)
            {
                MIN_Value.Content = (decimal)((double)min / 1000000);
                MIN_Scale.Content = "M";
            }
            else if (min < -999 || min > 999)
            {
                MIN_Value.Content = (decimal)((double)min / 1000);
                MIN_Scale.Content = "K";
            }
            else
            {
                MIN_Value.Content = (decimal)min;
                MIN_Scale.Content = "";
            }
        }

        private void updateMAX(decimal measurement)
        {
            max = measurement;
            if (max == 0)
            {
                MAX_Value.Content = (decimal)(max);
                MAX_Scale.Content = "";
            }
            else if (max < 1 & max > -1)
            {
                MAX_Value.Content = (decimal)((double)(max * 1000));
                MAX_Scale.Content = "m";
            }
            else if (max < -99999999999 || max > 99999999999)
            {
                MAX_Value.Content = "OVLD";
                if (Measurement_Selected == 2 || Measurement_Selected == 3 || Measurement_Selected == 8 || Measurement_Selected == 9)
                {
                    MAX_Scale.Content = "M";
                }
                else
                {
                    MAX_Scale.Content = "";
                }
            }
            else if (max < -999999 || max > 999999)
            {
                MAX_Value.Content = (decimal)((double)max / 1000000);
                MAX_Scale.Content = "M";
            }
            else if (max < -999 || max > 999)
            {
                MAX_Value.Content = (decimal)((double)max / 1000);
                MAX_Scale.Content = "K";
            }
            else
            {
                MAX_Value.Content = (decimal)max;
                MAX_Scale.Content = "";
            }
        }

        private void HP3456ACommunicateEvent(Object source, ElapsedEventArgs e)
        {
            try
            {
                if (isUserSendCommand == true)
                {
                    Serial_WriteQueue();
                    Measurement_Type_Select();
                    unlockControls();
                    isUserSendCommand = false;
                    if (UpdateSpeed > 2000)
                    {
                        Restore_Interval();
                    }
                }

                if (DataSampling == true)
                {
                    do
                    {
                        Read_Measurement();
                        Process_Data.Start();
                    } while (isSamplingOnly == true & DataSampling == true);
                }
                if (isUpdateSpeed_Changed == true)
                {
                    isUpdateSpeed_Changed = false;
                    insert_Log("Update Speed has been set to " + (UpdateSpeed / 1000) + " seconds.", 0);
                    DataTimer.Interval = UpdateSpeed;
                }
                DataTimer.Enabled = true;

            }
            catch (Exception Ex)
            {
                this.Dispatcher.Invoke(DispatcherPriority.ContextIdle, new ThreadStart(delegate
                {
                    if (Show_COM_Error.IsChecked == true)
                    {
                        insert_Log(Ex.Message, 2);
                        insert_Log("Could not get a measurement reading.", 2);
                        insert_Log("Don't worry. Trying again.", 2);
                        insert_Log("Slow the Update Speed if warning persists.", 2);
                    }
                }));
                if (isUpdateSpeed_Changed == true)
                {
                    isUpdateSpeed_Changed = false;
                    insert_Log("Update Speed has been set to " + (UpdateSpeed / 1000) + " seconds.", 0);
                    DataTimer.Interval = UpdateSpeed;
                }
                GPIB_Reconnect();
                DataTimer.Enabled = true;
            }
            finally
            {
                DataTimer.Enabled = true;
            }
        }

        private void Read_Measurement()
        {
            string data = Read().Trim();
            if (data.Length == 12)
            {
                if (NDigit3 == true)
                {
                    data = data.Remove(7, 2);
                    measurements.Add(data);
                }
                else if (NDigit4 == true)
                {
                    data = data.Remove(8, 1);
                    measurements.Add(data);
                }
                else
                {
                    measurements.Add(data);
                }
                Total_Samples++;
                if (saveMeasurements == true || save_to_Table == true || save_to_Graph == true || Save_to_N_Graph == true || Save_to_DateTime_Graph == true)
                {
                    Process_Measurement_Data(data);
                }
            }
            else
            {
                Invalid_Samples++;
            }
        }

        private void Process_Measurement_Data(string data)
        {
            string Date = DateTime.Now.ToString("yyyy-MM-dd h:mm:ss.fff tt");
            if (saveMeasurements == true)
            {
                switch (Selected_Measurement_type)
                {
                    case 0:
                        save_data_DCV.Add(Date + "," + data);
                        break;
                    case 1:
                        save_data_ACV.Add(Date + "," + data);
                        break;
                    case 2:
                        save_data_2Ohm.Add(Date + "," + data);
                        break;
                    case 3:
                        save_data_4Ohm.Add(Date + "," + data);
                        break;
                    case 4:
                        save_data_ACDCV.Add(Date + "," + data);
                        break;
                    case 5:
                        save_data_DCV_DCV.Add(Date + "," + data);
                        break;
                    case 6:
                        save_data_ACV_DCV.Add(Date + "," + data);
                        break;
                    case 7:
                        save_data_ACV_DCV_DCV.Add(Date + "," + data);
                        break;
                    case 8:
                        save_data_OCTwoOhms.Add(Date + "," + data);
                        break;
                    case 9:
                        save_data_OCFourOhms.Add(Date + "," + data);
                        break;
                    default:
                        insert_Log("Data was not saved. Something went wrong.", 0);
                        break;
                }
            }

            if (save_to_Table == true)
            {
                try
                {
                    HP3456A_Table.Table_Data_Queue.Add(Date + "," + data + "," + Current_Measurement_Unit);
                }
                catch (Exception)
                {
                    insert_Log("Could not add data to Table Window.", 2);
                    insert_Log("This could happen if the table window was opened or closed recently.", 2);
                }
            }

            if (save_to_Graph == true)
            {
                try
                {
                    HP3456A_Graph_Window.Data_Queue.Add(Date + "," + data);
                }
                catch (Exception)
                {
                    insert_Log("Could not add data to Graph Window.", 2);
                    insert_Log("This could happen if the Graph Window was opened or closed recently.", 2);
                }
            }

            if (Save_to_N_Graph == true)
            {
                try
                {
                    HP3456A_N_Graph_Window.Data_Queue.Add(Date + "," + data);
                }
                catch (Exception)
                {
                    insert_Log("Could not add data to N Sample Graph Window.", 2);
                    insert_Log("This could happen if the N Sample Graph Window was opened or closed recently.", 2);
                }
            }

            if (Save_to_DateTime_Graph == true)
            {
                try
                {
                    HP3456A_DateTime_Graph_Window.Data_Queue.Add(Date + "," + data);
                }
                catch (Exception)
                {
                    insert_Log("Could not add data to DateTime Graph Window.", 2);
                    insert_Log("This could happen if the DateTime Graph Window was opened or closed recently.", 2);
                }
            }
        }

        private void Serial_WriteQueue()
        {
            while (SerialWriteQueue.Count != 0)
            {
                string WriteCommand = SerialWriteQueue.Take();
                Serial_WriteProcess(WriteCommand);
            }
        }

        private void Serial_WriteProcess(string command)
        {
            switch (command)
            {
                case "0.01STI": //3
                    NPLC_indicator = 0;
                    NDigit3 = true;
                    NDigit4 = false;
                    Write(command);
                    break;
                case "0.1STI": //4
                    NPLC_indicator = 1;
                    NDigit3 = false;
                    NDigit4 = true;
                    Write(command);
                    break;
                case "1STI": //6
                    NPLC_indicator = 2;
                    NDigit3 = false;
                    NDigit4 = false;
                    Write(command);
                    break;
                case "10STI": //6
                    NPLC_indicator = 3;
                    NDigit3 = false;
                    NDigit4 = false;
                    Write(command);
                    break;
                case "100STI": //6
                    NPLC_indicator = 4;
                    NDigit3 = false;
                    NDigit4 = false;
                    Write(command);
                    break;
                case "S0F1R1":
                case "S0F2R1":
                case "S0F3R1":
                case "S0F4R1":
                case "S0F5R1":
                case "S1F1R1":
                case "S1F2R1":
                case "S1F3R1":
                    Write(command);
                    Interlocked.Exchange(ref resetMinMaxAvg, 1);
                    break;
                case "S1F4R1":
                case "S1F5R1":
                    Write(command);
                    Interlocked.Exchange(ref resetMinMaxAvg, 1);
                    break;
                case "FL1":
                    Filter_indicator = 1; //Filter On
                    Write(command);
                    break;
                case "FL0":
                    Filter_indicator = 0;
                    Write(command);
                    break;
                case "Z1":
                    AutoZero_indicator = 1;
                    Write(command);
                    break;
                case "Z0":
                    AutoZero_indicator = 0;
                    Write(command);
                    break;
                case "LOCAL_EXIT":
                    Application.Current.Dispatcher.Invoke(() => { Application.Current.Shutdown(); }, DispatcherPriority.Send);
                    break;
                default:
                    Write(command);
                    break;
            }
        }

        private void Measurement_Type_Select()
        {
            if (Measurement_Selected == 0)
            {
                Application.Current.Dispatcher.Invoke(DispatcherPriority.Background, new ThreadStart(delegate
                {
                    Measurement_Type.Content = "VDC";
                    Measurement_Scale.Content = "";
                    Measurement_Value.Content = "";
                    MIN_Type.Content = "VDC";
                    MAX_Type.Content = "VDC";
                    AVG_Type.Content = "VDC";
                    Current_Measurement_Unit = "VDC";
                }));
            }
            else if (Measurement_Selected == 1)
            {
                Application.Current.Dispatcher.Invoke(DispatcherPriority.Background, new ThreadStart(delegate
                {
                    Measurement_Type.Content = "VAC";
                    Measurement_Scale.Content = "";
                    Measurement_Value.Content = "";
                    MIN_Type.Content = "VAC";
                    MAX_Type.Content = "VAC";
                    AVG_Type.Content = "VAC";
                    Current_Measurement_Unit = "VAC";
                }));
            }
            else if (Measurement_Selected == 2)
            {
                Application.Current.Dispatcher.Invoke(DispatcherPriority.Background, new ThreadStart(delegate
                {
                    Measurement_Type.Content = "Ω";
                    Measurement_Scale.Content = "";
                    Measurement_Value.Content = "";
                    MIN_Type.Content = "Ω";
                    MAX_Type.Content = "Ω";
                    AVG_Type.Content = "Ω";
                    Current_Measurement_Unit = "Ω";
                }));
            }
            else if (Measurement_Selected == 3)
            {
                Application.Current.Dispatcher.Invoke(DispatcherPriority.Background, new ThreadStart(delegate
                {
                    Measurement_Type.Content = "Ω";
                    Measurement_Scale.Content = "";
                    Measurement_Value.Content = "";
                    MIN_Type.Content = "Ω";
                    MAX_Type.Content = "Ω";
                    AVG_Type.Content = "Ω";
                    Current_Measurement_Unit = "Ω";
                }));
            }
            else if (Measurement_Selected == 4)
            {
                Application.Current.Dispatcher.Invoke(DispatcherPriority.Background, new ThreadStart(delegate
                {
                    Measurement_Type.Content = "VAC";
                    Measurement_Scale.Content = "";
                    Measurement_Value.Content = "";
                    MIN_Type.Content = "VAC";
                    MAX_Type.Content = "VAC";
                    AVG_Type.Content = "VAC";
                    Current_Measurement_Unit = "VAC";
                }));
            }
            else if (Measurement_Selected == 5)
            {
                Application.Current.Dispatcher.Invoke(DispatcherPriority.Background, new ThreadStart(delegate
                {
                    Measurement_Type.Content = "VDC";
                    Measurement_Scale.Content = "";
                    Measurement_Value.Content = "";
                    MIN_Type.Content = "VDC";
                    MAX_Type.Content = "VDC";
                    AVG_Type.Content = "VDC";
                    Current_Measurement_Unit = "VDC";
                }));
            }
            else if (Measurement_Selected == 6)
            {
                Application.Current.Dispatcher.Invoke(DispatcherPriority.Background, new ThreadStart(delegate
                {
                    Measurement_Type.Content = "VAC";
                    Measurement_Scale.Content = "";
                    Measurement_Value.Content = "";
                    MIN_Type.Content = "VAC";
                    MAX_Type.Content = "VAC";
                    AVG_Type.Content = "VAC";
                    Current_Measurement_Unit = "VAC";
                }));
            }
            else if (Measurement_Selected == 7)
            {
                Application.Current.Dispatcher.Invoke(DispatcherPriority.Background, new ThreadStart(delegate
                {
                    Measurement_Type.Content = "VAC";
                    Measurement_Scale.Content = "";
                    Measurement_Value.Content = "";
                    MIN_Type.Content = "VAC";
                    MAX_Type.Content = "VAC";
                    AVG_Type.Content = "VAC";
                    Current_Measurement_Unit = "VAC";
                }));
            }
            else if (Measurement_Selected == 8)
            {
                Application.Current.Dispatcher.Invoke(DispatcherPriority.Background, new ThreadStart(delegate
                {
                    Measurement_Type.Content = "Ω";
                    Measurement_Scale.Content = "";
                    Measurement_Value.Content = "";
                    MIN_Type.Content = "Ω";
                    MAX_Type.Content = "Ω";
                    AVG_Type.Content = "Ω";
                    Current_Measurement_Unit = "Ω";
                }));
            }
            else if (Measurement_Selected == 9)
            {
                Application.Current.Dispatcher.Invoke(DispatcherPriority.Background, new ThreadStart(delegate
                {
                    Measurement_Type.Content = "Ω";
                    Measurement_Scale.Content = "";
                    Measurement_Value.Content = "";
                    MIN_Type.Content = "Ω";
                    MAX_Type.Content = "Ω";
                    AVG_Type.Content = "Ω";
                    Current_Measurement_Unit = "Ω";
                }));
            }
        }

        //Check if user input is a number and if it is then converts it from string to double.
        private (bool, double) isNumber(string Number)
        {
            bool isNum = double.TryParse(Number, out double number);
            return (isNum, number);
        }

        //inserts message to the output log
        private void insert_Log(string Message, int Code)
        {
            string date = DateTime.Now.ToString("yyyy-MM-dd h:mm:ss tt");
            SolidColorBrush Color;
            this.Dispatcher.Invoke(DispatcherPriority.ContextIdle, new ThreadStart(delegate
            {
                if (Output_Log.Inlines.Count >= Auto_Clear_Output_Log_Count)
                {
                    Output_Log.Text = String.Empty;
                    Output_Log.Inlines.Clear();
                    Output_Log.Inlines.Add(new Run("[" + date + "]" + " " + "Output Log has been auto cleared. \n") { Foreground = Brushes.Green });
                }
            }));
            string Status = "";
            switch (Code)
            {
                case 0:
                    Status = "[Success]";
                    Color = Brushes.Green;
                    break;
                case 1:
                    Status = "[Error]";
                    Color = Brushes.Red;
                    break;
                case 2:
                    Status = "[Warning]";
                    Color = Brushes.Orange;
                    break;
                case 3:
                    Status = "";
                    Color = Brushes.Blue;
                    break;
                case 4:
                    Status = "";
                    Color = Brushes.Black;
                    break;
                case 5:
                    Status = "";
                    Color = Brushes.BlueViolet;
                    break;
                default:
                    Status = "Unknown";
                    Color = Brushes.Black;
                    break;
            }
            Application.Current.Dispatcher.Invoke(DispatcherPriority.Background, new ThreadStart(delegate
            {
                Output_Log.Inlines.Add(new Run("[" + date + "]" + " " + Status + " " + Message + "\n") { Foreground = Color });
                if (AutoScroll.IsChecked == true)
                {
                    Output_Log_Scroll.ScrollToBottom();
                }
            }));
            //Saves output log to a text file
            if (saveOutputLog == true)
            {
                writeToFile("[" + date + "]" + " " + Status + " " + Message, GPIB_Address_Info.folder_Directory, GPIB_Address_Info.GPIB_Address.Replace(":", "") + "_" + "Output Log.txt", true);
            }
        }

        //Writes data to a file
        private void writeToFile(string data, string filePath, string fileName, bool append)
        {
            try
            {
                using (TextWriter datatotxt = new StreamWriter(filePath + @"\" + fileName, append))
                {
                    datatotxt.WriteLine(data.Trim());
                }
            }
            catch (Exception)
            {
                saveOutputLog = false;
                SaveOutputLog.IsChecked = false;
                insert_Log("Cannot write Output Log to text file.", 1);
                insert_Log("Save Output Log option disabled.", 1);
                insert_Log("Enable it again from Data Logger Menu if you wish to try again.", 1);
            }
        }

        //------------------------Config Options-----------------------------------------------

        private void Connect_Click(object sender, RoutedEventArgs e)
        {
            if (GPIB_Select == null)
            {
                GPIB_Select = new GPIB_Select_Window();
                GPIB_Select.Closed += (a, b) => { GPIB_Select = null; GPIB_COM_Selected(); };
                GPIB_Select.Owner = this;
                GPIB_Select.Show();
            }
            else
            {
                GPIB_Select.Show();
                insert_Log("GPIB Select Window is already open.", 2);
            }
        }

        private void Stop_Sampling_Click(object sender, RoutedEventArgs e)
        {
            if (Stop_Sampling.IsChecked == false)
            {
                DataSampling = false;
            }
            else
            {
                StartDateTime = DateTime.Now;
                DataSampling = true;
            }
            if (DataSampling == true)
            {
                insert_Log("Software is reading measurement data from multimeter.", 0);
            }
            else
            {
                insert_Log("Software will not read measurement data from multimeter.", 2);
            }
        }

        private void Sampling_Only_Click(object sender, RoutedEventArgs e)
        {
            if (Sampling_Only.IsChecked == true)
            {
                isSamplingOnly = true;
                Local_Exit.IsEnabled = false;
                lockControls();
            }
            else
            {
                isSamplingOnly = false;
                Local_Exit.IsEnabled = true;
                unlockControls();
            }
            if (isSamplingOnly == true)
            {
                insert_Log("Software will now only read measurements from the multimeter.", 2);
                insert_Log("All Write (front panel) operations are disabled.", 2);
            }
            else
            {
                insert_Log("Software will allow commands to be send to the multimeter.", 0);
                insert_Log("Sampling only mode disabled. Returned to normal mode.", 0);
            }
        }

        private void Exit_Click(object sender, RoutedEventArgs e)
        {
            Application.Current.Shutdown();
        }

        //-----------------------------------------------------------------------


        //---------------------------Data Logger--------------------------------------------

        private void OpenFolder_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                System.Diagnostics.Process.Start("explorer.exe", GPIB_Address_Info.folder_Directory);
            }
            catch (Exception)
            {
                insert_Log("Cannot open test files directory.", 1);
            }
        }

        private void SaveOutputLog_Click(object sender, RoutedEventArgs e)
        {
            if (SaveOutputLog.IsChecked == true)
            {
                saveOutputLog = true;
            }
            else
            {
                saveOutputLog = false;
            }
            if (saveOutputLog == true)
            {
                insert_Log("Output Log entries will be saved to a text file.", 0);
            }
            else
            {
                insert_Log("Output Log entries will not be saved.", 2);
            }
        }

        private void ClearOutputLog_Click(object sender, RoutedEventArgs e)
        {
            Output_Log.Text = String.Empty;
            Output_Log.Inlines.Clear();
        }

        private void SaveMeasurements_Click(object sender, RoutedEventArgs e)
        {
            if (SaveMeasurements.IsChecked == true)
            {
                saveMeasurements = true;
            }
            else
            {
                saveMeasurements = false;
            }
            if (saveMeasurements == true)
            {
                insert_Log("Measurement data will be saved.", 0);
                saveMeasurements_Timer.Enabled = true;
            }
            else
            {
                insert_Log("Measurement data will not be saved.", 2);
            }
        }

        private void SaveMeasurements_Interval_5Sec_Click(object sender, RoutedEventArgs e)
        {
            saveMeasurements_Timer.Interval = 5000;
            insert_Log("Save Measurement Interval set to 5 seconds.", 0);
            SaveMeasurements_IntervalSelected(5);
        }

        private void SaveMeasurements_Interval_10Sec_Click(object sender, RoutedEventArgs e)
        {
            saveMeasurements_Timer.Interval = 10000;
            insert_Log("Save Measurement Interval set to 10 seconds.", 0);
            SaveMeasurements_IntervalSelected(10);
        }

        private void SaveMeasurements_Interval_20Sec_Click(object sender, RoutedEventArgs e)
        {
            saveMeasurements_Timer.Interval = 20000;
            insert_Log("Save Measurement Interval set to 20 seconds.", 0);
            SaveMeasurements_IntervalSelected(20);
        }

        private void SaveMeasurements_Interval_40Sec_Click(object sender, RoutedEventArgs e)
        {
            saveMeasurements_Timer.Interval = 40000;
            insert_Log("Save Measurement Interval set to 40 seconds.", 0);
            SaveMeasurements_IntervalSelected(40);
        }

        private void SaveMeasurements_Interval_1Min_Click(object sender, RoutedEventArgs e)
        {
            saveMeasurements_Timer.Interval = 60000;
            insert_Log("Save Measurement Interval set to 1 Minute.", 0);
            SaveMeasurements_IntervalSelected(60);
        }

        private void SaveMeasurements_Interval_4Min_Click(object sender, RoutedEventArgs e)
        {
            saveMeasurements_Timer.Interval = 240000;
            insert_Log("Save Measurement Interval set to 4 Minutes.", 0);
            SaveMeasurements_IntervalSelected(240);
        }

        private void SaveMeasurements_Interval_8Min_Click(object sender, RoutedEventArgs e)
        {
            saveMeasurements_Timer.Interval = 480000;
            insert_Log("Save Measurement Interval set to 8 Minutes.", 0);
            SaveMeasurements_IntervalSelected(480);
        }

        private void SaveMeasurements_Interval_10Min_Click(object sender, RoutedEventArgs e)
        {
            saveMeasurements_Timer.Interval = 600000;
            insert_Log("Save Measurement Interval set to 10 Minutes.", 0);
            SaveMeasurements_IntervalSelected(600);
        }

        private void SaveMeasurements_IntervalSelected(int interval)
        {
            if (interval == 5)
            {
                SaveMeasurements_Interval_5Sec.IsChecked = true;
            }
            else
            {
                SaveMeasurements_Interval_5Sec.IsChecked = false;
            }
            if (interval == 10)
            {
                SaveMeasurements_Interval_10Sec.IsChecked = true;
            }
            else
            {
                SaveMeasurements_Interval_10Sec.IsChecked = false;
            }
            if (interval == 20)
            {
                SaveMeasurements_Interval_20Sec.IsChecked = true;
            }
            else
            {
                SaveMeasurements_Interval_20Sec.IsChecked = false;
            }
            if (interval == 40)
            {
                SaveMeasurements_Interval_40Sec.IsChecked = true;
            }
            else
            {
                SaveMeasurements_Interval_40Sec.IsChecked = false;
            }
            if (interval == 60)
            {
                SaveMeasurements_Interval_1Min.IsChecked = true;
            }
            else
            {
                SaveMeasurements_Interval_1Min.IsChecked = false;
            }
            if (interval == 240)
            {
                SaveMeasurements_Interval_4Min.IsChecked = true;
            }
            else
            {
                SaveMeasurements_Interval_4Min.IsChecked = false;
            }
            if (interval == 480)
            {
                SaveMeasurements_Interval_8Min.IsChecked = true;
            }
            else
            {
                SaveMeasurements_Interval_8Min.IsChecked = false;
            }
            if (interval == 600)
            {
                SaveMeasurements_Interval_10Min.IsChecked = true;
            }
            else
            {
                SaveMeasurements_Interval_10Min.IsChecked = false;
            }
        }

        private void Auto_Clear_20_Click(object sender, RoutedEventArgs e)
        {
            Interlocked.Exchange(ref Auto_Clear_Output_Log_Count, 20);
            insert_Log("Output Log will be cleared after " + Auto_Clear_Output_Log_Count + " logs are inserted into it.", 0);
            Auto_Clear_20.IsChecked = true;
            Auto_Clear_40.IsChecked = false;
            Auto_Clear_60.IsChecked = false;
        }

        private void Auto_Clear_40_Click(object sender, RoutedEventArgs e)
        {
            Interlocked.Exchange(ref Auto_Clear_Output_Log_Count, 40);
            insert_Log("Output Log will be cleared after " + Auto_Clear_Output_Log_Count + " logs are inserted into it.", 0);
            Auto_Clear_20.IsChecked = false;
            Auto_Clear_40.IsChecked = true;
            Auto_Clear_60.IsChecked = false;
        }

        private void Auto_Clear_60_Click(object sender, RoutedEventArgs e)
        {
            Interlocked.Exchange(ref Auto_Clear_Output_Log_Count, 60);
            insert_Log("Output Log will be cleared after " + Auto_Clear_Output_Log_Count + " logs are inserted into it.", 0);
            Auto_Clear_20.IsChecked = false;
            Auto_Clear_40.IsChecked = false;
            Auto_Clear_60.IsChecked = true;
        }

        //-----------------------------------------------------------------------

        //---------------------------------Graph Options--------------------------------------
        private void ShowMeasurementGraph_Click(object sender, RoutedEventArgs e)
        {
            if (HP3456A_Graph_Window == null)
            {
                Create_HP3456A_Graph_Window();
                ShowMeasurementGraph.IsChecked = true;
                AddDataGraph.IsChecked = true;
                save_to_Graph = true;
                Enable_AddDatatoGraph();
                insert_Log("HP3456A Graph Module has been opened.", 0);
            }
            else
            {
                ShowMeasurementGraph.IsChecked = true;
            }
        }

        private void Create_HP3456A_Graph_Window()
        {
            try
            {
                (string Measurement_Unit, string Graph_Y_Axis_Label) = MeasurementUnit_String();
                Thread Waveform_Thread = new Thread(new ThreadStart(() =>
                {
                    HP3456A_Graph_Window = new Graphing_Window(Measurement_Unit, Graph_Y_Axis_Label, "HP 3456A " + GPIB_Address_Info.GPIB_Address);
                    HP3456A_Graph_Window.Show();
                    HP3456A_Graph_Window.Closed += Close_Graph_Event;
                    Dispatcher.Run();
                }));
                Waveform_Thread.CurrentCulture = CultureInfo.CreateSpecificCulture("en-US");
                Waveform_Thread.CurrentUICulture = CultureInfo.CreateSpecificCulture("en-US");
                Waveform_Thread.SetApartmentState(ApartmentState.STA);
                Waveform_Thread.IsBackground = true;
                Waveform_Thread.Start();
            }
            catch (Exception Ex)
            {
                insert_Log(Ex.Message, 1);
                insert_Log("HP3456A Graph Window creation failed.", 1);
            }
        }

        private void Close_Graph_Event(object sender, EventArgs e)
        {
            HP3456A_Graph_Window.Dispatcher.InvokeShutdown();
            HP3456A_Graph_Window = null;
            Close_Graph_Module();
        }

        private void Close_Graph_Module()
        {
            this.Dispatcher.Invoke(() =>
            {
                if (HP3456A_Graph_Window == null & HP3456A_N_Graph_Window == null & HP3456A_DateTime_Graph_Window == null)
                {
                    Save_to_N_Graph = false;
                    Save_to_DateTime_Graph = false;
                    AddDataGraph.IsChecked = false;
                    insert_Log("No Graphs are opened, unchecking Add Data to Graphs option.", 2);
                }
                save_to_Graph = false;
                ShowMeasurementGraph.IsChecked = false;
                insert_Log("HP3456A Graph Module has been closed.", 0);
            });
        }

        private void Try_Graph_Reset()
        {
            try
            {
                (string Measurement_Unit, string Graph_Y_Axis_Label) = MeasurementUnit_String();
                if (HP3456A_Graph_Window != null)
                {
                    HP3456A_Graph_Window.Measurement_Unit = Measurement_Unit;
                    HP3456A_Graph_Window.Graph_Y_Axis_Label = Graph_Y_Axis_Label;
                    HP3456A_Graph_Window.Graph_Reset = true;
                }
                if (HP3456A_N_Graph_Window != null)
                {
                    HP3456A_N_Graph_Window.Measurement_Unit = Measurement_Unit;
                    HP3456A_N_Graph_Window.Graph_Y_Axis_Label = Graph_Y_Axis_Label;
                    HP3456A_N_Graph_Window.Graph_Reset = true;
                }
                if (HP3456A_DateTime_Graph_Window != null)
                {
                    HP3456A_DateTime_Graph_Window.Measurement_Unit = Measurement_Unit;
                    HP3456A_DateTime_Graph_Window.Graph_Y_Axis_Label = Graph_Y_Axis_Label;
                    HP3456A_DateTime_Graph_Window.Graph_Reset = true;
                }
            }
            catch (Exception)
            {
                insert_Log("Graph Reset may have failed, do a manual reset through the graph window.", 2);
            }
        }

        private void AddDataGraph_Click(object sender, RoutedEventArgs e)
        {
            if (AddDataGraph.IsChecked == true & HP3456A_Graph_Window != null)
            {
                save_to_Graph = true;
                insert_Log("Data will be added to Graph.", 0);
                AddDataGraph.IsChecked = true;
            }
            else
            {
                save_to_Graph = false;
            }
            if (AddDataGraph.IsChecked == true & HP3456A_N_Graph_Window != null)
            {
                Save_to_N_Graph = true;
                insert_Log("Data will be added to N Sample Waveform Graph.", 0);
                AddDataGraph.IsChecked = true;
            }
            else
            {
                Save_to_N_Graph = false;
            }
            if (AddDataGraph.IsChecked == true & HP3456A_DateTime_Graph_Window != null)
            {
                Save_to_DateTime_Graph = true;
                insert_Log("Data will be added to DateTime Graph.", 0);
                AddDataGraph.IsChecked = true;
            }
            else
            {
                Save_to_DateTime_Graph = false;
            }
            if (HP3456A_Graph_Window == null & HP3456A_N_Graph_Window == null & HP3456A_DateTime_Graph_Window == null)
            {
                save_to_Graph = false;
                Save_to_N_Graph = false;
                Save_to_DateTime_Graph = false;
                AddDataGraph.IsChecked = false;
                insert_Log("No Graphs are opened, unchecking Add Data to Graphs option.", 2);
            }
        }

        private void Enable_AddDatatoGraph()
        {
            if (HP3456A_Graph_Window != null)
            {
                save_to_Graph = true;
                AddDataGraph.IsChecked = true;
            }
            if (HP3456A_N_Graph_Window != null)
            {
                Save_to_N_Graph = true;
                AddDataGraph.IsChecked = true;
            }
            if (HP3456A_DateTime_Graph_Window != null)
            {
                Save_to_DateTime_Graph = true;
                AddDataGraph.IsChecked = true;
            }
        }

        private void N_Sample_Graph_Button_Click(object sender, RoutedEventArgs e)
        {
            (bool isValidNum, double N_Sample_Value) = Text_Num(N_Sample_Graph_Text.Text, false, true);
            if (isValidNum == true)
            {
                if (N_Sample_Value >= 10)
                {
                    if (HP3456A_N_Graph_Window == null)
                    {
                        Create_HP3456A_N_Sample_Graph_Window((int)N_Sample_Value);
                        Show_N_Sample_Graph.IsChecked = true;
                        AddDataGraph.IsChecked = true;
                        Save_to_N_Graph = true;
                        Enable_AddDatatoGraph();
                        insert_Log("HP3456A N Sample Graph Module has been opened.", 0);
                    }
                }
                else
                {
                    insert_Log("N Sample Graph Creation Value must be a positive integer greater than 10.", 2);
                }
            }
            else
            {
                insert_Log("N Sample Graph Creation Value must be a positive integer greater than 10.", 2);
            }
        }

        private void Create_HP3456A_N_Sample_Graph_Window(int N_Samples)
        {
            try
            {
                (string Measurement_Unit, string Graph_Y_Axis_Label) = MeasurementUnit_String();
                Thread Waveform_Thread = new Thread(new ThreadStart(() =>
                {
                    HP3456A_N_Graph_Window = new N_Sample_Graph_Window(N_Samples, Measurement_Unit, Graph_Y_Axis_Label, "HP 3456A " + GPIB_Address_Info.GPIB_Address);
                    HP3456A_N_Graph_Window.Show();
                    HP3456A_N_Graph_Window.Closed += N_Sample_Close_Graph_Event;
                    Dispatcher.Run();
                }));
                Waveform_Thread.CurrentCulture = CultureInfo.CreateSpecificCulture("en-US");
                Waveform_Thread.CurrentUICulture = CultureInfo.CreateSpecificCulture("en-US");
                Waveform_Thread.SetApartmentState(ApartmentState.STA);
                Waveform_Thread.IsBackground = true;
                Waveform_Thread.Start();
            }
            catch (Exception Ex)
            {
                insert_Log(Ex.Message, 1);
                insert_Log("HP3456A N Sample Graph Window creation failed.", 1);
            }
        }

        private void N_Sample_Close_Graph_Event(object sender, EventArgs e)
        {
            HP3456A_N_Graph_Window.Dispatcher.InvokeShutdown();
            HP3456A_N_Graph_Window = null;
            Close_N_Sample_Graph_Module();
        }

        private void Close_N_Sample_Graph_Module()
        {
            this.Dispatcher.Invoke(() =>
            {
                if (HP3456A_Graph_Window == null & HP3456A_N_Graph_Window == null & HP3456A_DateTime_Graph_Window == null)
                {
                    save_to_Graph = false;
                    Save_to_DateTime_Graph = false;
                    AddDataGraph.IsChecked = false;
                    insert_Log("No Graphs are opened, unchecking Add Data to Graphs option.", 2);
                }
                Save_to_N_Graph = false;
                Show_N_Sample_Graph.IsChecked = false;
                insert_Log("HP3456A N Sample Graph Module has been closed.", 0);
            });
        }

        private void Show_DateTime_Graph_Button_Click(object sender, RoutedEventArgs e)
        {
            if (HP3456A_DateTime_Graph_Window == null)
            {
                Create_HP3456A_DateTime_Graph_Window();
                ShowDateTimeGraph.IsChecked = true;
                AddDataGraph.IsChecked = true;
                Save_to_DateTime_Graph = true;
                Enable_AddDatatoGraph();
                insert_Log("HP3456A DateTime Graph Module has been opened.", 0);
            }
            else
            {
                ShowDateTimeGraph.IsChecked = true;
            }
        }

        private void Create_HP3456A_DateTime_Graph_Window()
        {
            try
            {
                (string Measurement_Unit, string Graph_Y_Axis_Label) = MeasurementUnit_String();
                Thread Waveform_Thread = new Thread(new ThreadStart(() =>
                {
                    HP3456A_DateTime_Graph_Window = new DateTime_Graph_Window(Measurement_Unit, Graph_Y_Axis_Label, "HP 3456A " + GPIB_Address_Info.GPIB_Address);
                    HP3456A_DateTime_Graph_Window.Show();
                    HP3456A_DateTime_Graph_Window.Closed += Close_DateTime_Graph_Event;
                    Dispatcher.Run();
                }));
                Waveform_Thread.CurrentCulture = CultureInfo.CreateSpecificCulture("en-US");
                Waveform_Thread.CurrentUICulture = CultureInfo.CreateSpecificCulture("en-US");
                Waveform_Thread.SetApartmentState(ApartmentState.STA);
                Waveform_Thread.IsBackground = true;
                Waveform_Thread.Start();
            }
            catch (Exception Ex)
            {
                insert_Log(Ex.Message, 1);
                insert_Log("HP3456A Graph Window creation failed.", 1);
            }
        }

        private void Close_DateTime_Graph_Event(object sender, EventArgs e)
        {
            HP3456A_DateTime_Graph_Window.Dispatcher.InvokeShutdown();
            HP3456A_DateTime_Graph_Window = null;
            Close_DateTime_Graph_Module();
        }

        private void Close_DateTime_Graph_Module()
        {
            this.Dispatcher.Invoke(() =>
            {
                if (HP3456A_Graph_Window == null & HP3456A_N_Graph_Window == null & HP3456A_DateTime_Graph_Window == null)
                {
                    save_to_Graph = false;
                    Save_to_N_Graph = false;
                    AddDataGraph.IsChecked = false;
                    insert_Log("No Graphs are opened, unchecking Add Data to Graphs option.", 2);
                }
                Save_to_DateTime_Graph = false;
                ShowDateTimeGraph.IsChecked = false;
                insert_Log("HP3456A DateTime Graph Module has been closed.", 0);
            });
        }

        //-----------------------------------------------------------------------

        //------------------------------Table Options-----------------------------------------

        private void ShowTable_Click(object sender, RoutedEventArgs e)
        {
            if (HP3456A_Table == null)
            {
                Create_HP3456A_Table_Window();
                AddDataTable.IsChecked = true;
                ShowTable.IsChecked = true;
                save_to_Table = true;
                insert_Log("HP3456A Table Window has been opened.", 0);
            }
            else
            {
                ShowTable.IsChecked = true;
            }
        }

        private void AddDataTable_Click(object sender, RoutedEventArgs e)
        {
            if (AddDataTable.IsChecked == true & HP3456A_Table != null)
            {
                save_to_Table = true;
                insert_Log("Data will be added to the table.", 0);
                AddDataTable.IsChecked = true;
            }
            else
            {
                save_to_Table = false;
                AddDataTable.IsChecked = false;
            }
        }

        private void Create_HP3456A_Table_Window()
        {
            try
            {
                Thread Table_Thread = new Thread(new ThreadStart(() =>
                {
                    HP3456A_Table = new Measurement_Data_Table("HP 3456A " + GPIB_Address_Info.GPIB_Address);
                    HP3456A_Table.Show();
                    HP3456A_Table.Closed += Close_Table_Event;
                    Dispatcher.Run();
                }));
                Table_Thread.CurrentCulture = CultureInfo.CreateSpecificCulture("en-US");
                Table_Thread.CurrentUICulture = CultureInfo.CreateSpecificCulture("en-US");
                Table_Thread.SetApartmentState(ApartmentState.STA);
                Table_Thread.IsBackground = true;
                Table_Thread.Start();
            }
            catch (Exception Ex)
            {
                insert_Log(Ex.Message, 1);
                insert_Log("HP3456A Table Window creation failed.", 1);
            }
        }

        private void Close_Table_Event(object sender, EventArgs e)
        {
            HP3456A_Table.Dispatcher.InvokeShutdown();
            HP3456A_Table = null;
            Close_Table_Window();
        }

        private void Close_Table_Window()
        {
            this.Dispatcher.Invoke(() =>
            {
                save_to_Table = false;
                AddDataTable.IsChecked = false;
                ShowTable.IsChecked = false;
                insert_Log("HP3456A Table Window has been closed.", 0);
            });
        }

        //-----------------------------------------------------------------------

        //----------------------------Speech Options-------------------------------------------

        private void EnableSpeech_Click(object sender, RoutedEventArgs e)
        {
            if (EnableSpeech.IsChecked == true)
            {
                Interlocked.Exchange(ref isSpeechActive, 1);
                insert_Log("The Speech Synthesizer is Enabled.", 4);
            }
            else
            {
                Interlocked.Exchange(ref isSpeechActive, 0);
                insert_Log("The Speech Synthesizer is Disabled.", 4);
            }
        }

        private void VoiceMale_Click(object sender, RoutedEventArgs e)
        {
            Voice.SelectVoiceByHints(VoiceGender.Male);
            VoiceMale.IsChecked = true;
            VoiceFemale.IsChecked = false;
            insert_Log("David will voice your measurements.", 0);
        }

        private void VoiceFemale_Click(object sender, RoutedEventArgs e)
        {
            Voice.SelectVoiceByHints(VoiceGender.Female);
            VoiceMale.IsChecked = false;
            VoiceFemale.IsChecked = true;
            insert_Log("Zira will voice your measurements.", 0);
        }

        private void VoiceSlow_Click(object sender, RoutedEventArgs e)
        {
            Voice.Rate = 0;
            VoiceSpeedSelected(0);
            insert_Log("Voice speed set to slow.", 4);
        }

        private void VoiceMedium_Click(object sender, RoutedEventArgs e)
        {
            Voice.Rate = 1;
            VoiceSpeedSelected(1);
            insert_Log("Voice speed set to medium.", 4);
        }

        private void VoiceFast_Click(object sender, RoutedEventArgs e)
        {
            Voice.Rate = 2;
            VoiceSpeedSelected(2);
            insert_Log("Voice speed set to fast.", 4);
        }

        private void VoiceVeryFast_Click(object sender, RoutedEventArgs e)
        {
            Voice.Rate = 3;
            VoiceSpeedSelected(3);
            insert_Log("Voice speed set to very fast.", 4);
        }

        private void VoiceFastest_Click(object sender, RoutedEventArgs e)
        {
            Voice.Rate = 4;
            VoiceSpeedSelected(4);
            insert_Log("Voice speed set to fastest.", 4);
        }

        private void VoiceSpeedSelected(int speed)
        {
            if (speed == 0)
            {
                VoiceSlow.IsChecked = true;
            }
            else
            {
                VoiceSlow.IsChecked = false;
            }
            if (speed == 1)
            {
                VoiceMedium.IsChecked = true;
            }
            else
            {
                VoiceMedium.IsChecked = false;
            }
            if (speed == 2)
            {
                VoiceFast.IsChecked = true;
            }
            else
            {
                VoiceFast.IsChecked = false;
            }
            if (speed == 3)
            {
                VoiceVeryFast.IsChecked = true;
            }
            else
            {
                VoiceVeryFast.IsChecked = false;
            }
            if (speed == 4)
            {
                VoiceFastest.IsChecked = true;
            }
            else
            {
                VoiceFastest.IsChecked = false;
            }
        }

        private void Voice_Volume_10_Click(object sender, RoutedEventArgs e)
        {
            Voice.Volume = 10;
            VoiceVolumeSelected(0);
            insert_Log("Voice volume set to 10%.", 4);
        }

        private void Voice_Volume_20_Click(object sender, RoutedEventArgs e)
        {
            Voice.Volume = 20;
            VoiceVolumeSelected(1);
            insert_Log("Voice volume set to 20%.", 4);
        }

        private void Voice_Volume_30_Click(object sender, RoutedEventArgs e)
        {
            Voice.Volume = 30;
            VoiceVolumeSelected(2);
            insert_Log("Voice volume set to 30%.", 4);
        }

        private void Voice_Volume_40_Click(object sender, RoutedEventArgs e)
        {
            Voice.Volume = 40;
            VoiceVolumeSelected(3);
            insert_Log("Voice volume set to 40%.", 4);
        }

        private void Voice_Volume_50_Click(object sender, RoutedEventArgs e)
        {
            Voice.Volume = 50;
            VoiceVolumeSelected(4);
            insert_Log("Voice volume set to 50%.", 4);
        }

        private void Voice_Volume_60_Click(object sender, RoutedEventArgs e)
        {
            Voice.Volume = 60;
            VoiceVolumeSelected(5);
            insert_Log("Voice volume set to 60%.", 4);
        }

        private void Voice_Volume_70_Click(object sender, RoutedEventArgs e)
        {
            Voice.Volume = 70;
            VoiceVolumeSelected(6);
            insert_Log("Voice volume set to 70%.", 4);
        }

        private void Voice_Volume_80_Click(object sender, RoutedEventArgs e)
        {
            Voice.Volume = 80;
            VoiceVolumeSelected(7);
            insert_Log("Voice volume set to 80%.", 4);
        }

        private void Voice_Volume_90_Click(object sender, RoutedEventArgs e)
        {
            Voice.Volume = 90;
            VoiceVolumeSelected(8);
            insert_Log("Voice volume set to 90%.", 4);
        }

        private void Voice_Volume_100_Click(object sender, RoutedEventArgs e)
        {
            Voice.Volume = 100;
            VoiceVolumeSelected(9);
            insert_Log("Voice volume set to 100%.", 4);
        }

        private void VoiceVolumeSelected(int volume)
        {
            if (volume == 0)
            {
                Voice_Volume_10.IsChecked = true;
            }
            else
            {
                Voice_Volume_10.IsChecked = false;
            }
            if (volume == 1)
            {
                Voice_Volume_20.IsChecked = true;
            }
            else
            {
                Voice_Volume_20.IsChecked = false;
            }
            if (volume == 2)
            {
                Voice_Volume_30.IsChecked = true;
            }
            else
            {
                Voice_Volume_30.IsChecked = false;
            }
            if (volume == 3)
            {
                Voice_Volume_40.IsChecked = true;
            }
            else
            {
                Voice_Volume_40.IsChecked = false;
            }
            if (volume == 4)
            {
                Voice_Volume_50.IsChecked = true;
            }
            else
            {
                Voice_Volume_50.IsChecked = false;
            }
            if (volume == 5)
            {
                Voice_Volume_60.IsChecked = true;
            }
            else
            {
                Voice_Volume_60.IsChecked = false;
            }
            if (volume == 6)
            {
                Voice_Volume_70.IsChecked = true;
            }
            else
            {
                Voice_Volume_70.IsChecked = false;
            }
            if (volume == 7)
            {
                Voice_Volume_80.IsChecked = true;
            }
            else
            {
                Voice_Volume_80.IsChecked = false;
            }
            if (volume == 8)
            {
                Voice_Volume_90.IsChecked = true;
            }
            else
            {
                Voice_Volume_90.IsChecked = false;
            }
            if (volume == 9)
            {
                Voice_Volume_100.IsChecked = true;
            }
            else
            {
                Voice_Volume_100.IsChecked = false;
            }
        }

        //-----------------------------------------------------------------------

        //----------------------------About Options-------------------------------------------

        private void DeviceSupport_Click(object sender, RoutedEventArgs e)
        {
            insert_Log("This software was created for HP 3456A.", 4);
            insert_Log("You will need an AR488 Arduino GPIB adapter.", 4);
        }

        private void Credits_Click(object sender, RoutedEventArgs e)
        {
            insert_Log("Created by Niravk Patel.", 4);
            insert_Log("Email: niravkp97@gmail.com", 4);
            insert_Log("This program was created using C# WPF .Net Framework 4.7.2", 4);
            insert_Log("Supports Windows 10, 8, 8.1, and 7", 4);
        }

        //-----------------------------------------------------------------------

        //--------------------------Measurements Options---------------------------------------------

        private void DCV_Button_Click(object sender, RoutedEventArgs e)
        {
            MesurementSelector(0);
            Range_Tab_Selector(0);
            insert_Log("DCV Measurement Selected.", 3);
            SerialWriteQueue.Add("S0F1R1");
            VDC_Range_Indicator(0);
            lockControls();
            isUserSendCommand = true;
            Try_Graph_Reset();
            Speedup_Interval();
        }

        private void ACV_Button_Click(object sender, RoutedEventArgs e)
        {
            MesurementSelector(1);
            Range_Tab_Selector(1);
            insert_Log("ACV Measurement Selected.", 3);
            SerialWriteQueue.Add("S0F2R1");
            VAC_Range_Indicator(0);
            lockControls();
            isUserSendCommand = true;
            Try_Graph_Reset();
            Speedup_Interval();
        }

        private void TwoOhms_Button_Click(object sender, RoutedEventArgs e)
        {
            if (Filter_indicator == 1)
            {
                insert_Log("Cannot select 2 Wire ohms measurement when the Filter is enabled.", 2);
            }
            else
            {
                MesurementSelector(2);
                Range_Tab_Selector(2);
                insert_Log("2 Wire Ohms Measurement Selected.", 3);
                SerialWriteQueue.Add("S0F4R1");
                Ohms_Range_Indicator(0);
                lockControls();
                isUserSendCommand = true;
                Try_Graph_Reset();
                Speedup_Interval();
            }
        }

        private void FourOhms_Button_Click(object sender, RoutedEventArgs e)
        {
            if (Filter_indicator == 1)
            {
                insert_Log("Cannot select 4 Wire ohms measurement when the Filter is enabled.", 2);
            }
            else
            {
                MesurementSelector(3);
                Range_Tab_Selector(2);
                insert_Log("4 Wire Ohms Measurement Selected.", 3);
                SerialWriteQueue.Add("S0F5R1");
                Ohms_Range_Indicator(0);
                lockControls();
                isUserSendCommand = true;
                Try_Graph_Reset();
                Speedup_Interval();
            }
        }

        private void ACDCV_Button_Click(object sender, RoutedEventArgs e)
        {
            MesurementSelector(4);
            Range_Tab_Selector(1);
            insert_Log("ACV + DCV Measurement Selected.", 3);
            SerialWriteQueue.Add("S0F3R1");
            VAC_Range_Indicator(0);
            lockControls();
            isUserSendCommand = true;
            Try_Graph_Reset();
            Speedup_Interval();
        }

        private void DCV_DCV_Button_Click(object sender, RoutedEventArgs e)
        {
            if (AutoZero_indicator == 0)
            {
                insert_Log("Cannot select DCV/DCV when Auto Zero is disabled.", 2);
            }
            else
            {
                MesurementSelector(5);
                Range_Tab_Selector(0);
                insert_Log("DCV / DCV Measurement Selected.", 3);
                SerialWriteQueue.Add("S1F1R1");
                VDC_Range_Indicator(0);
                lockControls();
                isUserSendCommand = true;
                Try_Graph_Reset();
                Speedup_Interval();
            }
        }

        private void ACV_DCV_Button_Click(object sender, RoutedEventArgs e)
        {
            if (AutoZero_indicator == 0)
            {
                insert_Log("Cannot select ACV / DCV when Auto Zero is disabled.", 2);
            }
            else
            {
                MesurementSelector(6);
                Range_Tab_Selector(1);
                insert_Log("ACV / DCV Measurement Selected.", 3);
                SerialWriteQueue.Add("S1F2R1");
                VAC_Range_Indicator(0);
                lockControls();
                isUserSendCommand = true;
                Try_Graph_Reset();
                Speedup_Interval();
            }
        }

        private void ACV_DCV_DCV_Button_Click(object sender, RoutedEventArgs e)
        {
            if (AutoZero_indicator == 0)
            {
                insert_Log("Cannot select ACV + DCV / DCV when Auto Zero is disabled.", 2);
            }
            else
            {
                MesurementSelector(7);
                Range_Tab_Selector(1);
                insert_Log("ACV + DCV / DCV Measurement Selected.", 3);
                SerialWriteQueue.Add("S1F3R1");
                VAC_Range_Indicator(0);
                lockControls();
                isUserSendCommand = true;
                Try_Graph_Reset();
                Speedup_Interval();
            }
        }

        private void OCTwoOhms_Button_Click(object sender, RoutedEventArgs e)
        {
            if (NPLC_indicator == 3 || NPLC_indicator == 4)
            {
                insert_Log("Cannot select OC 2 Wire Ohms measurement when NPLC 10 or 100 is selected.", 2);
            }
            else if (Filter_indicator == 1)
            {
                insert_Log("Cannot select OC 2 Wire ohms measurement when the Filter is enabled.", 2);
            }
            else if (AutoZero_indicator == 0)
            {
                insert_Log("Cannot select OC 2 Wire ohms measurement when the AutoZero is disabled.", 2);
            }
            else
            {
                MesurementSelector(8);
                Range_Tab_Selector(2);
                insert_Log("OC 2 Wire Ohms Measurement Selected.", 3);
                SerialWriteQueue.Add("S1F4R1");
                Ohms_Range_Indicator(0);
                lockControls();
                isUserSendCommand = true;
                Try_Graph_Reset();
                Speedup_Interval();
            }
        }

        private void OCFourOhms_Button_Click(object sender, RoutedEventArgs e)
        {
            if (NPLC_indicator == 3 || NPLC_indicator == 4)
            {
                insert_Log("Cannot select OC 4 Wire Ohms measurement when NPLC 10 or 100 is selected.", 2);
            }
            else if (Filter_indicator == 1)
            {
                insert_Log("Cannot select OC 4 Wire ohms measurement when the Filter is enabled.", 2);
            }
            else if (AutoZero_indicator == 0)
            {
                insert_Log("Cannot select OC 4 Wire ohms measurement when the AutoZero is disabled.", 2);
            }
            else
            {
                MesurementSelector(9);
                Range_Tab_Selector(2);
                insert_Log("OC 4 Wire Ohms Measurement Selected.", 3);
                SerialWriteQueue.Add("S1F5R1");
                Ohms_Range_Indicator(0);
                lockControls();
                isUserSendCommand = true;
                Try_Graph_Reset();
                Speedup_Interval();
            }
        }

        private void Range_Tab_Selector(int RangeType)
        {
            if (RangeType == 0) //VDC
            {
                DCV_Tab.IsSelected = true;
            }
            else
            {
                DCV_Tab.IsSelected = false;
            }
            if (RangeType == 1) //VAC
            {
                ACV_Tab.IsSelected = true;
            }
            else
            {
                ACV_Tab.IsSelected = false;
            }
            if (RangeType == 2) //Ohms
            {
                Ohms_Tab.IsSelected = true;
            }
            else
            {
                Ohms_Tab.IsSelected = false;
            }
        }

        private void MesurementSelector(int MeasurementChoice)
        {
            if (MeasurementChoice == 0) //VDC
            {
                DCV_Border.BorderBrush = Selected;
                Measurement_Selected = 0;
                Interlocked.Exchange(ref Selected_Measurement_type, 0);
            }
            else
            {
                DCV_Border.BorderBrush = Deselected;
            }
            if (MeasurementChoice == 1) //VAC
            {
                ACV_Border.BorderBrush = Selected;
                Measurement_Selected = 1;
                Interlocked.Exchange(ref Selected_Measurement_type, 1);
            }
            else
            {
                ACV_Border.BorderBrush = Deselected;
            }
            if (MeasurementChoice == 2) //2 Wire Ohms
            {
                TwoOhms_Border.BorderBrush = Selected;
                Measurement_Selected = 2;
                Interlocked.Exchange(ref Selected_Measurement_type, 2);
            }
            else
            {
                TwoOhms_Border.BorderBrush = Deselected;
            }
            if (MeasurementChoice == 3) //4 Wire Ohms
            {
                FourOhms_Border.BorderBrush = Selected;
                Measurement_Selected = 3;
                Interlocked.Exchange(ref Selected_Measurement_type, 3);
            }
            else
            {
                FourOhms_Border.BorderBrush = Deselected;
            }
            if (MeasurementChoice == 4) //ACDCV
            {
                ACDCV_Border.BorderBrush = Selected;
                Measurement_Selected = 4;
                Interlocked.Exchange(ref Selected_Measurement_type, 4);
            }
            else
            {
                ACDCV_Border.BorderBrush = Deselected;
            }
            if (MeasurementChoice == 5) //DCV_DCV
            {
                DCV_DCV_Border.BorderBrush = Selected;
                Measurement_Selected = 5;
                Interlocked.Exchange(ref Selected_Measurement_type, 5);
            }
            else
            {
                DCV_DCV_Border.BorderBrush = Deselected;
            }
            if (MeasurementChoice == 6) //ACV_DCV
            {
                ACV_DCV_Border.BorderBrush = Selected;
                Measurement_Selected = 6;
                Interlocked.Exchange(ref Selected_Measurement_type, 6);
            }
            else
            {
                ACV_DCV_Border.BorderBrush = Deselected;
            }
            if (MeasurementChoice == 7) //ACV_DCV_DCV
            {
                ACV_DCV_DCV_Border.BorderBrush = Selected;
                Measurement_Selected = 7;
                Interlocked.Exchange(ref Selected_Measurement_type, 7);
            }
            else
            {
                ACV_DCV_DCV_Border.BorderBrush = Deselected;
            }
            if (MeasurementChoice == 8) //OCTwoOhms
            {
                OCTwoOhms_Border.BorderBrush = Selected;
                Measurement_Selected = 8;
                Interlocked.Exchange(ref Selected_Measurement_type, 8);
            }
            else
            {
                OCTwoOhms_Border.BorderBrush = Deselected;
            }
            if (MeasurementChoice == 9) //OCFourOhms
            {
                OCFourOhms_Border.BorderBrush = Selected;
                Measurement_Selected = 9;
                Interlocked.Exchange(ref Selected_Measurement_type, 9);
            }
            else
            {
                OCFourOhms_Border.BorderBrush = Deselected;
            }
        }

        //-----------------------------------------------------------------------

        //------------------------------VDC Range-----------------------------------------

        private void DCV_Auto_Button_Click(object sender, RoutedEventArgs e)
        {
            if (Measurement_Selected == 0 || Measurement_Selected == 5)
            {
                SerialWriteQueue.Add(VDC_Range_Indicator(0));
                lockControls();
                isUserSendCommand = true;
                Speedup_Interval();
                insert_Log("DCV Range set to Auto.", 5);
            }
            else
            {
                insert_Log("Cannot set DCV Range when DCV Measurement is not selected.", 2);
            }
        }

        private void DCV_100mV_Button_Click(object sender, RoutedEventArgs e)
        {
            if (Measurement_Selected == 0 || Measurement_Selected == 5)
            {
                SerialWriteQueue.Add(VDC_Range_Indicator(1));
                lockControls();
                isUserSendCommand = true;
                Speedup_Interval();
                insert_Log("DCV Range set to 100mV.", 5);
            }
            else
            {
                insert_Log("Cannot set DCV Range when DCV Measurement is not selected.", 2);
            }
        }

        private void DCV_1000mV_Button_Click(object sender, RoutedEventArgs e)
        {
            if (Measurement_Selected == 0 || Measurement_Selected == 5)
            {
                SerialWriteQueue.Add(VDC_Range_Indicator(2));
                lockControls();
                isUserSendCommand = true;
                Speedup_Interval();
                insert_Log("DCV Range set to 1000mV.", 5);
            }
            else
            {
                insert_Log("Cannot set DCV Range when DCV Measurement is not selected.", 2);
            }
        }

        private void DCV_10V_Button_Click(object sender, RoutedEventArgs e)
        {
            if (Measurement_Selected == 0 || Measurement_Selected == 5)
            {
                SerialWriteQueue.Add(VDC_Range_Indicator(3));
                lockControls();
                isUserSendCommand = true;
                Speedup_Interval();
                insert_Log("DCV Range set to 10V.", 5);
            }
            else
            {
                insert_Log("Cannot set DCV Range when DCV Measurement is not selected.", 2);
            }
        }

        private void DCV_100V_Button_Click(object sender, RoutedEventArgs e)
        {
            if (Measurement_Selected == 0 || Measurement_Selected == 5)
            {
                SerialWriteQueue.Add(VDC_Range_Indicator(4));
                lockControls();
                isUserSendCommand = true;
                Speedup_Interval();
                insert_Log("DCV Range set to 100V.", 5);
            }
            else
            {
                insert_Log("Cannot set DCV Range when DCV Measurement is not selected.", 2);
            }
        }

        private void DCV_1000V_Button_Click(object sender, RoutedEventArgs e)
        {
            if (Measurement_Selected == 0 || Measurement_Selected == 5)
            {
                SerialWriteQueue.Add(VDC_Range_Indicator(5));
                lockControls();
                isUserSendCommand = true;
                Speedup_Interval();
                insert_Log("DCV Range set to 1000V.", 5);
            }
            else
            {
                insert_Log("Cannot set DCV Range when DCV Measurement is not selected.", 2);
            }
        }

        private string VDC_Range_Indicator(int Range)
        {
            string RangeCommand = "R1";
            if (Range == 0)
            {
                DCV_Auto_Border.BorderBrush = Selected;
                RangeCommand = "R1";
            }
            else
            {
                DCV_Auto_Border.BorderBrush = Deselected;
            }
            if (Range == 1)
            {
                DCV_100mV_Border.BorderBrush = Selected;
                RangeCommand = "R2";
            }
            else
            {
                DCV_100mV_Border.BorderBrush = Deselected;
            }
            if (Range == 2)
            {
                DCV_1000mV_Border.BorderBrush = Selected;
                RangeCommand = "R3";
            }
            else
            {
                DCV_1000mV_Border.BorderBrush = Deselected;
            }
            if (Range == 3)
            {
                DCV_10V_Border.BorderBrush = Selected;
                RangeCommand = "R4";
            }
            else
            {
                DCV_10V_Border.BorderBrush = Deselected;
            }
            if (Range == 4)
            {
                DCV_100V_Border.BorderBrush = Selected;
                RangeCommand = "R5";
            }
            else
            {
                DCV_100V_Border.BorderBrush = Deselected;
            }
            if (Range == 5)
            {
                DCV_1000V_Border.BorderBrush = Selected;
                RangeCommand = "R6";
            }
            else
            {
                DCV_1000V_Border.BorderBrush = Deselected;
            }
            return RangeCommand;
        }

        //-----------------------------------------------------------------------

        //------------------------VAC Range-----------------------------------------------

        private void ACV_Auto_Button_Click(object sender, RoutedEventArgs e)
        {
            if (Measurement_Selected == 1 || Measurement_Selected == 4 || Measurement_Selected == 6 || Measurement_Selected == 7)
            {
                SerialWriteQueue.Add(VAC_Range_Indicator(0));
                lockControls();
                isUserSendCommand = true;
                Speedup_Interval();
                insert_Log("ACV Range set to Auto.", 5);
            }
            else
            {
                insert_Log("Cannot set ACV Range when ACV Measurement is not selected.", 2);
            }
        }

        private void ACV_1000mV_Button_Click(object sender, RoutedEventArgs e)
        {
            if (Measurement_Selected == 1 || Measurement_Selected == 4 || Measurement_Selected == 6 || Measurement_Selected == 7)
            {
                SerialWriteQueue.Add(VAC_Range_Indicator(2));
                lockControls();
                isUserSendCommand = true;
                Speedup_Interval();
                insert_Log("ACV Range set to 1000mV.", 5);
            }
            else
            {
                insert_Log("Cannot set ACV Range when ACV Measurement is not selected.", 2);
            }
        }

        private void ACV_10V_Button_Click(object sender, RoutedEventArgs e)
        {
            if (Measurement_Selected == 1 || Measurement_Selected == 4 || Measurement_Selected == 6 || Measurement_Selected == 7)
            {
                SerialWriteQueue.Add(VAC_Range_Indicator(3));
                lockControls();
                isUserSendCommand = true;
                Speedup_Interval();
                insert_Log("ACV Range set to 10V.", 5);
            }
            else
            {
                insert_Log("Cannot set ACV Range when ACV Measurement is not selected.", 2);
            }
        }

        private void ACV_100V_Button_Click(object sender, RoutedEventArgs e)
        {
            if (Measurement_Selected == 1 || Measurement_Selected == 4 || Measurement_Selected == 6 || Measurement_Selected == 7)
            {
                SerialWriteQueue.Add(VAC_Range_Indicator(4));
                lockControls();
                isUserSendCommand = true;
                Speedup_Interval();
                insert_Log("ACV Range set to 100V.", 5);
            }
            else
            {
                insert_Log("Cannot set ACV Range when ACV Measurement is not selected.", 2);
            }
        }

        private void ACV_1000V_Button_Click(object sender, RoutedEventArgs e)
        {
            if (Measurement_Selected == 1 || Measurement_Selected == 4 || Measurement_Selected == 6 || Measurement_Selected == 7)
            {
                SerialWriteQueue.Add(VAC_Range_Indicator(5));
                lockControls();
                isUserSendCommand = true;
                Speedup_Interval();
                insert_Log("ACV Range set to 1000V.", 5);
            }
            else
            {
                insert_Log("Cannot set ACV Range when ACV Measurement is not selected.", 2);
            }
        }

        private string VAC_Range_Indicator(int Range)
        {
            string RangeCommand = "R1";
            if (Range == 0)
            {
                ACV_Auto_Border.BorderBrush = Selected;
                RangeCommand = "R1";
            }
            else
            {
                ACV_Auto_Border.BorderBrush = Deselected;
            }
            if (Range == 2)
            {
                ACV_1000mV_Border.BorderBrush = Selected;
                RangeCommand = "R3";
            }
            else
            {
                ACV_1000mV_Border.BorderBrush = Deselected;
            }
            if (Range == 3)
            {
                ACV_10V_Border.BorderBrush = Selected;
                RangeCommand = "R4";
            }
            else
            {
                ACV_10V_Border.BorderBrush = Deselected;
            }
            if (Range == 4)
            {
                ACV_100V_Border.BorderBrush = Selected;
                RangeCommand = "R5";
            }
            else
            {
                ACV_100V_Border.BorderBrush = Deselected;
            }
            if (Range == 5)
            {
                ACV_1000V_Border.BorderBrush = Selected;
                RangeCommand = "R6";
            }
            else
            {
                ACV_1000V_Border.BorderBrush = Deselected;
            }
            return RangeCommand;
        }

        //-----------------------------------------------------------------------

        //-------------------------Ohms Range----------------------------------------------

        private void Ohms_Auto_Button_Click(object sender, RoutedEventArgs e)
        {
            if (Measurement_Selected == 2 || Measurement_Selected == 3 || Measurement_Selected == 8 || Measurement_Selected == 9)
            {
                SerialWriteQueue.Add(Ohms_Range_Indicator(0));
                lockControls();
                isUserSendCommand = true;
                Speedup_Interval();
                insert_Log("Ohms Range set to Auto.", 5);
            }
            else
            {
                insert_Log("Cannot set Ω Range when 2/4 Wire Ω Measurement is not selected.", 2);
            }
        }

        private void Ohms_0_1K_Button_Click(object sender, RoutedEventArgs e)
        {
            if (Measurement_Selected == 2 || Measurement_Selected == 3 || Measurement_Selected == 8 || Measurement_Selected == 9)
            {
                SerialWriteQueue.Add(Ohms_Range_Indicator(1));
                lockControls();
                isUserSendCommand = true;
                Speedup_Interval();
                insert_Log("Ohms Range set to 0.1KΩ.", 5);
            }
            else
            {
                insert_Log("Cannot set Ω Range when 2/4 Wire Ω Measurement is not selected.", 2);
            }
        }

        private void Ohms_1K_Button_Click(object sender, RoutedEventArgs e)
        {
            if (Measurement_Selected == 2 || Measurement_Selected == 3 || Measurement_Selected == 8 || Measurement_Selected == 9)
            {
                SerialWriteQueue.Add(Ohms_Range_Indicator(2));
                lockControls();
                isUserSendCommand = true;
                Speedup_Interval();
                insert_Log("Ohms Range set to 1KΩ.", 5);
            }
            else
            {
                insert_Log("Cannot set Ω Range when 2/4 Wire Ω Measurement is not selected.", 2);
            }
        }

        private void Ohms_10K_Button_Click(object sender, RoutedEventArgs e)
        {
            if (Measurement_Selected == 2 || Measurement_Selected == 3 || Measurement_Selected == 8 || Measurement_Selected == 9)
            {
                SerialWriteQueue.Add(Ohms_Range_Indicator(3));
                lockControls();
                isUserSendCommand = true;
                Speedup_Interval();
                insert_Log("Ohms Range set to 10KΩ.", 5);
            }
            else
            {
                insert_Log("Cannot set Ω Range when 2/4 Wire Ω Measurement is not selected.", 2);
            }
        }

        private void Ohms_100K_Button_Click(object sender, RoutedEventArgs e)
        {
            if (Measurement_Selected == 2 || Measurement_Selected == 3 || Measurement_Selected == 8 || Measurement_Selected == 9)
            {
                SerialWriteQueue.Add(Ohms_Range_Indicator(4));
                lockControls();
                isUserSendCommand = true;
                Speedup_Interval();
                insert_Log("Ohms Range set to 100KΩ.", 5);
            }
            else
            {
                insert_Log("Cannot set Ω Range when 2/4 Wire Ω Measurement is not selected.", 2);
            }
        }

        private void Ohms_1M_Button_Click(object sender, RoutedEventArgs e)
        {
            if (Measurement_Selected == 2 || Measurement_Selected == 3)
            {
                SerialWriteQueue.Add(Ohms_Range_Indicator(5));
                lockControls();
                isUserSendCommand = true;
                Speedup_Interval();
                insert_Log("Ohms Range set to 1MΩ.", 5);
            }
            else
            {
                insert_Log("Cannot set Ω Range when 2/4 Wire Ω Measurement is not selected.", 2);
                insert_Log("1M Ω Range not available for OC 2 and 4 Wire Ohms measurements.", 2);
            }
        }

        private void Ohms_10M_Button_Click(object sender, RoutedEventArgs e)
        {
            if (Measurement_Selected == 2 || Measurement_Selected == 3)
            {
                SerialWriteQueue.Add(Ohms_Range_Indicator(6));
                lockControls();
                isUserSendCommand = true;
                Speedup_Interval();
                insert_Log("Ohms Range set to 10MΩ.", 5);
            }
            else
            {
                insert_Log("Cannot set Ω Range when 2/4 Wire Ω Measurement is not selected.", 2);
                insert_Log("10M Ω Range not available for OC 2 and 4 Wire Ohms measurements.", 2);
            }
        }

        private void Ohms_100M_Button_Click(object sender, RoutedEventArgs e)
        {
            if (Measurement_Selected == 2 || Measurement_Selected == 3)
            {
                SerialWriteQueue.Add(Ohms_Range_Indicator(7));
                lockControls();
                isUserSendCommand = true;
                Speedup_Interval();
                insert_Log("Ohms Range set to 100MΩ.", 5);
            }
            else
            {
                insert_Log("Cannot set Ω Range when 2/4 Wire Ω Measurement is not selected.", 2);
                insert_Log("100M Ω Range not available for OC 2 and 4 Wire Ohms measurements.", 2);
            }
        }

        private void Ohms_1000M_Button_Click(object sender, RoutedEventArgs e)
        {
            if (Measurement_Selected == 2 || Measurement_Selected == 3)
            {
                SerialWriteQueue.Add(Ohms_Range_Indicator(8));
                lockControls();
                isUserSendCommand = true;
                Speedup_Interval();
                insert_Log("Ohms Range set to 1000MΩ.", 5);
            }
            else
            {
                insert_Log("Cannot set Ω Range when 2/4 Wire Ω Measurement is not selected.", 2);
                insert_Log("1000M Ω Range not available for OC 2 and 4 Wire Ohms measurements.", 2);
            }
        }

        private string Ohms_Range_Indicator(int Range)
        {
            string RangeCommand = "R1";
            if (Range == 0)
            {
                Ohms_Auto_Border.BorderBrush = Selected;
                RangeCommand = "R1";
            }
            else
            {
                Ohms_Auto_Border.BorderBrush = Deselected;
            }
            if (Range == 1)
            {
                Ohms_0_1K_Border.BorderBrush = Selected;
                RangeCommand = "R2";
            }
            else
            {
                Ohms_0_1K_Border.BorderBrush = Deselected;
            }
            if (Range == 2)
            {
                Ohms_1K_Border.BorderBrush = Selected;
                RangeCommand = "R3";
            }
            else
            {
                Ohms_1K_Border.BorderBrush = Deselected;
            }
            if (Range == 3)
            {
                Ohms_10K_Border.BorderBrush = Selected;
                RangeCommand = "R4";
            }
            else
            {
                Ohms_10K_Border.BorderBrush = Deselected;
            }
            if (Range == 4)
            {
                Ohms_100K_Border.BorderBrush = Selected;
                RangeCommand = "R5";
            }
            else
            {
                Ohms_100K_Border.BorderBrush = Deselected;
            }
            if (Range == 5)
            {
                Ohms_1M_Border.BorderBrush = Selected;
                RangeCommand = "R6";
            }
            else
            {
                Ohms_1M_Border.BorderBrush = Deselected;
            }
            if (Range == 6)
            {
                Ohms_10M_Border.BorderBrush = Selected;
                RangeCommand = "R7";
            }
            else
            {
                Ohms_10M_Border.BorderBrush = Deselected;
            }
            if (Range == 7)
            {
                Ohms_100M_Border.BorderBrush = Selected;
                RangeCommand = "R8";
            }
            else
            {
                Ohms_100M_Border.BorderBrush = Deselected;
            }
            if (Range == 8)
            {
                Ohms_1000M_Border.BorderBrush = Selected;
                RangeCommand = "R9";
            }
            else
            {
                Ohms_1000M_Border.BorderBrush = Deselected;
            }
            return RangeCommand;
        }

        //-----------------------------------------------------------------------

        //----------------------Meter Resolution-------------------------------------------------

        private void Digit_Three_Button_Click(object sender, RoutedEventArgs e)
        {
            SerialWriteQueue.Add(NDigit_Selector(0));
            lockControls();
            insert_Log("N Digit set to 3½.", 3);
            isUserSendCommand = true;
            Speedup_Interval();
        }

        private void Digit_Four_Button_Click(object sender, RoutedEventArgs e)
        {
            SerialWriteQueue.Add(NDigit_Selector(1));
            lockControls();
            insert_Log("N Digit set to 4½.", 3);
            isUserSendCommand = true;
            Speedup_Interval();
        }

        private void Digit_Five_Button_Click(object sender, RoutedEventArgs e)
        {
            SerialWriteQueue.Add(NDigit_Selector(2));
            lockControls();
            insert_Log("N Digit set to 5½.", 3);
            isUserSendCommand = true;
            Speedup_Interval();
        }

        private void Digit_Six_Button_Click(object sender, RoutedEventArgs e)
        {
            SerialWriteQueue.Add(NDigit_Selector(3));
            lockControls();
            insert_Log("N Digit set to 6½.", 3);
            isUserSendCommand = true;
            Speedup_Interval();
        }

        private string NDigit_Selector(int Digit)
        {
            string Command = "5STG";
            if (Digit == 0)
            {
                Digit_Three_Border.BorderBrush = Selected;
                Command = "3STG";
            }
            else
            {
                Digit_Three_Border.BorderBrush = Deselected;
            }
            if (Digit == 1)
            {
                Digit_Four_Border.BorderBrush = Selected;
                Command = "4STG";
            }
            else
            {
                Digit_Four_Border.BorderBrush = Deselected;
            }
            if (Digit == 2)
            {
                Digit_Five_Border.BorderBrush = Selected;
                Command = "5STG";
            }
            else
            {
                Digit_Five_Border.BorderBrush = Deselected;
            }
            if (Digit == 3)
            {
                Digit_Six_Border.BorderBrush = Selected;
                Command = "6STG";
            }
            else
            {
                Digit_Six_Border.BorderBrush = Deselected;
            }
            return Command;
        }

        //-----------------------------------------------------------------------

        //-----------------------Auto Zero Options------------------------------------------------

        private void AutoZero_On_Button_Click(object sender, RoutedEventArgs e)
        {
            SerialWriteQueue.Add(AutoZero_Selector(0));
            lockControls();
            isUserSendCommand = true;
            Speedup_Interval();
            insert_Log("AutoZero is Enabled.", 3);
        }

        private void AutoZero_Off_Button_Click(object sender, RoutedEventArgs e)
        {
            SerialWriteQueue.Add(AutoZero_Selector(1));
            lockControls();
            isUserSendCommand = true;
            Speedup_Interval();
            insert_Log("AutoZero is Disabled.", 3);
        }

        private string AutoZero_Selector(int AutoZero)
        {
            string Command = "Z1";
            if (AutoZero == 0)
            {
                AutoZero_On_Border.BorderBrush = Selected;
                Command = "Z1";
            }
            else
            {
                AutoZero_On_Border.BorderBrush = Deselected;
            }
            if (AutoZero == 1)
            {
                AutoZero_Off_Border.BorderBrush = Selected;
                Command = "Z0";
            }
            else
            {
                AutoZero_Off_Border.BorderBrush = Deselected;
            }
            return Command;
        }

        //-----------------------------------------------------------------------

        //-----------------------Filter Options------------------------------------------------

        private void Filter_On_Button_Click(object sender, RoutedEventArgs e)
        {
            if (Measurement_Selected == 2 || Measurement_Selected == 3 || Measurement_Selected == 8 || Measurement_Selected == 9)
            {
                insert_Log("Filter cannot be enabled when Ohms measurement is selected.", 2);
            }
            else
            {
                SerialWriteQueue.Add(Filter_Selector(0));
                lockControls();
                isUserSendCommand = true;
                Speedup_Interval();
                insert_Log("Filter is Enabled.", 3);
            }
        }

        private void Filter_Off_Button_Click(object sender, RoutedEventArgs e)
        {
            SerialWriteQueue.Add(Filter_Selector(1));
            lockControls();
            isUserSendCommand = true;
            Speedup_Interval();
            insert_Log("Filter is Disabled.", 3);
        }

        private string Filter_Selector(int Filter)
        {
            string Command = "FL0";
            if (Filter == 0)
            {
                Filter_On_Border.BorderBrush = Selected;
                Command = "FL1";
            }
            else
            {
                Filter_On_Border.BorderBrush = Deselected;
            }
            if (Filter == 1)
            {
                Filter_Off_Border.BorderBrush = Selected;
                Command = "FL0";
            }
            else
            {
                Filter_Off_Border.BorderBrush = Deselected;
            }
            return Command;
        }

        //-----------------------------------------------------------------------

        //--------------------------NPLC Options---------------------------------------------

        private void NPLC_001_Button_Click(object sender, RoutedEventArgs e)
        {
            SerialWriteQueue.Add(NPLC_Selector(0));
            lockControls();
            insert_Log("NPLC set to 0.01", 3);
            isUserSendCommand = true;
            Speedup_Interval();
        }

        private void NPLC_01_Button_Click(object sender, RoutedEventArgs e)
        {
            SerialWriteQueue.Add(NPLC_Selector(1));
            lockControls();
            insert_Log("NPLC set to 0.1", 3);
            isUserSendCommand = true;
            Speedup_Interval();
        }

        private void NPLC_1_Button_Click(object sender, RoutedEventArgs e)
        {
            SerialWriteQueue.Add(NPLC_Selector(2));
            lockControls();
            insert_Log("NPLC set to 1", 3);
            isUserSendCommand = true;
            Speedup_Interval();
        }

        private void NPLC_10_Button_Click(object sender, RoutedEventArgs e)
        {
            if (Measurement_Selected == 8 || Measurement_Selected == 9)
            {
                insert_Log("Cannot set NPLC to 10 when OC 2Wire or OC 4Wire Ohms measurement is selected.", 2);
            }
            else
            {
                SerialWriteQueue.Add(NPLC_Selector(3));
                lockControls();
                insert_Log("NPLC set to 10", 3);
                isUserSendCommand = true;
                Speedup_Interval();
            }
        }

        private void NPLC_100_Button_Click(object sender, RoutedEventArgs e)
        {
            if (Measurement_Selected == 8 || Measurement_Selected == 9 || Filter_indicator == 1)
            {
                insert_Log("Cannot set NPLC to 100 when OC 2Wire or OC 4Wire Ohms measurement is selected.", 2);
                insert_Log("Or when the Filter is enabled.", 2);
            }
            else
            {
                SerialWriteQueue.Add(NPLC_Selector(4));
                lockControls();
                insert_Log("NPLC set to 100", 3);
                insert_Log("When NPLC is set to 100, 3456A may ignore GPIB write commands.", 2);
                insert_Log("Software will capture measurement data extremely slowly and timeout counter may go up.", 2);
                insert_Log("NPLC 100 is not recommended.", 2);
                isUserSendCommand = true;
                Speedup_Interval();
            }
        }

        private string NPLC_Selector(int Digit)
        {
            string Command = "10STI";
            if (Digit == 0)
            {
                NPLC_001_Border.BorderBrush = Selected;
                Command = "0.01STI";
            }
            else
            {
                NPLC_001_Border.BorderBrush = Deselected;
            }
            if (Digit == 1)
            {
                NPLC_01_Border.BorderBrush = Selected;
                Command = "0.1STI";
            }
            else
            {
                NPLC_01_Border.BorderBrush = Deselected;
            }
            if (Digit == 2)
            {
                NPLC_1_Border.BorderBrush = Selected;
                Command = "1STI";
            }
            else
            {
                NPLC_1_Border.BorderBrush = Deselected;
            }
            if (Digit == 3)
            {
                NPLC_10_Border.BorderBrush = Selected;
                Command = "10STI";
            }
            else
            {
                NPLC_10_Border.BorderBrush = Deselected;
            }
            if (Digit == 4)
            {
                NPLC_100_Border.BorderBrush = Selected;
                Command = "100STI";
            }
            else
            {
                NPLC_100_Border.BorderBrush = Deselected;
            }
            return Command;
        }

        //------------------------------------------------------------------------------------

        //--------------------------Trigger Options---------------------------------------------

        private void Trigger_Internal_Button_Click(object sender, RoutedEventArgs e)
        {
            SerialWriteQueue.Add(Trigger_Selector(0));
            lockControls();
            isUserSendCommand = true;
            Speedup_Interval();
            insert_Log("Internal Trigger Selected.", 3);
        }

        private void Trigger_External_Button_Click(object sender, RoutedEventArgs e)
        {
            SerialWriteQueue.Add(Trigger_Selector(1));
            lockControls();
            isUserSendCommand = true;
            Speedup_Interval();
            insert_Log("External Trigger Selected.", 3);
        }

        private void Trigger_Single_Button_Click(object sender, RoutedEventArgs e)
        {
            SerialWriteQueue.Add(Trigger_Selector(2));
            lockControls();
            isUserSendCommand = true;
            Speedup_Interval();
            insert_Log("Single Trigger Selected.", 3);
        }

        private void Trigger_Hold_Button_Click(object sender, RoutedEventArgs e)
        {
            SerialWriteQueue.Add(Trigger_Selector(3));
            lockControls();
            isUserSendCommand = true;
            Speedup_Interval();
            insert_Log("Hold Trigger Selected.", 3);
        }

        private string Trigger_Selector(int Trigger)
        {
            string Command = "T1";
            if (Trigger == 0)
            {
                Trigger_Internal_Border.BorderBrush = Selected;
                Command = "T1";
            }
            else
            {
                Trigger_Internal_Border.BorderBrush = Deselected;
            }
            if (Trigger == 1)
            {
                Trigger_External_Border.BorderBrush = Selected;
                Command = "T2";
            }
            else
            {
                Trigger_External_Border.BorderBrush = Deselected;
            }
            if (Trigger == 2)
            {
                Trigger_Single_Border.BorderBrush = Selected;
                Command = "T3";
            }
            else
            {
                Trigger_Single_Border.BorderBrush = Deselected;
            }
            if (Trigger == 3)
            {
                Trigger_Hold_Border.BorderBrush = Selected;
                Command = "T4";
            }
            else
            {
                Trigger_Hold_Border.BorderBrush = Deselected;
            }
            return Command;
        }

        //-----------------------------------------------------------------------

        //--------------------------------Display Options---------------------------------------

        private void Display_On_Button_Click(object sender, RoutedEventArgs e)
        {
            SerialWriteQueue.Add(Display_Selector(0));
            lockControls();
            isUserSendCommand = true;
            Speedup_Interval();
            insert_Log("Front Panel Display is On.", 3);
        }

        private void Display_Off_Button_Click(object sender, RoutedEventArgs e)
        {
            SerialWriteQueue.Add(Display_Selector(1));
            lockControls();
            isUserSendCommand = true;
            Speedup_Interval();
            insert_Log("Front Panel Display is Off.", 3);
        }

        private string Display_Selector(int Display)
        {
            string Command = "D1";
            if (Display == 0)
            {
                Display_On_Border.BorderBrush = Selected;
                Command = "D1";
            }
            else
            {
                Display_On_Border.BorderBrush = Deselected;
            }
            if (Display == 1)
            {
                Display_Off_Border.BorderBrush = Selected;
                Command = "D0";
            }
            else
            {
                Display_Off_Border.BorderBrush = Deselected;
            }
            return Command;
        }

        //-----------------------------------------------------------------------

        //----------------------------Speech Setup-------------------------------------------

        private void Speech_Continuous_Set_Button_Click(object sender, RoutedEventArgs e)
        {
            if (isSpeechActive == 1)
            {
                (bool isNum, double value) = isNumber(Speech_Continuous_Value.Text);
                if (isNum == true)
                {
                    if (value > 0)
                    {
                        Interlocked.Exchange(ref isSpeechContinuous, 1);
                        Speech_Measurement_Interval.Interval = (value * 60000);
                        Continuous_Selector(0);
                        insert_Log("Continuously voice measurement every " + value + " minutes.", 0);
                        Speech_Measurement_Interval.Start();
                    }
                    else
                    {
                        insert_Log("Continuous voice value must be a positive number.", 1);
                    }
                }
                else
                {
                    insert_Log("Continuous voice value must be a positive number.", 1);
                }
            }
            else
            {
                insert_Log("The Speech Synthesizer is not enabled. Enable it from Speech Menu.", 2);
            }
        }

        private void Speech_Continuous_Clear_Button_Click(object sender, RoutedEventArgs e)
        {
            if (isSpeechActive == 1)
            {
                Interlocked.Exchange(ref isSpeechContinuous, 0);
                Speech_Continuous_Value.Text = string.Empty;
                Continuous_Selector(1);
                insert_Log("Continuous voice measurement is cleared.", 0);
            }
            else
            {
                insert_Log("The Speech Synthesizer is not enabled. Enable it from Speech Menu.", 2);
            }

        }

        private void Continuous_Selector(int status)
        {
            if (status == 0)
            {
                Speech_Continuous_Set_Border.BorderBrush = Selected;
            }
            else
            {
                Speech_Continuous_Set_Border.BorderBrush = Deselected;
            }
            if (status == 1)
            {
                Speech_Continuous_Clear_Border.BorderBrush = Selected;
            }
            else
            {
                Speech_Continuous_Clear_Border.BorderBrush = Deselected;
            }
        }

        private void Speech_MIN_Button_Click(object sender, RoutedEventArgs e)
        {
            if (isSpeechActive == 1)
            {
                (bool isNum, double value) = isNumber(Speech_MIN_Value.Text);
                if (isNum == true)
                {
                    Interlocked.Exchange(ref Speech_min_value, value);
                    Interlocked.Exchange(ref isSpeechMIN, 1);
                    MIN_Selector(0);
                    insert_Log("Voice measurement less than " + value, 0);
                    Speech_MIN_Max.Start();
                }
                else
                {
                    insert_Log("MIN voice value must be a number.", 1);
                }
            }
            else
            {
                insert_Log("The Speech Synthesizer is not enabled. Enable it from Speech Menu.", 2);
            }
        }

        private void Speech_MIN_Clear_Button_Click(object sender, RoutedEventArgs e)
        {
            if (isSpeechActive == 1)
            {
                Interlocked.Exchange(ref isSpeechMIN, 0);
                Interlocked.Exchange(ref Speech_min_value, 0);
                Speech_MIN_Value.Text = string.Empty;
                MIN_Selector(1);
                insert_Log("MIN voice measurement is cleared.", 0);
            }
            else
            {
                insert_Log("The Speech Synthesizer is not enabled. Enable it from Speech Menu.", 2);
            }
        }

        private void MIN_Selector(int status)
        {
            if (status == 0)
            {
                Speech_MIN_Border.BorderBrush = Selected;
            }
            else
            {
                Speech_MIN_Border.BorderBrush = Deselected;
            }
            if (status == 1)
            {
                Speech_MIN_Clear_Border.BorderBrush = Selected;
            }
            else
            {
                Speech_MIN_Clear_Border.BorderBrush = Deselected;
            }
        }

        private void Speech_MAX_Button_Click(object sender, RoutedEventArgs e)
        {
            if (isSpeechActive == 1)
            {
                (bool isNum, double value) = isNumber(Speech_MAX_Value.Text);
                if (isNum == true)
                {
                    Interlocked.Exchange(ref Speech_max_value, value);
                    Interlocked.Exchange(ref isSpeechMAX, 1);
                    MAX_Selector(0);
                    insert_Log("Voice measurement greater than " + value, 0);
                    Speech_MIN_Max.Start();
                }
                else
                {
                    insert_Log("MAX voice value must be a number.", 1);
                }
            }
            else
            {
                insert_Log("The Speech Synthesizer is not enabled. Enable it from Speech Menu.", 2);
            }
        }

        private void Speech_MAX_Clear_Button_Click(object sender, RoutedEventArgs e)
        {
            if (isSpeechActive == 1)
            {
                Interlocked.Exchange(ref isSpeechMAX, 0);
                Interlocked.Exchange(ref Speech_max_value, 0);
                Speech_MAX_Value.Text = string.Empty;
                MAX_Selector(1);
                insert_Log("MAX voice measurement is cleared.", 0);
            }
            else
            {
                insert_Log("The Speech Synthesizer is not enabled. Enable it from Speech Menu.", 2);
            }
        }

        private void MAX_Selector(int status)
        {
            if (status == 0)
            {
                Speech_MAX_Border.BorderBrush = Selected;
            }
            else
            {
                Speech_MAX_Border.BorderBrush = Deselected;
            }
            if (status == 1)
            {
                Speech_MAX_Clear_Border.BorderBrush = Selected;
            }
            else
            {
                Speech_MAX_Clear_Border.BorderBrush = Deselected;
            }
        }

        //-----------------------------------------------------------------------

        //-----------------------------Update Speed Options------------------------------------------

        private void UpdateSpeed_Value_Set_Button_Click(object sender, RoutedEventArgs e)
        {
            (bool isNum, double value) = isNumber(UpdateSpeed_Value.Text);
            if (isNum == true)
            {
                if (value > 0)
                {
                    insert_Log("You may need to wait for " + (UpdateSpeed / 1000) + " seconds before your new update speed takes effect.", 2);
                    insert_Log("Update Speed set to " + value + " seconds Command Send.", 5);
                    value = value * 1000;
                    UpdateSpeed = value;
                    UpdateSpeed_Selector(0);
                    isUpdateSpeed_Changed = true;
                }
                else
                {
                    insert_Log("Update Speed must be number greater than 0. Minimum value can be 0.01 seconds.", 1);
                }
            }
            else
            {
                insert_Log("Update Speed must be number.", 1);
            }
        }

        private void UpdateSpeed_Default_Set_Button_Click(object sender, RoutedEventArgs e)
        {
            insert_Log("You may to wait for " + (UpdateSpeed / 1000) + " seconds before your new update speed takes effect.", 2);
            UpdateSpeed = 1000;
            insert_Log("Update Speed set to " + (UpdateSpeed / 1000) + " seconds Command Send.", 5);
            UpdateSpeed_Selector(1);
            isUpdateSpeed_Changed = true;
        }

        private void UpdateSpeed_Fast_Set_Button_Click(object sender, RoutedEventArgs e)
        {
            insert_Log("You may to wait for " + (UpdateSpeed / 1000) + " seconds before your new update speed takes effect.", 2);
            UpdateSpeed = 10;
            insert_Log("Update Speed set to " + (UpdateSpeed / 1000) + " seconds Command Send.", 5);
            UpdateSpeed_Selector(2);
            isUpdateSpeed_Changed = true;
        }

        private void UpdateSpeed_Selector(int status)
        {
            if (status == 0)
            {
                UpdateSpeed_Value_Set_Border.BorderBrush = Selected;
            }
            else
            {
                UpdateSpeed_Value_Set_Border.BorderBrush = Deselected;
            }
            if (status == 1)
            {
                UpdateSpeed_Default_Set_Border.BorderBrush = Selected;
            }
            else
            {
                UpdateSpeed_Default_Set_Border.BorderBrush = Deselected;
            }
            if (status == 2)
            {
                UpdateSpeed_Fast_Set_Border.BorderBrush = Selected;
            }
            else
            {
                UpdateSpeed_Fast_Set_Border.BorderBrush = Deselected;
            }
        }

        //-----------------------------------------------------------------------

        private void Measurement_Green_Click(object sender, RoutedEventArgs e)
        {
            Measurement_Color_Checker(0);
            Measurement_Color("#FF00FF17");

        }

        private void Measurement_Blue_Click(object sender, RoutedEventArgs e)
        {
            Measurement_Color_Checker(1);
            Measurement_Color("#FF00C0FF");
        }

        private void Measurement_Red_Click(object sender, RoutedEventArgs e)
        {
            Measurement_Color_Checker(2);
            Measurement_Color("Red");
        }

        private void Measurement_Yellow_Click(object sender, RoutedEventArgs e)
        {
            Measurement_Color_Checker(3);
            Measurement_Color("#FFFFFF00");
        }

        private void Measurement_Orange_Click(object sender, RoutedEventArgs e)
        {
            Measurement_Color_Checker(4);
            Measurement_Color("DarkOrange");
        }

        private void Measurement_Pink_Click(object sender, RoutedEventArgs e)
        {
            Measurement_Color_Checker(5);
            Measurement_Color("DeepPink");
        }

        private void Measurement_White_Click(object sender, RoutedEventArgs e)
        {
            Measurement_Color_Checker(6);
            Measurement_Color("White");
        }

        private void Measurement_Black_Click(object sender, RoutedEventArgs e)
        {
            Measurement_Color_Checker(7);
            Measurement_Color("Black");
        }

        private void Measurement_Color(string HexValue)
        {
            SolidColorBrush Color = new SolidColorBrush((Color)ColorConverter.ConvertFromString(HexValue));
            Measurement_Value.Foreground = Color;
            Measurement_Scale.Foreground = Color;
            Measurement_Type.Foreground = Color;
        }

        private void Measurement_Color_Checker(int Check)
        {
            if (Check == 0)
            {
                Measurement_Green.IsChecked = true;
            }
            else
            {
                Measurement_Green.IsChecked = false;
            }
            if (Check == 1)
            {
                Measurement_Blue.IsChecked = true;
            }
            else
            {
                Measurement_Blue.IsChecked = false;
            }
            if (Check == 2)
            {
                Measurement_Red.IsChecked = true;
            }
            else
            {
                Measurement_Red.IsChecked = false;
            }
            if (Check == 3)
            {
                Measurement_Yellow.IsChecked = true;
            }
            else
            {
                Measurement_Yellow.IsChecked = false;
            }
            if (Check == 4)
            {
                Measurement_Orange.IsChecked = true;
            }
            else
            {
                Measurement_Orange.IsChecked = false;
            }
            if (Check == 5)
            {
                Measurement_Pink.IsChecked = true;
            }
            else
            {
                Measurement_Pink.IsChecked = false;
            }
            if (Check == 6)
            {
                Measurement_White.IsChecked = true;
            }
            else
            {
                Measurement_White.IsChecked = false;
            }
            if (Check == 7)
            {
                Measurement_Black.IsChecked = true;
            }
            else
            {
                Measurement_Black.IsChecked = false;
            }
        }

        private void MIN_Green_Click(object sender, RoutedEventArgs e)
        {
            MIN_Color_Checker(0);
            MIN_Color("#FF00FF17");
        }

        private void MIN_Blue_Click(object sender, RoutedEventArgs e)
        {
            MIN_Color_Checker(1);
            MIN_Color("#FF00C0FF");
        }

        private void MIN_Red_Click(object sender, RoutedEventArgs e)
        {
            MIN_Color_Checker(2);
            MIN_Color("Red");
        }

        private void MIN_Yellow_Click(object sender, RoutedEventArgs e)
        {
            MIN_Color_Checker(3);
            MIN_Color("#FFFFFF00");
        }

        private void MIN_Orange_Click(object sender, RoutedEventArgs e)
        {
            MIN_Color_Checker(4);
            MIN_Color("DarkOrange");
        }

        private void MIN_Pink_Click(object sender, RoutedEventArgs e)
        {
            MIN_Color_Checker(5);
            MIN_Color("DeepPink");
        }

        private void MIN_White_Click(object sender, RoutedEventArgs e)
        {
            MIN_Color_Checker(6);
            MIN_Color("White");
        }

        private void MIN_Black_Click(object sender, RoutedEventArgs e)
        {
            MIN_Color_Checker(7);
            MIN_Color("Black");
        }

        private void MIN_Color(string HexValue)
        {
            SolidColorBrush Color = new SolidColorBrush((Color)ColorConverter.ConvertFromString(HexValue));
            MIN_Value.Foreground = Color;
            MIN_Scale.Foreground = Color;
            MIN_Type.Foreground = Color;
            MIN_Label.Foreground = Color;
        }

        private void MIN_Color_Checker(int Check)
        {
            if (Check == 0)
            {
                MIN_Green.IsChecked = true;
            }
            else
            {
                MIN_Green.IsChecked = false;
            }
            if (Check == 1)
            {
                MIN_Blue.IsChecked = true;
            }
            else
            {
                MIN_Blue.IsChecked = false;
            }
            if (Check == 2)
            {
                MIN_Red.IsChecked = true;
            }
            else
            {
                MIN_Red.IsChecked = false;
            }
            if (Check == 3)
            {
                MIN_Yellow.IsChecked = true;
            }
            else
            {
                MIN_Yellow.IsChecked = false;
            }
            if (Check == 4)
            {
                MIN_Orange.IsChecked = true;
            }
            else
            {
                MIN_Orange.IsChecked = false;
            }
            if (Check == 5)
            {
                MIN_Pink.IsChecked = true;
            }
            else
            {
                MIN_Pink.IsChecked = false;
            }
            if (Check == 6)
            {
                MIN_White.IsChecked = true;
            }
            else
            {
                MIN_White.IsChecked = false;
            }
            if (Check == 7)
            {
                MIN_Black.IsChecked = true;
            }
            else
            {
                MIN_Black.IsChecked = false;
            }
        }

        private void MAX_Green_Click(object sender, RoutedEventArgs e)
        {
            MAX_Color_Checker(0);
            MAX_Color("#FF00FF17");
        }

        private void MAX_Blue_Click(object sender, RoutedEventArgs e)
        {
            MAX_Color_Checker(1);
            MAX_Color("#FF00C0FF");
        }

        private void MAX_Red_Click(object sender, RoutedEventArgs e)
        {
            MAX_Color_Checker(2);
            MAX_Color("Red");
        }

        private void MAX_Yellow_Click(object sender, RoutedEventArgs e)
        {
            MAX_Color_Checker(3);
            MAX_Color("#FFFFFF00");
        }

        private void MAX_Orange_Click(object sender, RoutedEventArgs e)
        {
            MAX_Color_Checker(4);
            MAX_Color("DarkOrange");
        }

        private void MAX_Pink_Click(object sender, RoutedEventArgs e)
        {
            MAX_Color_Checker(5);
            MAX_Color("DeepPink");
        }

        private void MAX_White_Click(object sender, RoutedEventArgs e)
        {
            MAX_Color_Checker(6);
            MAX_Color("White");
        }

        private void MAX_Black_Click(object sender, RoutedEventArgs e)
        {
            MAX_Color_Checker(7);
            MAX_Color("Black");
        }

        private void MAX_Color(string HexValue)
        {
            SolidColorBrush Color = new SolidColorBrush((Color)ColorConverter.ConvertFromString(HexValue));
            MAX_Value.Foreground = Color;
            MAX_Scale.Foreground = Color;
            MAX_Type.Foreground = Color;
            MAX_Label.Foreground = Color;
        }

        private void MAX_Color_Checker(int Check)
        {
            if (Check == 0)
            {
                MAX_Green.IsChecked = true;
            }
            else
            {
                MAX_Green.IsChecked = false;
            }
            if (Check == 1)
            {
                MAX_Blue.IsChecked = true;
            }
            else
            {
                MAX_Blue.IsChecked = false;
            }
            if (Check == 2)
            {
                MAX_Red.IsChecked = true;
            }
            else
            {
                MAX_Red.IsChecked = false;
            }
            if (Check == 3)
            {
                MAX_Yellow.IsChecked = true;
            }
            else
            {
                MAX_Yellow.IsChecked = false;
            }
            if (Check == 4)
            {
                MAX_Orange.IsChecked = true;
            }
            else
            {
                MAX_Orange.IsChecked = false;
            }
            if (Check == 5)
            {
                MAX_Pink.IsChecked = true;
            }
            else
            {
                MAX_Pink.IsChecked = false;
            }
            if (Check == 6)
            {
                MAX_White.IsChecked = true;
            }
            else
            {
                MAX_White.IsChecked = false;
            }
            if (Check == 7)
            {
                MAX_Black.IsChecked = true;
            }
            else
            {
                MAX_Black.IsChecked = false;
            }
        }

        private void AVG_Green_Click(object sender, RoutedEventArgs e)
        {
            AVG_Color_Checker(0);
            AVG_Color("#FF00FF17");
        }

        private void AVG_Blue_Click(object sender, RoutedEventArgs e)
        {
            AVG_Color_Checker(1);
            AVG_Color("#FF00C0FF");
        }

        private void AVG_Red_Click(object sender, RoutedEventArgs e)
        {
            AVG_Color_Checker(2);
            AVG_Color("Red");
        }

        private void AVG_Yellow_Click(object sender, RoutedEventArgs e)
        {
            AVG_Color_Checker(3);
            AVG_Color("#FFFFFF00");
        }

        private void AVG_Orange_Click(object sender, RoutedEventArgs e)
        {
            AVG_Color_Checker(4);
            AVG_Color("DarkOrange");
        }

        private void AVG_Pink_Click(object sender, RoutedEventArgs e)
        {
            AVG_Color_Checker(5);
            AVG_Color("DeepPink");
        }

        private void AVG_White_Click(object sender, RoutedEventArgs e)
        {
            AVG_Color_Checker(6);
            AVG_Color("White");
        }

        private void AVG_Black_Click(object sender, RoutedEventArgs e)
        {
            AVG_Color_Checker(7);
            AVG_Color("Black");
        }

        private void AVG_Color(string HexValue)
        {
            SolidColorBrush Color = new SolidColorBrush((Color)ColorConverter.ConvertFromString(HexValue));
            AVG_Value.Foreground = Color;
            AVG_Scale.Foreground = Color;
            AVG_Type.Foreground = Color;
            AVG_Label.Foreground = Color;
        }

        private void AVG_Color_Checker(int Check)
        {
            if (Check == 0)
            {
                AVG_Green.IsChecked = true;
            }
            else
            {
                AVG_Green.IsChecked = false;
            }
            if (Check == 1)
            {
                AVG_Blue.IsChecked = true;
            }
            else
            {
                AVG_Blue.IsChecked = false;
            }
            if (Check == 2)
            {
                AVG_Red.IsChecked = true;
            }
            else
            {
                AVG_Red.IsChecked = false;
            }
            if (Check == 3)
            {
                AVG_Yellow.IsChecked = true;
            }
            else
            {
                AVG_Yellow.IsChecked = false;
            }
            if (Check == 4)
            {
                AVG_Orange.IsChecked = true;
            }
            else
            {
                AVG_Orange.IsChecked = false;
            }
            if (Check == 5)
            {
                AVG_Pink.IsChecked = true;
            }
            else
            {
                AVG_Pink.IsChecked = false;
            }
            if (Check == 6)
            {
                AVG_White.IsChecked = true;
            }
            else
            {
                AVG_White.IsChecked = false;
            }
            if (Check == 7)
            {
                AVG_Black.IsChecked = true;
            }
            else
            {
                AVG_Black.IsChecked = false;
            }
        }

        private void Background_White_Click(object sender, RoutedEventArgs e)
        {
            Background_Color_Checker(0);
            Background_Color("White");
        }

        private void Background_Black_Click(object sender, RoutedEventArgs e)
        {
            Background_Color_Checker(1);
            Background_Color("Black");
        }

        private void Background_Color(string HexValue)
        {
            SolidColorBrush Color = new SolidColorBrush((Color)ColorConverter.ConvertFromString(HexValue));
            DisplayPanel_Background.Background = Color;
        }

        private void Background_Color_Checker(int Check)
        {
            if (Check == 0)
            {
                Background_White.IsChecked = true;
            }
            else
            {
                Background_White.IsChecked = false;
            }
            if (Check == 1)
            {
                Background_Black.IsChecked = true;
            }
            else
            {
                Background_Black.IsChecked = false;
            }
        }

        private void Old_LCD_Click(object sender, RoutedEventArgs e)
        {
            Original_Display = true;
            Modern_LCD.IsChecked = false;
            Old_LCD.IsChecked = true;
        }

        private void Modern_LCD_Click(object sender, RoutedEventArgs e)
        {
            Original_Display = false;
            Modern_LCD.IsChecked = true;
            Old_LCD.IsChecked = false;
        }

        private void Calculate_AVG_Click(object sender, RoutedEventArgs e)
        {
            if (Calculate_AVG.IsChecked == true)
            {
                Interlocked.Exchange(ref AVG_Calculate, 1);
                insert_Log("Average will be calculated.", 0);
            }
            else
            {
                Interlocked.Exchange(ref AVG_Calculate, 0);
                insert_Log("Average will not be calculated.", 2);
            }
        }

        private void AVG_Res_2_Click(object sender, RoutedEventArgs e)
        {
            Interlocked.Exchange(ref avg_resolution, 2);
            insert_Log("Average's resolution set to " + avg_resolution + ".", 0);
            AVG_Res_Selected();
        }

        private void AVG_Res_3_Click(object sender, RoutedEventArgs e)
        {
            Interlocked.Exchange(ref avg_resolution, 3);
            insert_Log("Average's resolution set to " + avg_resolution + ".", 0);
            AVG_Res_Selected();
        }

        private void AVG_Res_4_Click(object sender, RoutedEventArgs e)
        {
            Interlocked.Exchange(ref avg_resolution, 4);
            insert_Log("Average's resolution set to " + avg_resolution + ".", 0);
            AVG_Res_Selected();
        }

        private void AVG_Res_5_Click(object sender, RoutedEventArgs e)
        {
            Interlocked.Exchange(ref avg_resolution, 5);
            insert_Log("Average's resolution set to " + avg_resolution + ".", 0);
            AVG_Res_Selected();
        }

        private void AVG_Res_6_Click(object sender, RoutedEventArgs e)
        {
            Interlocked.Exchange(ref avg_resolution, 6);
            insert_Log("Average's resolution set to " + avg_resolution + ".", 0);
            AVG_Res_Selected();
        }

        private void AVG_Res_Selected()
        {
            if (avg_resolution == 2)
            {
                AVG_Res_2.IsChecked = true;
            }
            else
            {
                AVG_Res_2.IsChecked = false;
            }
            if (avg_resolution == 3)
            {
                AVG_Res_3.IsChecked = true;
            }
            else
            {
                AVG_Res_3.IsChecked = false;
            }
            if (avg_resolution == 4)
            {
                AVG_Res_4.IsChecked = true;
            }
            else
            {
                AVG_Res_4.IsChecked = false;
            }
            if (avg_resolution == 5)
            {
                AVG_Res_5.IsChecked = true;
            }
            else
            {
                AVG_Res_5.IsChecked = false;
            }
            if (avg_resolution == 6)
            {
                AVG_Res_6.IsChecked = true;
            }
            else
            {
                AVG_Res_6.IsChecked = false;
            }
        }

        private void Factor_50_Click(object sender, RoutedEventArgs e)
        {
            Interlocked.Exchange(ref avg_factor, 50);
            insert_Log("Average's factor set to " + avg_factor + ".", 0);
            AVG_Fac_Selected();
        }

        private void Factor_100_Click(object sender, RoutedEventArgs e)
        {
            Interlocked.Exchange(ref avg_factor, 100);
            insert_Log("Average's factor set to " + avg_factor + ".", 0);
            AVG_Fac_Selected();
        }

        private void Factor_200_Click(object sender, RoutedEventArgs e)
        {
            Interlocked.Exchange(ref avg_factor, 200);
            insert_Log("Average's factor set to " + avg_factor + ".", 0);
            AVG_Fac_Selected();
        }

        private void Factor_400_Click(object sender, RoutedEventArgs e)
        {
            Interlocked.Exchange(ref avg_factor, 400);
            insert_Log("Average's factor set to " + avg_factor + ".", 0);
            AVG_Fac_Selected();
        }

        private void Factor_800_Click(object sender, RoutedEventArgs e)
        {
            Interlocked.Exchange(ref avg_factor, 800);
            insert_Log("Average's factor set to " + avg_factor + ".", 0);
            AVG_Fac_Selected();
        }

        private void Factor_1000_Click(object sender, RoutedEventArgs e)
        {
            Interlocked.Exchange(ref avg_factor, 1000);
            insert_Log("Average's factor set to " + avg_factor + ".", 0);
            AVG_Fac_Selected();
        }

        private void AVG_Fac_Selected()
        {
            if (avg_factor == 50)
            {
                Factor_50.IsChecked = true;
            }
            else
            {
                Factor_50.IsChecked = false;
            }
            if (avg_factor == 100)
            {
                Factor_100.IsChecked = true;
            }
            else
            {
                Factor_100.IsChecked = false;
            }
            if (avg_factor == 200)
            {
                Factor_200.IsChecked = true;
            }
            else
            {
                Factor_200.IsChecked = false;
            }
            if (avg_factor == 400)
            {
                Factor_400.IsChecked = true;
            }
            else
            {
                Factor_400.IsChecked = false;
            }
            if (avg_factor == 800)
            {
                Factor_800.IsChecked = true;
            }
            else
            {
                Factor_800.IsChecked = false;
            }
            if (avg_factor == 1000)
            {
                Factor_1000.IsChecked = true;
            }
            else
            {
                Factor_1000.IsChecked = false;
            }
        }

        private void Voice_Precision_0_Click(object sender, RoutedEventArgs e)
        {
            Interlocked.Exchange(ref Speech_Value_Precision, 0);
            insert_Log("Speech's precision is set to " + Speech_Value_Precision + ".", 0);
            Voice_Precision_Selected();
        }

        private void Voice_Precision_1_Click(object sender, RoutedEventArgs e)
        {
            Interlocked.Exchange(ref Speech_Value_Precision, 1);
            insert_Log("Speech's precision is set to " + Speech_Value_Precision + ".", 0);
            Voice_Precision_Selected();
        }

        private void Voice_Precision_2_Click(object sender, RoutedEventArgs e)
        {
            Interlocked.Exchange(ref Speech_Value_Precision, 2);
            insert_Log("Speech's precision is set to " + Speech_Value_Precision + ".", 0);
            Voice_Precision_Selected();
        }

        private void Voice_Precision_3_Click(object sender, RoutedEventArgs e)
        {
            Interlocked.Exchange(ref Speech_Value_Precision, 3);
            insert_Log("Speech's precision is set to " + Speech_Value_Precision + ".", 0);
            Voice_Precision_Selected();
        }

        private void Voice_Precision_4_Click(object sender, RoutedEventArgs e)
        {
            Interlocked.Exchange(ref Speech_Value_Precision, 4);
            insert_Log("Speech's precision is set to " + Speech_Value_Precision + ".", 0);
            Voice_Precision_Selected();
        }

        private void Voice_Precision_5_Click(object sender, RoutedEventArgs e)
        {
            Interlocked.Exchange(ref Speech_Value_Precision, 5);
            insert_Log("Speech's precision is set to " + Speech_Value_Precision + ".", 0);
            Voice_Precision_Selected();
        }

        private void Voice_Precision_6_Click(object sender, RoutedEventArgs e)
        {
            Interlocked.Exchange(ref Speech_Value_Precision, 6);
            insert_Log("Speech's precision is set to " + Speech_Value_Precision + ".", 0);
            Voice_Precision_Selected();
        }

        private void Voice_Precision_7_Click(object sender, RoutedEventArgs e)
        {
            Interlocked.Exchange(ref Speech_Value_Precision, 7);
            insert_Log("Speech's precision is set to " + Speech_Value_Precision + ".", 0);
            Voice_Precision_Selected();
        }

        private void Voice_Precision_Selected()
        {
            if (Speech_Value_Precision == 0)
            {
                Voice_Precision_0.IsChecked = true;
            }
            else
            {
                Voice_Precision_0.IsChecked = false;
            }
            if (Speech_Value_Precision == 1)
            {
                Voice_Precision_1.IsChecked = true;
            }
            else
            {
                Voice_Precision_1.IsChecked = false;
            }
            if (Speech_Value_Precision == 2)
            {
                Voice_Precision_2.IsChecked = true;
            }
            else
            {
                Voice_Precision_2.IsChecked = false;
            }
            if (Speech_Value_Precision == 3)
            {
                Voice_Precision_3.IsChecked = true;
            }
            else
            {
                Voice_Precision_3.IsChecked = false;
            }
            if (Speech_Value_Precision == 4)
            {
                Voice_Precision_4.IsChecked = true;
            }
            else
            {
                Voice_Precision_4.IsChecked = false;
            }
            if (Speech_Value_Precision == 5)
            {
                Voice_Precision_5.IsChecked = true;
            }
            else
            {
                Voice_Precision_5.IsChecked = false;
            }
            if (Speech_Value_Precision == 6)
            {
                Voice_Precision_6.IsChecked = true;
            }
            else
            {
                Voice_Precision_6.IsChecked = false;
            }
            if (Speech_Value_Precision == 7)
            {
                Voice_Precision_7.IsChecked = true;
            }
            else
            {
                Voice_Precision_7.IsChecked = false;
            }
        }

        //converts a string into a number
        private (bool, double) Text_Num(string text, bool allowNegative, bool isInteger)
        {
            if (isInteger == true)
            {
                bool isValid = int.TryParse(text, out int value);
                if (isValid == true)
                {
                    if (allowNegative == false)
                    {
                        if (value < 0)
                        {
                            return (false, 0);
                        }
                        else
                        {
                            return (true, value);
                        }
                    }
                    else
                    {
                        return (true, value);
                    }
                }
                else
                {
                    return (false, 0);
                }
            }
            else
            {
                bool isValid = double.TryParse(text, out double value);
                if (isValid == true)
                {
                    if (allowNegative == false)
                    {
                        if (value < 0)
                        {
                            return (false, 0);
                        }
                        else
                        {
                            return (true, value);
                        }
                    }
                    else
                    {
                        return (true, value);
                    }
                }
                else
                {
                    return (false, 0);
                }
            }
        }

        private void Local_Exit_Click(object sender, RoutedEventArgs e)
        {
            SerialWriteQueue.Add("LOCAL_EXIT");
            lockControls();
            isUserSendCommand = true;
            Speedup_Interval();
        }

        private void Load_Main_Window_Settings()
        {
            try
            {
                List<String> Config_Lines = new List<string>();
                string Software_Location = System.IO.Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + @"\" + "Settings.txt";
                string[] Config_Parts;
                using (var readFile = new StreamReader(Software_Location))
                {
                    Config_Parts = readFile.ReadLine().Split(',');
                    Set_Measurement_Color(int.Parse(Config_Parts[0]), int.Parse(Config_Parts[1]), int.Parse(Config_Parts[2]));
                    Config_Parts = readFile.ReadLine().Split(',');
                    Set_MIN_Color(int.Parse(Config_Parts[0]), int.Parse(Config_Parts[1]), int.Parse(Config_Parts[2]));
                    Config_Parts = readFile.ReadLine().Split(',');
                    Set_MAX_Color(int.Parse(Config_Parts[0]), int.Parse(Config_Parts[1]), int.Parse(Config_Parts[2]));
                    Config_Parts = readFile.ReadLine().Split(',');
                    Set_AVG_Color(int.Parse(Config_Parts[0]), int.Parse(Config_Parts[1]), int.Parse(Config_Parts[2]));
                    Config_Parts = readFile.ReadLine().Split(',');
                    Choose_GPIB_Lock(Config_Parts[0].ToUpper().Trim());
                    insert_Log("Settings.txt file loaded.", 0);
                }

            }
            catch (Exception Ex)
            {
                insert_Log(Ex.Message, 2);
                insert_Log("Could not load Settings.txt file, try again.", 2);
            }
        }

        private void Choose_GPIB_Lock(string Set)
        {
            if (Set == "TRUE")
            {
                insert_Log("GPIB Lock set to Exlusive Lock. Other software cannot communicate with HP3456A.", 0);
                GPIB_Lock = 1;
            }
            else if (Set == "FALSE")
            {
                insert_Log("No GPIB Lock set. Other software can communicate with HP3456A.", 0);
                GPIB_Lock = 0;
            }
            else
            {
                insert_Log("Bad String: " + "No GPIB Lock set. Other software can communicate with HP3456A.", 2);
                GPIB_Lock = 0;
            }
        }

        private void Randomize_Display_Colors(object sender, RoutedEventArgs e)
        {
            Random RGB_Value = new Random();
            int Value_Red = RGB_Value.Next(0, 255);
            int Value_Green = RGB_Value.Next(0, 255);
            int Value_Blue = RGB_Value.Next(0, 255);

            Set_Measurement_Color(Value_Red, Value_Green, Value_Blue);
            Set_MIN_Color(Value_Red, Value_Green, Value_Blue);
            Set_MAX_Color(Value_Red, Value_Green, Value_Blue);
            Set_AVG_Color(Value_Red, Value_Green, Value_Blue);

            insert_Log(Value_Red + "," + Value_Green + "," + Value_Blue + "," + "Measurement_Colors_Selected_RGB", 4);
        }

        private void Set_Measurement_Color(int Red, int Green, int Blue)
        {
            Measurement_Color_Checker(9);
            if ((Red <= 255 & Red >= 0) & (Green <= 255 & Green >= 0) & (Blue <= 255 & Blue >= 0))
            {
                Measurement_Value.Foreground = new SolidColorBrush(Color.FromArgb(255, (byte)(Red), (byte)(Green), (byte)(Blue)));
                Measurement_Scale.Foreground = new SolidColorBrush(Color.FromArgb(255, (byte)(Red), (byte)(Green), (byte)(Blue)));
                Measurement_Type.Foreground = new SolidColorBrush(Color.FromArgb(255, (byte)(Red), (byte)(Green), (byte)(Blue)));
            }
            else
            {
                insert_Log("Measurement_Value_Color: RGB Values must be between 0 to 255, try again.", 2);
            }
        }

        private void Set_MIN_Color(int Red, int Green, int Blue)
        {
            MIN_Color_Checker(9);
            if ((Red <= 255 & Red >= 0) & (Green <= 255 & Green >= 0) & (Blue <= 255 & Blue >= 0))
            {
                MIN_Value.Foreground = new SolidColorBrush(Color.FromArgb(255, (byte)(Red), (byte)(Green), (byte)(Blue)));
                MIN_Scale.Foreground = new SolidColorBrush(Color.FromArgb(255, (byte)(Red), (byte)(Green), (byte)(Blue)));
                MIN_Type.Foreground = new SolidColorBrush(Color.FromArgb(255, (byte)(Red), (byte)(Green), (byte)(Blue)));
                MIN_Label.Foreground = new SolidColorBrush(Color.FromArgb(255, (byte)(Red), (byte)(Green), (byte)(Blue)));
            }
            else
            {
                insert_Log("MIN_Value_Color: RGB Values must be between 0 to 255, try again.", 2);
            }
        }

        private void Set_MAX_Color(int Red, int Green, int Blue)
        {
            MAX_Color_Checker(9);
            if ((Red <= 255 & Red >= 0) & (Green <= 255 & Green >= 0) & (Blue <= 255 & Blue >= 0))
            {
                MAX_Value.Foreground = new SolidColorBrush(Color.FromArgb(255, (byte)(Red), (byte)(Green), (byte)(Blue)));
                MAX_Scale.Foreground = new SolidColorBrush(Color.FromArgb(255, (byte)(Red), (byte)(Green), (byte)(Blue)));
                MAX_Type.Foreground = new SolidColorBrush(Color.FromArgb(255, (byte)(Red), (byte)(Green), (byte)(Blue)));
                MAX_Label.Foreground = new SolidColorBrush(Color.FromArgb(255, (byte)(Red), (byte)(Green), (byte)(Blue)));
            }
            else
            {
                insert_Log("MAX_Value_Color: RGB Values must be between 0 to 255, try again.", 2);
            }
        }

        private void Set_AVG_Color(int Red, int Green, int Blue)
        {
            AVG_Color_Checker(9);
            if ((Red <= 255 & Red >= 0) & (Green <= 255 & Green >= 0) & (Blue <= 255 & Blue >= 0))
            {
                AVG_Value.Foreground = new SolidColorBrush(Color.FromArgb(255, (byte)(Red), (byte)(Green), (byte)(Blue)));
                AVG_Scale.Foreground = new SolidColorBrush(Color.FromArgb(255, (byte)(Red), (byte)(Green), (byte)(Blue)));
                AVG_Type.Foreground = new SolidColorBrush(Color.FromArgb(255, (byte)(Red), (byte)(Green), (byte)(Blue)));
                AVG_Label.Foreground = new SolidColorBrush(Color.FromArgb(255, (byte)(Red), (byte)(Green), (byte)(Blue)));
            }
            else
            {
                insert_Log("AVG_Value_Color: RGB Values must be between 0 to 255, try again.", 2);
            }
        }

        private void Window_Closed(object sender, EventArgs e)
        {
            try
            {
                if (GPIB_Address_Info.isConnected == true)
                {
                    if (GPIB_Lock == 1)
                    {
                        session.UnlockResource();
                    }
                    session.Dispose();
                    session = null;
                }
            }
            catch (Exception)
            {

            }
        }
    }
}
