using System.Text;
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
using System;
using System.Collections.ObjectModel;
using System.Security.Cryptography.X509Certificates;
using System.Runtime.CompilerServices;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using System.IO.Pipes;
using System.Text.RegularExpressions;
using System.DirectoryServices.ActiveDirectory;
using System.Runtime.InteropServices;
using System.Windows.Interop;
using System.Diagnostics;



public class SpeechExample
{
    // The _synthesizer variable is declared at the class level, allowing it to be accessed by any method within the class.
    private SpeechSynthesizer _synthesizer;


    // Constructor
    public SpeechExample()
    {
        _synthesizer = new SpeechSynthesizer();
    }

    // Public property, This property returns the _synthesizer instance, allowing external classes to access it.
    public SpeechSynthesizer Synthesizer
    {
        get { return _synthesizer; }
    }

    // Method to speak a given text
    public void Speak(string text)
    {
        if (_synthesizer != null)
        {
            _synthesizer.Speak(text);
        }
    }
}

namespace CK3_Reader
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    /// 

    public partial class MainWindow : Window
    {
        int rate = Properties.Settings.Default.Rate;
        int volume = Properties.Settings.Default.Volume;
        int PercentRate = 0;
        string TextRate = "";
        string TextVolume = "";
        string voice = "";
        bool launched = false;
        bool reading = true;
        bool clipAll = false;
        bool textShow = true;
        bool lineShow = true;

        private SpeechExample _speechExample;

        private Thread _udpThread;

        string documents = "";
        string logPath = "";
        string errorLog = "";
        string debugLog = "";
        string selectedLog = "";

        string line = "";
        string eventText = "";
        string[] formatting = [
            @"\bL\b",
                @"indent_newline:\d",
                @"TOOLTIP:SCALED_STATIC_MODIFIER,\w+,\d+\.\d+,\w+,\w+",
                @"TOOLTIP:\w+,\w+,\d+",
                @"ONCLICK:\w+,\d+",
                @"ONCLICK:\w+,\w+",
                @"TOOLTIP:\w+,\w+",
                @"TOOLTIP:\w+,\d+",
                @"positive_value",
                @"negative_value",
                @"COLOR_\w_\w",
                @"COLOR_\w",
                @"portrait_punishment_icon!",
                @"death_icon!",
                @"skill_\w+_icon!",
                @"_icon_\w+!",
                @"_icon!",
                @"skill_",
                @"\w+ ",
                @"\w;",
                @"\w+",
                @". ",
                @".",
                @"!",
                @"\w;",
                @";",
                @"stress_\w+",
                @"_",
                "   ",
                "  "
        ];

        private CancellationTokenSource _cancellationTokenSource;

        // Win32 API constants and imports for hotkeys
        const int WM_HOTKEY = 0x0312;

        [DllImport("user32.dll")]
        public static extern bool RegisterHotKey(IntPtr hWnd, int id, uint fsModifiers, uint vk);

        [DllImport("user32.dll")]
        public static extern bool UnregisterHotKey(IntPtr hWnd, int id);

        // Win32 API imports for getting the active window title
        [DllImport("user32.dll")]
        private static extern IntPtr GetForegroundWindow();

        [DllImport("user32.dll", SetLastError = true)]
        private static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);

        public static string GetActiveWindowTitle()
        {
            const int nChars = 256;
            StringBuilder Buff = new StringBuilder(nChars);
            IntPtr handle = GetForegroundWindow();

            if (GetWindowText(handle, Buff, nChars) > 0)
            {
                return Buff.ToString();
            }
            return null;
        }

        // Win32 API constants and imports
        private const int WM_CLIPBOARDUPDATE = 0x031D;

        [DllImport("user32.dll")]
        private static extern bool AddClipboardFormatListener(IntPtr hwnd);


        [DllImport("user32.dll")]
        private static extern bool RemoveClipboardFormatListener(IntPtr hwnd);

        private HwndSource _hwndSource;

        protected override void OnSourceInitialized(EventArgs e)
        {
            base.OnSourceInitialized(e);
            _hwndSource = PresentationSource.FromVisual(this) as HwndSource;
            _hwndSource.AddHook(WndProc);
            AddClipboardFormatListener(_hwndSource.Handle);
            var helper = new WindowInteropHelper(this);
            RegisterHotKey(helper.Handle, 1, 0x1, 0x43); //Alt C
            HwndSource source = HwndSource.FromHwnd(helper.Handle);
            source.AddHook(HwndHook);
        }

        private IntPtr HwndHook(IntPtr hwnd, int msg, IntPtr wParam, IntPtr lParam, ref bool handled)
        {
            if (msg == WM_HOTKEY)
            {
                int id = wParam.ToInt32();
                if (id == 1)
                {
                    // CTRL + G was pressed
                    StopSpeech();
                }
                handled = true;
            }
            return IntPtr.Zero;
        }

        protected override void OnClosed(EventArgs e)
        {
            RemoveClipboardFormatListener(_hwndSource.Handle);
            base.OnClosed(e);
            var helper = new WindowInteropHelper(this);
            UnregisterHotKey(helper.Handle, 1);
        }

        private IntPtr WndProc(IntPtr hwnd, int msg, IntPtr wParam, IntPtr lParam, ref bool handled)
        {
            if (msg == WM_CLIPBOARDUPDATE)
            {
                try
                {
                    string clipboardText = null;

                    // Retry in case another process holds the clipboard open
                    for (int attempt = 0; attempt < 10; attempt++)
                    {
                        try
                        {
                            if (Clipboard.ContainsText())
                                clipboardText = Clipboard.GetText();
                            break;
                        }
                        catch (System.Runtime.InteropServices.ExternalException)
                        {
                            Thread.Sleep(10);
                        }
                    }

                    if (clipboardText != null)
                    {
                        string activeTitle = GetActiveWindowTitle();
                        if (clipAll == true || (activeTitle == "Crusader Kings III" || activeTitle == "Europa Universalis V"))
                        {
                            StopSpeech();

                            eventText = clipboardText + "\n";

                            foreach (var format in formatting)
                            {
                                eventText = Regex.Replace(eventText, format, " ");
                            }

                            // Update the UI
                            Dispatcher.Invoke(() =>
                            {
                                txtLastLine.Text = "Copied from clipboard";
                                txtEvent.Text = eventText;
                            });
                            _speechExample.Synthesizer.SpeakAsync(eventText);
                        }
                    }
                }
                catch (Exception ex)
                {
                    Dispatcher.Invoke(() =>
                    {
                        txtLastLine.Text = "Clipboard error: " + ex.Message;
                    });
                }

                handled = true;
            }
            return IntPtr.Zero;
        }


        public MainWindow()
        {
            InitializeComponent();

            ShowTextCheckbox.IsChecked = true;
            ShowLastLineCheckbox.IsChecked = true;

            //DebugButton.IsChecked = (Properties.Settings.Default.log == "debug");
            //ErrorButton.IsChecked = (Properties.Settings.Default.log == "error");
            rate = Properties.Settings.Default.Rate;
            volume = Properties.Settings.Default.Volume;
            volumeSlider.Value = Properties.Settings.Default.Volume;
            speechSlider.Value = Properties.Settings.Default.Rate;

            DisplayRate();

            _speechExample = new SpeechExample();
            _speechExample.Synthesizer.Rate = rate;
            _speechExample.Synthesizer.Volume = volume;

            _speechExample.Synthesizer.SetOutputToDefaultAudioDevice();
            int voiceCounter = -1;

            foreach (InstalledVoice voice in _speechExample.Synthesizer.GetInstalledVoices())
            {
                VoiceBox.Items.Add(voice.VoiceInfo.Name);
                voiceCounter++;

                if (Properties.Settings.Default.Voice.Length > 0)
                {
                    if (Properties.Settings.Default.Voice == voice.VoiceInfo.Name)
                    {
                        VoiceBox.SelectedIndex = voiceCounter;
                    }
                }
                else
                {
                    VoiceBox.SelectedIndex = 0;
                }
            }

            if (Properties.Settings.Default.Voice.Length > 0 )
            {
                _speechExample.Synthesizer.SelectVoice(Properties.Settings.Default.Voice);
            }

            _speechExample.Synthesizer.SpeakAsync("Launching CK3 Reader");

            launched = true;

            string documents = System.Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            string logPath = documents + "\\Paradox Interactive\\Crusader Kings III\\logs\\";
            string errorLog = logPath + "error.log";
            string debugLog = logPath + "debug.log";
            //string selectedLog = logPath + Properties.Settings.Default.log + ".log";
            string selectedLog = debugLog;

            if (System.IO.File.Exists(selectedLog))
            {
                //lblStatus.Text = "✔️ ready, reading "+ Properties.Settings.Default.log + " log";
                lblStatus.Text = "✔️ ready, reading game log";
                lblStatus.Foreground = Brushes.Green;

            }

            Loaded += MainWindow_Loaded; // Subscribe to the Loaded event
        }

        private async void MainWindow_Loaded(object sender, RoutedEventArgs e)
        {
            _cancellationTokenSource = new CancellationTokenSource();
            await Task.Run(() => RunLoop(_cancellationTokenSource.Token));
        }

        private async void Restart_Click(object sender, RoutedEventArgs e)
        {
            _cancellationTokenSource = new CancellationTokenSource();
            StopLoop();
            await Task.Run(() => RunLoop(_cancellationTokenSource.Token));
        }

        private void RunLoop(CancellationToken token)
        {
            string counter = string.Empty;
            string documents = System.Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            string logPath = documents + "\\Paradox Interactive\\Crusader Kings III\\logs\\";
            string errorLog = logPath + "error.log";
            string debugLog = logPath + "debug.log";
            //string selectedLog = logPath + Properties.Settings.Default.log + ".log";
            string selectedLog = debugLog;
            string beginPattern = "<event-text>";
            string endPattern = "</event-text>";
            string stopReading = "<stop reading>";
            bool startMessage = false;


            // Update the variable

            using (FileStream stream = new FileStream(selectedLog, FileMode.Open, FileAccess.ReadWrite, FileShare.ReadWrite))
            {
                StreamReader reader = new StreamReader(stream);
                stream.Seek(0, SeekOrigin.End);

     
                //long fileLength = stream.Length;
                //if (fileLength == 0)
                //{
                //    Console.WriteLine("The file is empty.");
                //    return;
                //}

                //// Start from the end of the file
                //long fileLength = stream.Length;


                //// Buffer to hold the current line
                //StringBuilder currentLine = new StringBuilder();
                //int lineCount = 0;

                while (true)
                {
                    line = reader.ReadLine();

                    // this is just a little counter showing that it's reading lines, no other effect
                    counter += ".";
                    if (counter.Length > 6)
                    {
                        counter = string.Empty;
                    }


                    // Check if cancellation is requested
                    if (token.IsCancellationRequested)
                    {
                        break;
                    }

                    if (line != null)
                    {
                        if (line.Contains("<stop reading>"))
                        {
                            StopSpeech();
                        }
                        else
                        {
                            if (line.Contains(beginPattern))
                            {
                                eventText = string.Empty;
                                startMessage = true;   
                            }
                            if (startMessage == true)
                            {
                                eventText += "\n";
                                eventText += Regex.Replace(line, @".*\<event-text\>", "");

                                foreach (var format in formatting)
                                {
                                    eventText = Regex.Replace(eventText, format, " ");
                                }
                                if (line.EndsWith(endPattern) || eventText.Length > 10000)
                                {
                                    eventText = eventText.Replace(endPattern, "");
                                    startMessage = false;
                                    _speechExample.Synthesizer.SpeakAsync(eventText);
                                }
                            }
                        }

                        // Update the UI
                        Dispatcher.Invoke(() =>
                        {
                            txtLastLine.Text = "Last read line: " + line;
                            txtEvent.Text = eventText;
                        });
                    }

                    // If we haven't found an event, we wait 100ms (or 10)
                    // Without this it would be way too fast and eat up CPU
                    if (startMessage == false)
                    {
                        Thread.Sleep(Properties.Settings.Default.refresh);
                        //Thread.Sleep(10);
                    }

                    Dispatcher.Invoke(() =>
                    {
                        txtCounter.Text = counter;
                    });
                }
            }

            if (eventText.Length > 10000)
            {
                startMessage = false;
            }

            //string[] lines = System.IO.File.ReadAllLines(errorLog);
            //if (lines.Length > 0)
            //{
            //    line = lines[lines.Length-1];
            //}

        }

        // Optional: You can add a method to stop the loop if needed
        private void StopLoop()
        {
            _cancellationTokenSource?.Cancel();
        }

        public void DisplayRate()
        {
            PercentRate = rate * 10 + 100;
            TextRate = "Speech Rate: " + PercentRate.ToString() + "%";
            SpeechRate.Content = TextRate;
        }

        public void DisplayVolume()
        {
            TextVolume = "Speech Volume: " + volume.ToString() + "%";
            SpeechVolume.Content = TextVolume;
        }

        public void StopSpeech()
        {
            _speechExample.Synthesizer.SpeakAsyncCancelAll();
            reading = false;
            //_speechExample.Synthesizer.Dispose();
        }

        // controls

        

        public void StopButton_Click(object sender, RoutedEventArgs e)
        {
            StopSpeech();
            //StopButton.IsEnabled = (_speechExample.Synthesizer.State.ToString() == "Speaking");
        }

        public void Stop_Click(object sender, RoutedEventArgs e)
        {
            StopSpeech();
            StopLoop();
        }

        public void Button_Click(object sender, RoutedEventArgs e)
        {

            _speechExample.Synthesizer.SpeakAsyncCancelAll();

            _speechExample.Synthesizer.SpeakAsync("Crusader Kings 3");

        }

        public void Error_radio(object sender, RoutedEventArgs e)
        {
            Properties.Settings.Default.log = "error";
            Properties.Settings.Default.Save();
            ErrorButton.IsChecked = (Properties.Settings.Default.log == "error");
            lblStatus.Text = "✔️ ready, reading " + Properties.Settings.Default.log + " log";
            lblStatus.Foreground = Brushes.Green;
        }

        public void Debug_radio(object sender, RoutedEventArgs e)
        {
            Properties.Settings.Default.log = "debug";
            Properties.Settings.Default.Save();
            DebugButton.IsChecked = (Properties.Settings.Default.log == "debug");
            lblStatus.Text = "✔️ ready, reading " + Properties.Settings.Default.log + " log";
            lblStatus.Foreground = Brushes.Green;
        }

        private void speechSlider_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
        {
            rate = (int)e.NewValue;
            Properties.Settings.Default.Rate = rate;
            Properties.Settings.Default.Save();
            DisplayRate();
            if (_speechExample != null)
            {
                _speechExample.Synthesizer.Rate = rate;
                _speechExample.Synthesizer.SpeakAsyncCancelAll();
                _speechExample.Synthesizer.SpeakAsync(PercentRate.ToString());
            }
        }
        private void volumeSlider_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
        {
            volume = (int)e.NewValue;
            Properties.Settings.Default.Volume = volume;
            Properties.Settings.Default.Save();
            DisplayVolume();
            if (_speechExample != null)
            {
                _speechExample.Synthesizer.Volume = volume;
                _speechExample.Synthesizer.SpeakAsyncCancelAll();
                _speechExample.Synthesizer.SpeakAsync(volume.ToString());
            }
        }

        public void VoiceBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            voice = VoiceBox.SelectedItem.ToString();
            _speechExample.Synthesizer.SelectVoice(voice);
            if (launched == true)
            {
                _speechExample.Synthesizer.SpeakAsync(voice);
            }
            Properties.Settings.Default.Voice = voice;
            Properties.Settings.Default.Save();
            //synth.Dispose();

        }

        private void Window_closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            StopSpeech();
            _speechExample.Synthesizer.Dispose();
            StopLoop();
        }

        public void Clip_Checked(object sender, RoutedEventArgs e)
        {
            clipAll = false;
        }

        public void Clip_Unchecked(object sender, RoutedEventArgs e)
        {
            clipAll = true;
        }

        private void RefreshNormal_Checked(object sender, RoutedEventArgs e)
        {
            Properties.Settings.Default.refresh = 20;
            Properties.Settings.Default.Save();
            RefreshNormal.IsChecked = (Properties.Settings.Default.refresh == 20);
        }

        private void RefreshFast_Checked(object sender, RoutedEventArgs e)
        {
            Properties.Settings.Default.refresh = 5;
            Properties.Settings.Default.Save();
            RefreshFast.IsChecked = (Properties.Settings.Default.refresh == 5);
        }

        public void StopTalkingG(Object sender, ExecutedRoutedEventArgs e)
        {
            StopSpeech();
        }

        private void Line_Checked(object sender, RoutedEventArgs e)
        {
            txtLastLine.Visibility = Visibility.Visible;
        }
        private void Line_Unchecked(object sender, RoutedEventArgs e)
        {
            txtLastLine.Visibility = Visibility.Collapsed;
        }

        private void Text_Checked(object sender, RoutedEventArgs e)
        {
            RightScrollbox.Visibility = Visibility.Visible;
            this.Width = 830;
        }
        private void Text_Unchecked(object sender, RoutedEventArgs e)
        {
            RightScrollbox.Visibility = Visibility.Collapsed;
            this.Width = 360;
        }
    }
}