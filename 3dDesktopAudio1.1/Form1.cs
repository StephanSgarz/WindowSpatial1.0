using NAudio.CoreAudioApi.Interfaces;
using NAudio.CoreAudioApi;
using System.Windows.Forms;
using NAudio.SoundFont;
using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;
using static WindowBinauralizer._1.Form1;
using NAudio.Dmo;
using static System.Collections.Specialized.BitVector32;
using Microsoft.VisualBasic.Devices;
using Microsoft.VisualBasic.ApplicationServices;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using System.Reflection.Emit;
using System.Reflection.Metadata;
using System.Data;
using System.Security.Cryptography;
using System.Net;
using System.Runtime.CompilerServices;
using System.ComponentModel.Design;
using System.Threading;
using System.Text.RegularExpressions;
using System.IO;
using System.Reflection;


namespace WindowBinauralizer._1
{
    
    public partial class Form1 : Form
    {
        //calculates the filepath for the config file
        public string filePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "config.txt");

        public String reaperProjectLocation = "";
        public String reaperIniPath = "";
        public String reaperPath = "";
        public String reaperScriptPath = "";
        public String reaperProject1 = "";
        public String reaperProject2 = "";
        public String reaperProject3 = "";
        public String reaperProject4 = "";
        public String reaperProject5 = "";
        public String soundvolumeViewPath = "";



        //Screendata (this includes multiple screens as well)
        private int minX = Screen.AllScreens.Min(screen => screen.Bounds.Left);
        private int minY = Screen.AllScreens.Min(screen => screen.Bounds.Top);
        private int maxX = Screen.AllScreens.Max(screen => screen.Bounds.Right);
        private int maxY = Screen.AllScreens.Max(screen => screen.Bounds.Bottom);

        private System.ComponentModel.IContainer components;
        private CancellationTokenSource cts;
        private Task monitoringTask;

        //PInvokes
        //Signature for EnumWindows
        [DllImport("user32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);

        //Signature for GetWindowPlacement
        [DllImport("user32.dll", SetLastError = true)]
        static extern bool GetWindowPlacement(IntPtr hWnd, ref WINDOWPLACEMENT lpwndpl);

        //Signature for GetWindowText
        [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
        static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);

        //EnumWindows callback
        delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

        //Window structure definition
        [Serializable]
        [StructLayout(LayoutKind.Sequential)]
        struct WINDOWPLACEMENT
        {
            public int length;
            public int flags;
            public int showCmd;
            public System.Drawing.Point ptMinPosition;
            public System.Drawing.Point ptMaxPosition;
            public System.Drawing.Rectangle rcNormalPosition;
        }

        private const int SW_SHOWMINIMIZED = 2;

        //Initialization of the Form (Provides Data and information)
        public Form1()
        {
            InitializeComponent();
            Load += Form1_Load;
            this.FormClosing += new FormClosingEventHandler(newFormClosing);
        }
        private void Form1_Load(object sender, EventArgs e)
        {
            cts = new CancellationTokenSource();
            calculateLocations(filePath);
            monitoringTask = MonitorAudioSessionsAsync(cts.Token);
        }

        //Shutdown Action
        private void button1_Click(object sender, EventArgs e)
        {
            Close();
        }
        private ListBox listBox1;
        private System.Windows.Forms.Button button1;

        //gives the component its structure
        private void InitializeComponent()
        {
            listBox1 = new ListBox();
            button1 = new System.Windows.Forms.Button();
            SuspendLayout();
            // 
            // listBox1
            // 
            listBox1.AllowDrop = true;
            listBox1.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            listBox1.Font = new Font("Segoe UI", 6.75F, FontStyle.Regular, GraphicsUnit.Point, 0);
            listBox1.FormattingEnabled = true;
            listBox1.ItemHeight = 12;
            listBox1.Location = new Point(12, 7);
            listBox1.Name = "listBox1";
            listBox1.Size = new Size(296, 160);
            listBox1.TabIndex = 0;
            // 
            // button1
            // 
            button1.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
            button1.Location = new Point(233, 179);
            button1.Name = "button1";
            button1.Size = new Size(75, 29);
            button1.TabIndex = 1;
            button1.Text = "Close";
            button1.UseVisualStyleBackColor = true;
            button1.Click += button1_Click;
            // 
            // Form1
            // 
            ClientSize = new Size(318, 220);
            Controls.Add(button1);
            Controls.Add(listBox1);
            Name = "Form1";
            ResumeLayout(false);
        }

        //closes the programm and shuts down the reaper processes.
        //also routs all audiotabs back to the default audio
        private async void newFormClosing(object sender, FormClosingEventArgs e)
        {
            cts.Cancel();
            Thread.Sleep(200);

            var deviceEnumerator = new MMDeviceEnumerator();
            var defaultEndpoint = deviceEnumerator.GetDefaultAudioEndpoint(DataFlow.Render, Role.Multimedia);
            var defaultName = defaultEndpoint.DeviceFriendlyName;
            var endpoints = deviceEnumerator.EnumerateAudioEndPoints(DataFlow.Render, DeviceState.Active);
            foreach (var endpoint in endpoints)
            {
                if (endpoint.FriendlyName.Contains("CABLE"))
                {
                    var sessions = endpoint.AudioSessionManager.Sessions;
                    for (int i = 0; i < sessions.Count; i++)
                    {
                        var session = sessions[i];
                        if (session.State == AudioSessionState.AudioSessionStateActive)
                        {
                            var process = Process.GetProcessById((int)session.GetProcessID);
                            reroutAudioSource(process.Id.ToString(), defaultName, soundvolumeViewPath);
                        }
                    }
                }
            }


            var defaultSessions = defaultEndpoint.AudioSessionManager.Sessions;
            for (int i = 0; i < defaultSessions.Count; i++)
            {
                var session = defaultSessions[i];
                var process = Process.GetProcessById((int)session.GetProcessID);
                if (process.ProcessName == "reaper")
                {
                    process.Kill();
                }


            }
        }

        //opens the reaper tabs and calls the main loop on startup.
        private async Task MonitorAudioSessionsAsync(CancellationToken cancellationToken)
        {

            //load location data out of config file
            //await calculateLocations(filePath);

            //Activate all 5 Reaper streams that are already set to listen to different Cables and update their data via a single script repetetively 

            reaperIniInstance("CABLE Output (VB-Audio Virtual", reaperProjectLocation + reaperProject1);
            await loadReaper();
            await Task.Delay(5000);
            reaperIniInstance("CABLE-A Output (VB-Audio Cable", reaperProjectLocation + reaperProject2);
            await loadReaper();
            await Task.Delay(4000);
            reaperIniInstance("CABLE-B Output (VB-Audio Cable", reaperProjectLocation + reaperProject3);
            await loadReaper();
            await Task.Delay(4000);
            reaperIniInstance("CABLE-C Output (VB-Audio Cable", reaperProjectLocation + reaperProject4);
            await loadReaper();
            await Task.Delay(4000);
            reaperIniInstance("CABLE-D Output (VB-Audio Cable", reaperProjectLocation + reaperProject5);
            await loadReaper();
            await Task.Delay(4000);
            ListenerLoop(cancellationToken);
        }

        //Loops continuously until canceled, and changes the displayed information
        //Runs a seperate Task to calculate active audio tabs
        private void ListenerLoop (CancellationToken cancellationToken) 
        {
            Task.Run(() =>
            {
                //Loops the main function and displays information about Audiotabs
                while (!cancellationToken.IsCancellationRequested)
                {
                    //main function
                    List<string> currentSessions = GetActiveAudioSessions();

                    BeginInvoke(new Action(() =>
                    {

                        listBox1.Items.Clear();
                        if (currentSessions != null)
                        {
                            foreach (var session in currentSessions)
                            {
                                listBox1.Items.Add(session);
                            }
                        }
                    }));
                    //Thread.Sleep(100); //could maybe be uncapped
                }
            }, cancellationToken);
        }

        //Checks for active audiosessions and delegates them to a free Cable if any are free.
        //also calls the change-script for Reaper to adapt the position of the tab.
        private List<string> GetActiveAudioSessions()
        {


            List<string> activeSessions = new List<string>();
            Queue<string> freeCables = new Queue<string>();
            var deviceEnumerator = new MMDeviceEnumerator();
            var defaultEndpoint = deviceEnumerator.GetDefaultAudioEndpoint(DataFlow.Render, Role.Multimedia);
            var defaultName = defaultEndpoint.DeviceFriendlyName;

            // Shows which cabels have currently running active audios and returns all additional data
            // Cables that dont have any tasks will be saved in the freeCbales list.
            var endpoints = deviceEnumerator.EnumerateAudioEndPoints(DataFlow.Render, DeviceState.Active);
            foreach (var endpoint in endpoints)
            {
                //Only Check the Virtual Cables
                if (endpoint.FriendlyName.Contains("CABLE"))
                {
                    var sessions = endpoint.AudioSessionManager.Sessions;

                    int count = 0; //checks whether this cable is in use or not

                    for (int i = 0; i < sessions.Count; i++)

                    {
                        var session = sessions[i];

                        //if this session is an active audiosession
                        if (session.State == AudioSessionState.AudioSessionStateActive)
                        {

                            activeSessions.Add(endpoint.FriendlyName + ":");
                            var process = Process.GetProcessById((int)session.GetProcessID);
                            //if the programm is reaper or already in use rerout it back to standard audio output
                            if (process.ProcessName == "reaper" || count == 1)
                            {

                                reroutAudioSource(process.Id.ToString(), defaultName, soundvolumeViewPath);
                            }
                            //else calculate position of the tab and change the script for reaper calls
                            //also display information
                            else
                            {
                                //function to gather information about the tab-position and size
                                var displayName = EnumerateWindows((int)session.GetProcessID);
                                activeSessions.Add(displayName);

                                //translate the tabdata and position to a binaural position
                                CalculateBinauralChange(displayName, endpoint);
                                activeSessions.Add("");
                                count++;        //signal that this cable is in use
                            }
                        }
                    }
                    if (count == 0) freeCables.Enqueue(endpoint.FriendlyName);

                }
            }


            // Checks the defaultaudio ouput for any incoming audios and then switches them over to a free cable, which was calculated above
            var defaultDevice = deviceEnumerator.GetDefaultAudioEndpoint(DataFlow.Render, Role.Multimedia);
            var defaultSessions = defaultDevice.AudioSessionManager.Sessions;

            for (int i = 0; i < defaultSessions.Count; i++)
            {
                var session = defaultSessions[i];
                if (session.State == AudioSessionState.AudioSessionStateActive)
                {
                    var process = Process.GetProcessById((int)session.GetProcessID);
                    string pid = process.Id.ToString();
                    //If the process is a reaper-process then ignore it (here one can add additional programms that should be ignored)
                    if (freeCables.Any() && (process.ProcessName != "reaper") && (process.ProcessName != "zoom") && (process.ProcessName != "Zoom")
                        && (process.ProcessName != "obs64"))
                    {
                        string cable = freeCables.Dequeue();
                        cable = cable.Split('(')[0];
                        cable = cable.TrimEnd(' ');
                        reroutAudioSource(process.Id.ToString(), cable, soundvolumeViewPath);

                    }
                }
            }
            return activeSessions;
        }

        //calculate the position of the tab relative to the screen and then transform into aszimuth and elevation
        //change the binaural-script afterwards
        void CalculateBinauralChange(String displayName, MMDevice endpoint)
        {
            double x = 0;
            double y = 0;
            double width = 0;
            double height = 0;

            string pattern = @"X=([-]?\d+),Y=([-]?\d+),Width=([-]?\d+),Height=([-]?\d+)";
            Regex regex = new Regex(pattern);
            Match match = regex.Match(displayName);
            if (match.Success)
            {
                x = double.Parse(match.Groups[1].Value);
                y = double.Parse(match.Groups[2].Value);
                width = double.Parse(match.Groups[3].Value);
                height = double.Parse(match.Groups[4].Value);
            }

            //calculate the total width and height of the screen
            double totalWidth = maxX - minX;
            double totalHeight = maxY - minY;

            if (totalWidth < 0) totalWidth = 1;
            if (totalHeight < 0) totalHeight = 1;

            //shows the center of the window
            double newX = x-minX;
            double newY = y-minY;
            double newWidth = width-minX;
            double newHeight = height-minY;

            double centeredX = newX;
            if ((newWidth - newX) != 0) centeredX = (newX + (newWidth-newX)/2);
            double centeredY = newY;
            if ((newHeight - newY) != 0) centeredY = (newY + (newHeight - newY) / 2);

           

            //gives the relative positions of the window in regards to the screen
            double relativeX = centeredX / totalWidth;
            double relativeY = centeredY / totalHeight;

            //Transforms the tabdata to azimuth and elevation
            String azi = (1 - relativeX).ToString();
            if (azi.ToString().Length > 3) azi = (1 - relativeX).ToString().Substring(0, 3);
            azi = azi.Replace(",", ".");


            relativeY = ((1 - relativeY) * 0.5);
            if(relativeY < 0) relativeY = 0;

            //if process is minimized place tab behind the head
            if (!displayName.Contains("minimized"))
            {
                relativeY = relativeY + 0.5;
            }
            else
            {
                relativeY = 0.5 - relativeY ;
            }
            String elev = relativeY.ToString();
            if (elev.Length > 3) elev = relativeY.ToString().Substring(0, 3);
            elev = elev.Replace(",", ".");


            //call the correct reaper programm
            string reaperChannel = "";
            if (endpoint.FriendlyName.Contains("CABLE-A")) reaperChannel = reaperProject2;
            else
            if (endpoint.FriendlyName.Contains("CABLE-B")) reaperChannel = reaperProject3;
            else
            if (endpoint.FriendlyName.Contains("CABLE-C")) reaperChannel = reaperProject4;
            else
            if (endpoint.FriendlyName.Contains("CABLE-D")) reaperChannel = reaperProject5;
            else
                reaperChannel = reaperProject1;

            //change the script
            ChangeBinauralPosition(azi, elev, reaperChannel);
        }

        //with the programm id, search all open tabs for a match and then collect the tab data from it
        //returns the name of the tab, the size and the position
        static string EnumerateWindows(int pid)
        {
            string str = "";
            //enumerates through all windows
            EnumWindows((hWnd, lParam) =>
            {
                WINDOWPLACEMENT placement = new WINDOWPLACEMENT();
                placement.length = Marshal.SizeOf(typeof(WINDOWPLACEMENT));

                // Get window placement
                if (GetWindowPlacement(hWnd, ref placement))
                {

                    // Get title
                    StringBuilder titleBuilder = new StringBuilder(256);
                    GetWindowText(hWnd, titleBuilder, titleBuilder.Capacity);
                    string title = titleBuilder.ToString();

                    // get window placement information
                    if (titleMatches(pid, title))
                    {
                        str += $"{title} ";
                        str += $"{placement.rcNormalPosition}";
                        if(placement.showCmd == SW_SHOWMINIMIZED) str += $" - minimized";
                        return false; //cancel enumeration
                    }
                }
                return true; // Continue enumeration
            }, IntPtr.Zero);
            return str;
        }

        //checks if the process ID belongs to the same process as the title given by a window
        static bool titleMatches(int pid, string title)
        {
            Process targetProcess = Process.GetProcessById(pid);

            // find all processes with the same name
            Process[] processesWithSameName = Process.GetProcessesByName(targetProcess.ProcessName);

            // returns true if the name is not empty and matches with the given window name
            foreach (var process in processesWithSameName)
            {
                if (process.MainWindowTitle.Equals(title) && title != "")
                {
                    return true;
                }
            }
            return false;
        }

        //prepares the ini reaper file for the correct channel and project 
        private async Task reaperIniInstance(String channel, String project)
        {
            //Write BinauralScript

            if (!File.Exists(reaperIniPath))
            {
                Console.WriteLine("REAPER.ini does not exist.");
                return;
            }

            string[] lines = File.ReadAllLines(reaperIniPath);
            bool inAudioConfigSection = false;
            bool projectSet = false;

            for (int i = 0; i < lines.Length; i++)
            {
                string line = lines[i].Trim();

                if (!projectSet)
                {

                    if (line.StartsWith("lastproject=", StringComparison.OrdinalIgnoreCase))
                    {
                        lines[i] = $"lastproject={project}"; //Set a New Project

                    }
                    if (line.StartsWith("projecttab1=", StringComparison.OrdinalIgnoreCase))
                    {
                        lines[i] = $"projecttab1={project}"; //Set a New Project
                        projectSet = true;
                    }
                }

                // Detect the [audioconfig] section
                if (line.Equals("[audioconfig]", StringComparison.OrdinalIgnoreCase))
                {
                    inAudioConfigSection = true;
                    continue;
                }

                // Exit the section when another section starts
                if (inAudioConfigSection && line.StartsWith("[") && line.EndsWith("]"))
                {
                    break;
                }

                if (inAudioConfigSection)
                {
                    if (line.StartsWith("waveout_driver_in=", StringComparison.OrdinalIgnoreCase))
                    {
                        lines[i] = $"waveout_driver_in=\"{channel}\""; // Set the input device index
                    }
                }
            }

            // Write the changes back to the file
            File.WriteAllLines(reaperIniPath, lines);

            return;
        }

        //starts reaper minimized
        private async Task loadReaper()
        {
            ProcessStartInfo startInfo = new ProcessStartInfo();
            startInfo.FileName = reaperPath;
            startInfo.WindowStyle = ProcessWindowStyle.Minimized;
            Process reaper = Process.Start(startInfo);
            await Task.Delay(1000);
            return;
        }


        //changes the reaper binaural script to the given azi, elev and channel
        private async void ChangeBinauralPosition(String azimuth, String elevation, String reaperChannel)
        {

            string lines = "-- Get the master track\r\n" +
                "local targetProjectName = \"" + reaperChannel + "\"\r\n" +
                "local currentProjectName = reaper.GetProjectName(0, \"\")\r\n" +
                "if currentProjectName == targetProjectName then\r\n" +
                "  local track = reaper.GetMasterTrack(0)\r\n" +
                "\r\n" +
                "  -- Find the SPARTA Binauraliser FX (assuming it's the first FX, index 0)\r\n" +
                "  local fx_index = 0\r\n\r\n" +
                "  -- Parameter indices for azimuth and elevation (these might vary; adjust if needed)\r\n" +
                "  local azimuth_param_index = 9 -- Adjust if azimuth is on a different parameter index\r\n" +
                "  local elevation_param_index = 10 -- Adjust if elevation is on a different parameter index\r\n" +
                "\r\n" +
                "  -- New values for azimuth and elevation\r\n  " +
                "local azimuth_value = " + azimuth + " -- Range is typically 0.0 to 1.0\r\n" +
                "  local elevation_value = " + elevation + " -- Range is typically 0.0 to 1.0\r\n" +
                "\r\n" +
                "  -- Set the azimuth parameter\r\n" +
                "  reaper.TrackFX_SetParam(track, fx_index, azimuth_param_index, azimuth_value)\r\n" +
                "\r\n" +
                "  -- Set the elevation parameter\r\n" +
                "  reaper.TrackFX_SetParam(track, fx_index, elevation_param_index, elevation_value)\r\n" +
                "\r\n  reaper.UpdateArrange() -- Update the arrangement view to reflect changes\r\n" +
                "  end\r\n";
            File.WriteAllText(reaperScriptPath, lines);
            await Task.Delay(100);
        }



        //rerouts an audiotab to a specific cable via cmd commands
        private static void reroutAudioSource(string pid, string cable, string path)
        {
            string soundvolumeViewPath = path;

            using (Process process = new Process())
            {
                process.StartInfo.FileName = "cmd.exe";
                process.StartInfo.RedirectStandardOutput = true;
                process.StartInfo.RedirectStandardError = true;
                process.StartInfo.UseShellExecute = false;
                process.StartInfo.CreateNoWindow = true;
                process.StartInfo.Arguments = "/C " + soundvolumeViewPath + " /SetAppDefault \"" + cable + "\" all " + pid;        // Start the process.
                process.Start();        // Wait for the process to finish.
                process.WaitForExit();
            }
        }

        //reads the config file for all locations and overrides the variables
        private void calculateLocations(string filePath)
        {

            if (!File.Exists(filePath))
            {
                return;
            }

            foreach (string line in File.ReadLines(filePath))
            {
                
                if (line.Contains("="))
                {
                    string[] parts = line.Split(new[] { '=' }, 2);
                    if (parts.Length == 2)
                    {
                        string variable = parts[0].Trim();
                        string value = parts[1].Trim();
                        Type type = typeof(Form1);

                        FieldInfo fieldInfo = type.GetField(variable, BindingFlags.Public | BindingFlags.Instance);

                        if (fieldInfo != null)
                        {
                            // Set the new value
                            fieldInfo.SetValue(this, value);
                        }
                        else
                        {
                            return;
                        }
                    }
                }
            }
            return;
        }
    }
}

