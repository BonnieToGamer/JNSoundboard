﻿using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Windows.Forms;
using System.Linq;
using System.Threading;
using System.Media;
using NAudio.Wave;
using System.Diagnostics;

namespace JNSoundboard
{
    public partial class MainForm : Form
    {
        internal class ListViewItemComparer : IComparer
        {
            private readonly int col;

            public ListViewItemComparer()
            {
                col = 0;
            }

            public ListViewItemComparer(int column)
            {
                col = column;
            }

            public int Compare(object x, object y)
            {
                return string.Compare(((ListViewItem)x).SubItems[col].Text, ((ListViewItem)y).SubItems[col].Text);
            }
        }

        //There might be a smarter way to output the sound to two devices, but this is quick and it works.

        //Generally the virtual cable output
        readonly AudioPlaybackEngine playbackEngine1 = new AudioPlaybackEngine();

        //A second output to also output the sound to your headphones or speaker.
        readonly AudioPlaybackEngine playbackEngine2 = new AudioPlaybackEngine();

        //Linear volume for sounds sent to AudioPlaybackEngine (doesn't affect microphone loopback volume)
        private float soundVolume;
        private int SelectedPlaybackDevice1
        {
            get
            {
                if (cbPlaybackDevices1.Items.Count > 0)
                {
                    return cbPlaybackDevices1.SelectedIndex;
                }
                else
                {
                    return -1;
                }
            }
            set
            {
                if (value >= 0 && value <= cbPlaybackDevices1.Items.Count - 1)
                {
                    cbPlaybackDevices1.SelectedIndex = value;
                }
            }
        }

        private int SelectedPlaybackDevice2
        {
            get
            {
                if (cbPlaybackDevices2.Items.Count > 0)
                {
                    //minus one to account for null entry
                    return cbPlaybackDevices2.SelectedIndex - 1;
                }
                else
                {
                    return -1;
                }
            }
            set
            {
                if (value >= -1 && value <= cbPlaybackDevices2.Items.Count - 2)
                {
                    cbPlaybackDevices2.SelectedIndex = value + 1;
                }
            }
        }

        WaveIn loopbackSourceStream = null;
        BufferedWaveProvider loopbackWaveProvider = null;
        WaveOut loopbackWaveOut = null;
        readonly Random rand = new Random();

        const string DO_NOT_USE = "[Do not use]";

        bool keyUpPushToTalkKey = false;

        internal List<XMLSettings.SoundHotkey> soundHotkeys = new List<XMLSettings.SoundHotkey>();

        Keyboard.Keys pushToTalkKey;

        string xmlLocation;

        const string PLAYBACK1_TOOLTIP =
@"The playback device through which to play the audio.
Generally the virtual audio cable.
(Or your speakers or headphones for local playback only.)";

        const string PLAYBACK2_TOOLTIP =
@"A second playback device through which to play the audio. 
Generally your headset, so that you can hear the sounds being played, too.";

        const string LOOPBACK_TOOLTIP =
@"The input device which will also be sent into the above playback device.
Generally your real microphone to speak through.
DO NOT choose the virtual auidio cable!";

        const string SOUND_VOLUME_TOOLTIP =
@"The volume of soundboard audio files.
Doesn't affect sounds with custom volumes or that are currently playing.";

        private bool allowVisible = true;
        public MainForm()
        {
            InitializeComponent();

            ToolTip tooltip = new ToolTip();

            tooltip.SetToolTip(btnReloadDevices, "Refresh sound devices");
            tooltip.SetToolTip(btnReloadWindows, "Reload windows");
            tooltip.SetToolTip(cbPlaybackDevices1, PLAYBACK1_TOOLTIP);
            tooltip.SetToolTip(cbPlaybackDevices2, PLAYBACK2_TOOLTIP);
            tooltip.SetToolTip(cbLoopbackDevices, LOOPBACK_TOOLTIP);
            tooltip.SetToolTip(lblPlayback1, PLAYBACK1_TOOLTIP);
            tooltip.SetToolTip(lblPlayback2, PLAYBACK2_TOOLTIP);
            tooltip.SetToolTip(lblLoopback, LOOPBACK_TOOLTIP);
            tooltip.SetToolTip(vsSoundVolume, SOUND_VOLUME_TOOLTIP);
            tooltip.SetToolTip(nSoundVolume, SOUND_VOLUME_TOOLTIP);

            XMLSettings.LoadSoundboardSettingsXML();

            //Disable change events for elements that would trigger settings changes and unnecessarily write to settings.xml
            DisableCheckboxChangeEvents();
            DisableSoundVolumeChangeEvents();

            LoadSoundDevices(false); //false argument keeps device change events disabled

            Helper.GetWindows(cbWindows);
            Helper.SelectWindow(cbWindows, XMLSettings.soundboardSettings.AutoPushToTalkWindow);

            if (XMLSettings.soundboardSettings.StartMinimised)
            {
                this.WindowState = FormWindowState.Minimized;

                if (XMLSettings.soundboardSettings.MinimiseToTray)
                {
                    this.HideFormToTray();
                }
            }

            Helper.SetStartup(XMLSettings.soundboardSettings.StartWithWindows);

            cbEnableHotkeys.Checked = XMLSettings.soundboardSettings.EnableHotkeys;
            cbEnableLoopback.Checked = XMLSettings.soundboardSettings.EnableLoopback;

            soundVolume = XMLSettings.soundboardSettings.SoundVolume;
            vsSoundVolume.Volume = soundVolume;
            nSoundVolume.Value = Helper.LinearVolumeToInteger(vsSoundVolume.Volume); //needed because change events are still disabled

            pushToTalkKey = XMLSettings.soundboardSettings.AutoPushToTalkKey;

            tbPushToTalkKey.Text = pushToTalkKey.ToString() == "None" ? "" : pushToTalkKey.ToString();

            cbEnablePushToTalk.Checked = XMLSettings.soundboardSettings.EnableAutoPushToTalk;
            tbPushToTalkKey.Enabled = !cbEnablePushToTalk.Checked;
            clearHotkey.Enabled = !cbEnablePushToTalk.Checked;

            if (File.Exists(XMLSettings.soundboardSettings.LastXMLFile))
            {
                //loadXMLFile() returns true if error occurred
                if (LoadXMLFile(XMLSettings.soundboardSettings.LastXMLFile))
                {
                    XMLSettings.soundboardSettings.LastXMLFile = "";
                    XMLSettings.SaveSoundboardSettingsXML();
                }
            }

            //Add events after settings have been loaded
            EnableCheckboxChangeEvents();
            EnableSoundVolumeChangeEvents();
            EnableDeviceChangeEvents();

            mainTimer.Enabled = cbEnableHotkeys.Checked;
            InitAudioPlaybackEngine1();
            InitAudioPlaybackEngine2();
            RestartLoopback();

            //When sound stops, fire event which lets go of push-to-talk key.
            playbackEngine1.AllInputEnded += OnAllInputEnded;
            //Don't need to stop holding the push-to-talk key when engine2 stops playing, that's just our in-ear echo.
        }

        protected override void SetVisibleCore(bool value)
        {
            base.SetVisibleCore(allowVisible);
        }

        private void DisableCheckboxChangeEvents()
        {
            cbEnableHotkeys.CheckedChanged -= CbEnableHotkeys_CheckedChanged;
            cbEnableLoopback.CheckedChanged -= CbEnableLoopback_CheckedChanged;
            cbEnablePushToTalk.CheckedChanged -= CbEnablePushToTalk_CheckedChanged;
        }

        private void EnableCheckboxChangeEvents()
        {
            cbEnableHotkeys.CheckedChanged += CbEnableHotkeys_CheckedChanged;
            cbEnableLoopback.CheckedChanged += CbEnableLoopback_CheckedChanged;
            cbEnablePushToTalk.CheckedChanged += CbEnablePushToTalk_CheckedChanged;
        }

        private void DisableSoundVolumeChangeEvents()
        {
            vsSoundVolume.VolumeChanged -= VsSoundVolume_VolumeChanged;
            nSoundVolume.ValueChanged -= NSoundVolume_ValueChanged;
        }

        private void EnableSoundVolumeChangeEvents()
        {
            vsSoundVolume.VolumeChanged += VsSoundVolume_VolumeChanged;
            nSoundVolume.ValueChanged += NSoundVolume_ValueChanged;
        }

        private void DisableDeviceChangeEvents()
        {
            cbPlaybackDevices1.SelectedIndexChanged -= CbPlaybackDevices1_SelectedIndexChanged;
            cbPlaybackDevices2.SelectedIndexChanged -= CbPlaybackDevices2_SelectedIndexChanged;
            cbLoopbackDevices.SelectedIndexChanged -= CbLoopbackDevices_SelectedIndexChanged;
        }

        private void EnableDeviceChangeEvents()
        {
            cbPlaybackDevices1.SelectedIndexChanged += CbPlaybackDevices1_SelectedIndexChanged;
            cbPlaybackDevices2.SelectedIndexChanged += CbPlaybackDevices2_SelectedIndexChanged;
            cbLoopbackDevices.SelectedIndexChanged += CbLoopbackDevices_SelectedIndexChanged;
        }

        private void OnAllInputEnded(object sender, EventArgs e)
        {
            if (keyUpPushToTalkKey)
            {
                keyUpPushToTalkKey = false;
                Keyboard.SendKey(pushToTalkKey, false);
            }
        }

        private void InitAudioPlaybackEngine1()
        {
            try
            {
                playbackEngine1.Init(SelectedPlaybackDevice1);
            }
            catch (NAudio.MmException ex)
            {
                SystemSounds.Beep.Play();

                string msg = ex.ToString();

                if (msg.Contains("AlreadyAllocated calling waveOutOpen"))
                {
                    msg = "Failed to open device. Already in exclusive use by another application? \n\n" + msg;
                }

                MessageBox.Show("Playback 1" + msg);
            }
        }

        private void InitAudioPlaybackEngine2()
        {
            //Don't init if the null entry is selected
            if (SelectedPlaybackDevice2 >= 0)
            {
                try
                {
                    playbackEngine2.Init(SelectedPlaybackDevice2);
                }
                catch (NAudio.MmException ex)
                {
                    SystemSounds.Beep.Play();

                    string msg = ex.ToString();

                    if (msg.Contains("AlreadyAllocated calling waveOutOpen"))
                    {
                        msg = "Failed to open device. Already in exclusive use by another application? \n\n" + msg;
                    }

                    MessageBox.Show("Playback 2" + msg);
                }
            }
        }

        private void LoadSoundDevices(bool enableEvents = true)
        {
            DisableDeviceChangeEvents(); //avoid audio related errors

            List<WaveOutCapabilities> playbackSources = new List<WaveOutCapabilities>();
            List<WaveInCapabilities> loopbackSources = new List<WaveInCapabilities>();

            for (int i = 0; i < WaveOut.DeviceCount; i++)
            {
                playbackSources.Add(WaveOut.GetCapabilities(i));
            }

            for (int i = 0; i < WaveIn.DeviceCount; i++)
            {
                loopbackSources.Add(WaveIn.GetCapabilities(i));
            }

            cbPlaybackDevices1.Items.Clear();
            cbPlaybackDevices2.Items.Clear();
            cbLoopbackDevices.Items.Clear();

            foreach (WaveOutCapabilities source in playbackSources)
            {
                cbPlaybackDevices1.Items.Add(source.ProductName);
                cbPlaybackDevices2.Items.Add(source.ProductName);
            }

            cbLoopbackDevices.Items.Insert(0, DO_NOT_USE);
            cbPlaybackDevices2.Items.Insert(0, DO_NOT_USE);

            SelectedPlaybackDevice1 = -1;
            SelectedPlaybackDevice2 = -1;

            if (cbPlaybackDevices1.Items.Count > 0)
            {
                SelectedPlaybackDevice1 = 0;
            }

            if (cbPlaybackDevices2.Items.Count > 0)
            {
                SelectedPlaybackDevice2 = -1;
            }

            foreach (WaveInCapabilities source in loopbackSources)
            {
                cbLoopbackDevices.Items.Add(source.ProductName);
            }

            cbLoopbackDevices.SelectedIndex = 0;

            if (enableEvents)
            {
                EnableDeviceChangeEvents();
            }

            if (cbPlaybackDevices1.Items.Contains(XMLSettings.soundboardSettings.LastPlaybackDevice)) cbPlaybackDevices1.SelectedItem = XMLSettings.soundboardSettings.LastPlaybackDevice;
            if (cbPlaybackDevices2.Items.Contains(XMLSettings.soundboardSettings.LastPlaybackDevice2)) cbPlaybackDevices2.SelectedItem = XMLSettings.soundboardSettings.LastPlaybackDevice2;
            if (cbLoopbackDevices.Items.Contains(XMLSettings.soundboardSettings.LastLoopbackDevice)) cbLoopbackDevices.SelectedItem = XMLSettings.soundboardSettings.LastLoopbackDevice;
        }

        private void RestartLoopback()
        {
            StopLoopback();

            //Subtract one from index to account for null entry.
            int deviceNumber = cbLoopbackDevices.SelectedIndex - 1;

            if (deviceNumber >= 0 && cbEnableLoopback.Checked)
            {
                if (loopbackSourceStream == null)
                    loopbackSourceStream = new WaveIn();
                loopbackSourceStream.DeviceNumber = deviceNumber;
                loopbackSourceStream.WaveFormat = new WaveFormat(44100, WaveIn.GetCapabilities(deviceNumber).Channels);
                loopbackSourceStream.BufferMilliseconds = 25;
                loopbackSourceStream.NumberOfBuffers = 5;
                loopbackSourceStream.DataAvailable += LoopbackSourceStream_DataAvailable;

                loopbackWaveProvider = new BufferedWaveProvider(loopbackSourceStream.WaveFormat)
                {
                    DiscardOnBufferOverflow = true
                };

                if (loopbackWaveOut == null)
                    loopbackWaveOut = new WaveOut();
                loopbackWaveOut.DeviceNumber = cbPlaybackDevices1.SelectedIndex;
                loopbackWaveOut.DesiredLatency = 125;
                loopbackWaveOut.Init(loopbackWaveProvider);

                loopbackSourceStream.StartRecording();
                loopbackWaveOut.Play();
            }
        }

        private void StopLoopback()
        {
            try
            {
                if (loopbackWaveOut != null)
                {
                    loopbackWaveOut.Stop();
                    loopbackWaveOut.Dispose();
                    loopbackWaveOut = null;
                }

                if (loopbackWaveProvider != null)
                {
                    loopbackWaveProvider.ClearBuffer();
                    loopbackWaveProvider = null;
                }

                if (loopbackSourceStream != null)
                {
                    loopbackSourceStream.StopRecording();
                    loopbackSourceStream.Dispose();
                    loopbackSourceStream = null;
                }
            }
            catch (Exception) { }
        }

        private void StopPlayback()
        {
            playbackEngine1.StopAllSounds();
            playbackEngine2.StopAllSounds();
        }

        private void PlaySound(string file, float soundVolume)
        {
            if (!XMLSettings.soundboardSettings.OverlapAudio)
                StopPlayback();

            try
            {
                playbackEngine1.PlaySound(file, soundVolume);

                //Don't try to play the sound if the device is not selected or the device is the same as #1.
                if (SelectedPlaybackDevice2 >= 0 && SelectedPlaybackDevice2 != SelectedPlaybackDevice1)
                {
                    playbackEngine2.PlaySound(file, soundVolume);
                }
            }
            catch (FormatException ex)
            {
                SystemSounds.Beep.Play();
                MessageBox.Show(ex.ToString());
            }
            catch (System.Runtime.InteropServices.COMException ex)
            {
                SystemSounds.Beep.Play();
                MessageBox.Show(ex.ToString());
            }
            catch (NAudio.MmException ex)
            {
                SystemSounds.Beep.Play();
                string msg = ex.ToString();
                MessageBox.Show((msg.Contains("UnspecifiedError calling waveOutOpen") ? "Something is wrong with either the sound you tried to play (" + file.Substring(file.LastIndexOf("\\") + 1) + ") (try converting it to another format) or your sound card driver\n\n" + msg : msg));
            }
        }

        private bool LoadXMLFile(string path)
        {
            bool errorOccurred = true;

            try
            {
                XMLSettings.Settings s = (XMLSettings.Settings)XMLSettings.ReadXML(typeof(XMLSettings.Settings), path);
                
                if (s != null && s.SoundHotkeys != null && s.SoundHotkeys.Length > 0)
                {
                    List<ListViewItem> items = new List<ListViewItem>();
                    string errorMessage = "";
                    string sameKeys = "";

                    for (int i = 0; i < s.SoundHotkeys.Length; i++)
                    {
                        int kLength = s.SoundHotkeys[i].Keys.Length;
                        bool keysNull = (kLength > 0 && !s.SoundHotkeys[i].Keys.Any(x => x != 0));
                        int sLength = s.SoundHotkeys[i].SoundLocations.Length;
                        bool soundsNotEmpty = s.SoundHotkeys[i].SoundLocations.All(x => !string.IsNullOrWhiteSpace(x)); //false if even one location is empty
                        Environment.CurrentDirectory = Path.GetDirectoryName(Application.ExecutablePath);
                        bool filesExist = s.SoundHotkeys[i].SoundLocations.All(x => File.Exists(x));

                        if (keysNull || sLength < 1 || !soundsNotEmpty || !filesExist) //error in XML file
                        {
                            string tempErr = "";

                            if (kLength == 0 && (sLength == 0 || !soundsNotEmpty)) tempErr = "entry is empty";
                            else if (!keysNull) tempErr = "one or more keys are null";
                            else if (sLength == 0) tempErr = "no sounds provided";
                            else if (!filesExist) tempErr = "one or more sounds do not exist";

                            errorMessage += "Entry #" + (i + 1).ToString() + " has an error: " + tempErr + "\r\n";
                        }

                        string keys = (kLength < 1 ? "" : Helper.KeysArrayToString(s.SoundHotkeys[i].Keys));

                        if (keys != "" && items.Count > 0 && items[items.Count - 1].Text == keys && !sameKeys.Contains(keys))
                        {
                            sameKeys += (sameKeys != "" ? ", " : "") + keys;
                        }

                        string windowString = string.IsNullOrWhiteSpace(s.SoundHotkeys[i].WindowTitle) ? "" : s.SoundHotkeys[i].WindowTitle;
                        string volumeString = s.SoundHotkeys[i].SoundVolume == 1 ? "" : Helper.LinearVolumeToString(s.SoundHotkeys[i].SoundVolume);
                        string soundLocations = sLength < 1 ? "" : Helper.FileLocationsArrayToString(s.SoundHotkeys[i].SoundLocations);

                        ListViewItem temp = new ListViewItem(keys);
                        temp.SubItems.Add(volumeString);
                        temp.SubItems.Add(windowString);
                        temp.SubItems.Add(Path.GetFileNameWithoutExtension(soundLocations));
                        temp.SubItems.Add(soundLocations);

                        temp.ToolTipText = Helper.GetFileNamesTooltip(s.SoundHotkeys[i].SoundLocations); //blank tooltips are not displayed

                        items.Add(temp); //add even if there was an error, so that the user can fix within the app
                    }

                    if (items.Count > 0)
                    {
                        if (errorMessage != "")
                        {
                            MessageBox.Show((errorMessage == "" ? "" : errorMessage));
                        }
                        else
                        {
                            errorOccurred = false;
                        }

                        if (sameKeys != "")
                        {
                            MessageBox.Show("Multiple entries using the same keys. The keys being used multiple times are: " + sameKeys);
                        }

                        soundHotkeys.Clear();
                        soundHotkeys.AddRange(s.SoundHotkeys);

                        lvKeySounds.Items.Clear();
                        lvKeySounds.Items.AddRange(items.ToArray());

                        SortHotkeys();

                        xmlLocation = path;
                    }
                    else
                    {
                        SystemSounds.Beep.Play();
                        MessageBox.Show("No entries found, or all entries had errors in them (key being None, sound location behind empty or non-existant)");
                    }
                }
                else
                {
                    SystemSounds.Beep.Play();
                    MessageBox.Show("No entries found, or there was an error reading the settings file");
                }
            }
            catch
            {
                MessageBox.Show("Settings file structure is incorrect");
            }

            return errorOccurred;
        }

        public void SortHotkeys()
        {
            lvKeySounds.ListViewItemSorter = new ListViewItemComparer(0);
            lvKeySounds.Sort();

            soundHotkeys.Sort((XMLSettings.SoundHotkey x, XMLSettings.SoundHotkey y) =>
            {
                if (x.Keys == null && y.Keys == null) return 0;
                else if (x.Keys == null) return -1;
                else if (y.Keys == null) return 1;
                else return Helper.KeysArrayToString(x.Keys).CompareTo(Helper.KeysArrayToString(y.Keys));
            });

            lvKeySounds.AutoResizeColumns(ColumnHeaderAutoResizeStyle.ColumnContent);
            lvKeySounds.AutoResizeColumns(ColumnHeaderAutoResizeStyle.HeaderSize);
        }

        private void EditSelectedSoundHotkey()
        {
            if (lvKeySounds.SelectedItems.Count > 0)
            {
                AddEditHotkeyForm form = new AddEditHotkeyForm();

                int selectedIndex = lvKeySounds.SelectedIndices[0];

                form.editStrings = new string[4];

                form.editStrings[0] = Helper.KeysArrayToString(soundHotkeys[selectedIndex].Keys);
                form.editStrings[1] = soundHotkeys[selectedIndex].WindowTitle;
                form.editStrings[2] = Helper.FileLocationsArrayToString(soundHotkeys[selectedIndex].SoundLocations);
                form.editVolume = soundHotkeys[selectedIndex].SoundVolume;

                form.editIndex = lvKeySounds.SelectedIndices[0];

                form.ShowDialog();
            }
        }

        private void LoopbackSourceStream_DataAvailable(object sender, WaveInEventArgs e)
        {
            if (loopbackWaveProvider != null && loopbackWaveProvider.BufferedDuration.TotalMilliseconds <= 100)
                loopbackWaveProvider.AddSamples(e.Buffer, 0, e.BytesRecorded);
        }

        private void SettingsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            SettingsForm form = new SettingsForm();
            form.ShowDialog();
        }

        private void TexttospeechToolStripMenuItem_Click(object sender, EventArgs e)
        {
            TextToSpeechForm form = new TextToSpeechForm();
            form.ShowDialog();
        }

        private void CheckForUpdateToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //Jitnaught is inactive, so let's have some fun here 
            Process.Start("https://www.youtube.com/watch?v=dQw4w9WgXcQ");
        }

        private void BtnAdd_Click(object sender, EventArgs e)
        {
            AddEditHotkeyForm form = new AddEditHotkeyForm();
            form.ShowDialog();
        }

        private void BtnEdit_Click(object sender, EventArgs e)
        {
            EditSelectedSoundHotkey();
        }

        private void BtnRemove_Click(object sender, EventArgs e)
        {
            if (lvKeySounds.SelectedItems.Count > 0 && MessageBox.Show("Are you sure remove that item?", "Remove", MessageBoxButtons.YesNo) == DialogResult.Yes)
            {
                soundHotkeys.RemoveAt(lvKeySounds.SelectedIndices[0]);
                lvKeySounds.Items.Remove(lvKeySounds.SelectedItems[0]);

                if (lvKeySounds.Items.Count == 0) cbEnableHotkeys.Checked = false;
            }
        }

        private void BtnClear_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Are you sure you want to clear all items?", "Clear", MessageBoxButtons.YesNo) == DialogResult.Yes)
            {
                soundHotkeys.Clear();
                lvKeySounds.Items.Clear();

                cbEnableHotkeys.Checked = false;
            }
        }

        private void BtnPlaySound_Click(object sender, EventArgs e)
        {
            if (lvKeySounds.SelectedItems.Count > 0) PlayKeySound(soundHotkeys[lvKeySounds.SelectedIndices[0]]);
        }

        private void BtnStopAllSounds_Click(object sender, EventArgs e)
        {
            StopPlayback();
        }

        private void BtnLoad_Click(object sender, EventArgs e)
        {
            OpenFileDialog diag = new OpenFileDialog
            {
                Filter = "XML file containing keys and sounds|*.xml"
            };

            DialogResult result = diag.ShowDialog();

            if (result == DialogResult.OK)
            {
                string path = diag.FileName;

                //loading hotkeys file and saving soundboard settings
                XMLSettings.soundboardSettings.LastXMLFile = LoadXMLFile(path) ? "" : path;
                SaveSettings();
            }
        }

        private void BtnSave_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(xmlLocation))
            {
                SaveHotkeysAs();
            }
            else
            {
                SaveHotkeys();
            }
        }

        private void BtnSaveAs_Click(object sender, EventArgs e)
        {
            SaveHotkeysAs();
        }

        private void SaveHotkeys()
        {
            XMLSettings.WriteXML(new XMLSettings.Settings() { SoundHotkeys = soundHotkeys.ToArray() }, xmlLocation);

            XMLSettings.soundboardSettings.LastXMLFile = xmlLocation;
            SaveSettings();

            MessageBox.Show("Hotkeys saved");
        }

        private void SaveHotkeysAs()
        {
            string path = Helper.UserGetXmlLocation();

            if (string.IsNullOrWhiteSpace(path))
            {
                return;
            }
            else
            {
                xmlLocation = path;
                SaveHotkeys();
            }
        }

        private void BtnReloadDevices_Click(object sender, EventArgs e)
        {
            StopPlayback();
            StopLoopback();

            LoadSoundDevices();
        }

        private void CbEnableHotkeys_CheckedChanged(object sender, EventArgs e)
        {
            mainTimer.Enabled = cbEnableHotkeys.Checked;

            XMLSettings.soundboardSettings.EnableHotkeys = cbEnableHotkeys.Checked;
            SaveSettings();
        }

        private void CbEnableLoopback_CheckedChanged(object sender, EventArgs e)
        {
            RestartLoopback();

            XMLSettings.soundboardSettings.EnableLoopback = cbEnableLoopback.Checked;
            SaveSettings();
        }

        private void LvKeySounds_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            EditSelectedSoundHotkey();
        }


        Keyboard.Keys[] keysJustPressed = null;
        bool showingMsgBox = false;
        int lastIndex = -1;
        private void MainTimer_Tick(object sender, EventArgs e)
        {
            if (cbEnableHotkeys.Checked)
            {
                int keysPressed = 0;

                if (soundHotkeys.Count > 0) //check that required keys are pressed to play sound
                {
                    IntPtr foregroundWindow = Helper.GetForegroundWindow();

                    for (int i = 0; i < soundHotkeys.Count; i++)
                    {
                        keysPressed = 0;

                        if (soundHotkeys[i].Keys.Length == 0) continue;

                        if (soundHotkeys[i].WindowTitle != "" && !Helper.IsForegroundWindow(foregroundWindow, soundHotkeys[i].WindowTitle)) continue;

                        for (int j = 0; j < soundHotkeys[i].Keys.Length; j++)
                        {
                            if (Keyboard.IsKeyDown(soundHotkeys[i].Keys[j]))
                                keysPressed++;
                        }

                        if (keysPressed == soundHotkeys[i].Keys.Length)
                        {
                            if (keysJustPressed == soundHotkeys[i].Keys) continue;

                            if (soundHotkeys[i].Keys.Length > 0 && soundHotkeys[i].Keys.All(x => x != 0) && soundHotkeys[i].SoundLocations.Length > 0
                                && soundHotkeys[i].SoundLocations.Length > 0 && soundHotkeys[i].SoundLocations.Any(x => File.Exists(x)))
                            {
                                if (cbEnablePushToTalk.Checked && !keyUpPushToTalkKey && !Keyboard.IsKeyDown(pushToTalkKey)
                                    && (cbWindows.SelectedIndex == 0 || Helper.IsForegroundWindow(cbWindows.Text)))
                                {
                                    keyUpPushToTalkKey = true;
                                    bool result = Keyboard.SendKey(pushToTalkKey, true);
                                    Thread.Sleep(100);
                                }

                                PlayKeySound(soundHotkeys[i]);
                                return;
                            }
                        }
                        else if (keysJustPressed == soundHotkeys[i].Keys)
                            keysJustPressed = null;
                    }

                    keysPressed = 0;
                }

                if (XMLSettings.soundboardSettings.StopSoundKeys != null && XMLSettings.soundboardSettings.StopSoundKeys.Length > 0) //check that required keys are pressed to stop all sounds
                {
                    for (int i = 0; i < XMLSettings.soundboardSettings.StopSoundKeys.Length; i++)
                    {
                        if (Keyboard.IsKeyDown(XMLSettings.soundboardSettings.StopSoundKeys[i])) keysPressed++;
                    }

                    if (keysPressed == XMLSettings.soundboardSettings.StopSoundKeys.Length)
                    {
                        if (keysJustPressed == null || !keysJustPressed.Intersect(XMLSettings.soundboardSettings.StopSoundKeys).Any())
                        {
                            StopPlayback();

                            keysJustPressed = XMLSettings.soundboardSettings.StopSoundKeys;

                            return;
                        }
                    }
                    else if (keysJustPressed == XMLSettings.soundboardSettings.StopSoundKeys)
                    {
                        keysJustPressed = null;
                    }

                    keysPressed = 0;
                }

                if (XMLSettings.soundboardSettings.LoadXMLFiles != null && XMLSettings.soundboardSettings.LoadXMLFiles.Length > 0) //check that required keys are pressed to load XML file
                {
                    for (int i = 0; i < XMLSettings.soundboardSettings.LoadXMLFiles.Length; i++)
                    {
                        if (XMLSettings.soundboardSettings.LoadXMLFiles[i].Keys.Length == 0) continue;

                        keysPressed = 0;

                        for (int j = 0; j < XMLSettings.soundboardSettings.LoadXMLFiles[i].Keys.Length; j++)
                        {
                            if (Keyboard.IsKeyDown(XMLSettings.soundboardSettings.LoadXMLFiles[i].Keys[j])) keysPressed++;
                        }

                        if (keysPressed == XMLSettings.soundboardSettings.LoadXMLFiles[i].Keys.Length)
                        {
                            if (keysJustPressed == null || !keysJustPressed.Intersect(XMLSettings.soundboardSettings.LoadXMLFiles[i].Keys).Any())
                            {
                                if (!string.IsNullOrWhiteSpace(XMLSettings.soundboardSettings.LoadXMLFiles[i].XMLLocation) && File.Exists(XMLSettings.soundboardSettings.LoadXMLFiles[i].XMLLocation))
                                {
                                    keysJustPressed = XMLSettings.soundboardSettings.LoadXMLFiles[i].Keys;

                                    LoadXMLFile(XMLSettings.soundboardSettings.LoadXMLFiles[i].XMLLocation);
                                }

                                return;
                            }
                        }
                        else if (keysJustPressed == XMLSettings.soundboardSettings.LoadXMLFiles[i].Keys)
                        {
                            keysJustPressed = null;
                        }
                    }

                    keysPressed = 0;
                }

                if (keyUpPushToTalkKey)
                {
                    if (!Keyboard.IsKeyDown(pushToTalkKey)) keyUpPushToTalkKey = false;

                    if (cbWindows.SelectedIndex > 0 && !Helper.IsForegroundWindow(cbWindows.Text))
                    {
                        keyUpPushToTalkKey = false;
                        Keyboard.SendKey(pushToTalkKey, false);
                        Keyboard.SendKey(pushToTalkKey, false);
                    }
                }
            }
        }

        private void PlayKeySound(XMLSettings.SoundHotkey currentKeysSounds)
        {
            Environment.CurrentDirectory = Path.GetDirectoryName(Application.ExecutablePath);

            string path;

            if (currentKeysSounds.SoundLocations.Length > 1)
            {
                //get random sound
                int temp;

                while (true)
                {
                    temp = rand.Next(0, currentKeysSounds.SoundLocations.Length);

                    if (temp != lastIndex && (File.Exists(currentKeysSounds.SoundLocations[temp]) || currentKeysSounds.SoundLocations[temp].Contains("http"))) break;
                    Thread.Sleep(1);
                }

                lastIndex = temp;

                path = currentKeysSounds.SoundLocations[lastIndex];
            }
            else
                path = currentKeysSounds.SoundLocations[0]; //get first sound

            if (File.Exists(path) || path.Contains("http"))
            {
                float customSoundVolume = currentKeysSounds.SoundVolume;

                //use custom sound volume if the user changed it
                if (customSoundVolume < 1)
                {
                    PlaySound(path, customSoundVolume);
                }
                else
                {
                    PlaySound(path, soundVolume);
                }

                keysJustPressed = currentKeysSounds.Keys;
            }
            else if (!showingMsgBox) //dont run when already showing messagebox (don't want a bunch of these on your screen, do you?)
            {
                SystemSounds.Beep.Play();
                showingMsgBox = true;
                MessageBox.Show("File " + path + " does not exist");
                showingMsgBox = false;
            }
        }
      

        private void CbLoopbackDevices_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cbEnableLoopback.Checked) //start loopback on new device, or stop loopback
            {
                RestartLoopback();
            }

            string deviceName = (string)cbLoopbackDevices.SelectedItem;

            XMLSettings.soundboardSettings.LastLoopbackDevice = deviceName;
            SaveSettings();
        }

        private void CbPlaybackDevices1_SelectedIndexChanged(object sender, EventArgs e)
        {
            //start loopback on new device and stop all sounds playing
            if (loopbackWaveOut != null && loopbackSourceStream != null)
            {
                RestartLoopback();
            }

            StopPlayback();

            string deviceName = (string)cbPlaybackDevices1.SelectedItem;

            InitAudioPlaybackEngine1();

            XMLSettings.soundboardSettings.LastPlaybackDevice = deviceName;
            SaveSettings();
        }

        private void CbPlaybackDevices2_SelectedIndexChanged(object sender, EventArgs e)
        {
            //start loopback on new device and stop all sounds playing
            if (loopbackWaveOut != null && loopbackSourceStream != null)
            {
                RestartLoopback();
            }
                
            StopPlayback();

            string deviceName = (string)cbPlaybackDevices2.SelectedItem;

            InitAudioPlaybackEngine2();

            XMLSettings.soundboardSettings.LastPlaybackDevice2 = deviceName;
            SaveSettings();
        }

        private void FrmMain_Resize(object sender, EventArgs e)
        {
            if (this.WindowState == FormWindowState.Minimized && XMLSettings.soundboardSettings.MinimiseToTray)
            {
                this.HideFormToTray();
            }

            //deselect all controls (to set values)
            this.ActiveControl = null;
        }

        private void NotifyIcon1_MouseUp(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                notifyIcon1.Visible = false;

                this.ShowForm();
            }

            if (e.Button == MouseButtons.Right)
            {
                contextMenuStrip1.Show();
            }
        }

        private void Open_Click(object sender, EventArgs e)
        {
            this.ShowForm();
        }


        private void Exit_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void ShowForm()
        {
            allowVisible = true;

            this.Show();
            this.WindowState = FormWindowState.Normal;
        }

        private void HideFormToTray()
        {
            allowVisible = false;
            notifyIcon1.Visible = true;

            this.Hide();
        }

        
        private bool cbEnableHotkeysWasChecked = false;
        private void TbPushToTalkKey_Enter(object sender, EventArgs e)
        {
            if (cbEnableHotkeys.Checked)
            {
                cbEnableHotkeys.Checked = false;
                cbEnableHotkeysWasChecked = true;
            }


            pushToTalkKeyTimer.Enabled = true;
        }

        private void TbPushToTalkKey_Leave(object sender, EventArgs e)
        {
            //only check enable hotkeys if it was previously checked before changing this field 
            if (cbEnableHotkeysWasChecked)
            {
                cbEnableHotkeys.Checked = true;
                cbEnableHotkeysWasChecked = false;
            }

            pushToTalkKeyTimer.Enabled = false;
            pushToTalkKeyTimer.Interval = 100;

            XMLSettings.soundboardSettings.AutoPushToTalkKey = pushToTalkKey;
            SaveSettings();
        }

        private void PushToTalkKeyTimer_Tick(object sender, EventArgs e)
        {
            pushToTalkKeyTimer.Interval = 10; //lowering the interval to avoid missing key presses (e.g. when an input is corrected)

            if (Keyboard.IsKeyDown(Keyboard.Keys.Escape))
            {
                tbPushToTalkKey.Text = "";
                pushToTalkKey = default;
            }
            else
            {
                foreach (Keyboard.Keys key in Enum.GetValues(typeof(Keyboard.Keys)))
                {
                    if (Keyboard.IsKeyDown(key))
                    {
                        tbPushToTalkKey.Text = key.ToString();
                        pushToTalkKey = key;

                        break;
                    }
                }
            }
        }

        private void CbEnablePushToTalk_CheckedChanged(object sender, EventArgs e)
        {
            if (tbPushToTalkKey.Text == "")
            {
                cbEnablePushToTalk.Checked = false;
                MessageBox.Show("There is no push to talk key entered");

                return;
            }

            tbPushToTalkKey.Enabled = !cbEnablePushToTalk.Checked;
            clearHotkey.Enabled = !cbEnablePushToTalk.Checked;

            XMLSettings.soundboardSettings.EnableAutoPushToTalk = cbEnablePushToTalk.Checked;
            SaveSettings();
        }

        private void CbWindows_Leave(object sender, EventArgs e)
        {
            SaveAutoPushToTalkWindow();
        }

        private void BtnReloadWindows_Click(object sender, EventArgs e)
        {
            Helper.GetWindows(cbWindows);
            SaveAutoPushToTalkWindow();
        }

        private void SaveAutoPushToTalkWindow() {
            XMLSettings.soundboardSettings.AutoPushToTalkWindow = cbWindows.SelectedIndex == 0 ? "" : cbWindows.Text;
            SaveSettings();
        }

        private void ClearHotkey_Click( object sender, EventArgs e )
        {
            tbPushToTalkKey.Text = "";

            XMLSettings.soundboardSettings.AutoPushToTalkKey = default;
            SaveSettings();
        }


        private bool volumeChangedBySlider = false;
        private bool volumeChangedByField = false;
        public void VsSoundVolume_VolumeChanged(object sender, EventArgs e)
        {
            soundVolume = vsSoundVolume.Volume;

            XMLSettings.soundboardSettings.SoundVolume = soundVolume;
            SaveSettings();

            //prevent infinite or skipped changes
            if (volumeChangedByField)
            {
                volumeChangedByField = false;

                return;
            }

            volumeChangedBySlider = true;

            nSoundVolume.Value = Helper.LinearVolumeToInteger(vsSoundVolume.Volume);
        }

        public void VsSoundVolume_MouseWheel(object sender, MouseEventArgs e)
        {
            vsSoundVolume.Volume = Helper.GetNewSoundVolume(vsSoundVolume.Volume, e.Delta);
        }

        public void NSoundVolume_ValueChanged(object sender, EventArgs e)
        {
            //prevent infinite or skipped changes
            if (volumeChangedBySlider)
            {
                volumeChangedBySlider = false;

                return;
            }

            volumeChangedByField = true;

            vsSoundVolume.Volume = (float)(nSoundVolume.Value / 100);
        }

        private void SaveSettings()
        {
            saveSettingsTimer.Stop();
            saveSettingsTimer.Start();
        }

        private void SaveSettingsTimer_Tick(object sender, EventArgs e)
        {
            //only save settings after no setting changes have been made for one second
            saveSettingsTimer.Stop();
            XMLSettings.SaveSoundboardSettingsXML();
        }

        public void Form_Click(object sender, EventArgs e)
        {
            //deselect all controls (to set values)
            this.ActiveControl = null;
        }
    }
}