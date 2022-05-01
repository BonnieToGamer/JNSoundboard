using MediaToolkit;
using MediaToolkit.Model;
using System;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using VideoLibrary;

namespace JNSoundboard
{
    public partial class AddEditHotkeyForm : Form
    {
        MainForm mainForm;
        SettingsForm settingsForm;

        internal static string lastWindow = "";
        internal static float lastSoundVolume = 1;

        internal string[] editStrings = null;
        internal float editVolume;
        internal int editIndex = -1;

        public AddEditHotkeyForm()
        {
            InitializeComponent();
        }

        private void AddEditSoundKeys_Load(object sender, EventArgs e)
        {
            if (SettingsForm.addingEditingLoadXMLFile)
            {
                //resize and hide components unrelated to settings form
                this.Size = new Size(282, 176);
                pnAddEditSound.Visible = false;

                settingsForm = Application.OpenForms[1] as SettingsForm;

                this.Text = "Add/edit keys and XML location";

                if (editIndex != -1)
                {
                    tbKeys.Text = editStrings[0];
                    tbLocation.Text = editStrings[1];
                }
            }
            else
            {
                mainForm = Application.OpenForms[0] as MainForm;

                labelLocation.Text += "(s) \n(use ; to separate multiple locations)";
                labelKeys.Text += " (optional)";

                if (editIndex != -1)
                {
                    tbLocation.Text = editStrings[2];
                    tbKeys.Text = editStrings[0];
                }

                vsSoundVolume.Volume = (editIndex != -1) ? editVolume : lastSoundVolume;

                Helper.GetWindows(cbWindows);

                string windowToSelect = (editIndex != -1) ? editStrings[1] : lastWindow;

                Helper.SelectWindow(cbWindows, windowToSelect);
            }
        }

        private void BtnOK_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(tbLocation.Text))
            {
                MessageBox.Show("Location is empty");
                return;
            }

            if (SettingsForm.addingEditingLoadXMLFile && string.IsNullOrWhiteSpace(tbKeys.Text))
            {
                MessageBox.Show("No keys entered");
                return;
            }

            string[] soundLocations = null;
            string fileNames;

            if (!SettingsForm.addingEditingLoadXMLFile)
            {
                if (Helper.StringToFileLocationsArray(tbLocation.Text, out soundLocations, out string errorMessage))
                {
                    if (soundLocations.Any(x => string.IsNullOrWhiteSpace(x) || (!File.Exists(x) && !x.Contains("http"))))
                    {
                        MessageBox.Show("One or more files or URLs do not exist");

                        this.Close();

                        return;
                    }
                }

                if (soundLocations == null)
                {
                    MessageBox.Show(errorMessage);
                    return;
                }
            }

            if (!Helper.StringToKeysArray(tbKeys.Text, out Keyboard.Keys[] keysArray, out _)) keysArray = new Keyboard.Keys[] { };

            if (SettingsForm.addingEditingLoadXMLFile)
            {
                fileNames = Helper.FileLocationsArrayToString(new string[] { tbLocation.Text });

                if (editIndex != -1)
                {
                    settingsForm.lvKeysLocations.Items[editIndex].Text = tbKeys.Text;
                    settingsForm.lvKeysLocations.Items[editIndex].SubItems[1].Text = tbLocation.Text;

                    settingsForm.lvKeysLocations.Items[editIndex].ToolTipText = fileNames;

                    settingsForm.loadXMLFilesList[editIndex].Keys = keysArray;
                    settingsForm.loadXMLFilesList[editIndex].XMLLocation = tbLocation.Text;
                }
                else
                {
                    ListViewItem item = new ListViewItem(tbKeys.Text);
                    item.SubItems.Add(tbLocation.Text);

                    item.ToolTipText = fileNames;

                    settingsForm.lvKeysLocations.Items.Add(item);

                    settingsForm.loadXMLFilesList.Add(new XMLSettings.LoadXMLFile(keysArray, tbLocation.Text));
                }
            }
            else
            {
                fileNames = Helper.GetFileNamesTooltip(soundLocations);

                string volumeString = vsSoundVolume.Volume == 1 ? "" : Helper.LinearVolumeToString(vsSoundVolume.Volume);

                string windowText = (cbWindows.SelectedIndex > 0) ? cbWindows.Text : "";

                if (editIndex != -1)
                {
                    mainForm.lvKeySounds.Items[editIndex].SubItems[0].Text = tbKeys.Text;
                    mainForm.lvKeySounds.Items[editIndex].SubItems[1].Text = volumeString;
                    mainForm.lvKeySounds.Items[editIndex].SubItems[2].Text = windowText;
                    mainForm.lvKeySounds.Items[editIndex].SubItems[3].Text = tbLocation.Text;

                    mainForm.lvKeySounds.Items[editIndex].ToolTipText = fileNames;

                    mainForm.soundHotkeys[editIndex] = new XMLSettings.SoundHotkey(keysArray, vsSoundVolume.Volume, windowText, soundLocations);
                }

                else
                {
                    string title = tbLocation.Text.Contains("http") ? Helper.youTube.GetVideo(tbLocation.Text).Title : Path.GetFileName(tbLocation.Text);

                    ListViewItem newItem = new ListViewItem(tbKeys.Text);
                    newItem.SubItems.Add(volumeString);
                    newItem.SubItems.Add(windowText);
                    newItem.SubItems.Add(title);
                    newItem.SubItems.Add(tbLocation.Text);

                    newItem.ToolTipText = fileNames;

                    mainForm.lvKeySounds.Items.Add(newItem);

                    mainForm.soundHotkeys.Add(new XMLSettings.SoundHotkey(keysArray, vsSoundVolume.Volume, windowText, soundLocations));
                }

                mainForm.SortHotkeys();

                //remember last used options
                lastSoundVolume = vsSoundVolume.Volume;
                lastWindow = cbWindows.Text;
            }

            this.Close();
        }

        private void BtnCancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void BtnBrowseSoundLocation_Click(object sender, EventArgs e)
        {
            OpenFileDialog diag = new OpenFileDialog
            {
                Multiselect = !SettingsForm.addingEditingLoadXMLFile,

                Filter = (SettingsForm.addingEditingLoadXMLFile ? "XML file containing keys and sounds|*.xml" : "Supported audio formats|*.mp3;*.m4a;*.wav;*.wma;*.ac3;*.aiff;*.mp2|All files|*.*")
            };

            DialogResult result = diag.ShowDialog();

            if (result == DialogResult.OK)
            {
                string text = "";

                for (int i = 0; i < diag.FileNames.Length; i++)
                {
                    string fileName = diag.FileNames[i];

                    if (fileName != "") text += (i == 0 ? "" : ";") + fileName;
                }

                tbLocation.Text = text;
            }
        }

        private void TbKeys_Enter(object sender, EventArgs e)
        {
            keysTimer.Enabled = true;
        }

        private void TbKeys_Leave(object sender, EventArgs e)
        {
            keysTimer.Enabled = false;
            keysTimer.Interval = 100;
        }


        private int lastAmountPressed = 0;
        private void KeysTimer_Tick(object sender, EventArgs e)
        {
            keysTimer.Interval = 10; //lowering the interval to avoid missing key presses (e.g. when an input is corrected)

            Tuple<int, string> keysData = Keyboard.GetKeys(lastAmountPressed, tbKeys.Text);

            lastAmountPressed = keysData.Item1;
            tbKeys.Text = keysData.Item2;
        }

        private void BtnReloadWindows_Click(object sender, EventArgs e)
        {
            Helper.GetWindows(cbWindows);
        }

        private void ClearHotkey_Click( object sender, EventArgs e )
        {
            tbKeys.Text = "";
        }


        private bool volumeChangedBySlider = false;
        private bool volumeChangedByField = false;
        public void VsSoundVolume_VolumeChanged(object sender, EventArgs e)
        {
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
    }
}
