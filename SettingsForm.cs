using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace JNSoundboard
{
    public partial class SettingsForm : Form
    {
        internal List<XMLSettings.LoadXMLFile> loadXMLFilesList = new List<XMLSettings.LoadXMLFile>(XMLSettings.soundboardSettings.LoadXMLFiles); //list so can dynamically add/remove

        internal static bool addingEditingLoadXMLFile = false;

        public SettingsForm()
        {
            InitializeComponent();

            for (int i = 0; i < loadXMLFilesList.Count; i++)
            {
                string xmlLocation = loadXMLFilesList[i].XMLLocation;

                bool keysLengthCorrect = loadXMLFilesList[i].Keys.Length > 0;
                bool xmlLocationUnempty = !string.IsNullOrWhiteSpace(xmlLocation);

                if (!keysLengthCorrect && !xmlLocationUnempty) //remove if empty
                {
                    loadXMLFilesList.RemoveAt(i);
                    i--;
                    continue;
                }

                ListViewItem item = new ListViewItem((keysLengthCorrect ? string.Join("+", loadXMLFilesList[i].Keys) : ""));
                item.SubItems.Add((xmlLocationUnempty ? xmlLocation : ""));

                item.ToolTipText = Helper.GetFileNamesTooltip(new string[] { xmlLocation });

                lvKeysLocations.Items.Add(item);
            }

            tbStopSoundKeys.Text = Helper.KeysArrayToString(XMLSettings.soundboardSettings.StopSoundKeys);

            cbStartWithWindows.Checked = XMLSettings.soundboardSettings.StartWithWindows;

            cbStartMinimised.Checked = XMLSettings.soundboardSettings.StartMinimised;

            cbMinimiseToTray.Checked = XMLSettings.soundboardSettings.MinimiseToTray;
        }

        private void BtnAdd_Click(object sender, EventArgs e)
        {
            addingEditingLoadXMLFile = true;

            AddEditHotkeyForm form = new AddEditHotkeyForm();
            form.ShowDialog();

            addingEditingLoadXMLFile = false;
        }

        private void BtnEdit_Click(object sender, EventArgs e)
        {
            if (lvKeysLocations.SelectedIndices.Count > 0)
            {
                addingEditingLoadXMLFile = true;

                AddEditHotkeyForm form = new AddEditHotkeyForm
                {
                    editIndex = lvKeysLocations.SelectedIndices[0],
                    editStrings = new string[] { lvKeysLocations.SelectedItems[0].Text, lvKeysLocations.SelectedItems[0].SubItems[1].Text }
                };

                form.ShowDialog();

                addingEditingLoadXMLFile = false;
            }
        }

        private void BtnRemove_Click(object sender, EventArgs e)
        {
            if (lvKeysLocations.SelectedIndices.Count > 0 && MessageBox.Show("Are you sure?", "Are you sure?", MessageBoxButtons.YesNo) == DialogResult.Yes)
            {
                int index = lvKeysLocations.SelectedIndices[0];

                lvKeysLocations.Items.RemoveAt(index);
                loadXMLFilesList.RemoveAt(index);
            }
        }
        
        private void BtnOK_Click(object sender, EventArgs e)
        {
            if (!Helper.StringToKeysArray(tbStopSoundKeys.Text, out Keyboard.Keys[] keysArray, out _)) keysArray = new Keyboard.Keys[] { };

            if (loadXMLFilesList.Count == 0 || loadXMLFilesList.All(x => x.Keys.Length > 0 && !string.IsNullOrWhiteSpace(x.XMLLocation) && File.Exists(x.XMLLocation)))
            {
                XMLSettings.soundboardSettings.StopSoundKeys = keysArray;

                XMLSettings.soundboardSettings.LoadXMLFiles = loadXMLFilesList.ToArray();

                XMLSettings.soundboardSettings.StartWithWindows = cbStartWithWindows.Checked;
                Helper.SetStartup(XMLSettings.soundboardSettings.StartWithWindows);

                XMLSettings.soundboardSettings.StartMinimised = cbStartMinimised.Checked;

                XMLSettings.soundboardSettings.MinimiseToTray = cbMinimiseToTray.Checked;

                XMLSettings.SaveSoundboardSettingsXML();

                this.Close();
            }
            else
            {
                MessageBox.Show("One or more entries either have no keys added, the location is empty, or the file the location points to does not exist");
            }
        }

        private void BtnCancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void LvKeysLocations_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            BtnEdit_Click(null, null);
        }

        private void TbStopSoundKeys_Enter(object sender, EventArgs e)
        {
            keysTimer.Enabled = true;
        }

        private void TbStopSoundKeys_Leave(object sender, EventArgs e)
        {
            keysTimer.Enabled = false;
            keysTimer.Interval = 100;
        }


        private int lastAmountPressed = 0;
        private void KeysTimer_Tick(object sender, EventArgs e)
        {
            keysTimer.Interval = 10; //lowering the interval to avoid missing key presses (e.g. when an input is corrected)

            Tuple<int, string> keysData = Keyboard.GetKeys(lastAmountPressed, tbStopSoundKeys.Text);

            lastAmountPressed = keysData.Item1;
            tbStopSoundKeys.Text = keysData.Item2;
        }

        private void ClearHotkey_Click(object sender, EventArgs e)
        {
            tbStopSoundKeys.Text = "";
        }
    }
}