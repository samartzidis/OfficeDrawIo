using System;
using System.Windows.Forms;

namespace OfficeDrawIo
{
    public partial class SettingsForm : Form
    {
        public Action ResetSettingsAction { get; private set; }

        public SettingsForm(object settingsAdapter, Action resetSettingsAction = null)
        {
            InitializeComponent();

            propertyGrid.SelectedObject = settingsAdapter;
            ResetSettingsAction = resetSettingsAction;

            if (resetSettingsAction == null)
                defaultsButton.Visible = false;
        }

        private void defaultsButton_Click(object sender, EventArgs e)
        {
            if (ResetSettingsAction != null)
                ResetSettingsAction.Invoke();

            propertyGrid.Refresh();
        }

        private void okButton_Click(object sender, EventArgs e)
        {
            DialogResult = System.Windows.Forms.DialogResult.OK;
        }

        private void cancelButton_Click(object sender, EventArgs e)
        {
            DialogResult = System.Windows.Forms.DialogResult.Cancel;
        }
    }    
}