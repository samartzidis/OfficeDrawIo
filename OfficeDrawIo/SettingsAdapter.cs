using System;
using System.ComponentModel;
using System.Drawing.Design;
using System.Windows.Forms;

namespace OfficeDrawIo
{
    class SettingsAdapter
    {
        [Category("General"), DisplayName("Draw.io Path"), Description("Draw.io (desktop version) executable file path. You may have to install is from https://github.com/jgraph/drawio-desktop/releases if it does not exist on your PC.")]
        [PathEditor.OfdParams("Executable files (*.exe)|*.exe", "Selection")]
        [Editor(typeof(PathEditor), typeof(UITypeEditor))]
        public string DrawIoExePath
        {
            get => Properties.Settings.Default.DrawIoExePath;
            set => Properties.Settings.Default.DrawIoExePath = value;
        }

        public void Save()
        {
            Properties.Settings.Default.Save();
        }

        public void Reset()
        {
            Properties.Settings.Default.Reset();
        }

        #region UITypeEditors
        class PathEditor : UITypeEditor
        {
            //A class to hold our OpenFileDialog Settings
            public class OfdParamsAttribute : Attribute
            {
                public OfdParamsAttribute(string filter, string title)
                {
                    Filter = filter;
                    Title = title;
                }

                //The File Filter(s) of the open dialog
                public string Filter { get; set; }

                //The Title of the open dialog
                public string Title { get; set; }
            }

            //The default settings for the file dialog
            private OfdParamsAttribute _settings = new OfdParamsAttribute("All Files (*.*)|*.*", "Open");
            public OfdParamsAttribute Settings
            {
                get { return _settings; }
                set { _settings = value; }
            }

            //Define a modal editor style and capture the settings from the property
            public override UITypeEditorEditStyle GetEditStyle(ITypeDescriptorContext context)
            {
                if (context == null || context.Instance == null)
                    return base.GetEditStyle(context);

                //Retrieve our settings attribute (if one is specified)
                OfdParamsAttribute sa = (OfdParamsAttribute)context.PropertyDescriptor.Attributes[typeof(OfdParamsAttribute)];
                if (sa != null)
                    Settings = sa; //Store it in the editor

                return UITypeEditorEditStyle.Modal;
            }

            //Do the actual editing
            public override object EditValue(ITypeDescriptorContext context, IServiceProvider provider, object value)
            {
                if (context == null || context.Instance == null || provider == null)
                    return value;

                OpenFileDialog dlg = new OpenFileDialog();
                dlg.Filter = Settings.Filter;
                dlg.CheckFileExists = true;
                dlg.Title = Settings.Title;

                string newValue = (string)value;
                if (dlg.ShowDialog() == DialogResult.OK)
                    newValue = dlg.FileName;

                return newValue;
            }
        }
        #endregion
    }
}
