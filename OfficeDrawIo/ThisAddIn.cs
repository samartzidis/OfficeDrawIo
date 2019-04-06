using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Formatters.Binary;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using System.Xml;
using System.Xml.Linq;

namespace OfficeDrawIo
{
    public partial class ThisAddIn
    {
        public SynchronizationContext TheWindowsFormsSynchronizationContext { get; private set; }
        public Microsoft.Office.Tools.Word.PictureContentControl SelectedCtrl { get; private set; }

        private static ThisAddIn _addin;    
        private string _userTmpFilesDir;
        private SettingsAdapter _settings;
        private SettingsForm _sf;
        private string _drawioExportDir;
        private FileSystemWatcher _watcher;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            Trace.Listeners.Add(new TraceListener());

            Trace.WriteLine("ThisAddIn_Startup()");

            TheWindowsFormsSynchronizationContext = SynchronizationContext.Current ?? new WindowsFormsSynchronizationContext();            

            _addin = this;
            _drawioExportDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "drawio-export");
            _userTmpFilesDir = Path.Combine(Path.GetTempPath(), "OfficeDrawIo");
            _settings = new SettingsAdapter();
            _sf = new SettingsForm(_settings, () =>
            {
                Properties.Settings.Default.Reset();
            });

            Application.DocumentBeforeSave += Application_DocumentBeforeSave;
            Application.DocumentBeforeClose += Application_DocumentBeforeClose;
            Application.DocumentChange += Application_DocumentChange;


            if (!Directory.Exists(_userTmpFilesDir))
                Directory.CreateDirectory(_userTmpFilesDir);

            CreateFileWatcher(_userTmpFilesDir);  
        }

        private void Application_DocumentChange()
        {
            Trace.WriteLine("Application_DocumentChange()");

            // It may happen that there is an already active open document before the add-in has completed startup, so do this
            try
            {
                if (Application.ActiveDocument != null) // throws if there is no ActiveDocument
                    ManageDoc(Application.ActiveDocument);
            }
            catch
            {
            }
        }

        private void ManageDoc(Microsoft.Office.Interop.Word.Document doc)
        {
            if (doc == null)
                return;

            var vstoDoc = ActiveVstoDocument;
            if (vstoDoc == null)
                return;

            foreach (Microsoft.Office.Interop.Word.ContentControl nativeControl in doc.ContentControls)
            {
                if (nativeControl.Type == Microsoft.Office.Interop.Word.WdContentControlType.wdContentControlPicture)
                {
                    if (!IsDrawioTag(nativeControl.Tag))
                        continue;

                    // See: https://docs.microsoft.com/en-us/visualstudio/vsto/persisting-dynamic-controls-in-office-documents?view=vs-2017
                    var ctrl = vstoDoc.Controls.AddPictureContentControl(nativeControl, nativeControl.Tag);

                    ctrl.LockContents = true;

                    ctrl.Entering += PictureControl_Entering;
                    ctrl.Exiting += PictureControl_Exiting;
                    ctrl.Deleting += PictureControl_Deleting;
                }
            }

            vstoDoc.ContentControlAfterAdd += VstoDoc_ContentControlAfterAdd;
        }

        private void VstoDoc_ContentControlAfterAdd(Microsoft.Office.Interop.Word.ContentControl addedControl, bool inUndoRedo)
        {
            if (IsDrawioTag(addedControl.Tag)) // Did we copy from an existing draw.io specific PictureControl?
            {
                Microsoft.Office.Core.CustomXMLPart part;
                Microsoft.Office.Tools.Word.PictureContentControl ctrl;
                              
                var dataPartHelper = new DrawIoDataPartHelper(Application.ActiveDocument);
                var incomingId = GetDrawioTagId(addedControl.Tag);
                var data = dataPartHelper.GetDrawIoDataPart(incomingId); // Get PictureControl associated Draw.io image    
                
                if (data == null) // If data is null it means that the incoming draw.io specific PictureControl is not from this document
                {
                    var path = Path.Combine(_userTmpFilesDir, $"{incomingId}.png");
                    data = Helpers.LoadBinaryResource("Resources.new.png");         
                    part = dataPartHelper.AddDrawIoDataPart(data);

                    ctrl = ActiveVstoDocument.Controls.AddPictureContentControl(part.Id);

                    Image img;
                    using (var stream = new MemoryStream(data, false))
                        img = Image.FromStream(stream);

                    ctrl.Image = img;
                }
                else
                {
                    part = dataPartHelper.AddDrawIoDataPart(data);
                    ctrl = ActiveVstoDocument.Controls.AddPictureContentControl(addedControl, part.Id);

                    ctrl.LockContents = false;
                }

                ctrl.Tag = MakeDrawioTag(part.Id);
                ctrl.Title = $"Draw.io diagram {part.Id}";

                ctrl.LockContents = true;

                ctrl.Entering += PictureControl_Entering;
                ctrl.Exiting += PictureControl_Exiting;
                ctrl.Deleting += PictureControl_Deleting;

                SelectedCtrl = ctrl;
            }
        }

        public void AddDrawIoDiagramOnDocument()
        {            
            if (!ValidateDependencies())
                return;

            var dataPartHelper = new DrawIoDataPartHelper(Application.ActiveDocument);

            var data = Helpers.LoadBinaryResource("Resources.new.png");
            Image img;
            using (var stream = new MemoryStream(data, false))
                img = Image.FromStream(stream);
            var part = dataPartHelper.AddDrawIoDataPart(data);
            var id = part.Id;

            var ctrl = ActiveVstoDocument.Controls.AddPictureContentControl(id);
            ctrl.Title = $"Draw.io diagram {id}";
            ctrl.Tag = MakeDrawioTag(id);
            ctrl.Image = img;
            ctrl.LockContents = true;

            ctrl.Entering += PictureControl_Entering;
            ctrl.Exiting += PictureControl_Exiting;
            ctrl.Deleting += PictureControl_Deleting;

            SelectedCtrl = ctrl;
        }

        public void EditDrawIoDiagramOnDocument()
        {
            Trace.WriteLine("EditDrawIoDiagramOnDocument()");

            if (SelectedCtrl == null)
                return;

            if (!ValidateDependencies())
                return;

            var id = GetDrawioTagId(SelectedCtrl.Tag);
            if (id == null)
                return;

            var wnd = NativeWindowHelper.FindWindow($"{id}.drawio - draw.io");
            if (wnd != IntPtr.Zero)
            {
                NativeWindowHelper.RestoreFromMinimized(wnd);
                NativeWindowHelper.SetForegroundWindow(wnd);

                return;
            }

            var dataPartHelper = new DrawIoDataPartHelper(Application.ActiveDocument);

            var drawioFilePath = Path.Combine(_userTmpFilesDir, $"{id}.png");
            var data = dataPartHelper.GetDrawIoDataPart(id);
            if (data == null)
                return;

            try
            {
                File.WriteAllBytes(drawioFilePath, data);

                var process = new Process();
                process.StartInfo.FileName = _settings.DrawIoExePath;
                process.StartInfo.Arguments = drawioFilePath;
                process.StartInfo.WindowStyle = ProcessWindowStyle.Maximized;

                process.Start();
            }
            catch (Exception m)
            {
                var msg = $"Failed to start Draw.io Desktop application for file {drawioFilePath}. Error: {m.Message}.";
                MessageBox.Show(msg, Application.ActiveWindow.Caption, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }

        public void ExportDrawIoDiagram()
        {
            Trace.WriteLine("ExportDrawIoDiagram()");

            if (SelectedCtrl == null)
                return;

            if (!ValidateDependencies())
                return;

            var dlg = new SaveFileDialog();
            dlg.Filter = "Draw.io files (*.png)|*.png|All files (*.*)|*.*";
            dlg.DefaultExt = ".png";

            if (dlg.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    var dataPartHelper = new DrawIoDataPartHelper(Application.ActiveDocument);
                    var data = dataPartHelper.GetDrawIoDataPart(GetDrawioTagId(SelectedCtrl.Tag));

                    File.WriteAllBytes(dlg.FileName, data);
                }
                catch (Exception m)
                {
                    MessageBox.Show($"Failed to export Draw.io document: {m.Message}.",
                        Application.ActiveWindow.Caption, MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            
        }

        public void Settings()
        {
            if (_sf.ShowDialog() == DialogResult.OK)
                _settings.Save();
        }

        public void AddInNotifyChanged(string id)
        {
            Trace.WriteLine($"AddInNotifyChanged(partId = {id})");

            try
            {
                var dataPartHelper = new DrawIoDataPartHelper(Application.ActiveDocument);
                if (!dataPartHelper.ExistsDrawIoDataPart(id))
                    return;

                var path = Path.Combine(_userTmpFilesDir, $"{id}.png");
                if (!File.Exists(path))
                    return;

                var ctrl = FindPictureContentControlById(id);
                if (ctrl == null)
                    return;

                byte[] data = null;
                long lastModifiedTimestamp = 0;
                try
                {
                    data = File.ReadAllBytes(path);
                    lastModifiedTimestamp = File.GetLastWriteTime(path).Ticks;
                }
                catch(Exception m)
                {
                    Trace.WriteLine(m.Message);
                    return; // File may be locked by Draw.io desktop app.
                }

                Trace.WriteLine($"New image data length = {data.Length}");
                dataPartHelper.UpdateDrawIoDataPart(id, data);

                Image img = null;
                using (var stream = new MemoryStream(data, false))
                    img = Image.FromStream(stream);

                ctrl.LockContents = false;
                ctrl.Image = img;
                ctrl.LockContents = true;                
            }
            catch (Exception m)
            {
                MessageBox.Show($"Something went wrong while updating the the Draw.io diagram.\r\nError details:\r\n{m}",
                            Application.ActiveWindow.Caption, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }

        private Microsoft.Office.Tools.Word.PictureContentControl FindPictureContentControlById(string id)
        {
            foreach (var cc in ActiveVstoDocument.Controls)
                if (cc is Microsoft.Office.Tools.Word.PictureContentControl ctrl && GetDrawioTagId(ctrl.Tag) == id)
                    return ctrl;

            return null;
        }

        public void About()
        {
            var dlg = new AboutBox();
            dlg.ShowDialog();
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            _watcher?.Dispose();
        }

        private bool IsDrawioTag(string tag)
        {
            return tag != null && tag.StartsWith("OfficeDrawIo:");
        }

        private string GetDrawioTagId(string tag)
        {
            if (!IsDrawioTag(tag))
                return null;
            return tag.Split(':')[1];
        }

        private string MakeDrawioTag(string id)
        {
            return $"OfficeDrawIo:{id}";
        }

        private void Application_DocumentBeforeClose(Microsoft.Office.Interop.Word.Document doc, ref bool cancel)
        {
            // Cleanup temp files
            foreach (var cc in ActiveVstoDocument.Controls)
            {
                if (cc is Microsoft.Office.Tools.Word.PictureContentControl ctrl)
                {
                    var partId = GetDrawioTagId(ctrl.Tag);
                    if (partId == null)
                        continue;
                    try
                    {
                        var imageFilePath = Path.Combine(_userTmpFilesDir, $"{partId}.png");
                        if (File.Exists(imageFilePath))
                            File.Delete(imageFilePath);
                    }
                    catch { }
                }
            }
            
        }

        private Microsoft.Office.Tools.Word.Document ActiveVstoDocument
        {
            get
            {
                if (Application.ActiveDocument == null)
                    return null;
                try
                {
                    return Globals.Factory.GetVstoObject(Application.ActiveDocument);
                }
                catch
                {
                }

                return null;
            }
        }        


        private void PictureControl_Exiting(object sender, Microsoft.Office.Tools.Word.ContentControlExitingEventArgs e)
        {
            SelectedCtrl = null;
        }

        private void PictureControl_Entering(object sender, Microsoft.Office.Tools.Word.ContentControlEnteringEventArgs e)
        {
            SelectedCtrl = sender as Microsoft.Office.Tools.Word.PictureContentControl;
        }

        private void PictureControl_Deleting(object sender, Microsoft.Office.Tools.Word.ContentControlDeletingEventArgs e)
        {

        }

        private FileSystemWatcher CreateFileWatcher(string path)
        {
            if (_watcher != null)
                return _watcher;

            _watcher = new FileSystemWatcher(path);
            _watcher.NotifyFilter = NotifyFilters.LastAccess | NotifyFilters.LastWrite | NotifyFilters.FileName | NotifyFilters.DirectoryName;
            _watcher.Filter = "*.png";
            _watcher.IncludeSubdirectories = false;
            _watcher.Changed += OnFileChanged;
            //watcher.Created += OnFileChanged;
            //watcher.Deleted += OnFileChanged;
            _watcher.EnableRaisingEvents = true;

            return _watcher;
        }        

        private static void OnFileChanged(object source, FileSystemEventArgs e)
        {
            var xmlPartId = Path.GetFileNameWithoutExtension(e.FullPath);

            Globals.ThisAddIn.TheWindowsFormsSynchronizationContext.Send(d =>
            {
                _addin.AddInNotifyChanged(xmlPartId);
            }, null);
        }  

        private void Application_DocumentBeforeSave(Microsoft.Office.Interop.Word.Document doc, ref bool saveAsUi, ref bool cancel)
        {
            //foreach (var ctrl in ActiveVstoDocument.Controls)
            //{
            //    if (ctrl is Microsoft.Office.Tools.Word.PictureContentControl pcc)
            //    {
            //        RemoveCommentHeader(pcc.Image);
            //    }
            //}

            // Get list of all Draw.io PictureContentControls in document
            var ids = new HashSet<string>();
            foreach (Microsoft.Office.Interop.Word.ContentControl nativeControl in doc.ContentControls)
            {
                if (nativeControl.Type == Microsoft.Office.Interop.Word.WdContentControlType.wdContentControlPicture
                    && IsDrawioTag(nativeControl.Tag))
                {
                    ids.Add(GetDrawioTagId(nativeControl.Tag));
                }
            }

            // Clean up unreferenced OfficeDrawIo data parts
            foreach (Microsoft.Office.Core.CustomXMLPart part in Application.ActiveDocument.CustomXMLParts)
            {
                if (!ids.Contains(part.Id) && part.XML != null && part.XML.TrimStart().StartsWith("<OfficeDrawIo"))
                {
                    Trace.WriteLine($"Deleting orphaned Draw.io data part id: {part.Id}");
                    part.Delete();
                }
            }
        }

        private bool ValidateDependencies()
        {
            if (!File.Exists(_settings.DrawIoExePath))
            {
                MessageBox.Show($"Draw.io Desktop not found. Please download and install it from {Properties.Settings.Default.DrawIoUrl}. If you believe it is installed, then please check the Add-In settings.",
                    Application.ActiveWindow.Caption, MessageBoxButtons.OK, MessageBoxIcon.Error);

                Process.Start(Properties.Settings.Default.DrawIoUrl);

                return false;
            }

            return true;
        }

        private static void ActionTryEnter(object lck, Action action)
        {
            bool _lockTaken = false;
            try
            {
                Monitor.TryEnter(lck, ref _lockTaken);
                if (_lockTaken)
                {
                    action?.Invoke();
                }
            }
            finally
            {
                if (_lockTaken)
                    Monitor.Exit(lck);
            }
        }

        

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += ThisAddIn_Startup;
            this.Shutdown += ThisAddIn_Shutdown;
        }
        
        #endregion
    }

    
}
