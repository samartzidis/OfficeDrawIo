using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using System.Xml.Linq;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using Gma.System.MouseKeyHook;
using OfficeDrawIo;
using OfficeDrawIoUtil;

namespace OfficeDrawIoPpt
{
    public partial class ThisAddIn
    {
        private IKeyboardMouseEvents _globalHook;

        public EventHandler SelectionChanged;
        public SynchronizationContext TheWindowsFormsSynchronizationContext { get; private set; }
        public PowerPoint.Shape SelectedShape => GetCurrentSelection().FirstOrDefault();
        public bool SuppressFileWatcherNotifications { get; private set; }

        private static ThisAddIn _addin;
        private string _userTmpFilesDir;
        private SettingsAdapter _settings;
        private SettingsForm _sf;
        private FileSystemWatcher _watcher;
        private readonly BiDictionary<Guid, int> _editMap = new BiDictionary<Guid, int>();

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            Trace.Listeners.Clear();
            Trace.Listeners.Add(new OfficeDrawIo.TraceListener());
            Trace.WriteLine("ThisAddIn_Startup()");

            _addin = this;
            _userTmpFilesDir = Path.Combine(Path.GetTempPath(), $"OfficeDrawIo.{Process.GetCurrentProcess().Id}");
            Trace.WriteLine(_userTmpFilesDir);
            _settings = new SettingsAdapter();
            _sf = new SettingsForm(_settings, () => Properties.Settings.Default.Reset());

            TheWindowsFormsSynchronizationContext = SynchronizationContext.Current ?? new WindowsFormsSynchronizationContext();

            if (!Directory.Exists(_userTmpFilesDir))
                Directory.CreateDirectory(_userTmpFilesDir);

            _globalHook = Hook.AppEvents();
            _globalHook.MouseDownExt += _globalHook_MouseDownExt;

            Application.WindowSelectionChange += Application_WindowSelectionChange;

            EnsureStartFileWatcher();
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            
        }

        private void Application_WindowSelectionChange(PowerPoint.Selection sel)
        {
            var handler = SelectionChanged;
            handler?.Invoke(this, EventArgs.Empty);
        }

        private void _globalHook_MouseDownExt(object sender, MouseEventExtArgs e)
        {
            if (e.Clicks == 2)
            {
                Globals.ThisAddIn.TheWindowsFormsSynchronizationContext.Send(d =>
                {
                    if (SelectedShape != null)
                    {
                        EditDiagramShape(SelectedShape);
                    }

                }, null);

                e.Handled = true;
            }
        }

        private IEnumerable<PowerPoint.Shape> GetCurrentSelection()
        {
            PowerPoint.ShapeRange shapeRange = null;
            try
            {
                shapeRange = Application.ActiveWindow.Selection.ShapeRange;
            }
            catch 
            {
            }

            if (shapeRange == null)
                yield return null;

            foreach (PowerPoint.Shape shape in shapeRange)
            {
                if (shape.Title.IndexOf(Util.OfficeDrawIoPayloadPrexix, StringComparison.Ordinal) == 0)
                {
                    yield return shape;
                }
                yield return null;
            }
        }

        public void AddDiagramOnDocument()
        {
            if (!ValidateDependencies())
                return;

            HousekeepEditMap();

            try
            {
                var newPngData = Helpers.LoadBinaryResource(Assembly.GetExecutingAssembly(), "Resources.new.png");
                var newPngFileName = "new.png";
                var newPngFilePath = Path.Combine(_userTmpFilesDir, newPngFileName);
                File.WriteAllBytes(newPngFilePath, newPngData);

                var shape = AddDiagramShape(newPngFilePath);

                try { File.Delete(newPngFilePath); } catch { }
            }
            catch (Exception m)
            {
                var msg = $"Failed to add Draw.io diagram. Error: {m}.";
                MessageBox.Show(msg, Application.ActiveWindow.Caption, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private PowerPoint.Shape AddDiagramShape(string newPngFilePath, RectangleF? rect = null)
        {
            var slide = (PowerPoint.Slide)Application.ActiveWindow.View.Slide;

            PowerPoint.Shape shape;
            if (rect == null)
                shape = slide.Shapes.AddPicture(newPngFilePath, Office.MsoTriState.msoFalse, Office.MsoTriState.msoCTrue, 0, 0);
            else
                shape = slide.Shapes.AddPicture(newPngFilePath, Office.MsoTriState.msoFalse, Office.MsoTriState.msoCTrue, 
                    Left: rect.Value.Left, Top: rect.Value.Top);

            shape.Title = Util.EncodePngFile(newPngFilePath);
            shape.LockAspectRatio = Microsoft.Office.Core.MsoTriState.msoTrue;

            return shape;
        }

        public void FileNotifyChanged(string id)
        {
            Trace.WriteLine($"FileNotifyChanged({id})");
            if (!Guid.TryParse(id, out var editId))
                return;

            try
            {
                var oldShape = FindShape(editId);
                if (oldShape == null)
                    return;

                var rect = new RectangleF(oldShape.Left, oldShape.Top, oldShape.Width, oldShape.Height);
                var rotation = oldShape.Rotation;

                Trace.WriteLine($"FileNotifyChanged:oldShape.Id: {oldShape.Id}");

                oldShape.Delete();

                var editFilePath = Path.Combine(_userTmpFilesDir, $"{editId}.png");
                var newShape = AddDiagramShape(editFilePath, rect);
                newShape.Rotation = rotation;
                
                Trace.WriteLine($"FileNotifyChanged:newShape.Id: {newShape.Id}");

                _editMap.AddOrUpdate(editId, newShape.Id);
            }
            catch (Exception m)
            {
                MessageBox.Show($"Something went wrong while updating the the Draw.io diagram.\r\nError details:\r\n{m}",
                            Application.ActiveWindow.Caption, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void EditDiagramShape(PowerPoint.Shape shape)
        {
            Trace.WriteLine($"EditDiagramShape: Id: {shape.Id}");

            if (!_editMap.TryGetBySecond(shape.Id, out var editFileId))
            {
                editFileId = Guid.NewGuid();
                _editMap.AddOrUpdate(editFileId, shape.Id);
            }

            var wnd = NativeWindowHelper.FindWindow($"{editFileId}.png - draw.io");
            if (wnd != IntPtr.Zero)
            {
                NativeWindowHelper.RestoreFromMinimized(wnd);
                NativeWindowHelper.SetForegroundWindow(wnd);

                return;
            }

            var editFilePath = Path.Combine(_userTmpFilesDir, $"{editFileId}.png");

            try
            {
                var pngBytes = Util.DecodePngFile(shape.Title);
                File.WriteAllBytes(editFilePath, pngBytes);

                var process = new Process();
                process.StartInfo.FileName = _settings.DrawIoExePath;
                process.StartInfo.Arguments = editFilePath;
                process.StartInfo.WindowStyle = ProcessWindowStyle.Maximized;

                process.Start();
            }
            catch (Exception m)
            {
                var msg = $"Failed to start Draw.io Desktop application for file {editFilePath}. Error: {m}.";
                MessageBox.Show(msg, Application.ActiveWindow.Caption, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void ExportDrawIoDiagram()
        {
            Trace.WriteLine("ExportDrawIoDiagram()");

            var sels = GetCurrentSelection().ToList();
            if (sels.Count == 0)
                return; // Nothing selected

            var dlg = new FolderBrowserDialog();
            if (dlg.ShowDialog() != DialogResult.OK)
                return;

            try
            {
                var index = 1;
                foreach (var shape in sels)
                {
                    var filePath = Path.Combine(dlg.SelectedPath, $"{index++}.drawio.png");
                    var pngBytes = Util.DecodePngFile(shape.Title);
                    File.WriteAllBytes(filePath, pngBytes);
                }
            }
            catch (Exception m)
            {
                MessageBox.Show($@"Failed to export Draw.io document: {m.Message}.",
                    Application.ActiveWindow.Caption, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }


        private PowerPoint.Shape FindShape(Guid editId)
        {
            if (!_editMap.TryGetByFirst(editId, out var id))
                return null;

            return FindShapeById(id);
        }

        private PowerPoint.Shape FindShapeById(int id)
        {
            foreach (PowerPoint.Slide slide in Application.ActivePresentation.Slides)
            {
                foreach (PowerPoint.Shape shape in slide.Shapes)
                {
                    if (shape.Id == id)
                        return shape;
                }
            }

            return null;
        }

        private void HousekeepEditMap()
        {
            foreach (var anchorId in _editMap.SecondKeys.ToList())
            {
                if (FindShapeById(anchorId) == null)
                {
                    _editMap.TryRemoveBySecond(anchorId, out var editFileId);

                    var editFilePath = Path.Combine(_userTmpFilesDir, $"{editFileId}.png");
                    try
                    {
                        Trace.WriteLine($"HousekeepEditMap: Deleting {editFilePath}");
                        File.Delete(editFilePath);
                    }
                    catch { }
                }
            }
        }

        public void Settings()
        {
            if (_sf.ShowDialog() == DialogResult.OK)
                _settings.Save();
        }

        public void About()
        {
            var dlg = new AboutBox();
            dlg.ShowDialog();
        }

        private static void OnFileChanged(object source, FileSystemEventArgs e)
        {
            if (_addin.SuppressFileWatcherNotifications)
                return;

            var id = Path.GetFileNameWithoutExtension(e.FullPath);

            Globals.ThisAddIn.TheWindowsFormsSynchronizationContext.Send(d =>
            {
                using (new ScopedLambda(() => _addin.SuppressFileWatcherNotifications = true, () => _addin.SuppressFileWatcherNotifications = false))
                    _addin.FileNotifyChanged(id);

            }, null);
        }

        private void EnsureStartFileWatcher()
        {
            if (_watcher != null)
                return;

            _watcher = new FileSystemWatcher(_userTmpFilesDir);
            _watcher.NotifyFilter = NotifyFilters.LastAccess | NotifyFilters.LastWrite | NotifyFilters.FileName | NotifyFilters.DirectoryName;
            _watcher.Filter = "*.png";
            _watcher.IncludeSubdirectories = false;

            _watcher.Changed += OnFileChanged;

            _watcher.EnableRaisingEvents = true;
        }

        private void EnsureStopFileWatcher()
        {
            _watcher?.Dispose();
            _watcher = null;
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

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
