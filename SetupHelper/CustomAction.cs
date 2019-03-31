using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Text;
using System.Windows.Forms;
using Microsoft.Deployment.WindowsInstaller;

namespace SetupHelper
{
    public class CustomActions
    {     
        [CustomAction]
        public static ActionResult ExtractArchive(Session session)
        {
            try
            {
                session.Log("Begin ExtractArchive");

                var archivePath = session.CustomActionData["FILEPATH"];
                var extractFolder = session.CustomActionData["EXTRACTFOLDER"];

                session.Log($"FILEPATH = {archivePath}");
                session.Log($"EXTRACTFOLDER = {extractFolder}");

                if (!File.Exists(archivePath))
                    session.Log($"{archivePath} does not exist!");

                if (!Directory.Exists(extractFolder))
                    Directory.CreateDirectory(extractFolder);

                ZipFile.ExtractToDirectory(archivePath, extractFolder);

                return ActionResult.Success;
            }
            catch (Exception m)
            {
                session.Log(m.ToString());

                MessageBox.Show(m.ToString(), "OfficeDrawIoSetup.SetupHelper.ExtractArchive", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                session.Log("End ExtractArchive");
            }

            return ActionResult.Failure;  
        }

        [CustomAction]
        public static ActionResult DeleteFolder(Session session)
        {
            try
            {
                session.Log("Begin DeleteFolder");

                var path = session.CustomActionData["PATH"];
                session.Log($"PATH = {path}");

                if (Directory.Exists(path))
                    Directory.Delete(path, true);

                return ActionResult.Success;
            }
            catch (Exception m)
            {
                session.Log(m.ToString());

                MessageBox.Show(m.ToString(), "OfficeDrawIoSetup.SetupHelper.DeleteFolder", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                session.Log("End DeleteFolder");
            }

            return ActionResult.Failure;
        }
    }
}
