// This is an independent project of an individual developer. Dear PVS-Studio, please check it.
// PVS-Studio Static Code Analyzer for C, C++ and C#: http://www.viva64.com

/* Mover
 * ExternalCommand_Work.cs
 * https://google.com
 * © Vitaclick, 2018
 *
 * This file contains the methods which are used by the 
 * command.
 */
#region Namespaces
using System;
using System.IO;
using System.Collections.Generic;
using System.Diagnostics;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System.Resources;
using System.Reflection;
using System.Drawing;
using System.Windows.Media.Imaging;
using System.Windows.Interop;
using WPF = System.Windows;
using System.Linq;
using Bushman.RevitDevTools;
using Vitaclick.Mover.Properties;
#endregion


namespace Vitaclick.Mover
{

  public sealed partial class ExternalCommand
  {

    private bool DoWork(ExternalCommandData commandData,
        ref String message, ElementSet elements)
    {

      if (null == commandData)
      {

        throw new ArgumentNullException(nameof(
            commandData));
      }

      if (null == message)
      {

        throw new ArgumentNullException(nameof(message)
            );
      }

      if (null == elements)
      {

        throw new ArgumentNullException(nameof(elements
            ));
      }

      ResourceManager res_mng = new ResourceManager(
            GetType());
      ResourceManager def_res_mng = new ResourceManager(
          typeof(Properties.Resources));

      UIApplication ui_app = commandData.Application;
      UIDocument ui_doc = ui_app?.ActiveUIDocument;
      Application app = ui_app?.Application;
      Document doc = ui_doc?.Document;

      var tr_name = res_mng.GetString("_transaction_name");

      try
      {
        using (var tr = new Transaction(doc, tr_name))
        {
          if (TransactionStatus.Started == tr.Start())
          {
            var userProfilePath = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
            string[] data = System.IO.File.ReadAllLines(userProfilePath + @"\AppData\Roaming\Autodesk\Revit\Addins\2017\data.txt");

            OpenDetachedAndSaveAsNewCentral(ui_app, data);

            return TransactionStatus.Committed == tr.Commit();
          }
        }
      }
      catch (Exception ex)
      {
        /* TODO: Handle the exception here if you need 
         * or throw the exception again if you need. */
        throw ex;
      }
      finally
      {

        res_mng.ReleaseAllResources();
        def_res_mng.ReleaseAllResources();
      }

      return false;
    }
    public void OpenDetachedAndSaveAsNewCentral(UIApplication uiapp, string[] data)
    {
      // Open non-interactive document with options

      Application app = uiapp.Application;
            var targetFolder = data[0];
            var modelPaths = data.Skip(1).ToArray();

      var toOpen = GetWSAPIModelPath(uiapp, modelPaths);

            for (int i = 0; i < toOpen.Count; i++)
            {
                Document openedDoc = OpenDetached(app, toOpen[i]);

                SaveAsOptions options = new SaveAsOptions();

                WorksharingSaveAsOptions wsOptions = new WorksharingSaveAsOptions();

                wsOptions.SaveAsCentral = true;
                options.SetWorksharingOptions(wsOptions);
                var docName = openedDoc.PathName.Substring(0, openedDoc.PathName.Length - 16);

                try
                {
                    openedDoc.SaveAs(Path.Combine(targetFolder, docName + ".rvt"), options);

                    openedDoc.Close(false);
                }
                catch (Exception ex)
                {

                    continue;
                }
                
            }
      

      //ShowInfoOnOpenedWorksharedDocument(openedDoc);

      //String serverPathRoot = uiapp.Application.GetRevitServerNetworkHosts().First();
      //FileInfo modelPath = new FileInfo(Path.Combine(@"C:\Users\Vitaclick\Desktop\", fileName));

      //ModelPath modelPath = new ServerPath(@"C:\Users\Vitaclick\Desktop\", "central.rvt");


      
    }

    /// <summary>
    /// Helper to return a local path 
    /// for a target model file.
    /// </summary>
    static List<ModelPath> GetWSAPIModelPath(UIApplication uiapp, string[] fileNames)
    {
      String serverPathRoot = uiapp.Application.GetRevitServerNetworkHosts()[3];
        var modelPaths = new List<ModelPath>();
            //FileInfo filePath = new FileInfo(Path.Combine(serverPathRoot, fileName));
            foreach (var fileName in fileNames)
            {
                ModelPath modelPath = new ServerPath(serverPathRoot, fileName);
                modelPaths.Add(modelPath);
                //ModelPath mp = ModelPathUtils.ConvertUserVisiblePathToModelPath(filePath.FullName);
            }
      return modelPaths;
    }

    static Document OpenDetached(Application app, ModelPath modelPath)
    {
      OpenOptions opt = new OpenOptions();

      opt.DetachFromCentralOption = DetachFromCentralOption.DetachAndPreserveWorksets;

      Document openedDoc = app.OpenDocumentFile(modelPath, opt);

      return openedDoc;
    }

    /// <summary>
    /// Show popup with info about worksets 
    /// and worksharing status
    /// </summary>
    static void ShowInfoOnOpenedWorksharedDocument(Document doc)
    {
      String documentName = doc.Title;
      bool isWorkshared = doc.IsWorkshared;

      FilteredWorksetCollector fwc
        = new FilteredWorksetCollector(doc);

      fwc.OfKind(WorksetKind.UserWorkset);

      int wsCount = fwc.Count<Workset>();

      TaskDialog td = new TaskDialog("Opened document info");

      td.MainInstruction = "Application has opened the document " + documentName;

      string mainContent = "Workshared: "
        + isWorkshared
        + (isWorkshared
          ? "\nassociated to: "
            + ModelPathUtils
              .ConvertModelPathToUserVisiblePath(
                doc.GetWorksharingCentralModelPath())
          : "")
        + "\nWorkset count: " + wsCount + "\n"
        + String.Join("\n",
          fwc.Select<Workset, String>(
            ws => ws.Name + " - "
              + (ws.IsOpen ? "open" : "closed")));

      td.MainContent = mainContent;

      td.Show();
    }
  }
}
