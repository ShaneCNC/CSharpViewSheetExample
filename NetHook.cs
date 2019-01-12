// <copyright file="NetHook.cs" company="CNC Software, Inc.">
// Copyright (c) 2015 CNC Software, Inc.
// </copyright>
// <summary>Implements the NetHook3App interface</summary>

namespace ViewSheets
{
    using System.Text;
    using System.Windows.Forms;

    using Mastercam.App.Types;
    using Mastercam.IO;
    using Mastercam.Support;

    using ViewSheets.Localization;

    using ViewSheetsDemo.Services;

    #region Public Methods

    public class NETHOOK : Mastercam.App.NetHook3App
    {
        /// <summary> This method serves as the main entry point for the NetHook. </summary>
        ///
        /// <param name="param"> The parameter (optional). </param>
        ///
        /// <returns> A MCamReturn return type code. </returns>
        public override MCamReturn Run(int param)
        {
            const string SheetName = "MySheet";
            var result = this.RunViewSheetDemo(SheetName);

            var msg = result ? LocalizationStrings.DemoCompleted : LocalizationStrings.DemoFailed;
            MessageBox.Show(msg, LocalizationStrings.Title, MessageBoxButtons.OK, MessageBoxIcon.Information);

            return MCamReturn.NoErrors;
        }

        #endregion

        #region Private Methods

        /// <summary> Executes the view sheet demo. </summary>
        ///
        /// <param name="sheetName"> Name of the ViewSheet to create. </param>
        ///
        /// <returns> true if it succeeds, false if it fails. </returns>
        private bool RunViewSheetDemo(string sheetName)
        {
            // First we need to be sure that ViewSheets are enabled in Mastercam!
            if (!this.CheckViewSheetsEnabled())
            {
                return false;
            }

            // Now we'll create some ViewSheets
            if (!this.RunCreateViewSheetDemo(sheetName))
            {
                return false;
            }

            // We'll delete the sheet at this sheetIndex, later...
            var sheetIndex = this.RunCopyViewSheetDemo(sheetName);
            if (sheetIndex == -1)
            {
                return false;
            }

            // Which sheet is considered is considered the *main* sheet?
            if (!this.RunMainViewSheetDemo())
            {
                return false;
            }

            // Show getting and setting the *active* sheet.
            if (!this.RunActiveViewSheetDemo(sheetName))
            {
                return false;
            }

            //// Show the bookmark methods.
            if (!this.RunBookmarkViewSheetDemo(sheetName))
            {
                return false;
            }

            // Demo deleting a sheet
            if (!this.RunDeleteViewSheetDemo(sheetIndex))
            {
                return false;
            }

            // Show the current data for a sheet
            if (this.ShowViewSheetDataPackage(sheetName))
            {
                // Alter the data
                if (this.AlterViewSheetData(sheetName))
                {
                    // Show the updated data
                    this.ShowViewSheetDataPackage(sheetName);
                }
                else
                {
                    return false;
                }
            }
            else
            {
                return false;
            }

            // Show manipulating the "Save to ViewSheet" settings
            if (!this.ShowViewSheetSettingsData())
            {
                return false;
            }

            GeometryCreationService.Instance.Create();

            this.CycleActiveSheet();

            return true;
        }

        /// <summary> Cycle through the sheets making each to *active* sheet . </summary>
        private void CycleActiveSheet()
        {
            MessageBox.Show(
                LocalizationStrings.CycleActiveSheetMessage,
                LocalizationStrings.Title,
                MessageBoxButtons.OKCancel,
                MessageBoxIcon.Information);

            var sheetCount = ViewSheets.GetSheetCount();
            for (var i = 0; i < sheetCount; i++)
            {
                if (ViewSheets.SetActiveViewSheet(i))
                {
                    GraphicsManager.Repaint();
                    this.ShowLevels();
                }
            }
        }

        /// <summary> Shows the VisibleLevel data in the active ViewSheet. </summary>
        private void ShowLevels()
        {
            var sheetIndex = ViewSheets.GetActiveViewSheet();
            var name = ViewSheets.GetViewSheetName(sheetIndex);
            var data = new ViewSheets.ViewSheetData();

            if (ViewSheets.GetViewSheetData(sheetIndex, ref data))
            {
                var sb = new StringBuilder(name);
                foreach (var level in data.VisibleLevels)
                {
                    sb.AppendFormat("\nLevel {0}", level);
                }

                MessageBox.Show(sb.ToString(),
                    LocalizationStrings.Title,
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            }
            else
            {
                MessageBox.Show(LocalizationStrings.FailedToGetSheetData,
                    LocalizationStrings.Title,
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Exclamation);
            }
        }

        /// <summary> Tests to see if ViewSheets are enabled. </summary>
        ///
        /// <remarks> If there are not enabled, we will turn them on!</remarks>
        ///
        /// <returns> true if they are, false if not. </returns>
        private bool CheckViewSheetsEnabled()
        {
            if (!ViewSheets.ViewSheetsEnabled)
            {
                this.ShowMessage(LocalizationStrings.EnableViewSheets, LocalizationStrings.Title);
                ViewSheets.ViewSheetsEnabled = true;
            }

            return ViewSheets.ViewSheetsEnabled;
        }

        /// <summary> Executes the create view sheet demo operation. </summary>
        ///
        /// <param name="sheetName"> Name of the ViewSheet to create. </param>
        ///
        /// <returns> true if it succeeds, false if it fails. </returns>
        private bool RunCreateViewSheetDemo(string sheetName)
        {
            this.ShowMessage(LocalizationStrings.CreateViewSheets, LocalizationStrings.Title);

            // Now we'll create some new ViewSheets...
            // First a ViewSheet with a system generated name.
            // If ViewSheets are enabled, this really shouldn't fail.
            if (ViewSheets.CreateViewSheet())
            {
                // Now a ViewSheet with a name of our choosing.
                // If we specify an invalid name (e.g. one that's already used) this could return false.
                var result = ViewSheets.CreateViewSheet(sheetName);
                if (!result && !this.ShowMessage(LocalizationStrings.FailedCreateViewSheet, LocalizationStrings.Title, true))
                {
                    return false;
                }

                // Now we'll determine the sheetIndex of the "named" ViewSheet in the sheet list.
                var sheetIndex = ViewSheets.GetViewSheet(sheetName);
                if (sheetIndex == -1 && !this.ShowMessage("Retrieving a sheet named '" + sheetName + "' failed!", LocalizationStrings.Title, true))
                {
                    return false;
                }

                // Now we'll create another sheet with a system generated name and then rename it.
                // This will demonstrate the RenameViewSheet method.
                if (ViewSheets.CreateViewSheet())
                {
                    // New sheets are added to the end of the (zero-based) list, so we can determine it's sheetIndex this way.
                    sheetIndex = ViewSheets.GetSheetCount() - 1;
                    var oldName = ViewSheets.GetViewSheetName(sheetIndex);
                    this.ShowMessage($"Now we'll rename the ViewSheet {oldName} we just created.", LocalizationStrings.Title);

                    result = ViewSheets.RenameViewSheet(sheetIndex, sheetName + " #2");
                    if (!result && !this.ShowMessage($"Renaming {oldName} to {sheetName} #2 failed", LocalizationStrings.Title, true))
                    {
                        return false;
                    }

                    this.ShowMessage($"We've now added 3 new sheets to the list of {ViewSheets.GetSheetCount()} total sheets.", LocalizationStrings.Title);
                }
            }

            return true;
        }

        /// <summary> Executes the copy view sheet demo operation. </summary>
        ///
        /// <param name="sheetName"> Name of the ViewSheet to copy. </param>
        ///
        /// <returns> the sheetIndex of the copy, else -1 if the copy operation failed. </returns>
        private int RunCopyViewSheetDemo(string sheetName)
        {
            var sheetIndex = ViewSheets.GetViewSheet(sheetName);
            if (sheetIndex == -1 && !this.ShowMessage($"Could not find sheet {sheetName}.", LocalizationStrings.Title, true))
            {
                return -1;
            }

            sheetIndex = ViewSheets.CreateCopySheet(sheetIndex);
            if (sheetIndex == -1 && !this.ShowMessage($"Copying the sheet named {sheetName} failed.", LocalizationStrings.Title, true))
            {
                return -1;
            }

            this.ShowMessage($"Using CreateCopySheet we've made a copy of the {sheetName}.", LocalizationStrings.Title);

            return sheetIndex;
        }

        /// <summary> Executes the "active" view sheet demo operation. </summary>
        ///
        /// <param name="sheetName"> Name of the ViewSheet to set as the active sheet. </param>
        ///
        /// <returns> true if it succeeds, false if it fails. </returns>
        private bool RunActiveViewSheetDemo(string sheetName)
        {
            // We use 'sheetIndex" later on to set the *active* sheet.
            var sheetIndex = ViewSheets.GetViewSheet(sheetName);
            if (sheetIndex == -1 && !this.ShowMessage($"Could not find sheet {sheetName}.", LocalizationStrings.Title, true))
            {
                return false;
            }

            // If ViewSheets are enabled, this should not fail!
            // Get the current active sheet
            var currentSheetIndex = ViewSheets.GetActiveViewSheet();

            // Get the name of the *active* sheet
            var currentSheetName = ViewSheets.GetViewSheetName(currentSheetIndex);
            this.ShowMessage($"{currentSheetName} at sheet index {currentSheetIndex} is the current active sheet.", LocalizationStrings.Title);

            // Now set the requested sheet to be the *active* sheet.
            if (ViewSheets.SetActiveViewSheet(sheetIndex))
            {
                this.ShowMessage($"{sheetName} at sheet index {sheetIndex} is the new active sheet.", LocalizationStrings.Title);
            }

            return true;
        }

        /// <summary> Executes the "main" view sheet demo operation. </summary>
        ///
        /// <returns> true if it succeeds, false if it fails. </returns>
        private bool RunMainViewSheetDemo()
        {
            var sheetIndex = ViewSheets.GetActiveViewSheet();
            if (sheetIndex == -1 && !this.ShowMessage(LocalizationStrings.CouldNotFindSheet, LocalizationStrings.Title, true))
            {
                return false;
            }

            var sheetName = ViewSheets.GetViewSheetName(sheetIndex);
            if (string.IsNullOrWhiteSpace(sheetName) && !this.ShowMessage(LocalizationStrings.CouldNotFindSheet, LocalizationStrings.Title, true))
            {
                return false;
            }

            this.ShowMessage($"{sheetName} at sheet index {sheetIndex} is the main sheet.", LocalizationStrings.Title);

            return true;
        }

        /// <summary> Executes the view sheet bookmark demo operation. </summary>
        ///
        /// <param name="sheetName"> Name of the ViewSheet to manipulate it's bookmark. </param>
        ///
        /// <returns> true if it succeeds, false if it fails. </returns>
        private bool RunBookmarkViewSheetDemo(string sheetName)
        {
            var sheetIndex = ViewSheets.GetViewSheet(sheetName);
            if (sheetIndex == -1 && !this.ShowMessage($"{LocalizationStrings.CouldNotFindSheet} {sheetName}.", LocalizationStrings.Title, true))
            {
                return false;
            }

            var hasBookmark = ViewSheets.SheetHasBookmark(sheetIndex);
            if (hasBookmark)
            {
                this.ShowMessage($"{sheetName} has an existing bookmark", LocalizationStrings.Title, true);
            }
            else
            {
                this.ShowMessage($"{LocalizationStrings.AddingBookmark} {sheetName}", LocalizationStrings.Title);
                if (!ViewSheets.SaveBookmark(sheetIndex))
                {
                    this.ShowMessage(LocalizationStrings.FailedToSaveBookmark, LocalizationStrings.Title);
                    return false;
                }
            }

            this.ShowMessage($"Deleting a bookmark from {sheetName}", LocalizationStrings.Title);
            if (!ViewSheets.DeleteBookmark(sheetIndex))
            {
                this.ShowMessage(LocalizationStrings.FailedToDeleteBookmark, LocalizationStrings.Title);
                return false;
            }

            return true;
        }

        /// <summary> Executes the delete view sheet demo operation. </summary>
        ///
        /// <param name="sheetIndex"> The sheetIndex of the ViewSheet to delete. </param>
        ///
        /// <returns> true if it succeeds, false if it fails. </returns>
        private bool RunDeleteViewSheetDemo(int sheetIndex)
        {
            var sheetName = ViewSheets.GetViewSheetName(sheetIndex);
            if (string.IsNullOrEmpty(sheetName) && !this.ShowMessage($"Could not find a sheet at sheet index {sheetIndex}.", LocalizationStrings.Title, true))
            {
                return false;
            }

            this.ShowMessage($"We'll now delete {sheetName} at sheet index {sheetIndex}.", LocalizationStrings.Title);
            return ViewSheets.DeleteViewSheet(sheetIndex);
        }

        /// <summary> Shows the "data package" for a view sheet. </summary>
        ///
        /// <param name="sheetName"> Name of the ViewSheet to create. </param>
        /// <param name="label">     (Optional) A label to place on the output. </param>
        ///
        /// <returns> true if it succeeds, false if it fails. </returns>
        private bool ShowViewSheetDataPackage(string sheetName, string label = "")
        {
            var sheetIndex = ViewSheets.GetViewSheet(sheetName);
            if (string.IsNullOrEmpty(sheetName) && !this.ShowMessage($"Could not find a sheet named {sheetName}.", LocalizationStrings.Title, true))
            {
                return false;
            }

            // The data package attached to a ViewSheet
            var data = new ViewSheets.ViewSheetData();
            var result = ViewSheets.GetViewSheetData(sheetIndex, ref data);
            if (!result)
            {
                if (!this.ShowMessage($"An error occurred retrieving the data for sheet {sheetName} at sheet index {sheetIndex}.",
                    LocalizationStrings.Title,
                    true))
                {
                    return false;
                }
            }

            if (result)
            {
                // A worker class that display the data.
                var viewSheetDataPackages = new ViewSheetDataPackages();
                if (string.IsNullOrEmpty(label))
                {
                    viewSheetDataPackages.DumpViewSheetDataPackage(data, sheetName);
                }
                else
                {
                    viewSheetDataPackages.DumpViewSheetDataPackage(data, sheetName + "[" + label + "]");
                }
            }

            return result;
        }

        /// <summary> Alter and update the ViewSheetData package on a view sheet. </summary>
        ///
        /// <param name="sheetName"> Name of the ViewSheet to create. </param>
        ///
        /// <returns> true if it succeeds, false if it fails. </returns>
        private bool AlterViewSheetData(string sheetName)
        {
            var sheetIndex = ViewSheets.GetViewSheet(sheetName);
            if (string.IsNullOrEmpty(sheetName) && !this.ShowMessage($"Could not find a sheet named {sheetName}.",
                LocalizationStrings.Title,
                true))
            {
                return false;
            }

            // Get the data package attached to the ViewSheet
            var data = new ViewSheets.ViewSheetData();
            var result = ViewSheets.GetViewSheetData(sheetIndex, ref data);
            if (!result)
            {
                if (!this.ShowMessage($"An error occurred retrieving the data for sheet {sheetName} at sheet index {sheetIndex}.",
                    LocalizationStrings.Title,
                    true))
                {
                    return false;
                }
            }

            // Now make some changes and push those back to Mastercam...
            data.WireframeColor = 66;        // Change the wireframe color settings.
            data.ConstructionDepthZ = 0.25;  // Alter the construction depth.
            data.VisibleLevels.Add(101);     // Add a level to the level list for this sheet.

            // Update the sheet
            result = ViewSheets.SetViewSheetData(sheetIndex, data);
            if (!result)
            {
                if (!this.ShowMessage($"An error occurred updating the data for sheet {sheetName} at sheet index {sheetIndex}.",
                    LocalizationStrings.Title,
                    true))
                {
                    return false;
                }
            }

            return result;
        }

        /// <summary> Shows the "Save To ViewSheets" settings. </summary>
        ///
        /// <returns> true if it succeeds, false if it fails. </returns>
        private bool ShowViewSheetSettingsData()
        {
            // The data package with the active "Save To ViewSheets" settings.
            var settings = new ViewSheets.ViewSheetSettings();
            var result = ViewSheets.GetSaveToViewSheetSettings(ref settings);
            if (result)
            {
                this.ShowMessage(LocalizationStrings.ShowSaveToSettings, LocalizationStrings.Title);

                // A worker class that display the data.
                var viewSheetDataPackages = new ViewSheetDataPackages();
                viewSheetDataPackages.DumpViewSheetSettings(settings);

                this.ShowMessage(LocalizationStrings.ShowSettingsDialog, LocalizationStrings.Title);
                ViewSheets.ShowSheetSettings();

                result = ViewSheets.GetSaveToViewSheetSettings(ref settings);
                this.ShowMessage(LocalizationStrings.ShowChangedSettings, LocalizationStrings.Title);
                viewSheetDataPackages.DumpViewSheetSettings(settings);

                this.ShowMessage(LocalizationStrings.NowWeWillChangeSettings, LocalizationStrings.Title);

                // Turn them ALL ON
                settings = ViewSheets.SetAllSaveToViewSheetSettings(false);

                // Now set OFF just the option(s) we want do not want active...
                settings.WcsPlane = true;
                settings.Color = true;
                settings.Zdepth = true;

                // Now update these new settings into Mastercam.
                if (!ViewSheets.SetSaveToViewSheetSettings(settings))
                {
                    this.ShowMessage(LocalizationStrings.FailedToChangeSettings, LocalizationStrings.Title);
                }

                // And display the altered settings to the user.
                ViewSheets.ShowSheetSettings();
            }

            return result;
        }

        #endregion Private Methods

        /// <summary> Shows a message to the user. </summary>
        ///
        /// <param name="message">     The message. </param>
        /// <param name="title">       The title. </param>
        /// <param name="allowCancel"> (Optional) true to allow the user cancel. </param>
        ///
        /// <returns> true if it's OK to continue, else if not. </returns>
        private bool ShowMessage(string message, string title, bool allowCancel = false)
        {
            var result = MessageBox.Show(
                message,
                title,
                allowCancel ? MessageBoxButtons.OKCancel : MessageBoxButtons.OK,
                allowCancel ? MessageBoxIcon.Question : MessageBoxIcon.Information);

            return (!allowCancel || result != DialogResult.Cancel);
        }

        /// <summary> Query if 'name' is name valid. </summary>
        ///
        /// <param name="name"> The name. </param>
        ///
        /// <returns> True if name valid, false if not. </returns>
        private bool IsNameValid(string name)
        {
            var valid = ViewSheets.IsNameValid(name);

            var msg = $"Name [{name}] is{(valid ? " " : " not ")}valid {(valid ? " " : " as it is already used.")}";

            MessageBox.Show(msg,
                LocalizationStrings.Title,
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);
            return valid;
        }
    }
}