// <copyright file="ViewSheetDataPackages.cs" company="CNC Software, Inc.">
// Copyright (c) CNC Software, Inc.. All rights reserved.
// </copyright>

namespace ViewSheets
{
    using System.Windows.Forms;

    using ViewSheets.Localization;

    /// <summary> A class the displays the (2) ViewSheet data packages. </summary>
    public class ViewSheetDataPackages
    {
        #region Public methods

        /// <summary> Format a matrix. </summary>
        ///
        /// <param name="matrix"> The matrix to format. </param>
        ///
        /// <returns> The formatted matrix. </returns>
        public string FormatMatrix(Mastercam.Math.Matrix3D matrix)
        {
            var sb = new System.Text.StringBuilder();
            sb.AppendLine();
            sb.AppendFormat("{0:F4} : {1:F4} : {2:F4}", matrix.Row1.x, matrix.Row1.y, matrix.Row1.z);
            sb.AppendLine();
            sb.AppendFormat("{0:F4} : {1:F4} : {2:F4}", matrix.Row2.x, matrix.Row2.y, matrix.Row2.z);
            sb.AppendLine();
            sb.AppendFormat("{0:F4} : {1:F4} : {2:F4}", matrix.Row3.x, matrix.Row3.y, matrix.Row3.z);
            return sb.ToString();
        }

        /// <summary> Dumps the view sheet data package for review. </summary>
        ///
        /// <param name="data">    The data for the sheet. </param>
        /// <param name="label">   (Optional) A label to prefix to the report. </param>
        /// <param name="display"> (Optional) true to display the data on screen. </param>
        ///
        /// <returns> A string containing all of the data. </returns>
        public string DumpViewSheetDataPackage(Mastercam.Support.ViewSheets.ViewSheetData data, string label = "", bool display = true)
        {
            var sb = new System.Text.StringBuilder();
            if (!string.IsNullOrEmpty(label))
            {
                sb.AppendLine(label);
            }

            sb.AppendFormat("\nConstructionMode3D = {0}", data.ConstructionMode3D);
            sb.AppendFormat("\nWireframeColor = {0}", data.WireframeColor);
            sb.AppendFormat("\nSurfaceColor = {0}", data.SurfaceColor);
            sb.AppendFormat("\nSolidColor = {0}", data.SolidColor);

            sb.AppendFormat("\nConstructionPlaneMatrix => {0}", this.FormatMatrix(data.ConstructionPlaneMatrix));
            sb.AppendFormat("\nToolPlaneMatrix => {0}", this.FormatMatrix(data.ToolPlaneMatrix));
            sb.AppendFormat("\nGraphicsPlaneMatrix => {0}", this.FormatMatrix(data.GraphicsPlaneMatrix));
            sb.AppendFormat("\nWcsPlaneMatrix => {0}", this.FormatMatrix(data.WcsPlaneMatrix));

            sb.AppendFormat("\nViewportCenter = {0}", data.ViewportCenter);
            sb.AppendFormat("\nConstructionDepthZ = {0}", data.ConstructionDepthZ);
            sb.AppendFormat("\nLineWidth = {0}", data.LineWidth);

            sb.AppendFormat("\nSurfaceDensity = {0}", data.SurfaceDensity);
            sb.AppendFormat("\nActiveLevel = {0}", data.ActiveLevel);
            sb.AppendFormat("\nConstructionPlaneID = {0}", data.ConstructionPlaneID);
            sb.AppendFormat("\nTooltPlaneID = {0}", data.ToolPlaneID);
            sb.AppendFormat("\nGraphicsPlaneID = {0}", data.GraphicsPlaneID);
            sb.AppendFormat("\nWcsPlakneID = {0}", data.ToolPlaneID);

            var count = data.VisibleLevels.Count;
            sb.AppendFormat("\nVisibleLevels has [{0}] {1}  ->", count, (count == 1) ? "entry" : "entries");
            for (var i = 0; i < data.VisibleLevels.Count; i++)
            {
                sb.AppendFormat("\n{0}> {1}", i + 1, data.VisibleLevels[i]);
            }

            var dataDump = sb.ToString();
            if (display)
            {
                this.ShowData(dataDump);
            }

            return dataDump;
        }

        #endregion Public methods

        /// <summary> Dumps the active view sheet settings for review. </summary>
        ///
        /// <param name="settings"> The data for the sheet. </param>
        /// <param name="label">    (Optional) A label to prefix to the report. </param>
        /// <param name="display">  (Optional) true to display the data on screen. </param>
        ///
        /// <returns> A string containing all of the data. </returns>
        public string DumpViewSheetSettings(Mastercam.Support.ViewSheets.ViewSheetSettings settings, string label = "", bool display = true)
        {
            var sb = new System.Text.StringBuilder();
            if (!string.IsNullOrEmpty(label))
            {
                sb.AppendLine(label);
            }

            sb.AppendFormat("\nConstructionMode (3D) = {0}", settings.ConstructionMode);
            sb.AppendFormat("\nGraphicsView = {0}", settings.GraphicsView);
            sb.AppendFormat("\nPlanes = {0}", settings.Planes);
            sb.AppendFormat("\nToolplanes = {0}", settings.ToolPlanes);
            sb.AppendFormat("\nConstructionPlanes= {0}", settings.ConstructionPlanes);
            sb.AppendFormat("\nZdepth = {0}", settings.Zdepth);
            sb.AppendFormat("\nWcs = {0}", settings.WcsPlane);
            sb.AppendFormat("\nColor = {0}", settings.Color);
            sb.AppendFormat("\nActivelevel = {0}", settings.ActiveLevel);
            sb.AppendFormat("\nLinestyle= {0}", settings.LineStyle);
            sb.AppendFormat("\nPointstyle = {0}", settings.PointStyle);
            sb.AppendFormat("\nLineWidth = {0}", settings.LineWidth);
            sb.AppendFormat("\nSurfaceDensity = {0}", settings.SurfaceDensity);
            sb.AppendFormat("\nAutoRestore = {0}", settings.AutoRestore);

            var settingsDump = sb.ToString();
            if (display)
            {
                this.ShowData(settingsDump);
            }

            return settingsDump;
        }

        #region Private methods

        /// <summary> Shows the data on screen in a messagebox. </summary>
        ///
        /// <param name="data"> The data to be displayed. </param>
        private void ShowData(string data) => MessageBox.Show(data, LocalizationStrings.Title, MessageBoxButtons.OK, MessageBoxIcon.Information);

        #endregion Private methods
    }
}