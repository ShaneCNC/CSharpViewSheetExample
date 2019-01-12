// <copyright file="GeometryCreationService.cs" company="CNC Software, Inc.">
// Copyright (c) CNC Software, Inc.. All rights reserved.
// </copyright>

namespace ViewSheetsDemo.Services
{
    using Mastercam.BasicGeometry;
    using Mastercam.Curves;
    using Mastercam.IO;
    using Mastercam.IO.Types;
    using Mastercam.Math;
    using Mastercam.Support;

    public sealed class GeometryCreationService
    {
        #region Private Fields

        /// <summary> The synchronise root. </summary>
        private static readonly object SyncRoot = new object();

        /// <summary> The instance. </summary>
        private static volatile GeometryCreationService instance;

        #endregion

        #region Construction

        /// <summary>
        /// Initializes a new instance of the <see cref="GeometryCreationService"/> class.
        /// </summary>
        private GeometryCreationService()
        {
        }

        #endregion

        #region Public Properties

        public static GeometryCreationService Instance
        {
            get
            {
                if (instance == null)
                {
                    lock (SyncRoot)
                    {
                        if (instance == null)
                        {
                            instance = new GeometryCreationService();
                        }
                    }
                }

                return instance;
            }
        }

        #endregion

        #region Public Methods

        /// <summary> Creates this object. </summary>
        public void Create()
        {
            var level = LevelsManager.GetMainLevel();

            this.CreateLine1();
            this.CreateLine2();
            this.CreateLine3();

            LevelsManager.SetMainLevel(101);

            this.CreateArc1();
            this.CreateArc2();
            this.CreateArc3();
            this.CreateArc3();

            LevelsManager.SetMainLevel(level);
        }

        #endregion

        #region Private Methods

        /// <summary> Creates an point using the 1st PointGeometry constructor. </summary>
        private void CreatePoint1()
        {
            var pt = new PointGeometry();
            pt.Commit();
        }

        /// <summary> Creates an point using the 2nd PointGeometry constructor. </summary>
        private void CreatePoint2()
        {
            var pt = new PointGeometry(new Point3D(-1.0, 2.0, 0.0));
            pt.Commit();
        }

        /// <summary> Creates an point using the 3rd PointGeometry constructor. </summary>
        private void CreatePoint3()
        {
            var point = new PointGeometry(-3.0, 2.0, 0.0);
            point.Commit();
        }

        /// <summary> Creates an line using the 1st LineGeometry constructor. </summary>
        private void CreateLine1()
        {
            var line = new LineGeometry();
            line.Commit();
        }

        /// <summary> Creates an line using the 2nd LineGeometry constructor. </summary>
        private void CreateLine2()
        {
            var data = new Line3D(new Point3D(), new Point3D(1.0, 2.0, 0.0));
            var line = new LineGeometry(data);
            line.Commit();
        }

        /// <summary> Creates an line using the 3rd LineGeometry constructor. </summary>
        private void CreateLine3()
        {
            var line = new LineGeometry(new Point3D(), new Point3D(3.0, 2.0, 0.0));
            line.Commit();
        }

        /// <summary> Creates an arc using the 1st ArcGeometry constructor. </summary>
        private void CreateArc1()
        {
            var arc = new ArcGeometry();
            arc.Commit();
        }

        /// <summary> Creates an arc using the 2nd ArcGeometry constructor. </summary>
        private void CreateArc2()
        {
            var arc = new ArcGeometry(new Arc3D(new Point3D(), 1.0, 0.0, 180.0));
            arc.Commit();
        }

        /// <summary> Creates an arc using the 3rd ArcGeometry constructor. </summary>
        private void CreateArc3()
        {
            // Here the view is specified "by number".
            // It is suggested that you prefer to use the ArcGeometry constructor that takes a MCView object.
            var arc = new ArcGeometry(1, new Point3D(), 3.0, 0.0, 180.0);
            arc.Commit();
        }

        /// <summary> Creates an arc using the 4th ArcGeometry constructor. </summary>
        private void CreateArc4()
        {
            // Here the view to create the arc in is specified.
            // This is the preferred method of creating arcs that are not to be placed in the 
            // current Construction Plane.
            // Get the View that this arc is to be created in.
            var view = SearchManager.GetSystemView(SystemPlaneType.Right);
            var arc = new ArcGeometry(view, new Point3D(), -3.0, 0.0, 180.0);
            arc.Commit();
        }

        #endregion       
    }
}