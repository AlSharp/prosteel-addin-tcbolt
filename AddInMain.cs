/*--------------------------------------------------------------------------------------+
|
|     GSF - Garrison Steel Fabricators, Pell city, AL
|     Developer - Albert Sharapov, Steel Detailer
|
+--------------------------------------------------------------------------------------*/
using System;
using Bentley.MicroStation;

//  Provides access to adapters needed to use forms and controls
//  from System.Windows.Forms in MicroStation
using BMW = Bentley.MicroStation.WinForms;

//  Provides access to classes used to make forms dockable in MicroStation
using BWW = Bentley.Windowing.WindowManager;

//  The Primary Interop Assembley (PIA) for MicroStation's COM object
//  model uses the namespace Bentley.Interop.MicroStationDGN
using BCOM = Bentley.Interop.MicroStationDGN;

//  The InteropServices namespace contains utilities to simplify using 
//  COM object model.
using BMI = Bentley.MicroStation.InteropServices;

namespace GSFBolt
{
    /// <summary>When loading an AddIn MicroStation looks for a class
    /// derived from AddIn.</summary>
    // [Bentley.MicroStation.AddInAttribute(MdlTaskID = "GSFBolt", KeyinTree = "GSFBolt.GSFBolt.commands.xml")]
    [Bentley.MicroStation.AddIn(MdlTaskID = "GSFBolt", KeyinTree = "GSFBolt.GSFBolt.commands.xml")]
    internal sealed class AddInMain : Bentley.MicroStation.AddIn
    {
        private static AddInMain s_addin = null;
        private static BCOM.Application s_comApp = null;

        internal static Bentley.MicroStation.AddIn MyAddin
        {
            get { return s_addin; }
        }

        /// <summary>Static property that the rest of the application uses to
        /// get the reference to the COM object model's main application object.</summary>
        internal static BCOM.Application ComApp
        {
            get { return s_comApp; }
        }

        /// <summary>Private constructor required for all AddIn classes derived from 
        /// Bentley.MicroStation.AddIn.</summary>
        private AddInMain
        (
        System.IntPtr mdlDesc
        )
            : base(mdlDesc)
        {
            s_addin = this;
        }

        /// <summary>Handles MDL LOAD requests after the application has been loaded.
        /// </summary>
        private void AddInMain_ReloadEvent(Bentley.MicroStation.AddIn sender, ReloadEventArgs eventArgs)
        {
           BoltControl.ShowForm (this);
        }

        private void AddInMain_UnloadedEvent(Bentley.MicroStation.AddIn sender, UnloadedEventArgs eventArgs)
        {   
        }

        protected override void OnUnloading(UnloadingEventArgs eventArgs)
        {
            BoltControl.CloseForm ();
            base.OnUnloading (eventArgs);
        }

        /// <summary>The AddIn loader creates an instance of a class 
        /// derived from Bentley.MicroStation.AddIn and invokes Run.
        /// </summary>
        protected override int Run(System.String[] commandLine)
        {
            s_comApp = BMI.Utilities.ComApp;

            //  Register reload and unload events, and show the form
            this.ReloadEvent += new ReloadEventHandler(AddInMain_ReloadEvent);
            this.UnloadedEvent += new UnloadedEventHandler(AddInMain_UnloadedEvent);

            BoltControl.ShowForm (this);

            return 0;
        }
    }   // End of GSFBolt
}   // End of Namespace
