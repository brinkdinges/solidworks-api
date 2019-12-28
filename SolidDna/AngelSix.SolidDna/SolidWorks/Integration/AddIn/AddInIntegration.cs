using Dna;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using SolidWorks.Interop.sldworks;
using static Dna.FrameworkDI;

namespace AngelSix.SolidDna
{
    /// <summary>
    /// Provides functions related to SolidDna add-ins
    /// </summary>
    public static class AddInIntegration
    {
        #region Public Properties

        /// <summary>
        /// A list of all add-ins that are currently active.
        /// </summary>
        public static List<SolidAddIn> ActiveAddIns { get; } = new List<SolidAddIn>();

        /// <summary>
        /// Represents the current SolidWorks application
        /// </summary>
        public static SolidWorksApplication SolidWorks { get; private set; }

        #endregion

        #region Get add-in with a certain type

        /// <summary>
        /// Get one of the active add-ins by its type.
        /// </summary>
        /// <param name="type">The type of the add-in that contains the new taskpane</param>
        /// <returns>Returns the only add-in if only one is active. Otherwise returns the first add-in with the requested name or null.</returns>
        public static SolidAddIn GetOnlyAddInOrAddInWithType(Type type)
        {
            // If there is only one add-in (which will happen often), we return that one.
            if (ActiveAddIns.Count == 1)
                return ActiveAddIns.First();

            // If no match is found, return the first add-in
            var addInWithSameType = ActiveAddIns.FirstOrDefault(x => x.GetType() == type);
            if (addInWithSameType == null)
                return ActiveAddIns.First();

            // If a match is found, return it.
            return addInWithSameType;
        }

        #endregion

        #region Com Registration

        /// <summary>
        /// The COM registration call to add our registry entries to the SolidWorks add-in registry
        /// </summary>
        /// <param name="t"></param>
        [ComRegisterFunction]
        private static void ComRegister(Type t)
        {
            // Create new instance of ComRegister add-in to setup DI
            var addin = new ComRegisterAddInIntegration();

            try
            {
                // Get assembly name
                var assemblyName = t.Assembly.Location;

                // Log it
                Logger.LogInformationSource($"Registering {assemblyName}");

                // Get registry key path
                var keyPath = string.Format(@"SOFTWARE\SolidWorks\AddIns\{0:b}", t.GUID);

                // Create our registry folder for the add-in
                using (var rk = Microsoft.Win32.Registry.LocalMachine.CreateSubKey(keyPath))
                {
                    // Load add-in when SolidWorks opens
                    rk.SetValue(null, 1);

                    //
                    // IMPORTANT: 
                    //
                    //   In this special case, COM register won't load the wrong AngelSix.SolidDna.dll file 
                    //   as it isn't loading multiple instances and keeping them in memory
                    //            
                    //   So loading the path of the AngelSix.SolidDna.dll file that should be in the same
                    //   folder as the add-in dll right now will work fine to get the add-in path
                    //
                    var pluginPath = typeof(PlugInIntegration).CodeBaseNormalized();

                    // Force auto-discovering plug-in during COM registration
                    PlugInIntegration.AutoDiscoverPlugins = true;

                    Logger.LogInformationSource("Configuring plugins...");

                    // Let plug-ins configure title and descriptions
                    PlugInIntegration.ConfigurePlugIns(pluginPath, addin);
                    
                    // Set SolidWorks add-in title and description
                    rk.SetValue("Title", addin.SolidWorksAddInTitle);
                    rk.SetValue("Description", addin.SolidWorksAddInDescription);

                    Logger.LogInformationSource($"COM Registration successful. '{addin.SolidWorksAddInTitle}' : '{addin.SolidWorksAddInDescription}'");
                }
            }
            catch (Exception ex)
            {
                Logger.LogCriticalSource($"COM Registration error. {ex}");
                throw;
            }
        }

        /// <summary>
        /// The COM unregister call to remove our custom entries we added in the COM register function
        /// </summary>
        /// <param name="t"></param>
        [ComUnregisterFunction]
        private static void ComUnregister(Type t)
        {
            // Get registry key path
            var keyPath = string.Format(@"SOFTWARE\SolidWorks\AddIns\{0:b}", t.GUID);

            // Remove our registry entry
            Microsoft.Win32.Registry.LocalMachine.DeleteSubKeyTree(keyPath);
        }

        #endregion

        #region Connect to SolidWorks

        /// <summary>
        /// Attempts to set the SolidWorks property to the active SolidWorks instance
        /// Remember to call <see cref="TearDown"/> once done.
        /// </summary>
        /// <returns></returns>
        public static bool ConnectToActiveSolidWorksForStandAlone()
        {
            try
            {
                // Try and get the active SolidWorks instance
                SolidWorks = new SolidWorksApplication((SldWorks) Marshal.GetActiveObject("SldWorks.Application"), 0);

                // Log it
                Logger.LogDebugSource($"Acquired active instance SolidWorks in Stand-Alone mode");

                // Return if successful
                return SolidWorks != null;
            }
            // If we failed to get active instance...
            catch (COMException)
            {
                // Log it
                Logger.LogDebugSource("Failed to get active instance of SolidWorks in Stand-Alone mode");

                // Return failure
                return false;
            }
        }

        /// <summary>
        /// Attempts to set the SolidWorks property to the active SolidWorks instance
        /// </summary>
        /// <param name="version"></param>
        /// <param name="cookie"></param>
        /// <returns></returns>
        public static void ConnectToActiveSolidWorks(string version, int cookie)
        {
            try
            {
                // Get the version number (such as 25 for 2016)
                var postFix = "";
                if (version != null && version.Contains("."))
                    postFix = "." + version.Substring(0, version.IndexOf('.'));
                var solidWorksProgId = "SldWorks.Application" + postFix;

                SolidWorks = new SolidWorksApplication((SldWorks) Activator.CreateInstance(Type.GetTypeFromProgID(solidWorksProgId)), cookie);
            }
            catch (Exception e)
            {
                Logger.LogDebugSource("Failed to get active instance of SolidWorks in add-in mode", exception: e);
            }
        }

        #endregion

        #region Tear Down

        /// <summary>
        /// Cleans up the SolidWorks instance
        /// </summary>
        public static void TearDown()
        {
            if (ActiveAddIns.Count != 0)
                return;
            
            // If we have an reference...
            if (SolidWorks != null)
            {
                // Log it
                Logger.LogDebugSource($"Disposing SolidWorks COM reference...");

                // Dispose SolidWorks COM
                SolidWorks?.Dispose();
            }

            // Set to null
            SolidWorks = null;
        }

        #endregion
    }
}
