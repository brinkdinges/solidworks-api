using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using Dna;
using Microsoft.Extensions.DependencyInjection;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swpublished;
using static Dna.FrameworkDI;

namespace AngelSix.SolidDna
{
    /// <summary>
    /// Integrates into SolidWorks as an add-in and registers for callbacks provided by SolidWorks
    /// 
    /// IMPORTANT: The class that overrides <see cref="ISwAddin"/> MUST be the same class that 
    /// contains the ComRegister and ComUnregister functions due to how SolidWorks loads add-ins
    /// </summary>
    public abstract class SolidAddIn : ISwAddin
    {
        #region Protected Members

        /// <summary>
        /// A list of assemblies to use when resolving any missing references
        /// </summary>
        protected List<AssemblyName> mReferencedAssemblies = new List<AssemblyName>();

        #endregion

        #region Public properties

        /// <summary>
        /// The title displayed for this SolidWorks Add-in
        /// </summary>
        public string SolidWorksAddInTitle { get; set; } = "AngelSix SolidDna AddIn";

        /// <summary>
        /// The description displayed for this SolidWorks Add-in
        /// </summary>
        public string SolidWorksAddInDescription { get; set; } = "All your pixels are belong to us!";

        /// <summary>
        /// A list of available plug-ins loaded once SolidWorks has connected
        /// </summary>
        public List<SolidPlugIn> PlugIns { get; set; } = new List<SolidPlugIn>();

        /// <summary>
        /// Gets the list of all known reference assemblies in this solution
        /// </summary>
        public AssemblyName[] ReferencedAssemblies => mReferencedAssemblies.ToArray();

        #endregion

        #region Public Events

        /// <summary>
        /// Called once SolidWorks has loaded our add-in and is ready.
        /// Now is a good time to create taskpanes, menu bars or anything else.
        ///  
        /// NOTE: This call will be made twice, one in the default domain and one in the AppDomain as the SolidDna plug-ins
        /// </summary>
        public event Action ConnectedToSolidWorks = () => { };

        /// <summary>
        /// Called once SolidWorks has unloaded our add-in.
        /// Now is a good time to clean up taskpanes, menu bars or anything else.
        /// 
        /// NOTE: This call will be made twice, one in the default domain and one in the AppDomain as the SolidDna plug-ins
        /// </summary>
        public event Action DisconnectedFromSolidWorks = () => { };

        #endregion

        #region Constructor

        /// <summary>
        /// Default constructor
        /// </summary>
        /// <param name="standAlone">
        ///     If true, sets the SolidWorks Application to the active instance
        ///     (if available) so the environment can be used from a stand alone application.
        /// </param>
        public SolidAddIn(bool standAlone = false)
        {
            try
            {
                // Help resolve any assembly references
                AppDomain.CurrentDomain.AssemblyResolve += CurrentDomain_AssemblyResolve;

                // Get the path to this actual add-in dll
                var assemblyFilePath = this.AssemblyFilePath();
                var assemblyPath = this.AssemblyPath();

                // Setup IoC
                IoC.Setup(assemblyFilePath, construction =>
                {
                    //  Add SolidDna-specific services
                    // --------------------------------

                    // Add reference to the add-in integration
                    // Which can then be fetched anywhere with
                    // IoC.AddIn
                    construction.Services.AddSingleton(this);

                    // Add localization manager
                    construction.AddLocalizationManager();

                    //  Configure any services this class wants to add
                    // ------------------------------------------------
                    ConfigureServices(construction);
                });

                // Log details
                Logger.LogDebugSource($"DI Setup complete");
                Logger.LogDebugSource($"Assembly File Path {assemblyFilePath}");
                Logger.LogDebugSource($"Assembly Path {assemblyPath}");

                // If we are in stand-alone mode and the SolidWorks has not yet been instantiated..
                if (standAlone && AddInIntegration.SolidWorks == null)
                    // Connect to active SolidWorks
                    AddInIntegration.ConnectToActiveSolidWorksForStandAlone();

                AddInIntegration.ActiveAddIns.Add(this);
            }
            catch (Exception ex)
            {
                // Fall-back just write a static log directly
                File.AppendAllText(Path.ChangeExtension(this.AssemblyFilePath(), "fatal.log.txt"), $"\r\nUnexpected error: {ex}");
            }
        }

        #endregion

        #region Public Abstract Methods

        /// <summary>
        /// Specific application startup code when SolidWorks is connected 
        /// and before any plug-ins or listeners are informed
        /// 
        /// NOTE: This call will not be in the same AppDomain as the SolidDna plug-ins
        /// </summary>
        /// <returns></returns>
        public abstract void ApplicationStartup();

        /// <summary>
        /// Run immediately when <see cref="ConnectToSW(object, int)"/> is called
        /// to do any pre-setup such as <see cref="PlugInIntegration.UseDetachedAppDomain"/>
        /// </summary>
        public abstract void PreConnectToSolidWorks();

        /// <summary>
        /// Run before loading plug-ins.
        /// This call should be used to add plug-ins to be loaded, via <see cref="PlugInIntegration.AddPlugIn{T}"/>
        /// </summary>
        /// <returns></returns>
        public abstract void PreLoadPlugIns();

        /// <summary>
        /// Add any dependency injection items into the DI provider that you would like to use in your application
        /// </summary>
        /// <param name="construction"></param>
        public abstract void ConfigureServices(FrameworkConstruction construction);

        #endregion

        #region SolidWorks Add-in Callbacks

        /// <summary>
        /// Used to pass a callback message onto our plug-ins
        /// </summary>
        /// <param name="arg"></param>
        public void Callback(string arg)
        {
            // Log it
            Logger.LogDebugSource($"SolidWorks Callback fired {arg}");

            PlugInIntegration.OnCallback(arg);
        }

        /// <summary>
        /// Called when SolidWorks has loaded our add-in and wants us to do our connection logic
        /// </summary>
        /// <param name="thisSw">The current SolidWorks instance</param>
        /// <param name="cookie">The current SolidWorks cookie Id</param>
        /// <returns></returns>
        public bool ConnectToSW(object thisSw, int cookie)
        {
            try
            {
                // Get the directory path to this actual add-in dll
                var assemblyPath = this.AssemblyPath();

                // Log it
                Logger.LogDebugSource($"{SolidWorksAddInTitle} Connected to SolidWorks...");

                // Log it
                Logger.LogDebugSource($"Firing PreConnectToSolidWorks...");

                // Fire event
                PreConnectToSolidWorks();

                //
                //   NOTE: Do not need to create it here, as we now create it inside PlugInIntegration.Setup in it's own AppDomain
                //         If we change back to loading directly (not in an app domain) then uncomment this 
                //
                // Store a reference to the current SolidWorks instance
                // Initialize SolidWorks (SolidDNA class)
                //SolidWorks = new SolidWorksApplication((SldWorks)ThisSW, Cookie);

                // Log it
                Logger.LogDebugSource($"Setting AddinCallbackInfo...");

                // Setup callback info
                var ok = ((SldWorks) thisSw).SetAddinCallbackInfo2(0, this, cookie);

                // Log it
                Logger.LogDebugSource($"PlugInIntegration Setup...");

                // Setup plug-in application domain
                PlugInIntegration.Setup(assemblyPath, ((SldWorks) thisSw).RevisionNumber(), cookie);

                // Log it
                Logger.LogDebugSource($"Firing PreLoadPlugIns...");

                // Any pre-load steps
                PreLoadPlugIns();

                // Log it
                Logger.LogDebugSource($"Configuring PlugIns...");

                // Perform any plug-in configuration
                PlugInIntegration.ConfigurePlugIns(assemblyPath, this);

                // Log it
                Logger.LogDebugSource($"Firing ApplicationStartup...");

                // Call the application startup function for an entry point to the application
                ApplicationStartup();

                // Log it
                Logger.LogDebugSource($"Firing ConnectedToSolidWorks...");

                // Inform listeners
                ConnectedToSolidWorks();

                // Log it
                Logger.LogDebugSource($"PlugInIntegration ConnectedToSolidWorks...");

                // And plug-in domain listeners
                PlugInIntegration.ConnectedToSolidWorks(this);

                // Return ok
                return true;
            }
            catch (Exception ex)
            {
                // Log it
                Logger.LogCriticalSource($"Unexpected error: {ex}");

                return false;
            }
        }

        /// <summary>
        /// Called when SolidWorks is about to unload our add-in and wants us to do our disconnection logic
        /// </summary>
        /// <returns></returns>
        public bool DisconnectFromSW()
        {
            // Log it
            Logger.LogDebugSource($"{SolidWorksAddInTitle} Disconnected from SolidWorks...");

            // Log it
            Logger.LogDebugSource($"Firing DisconnectedFromSolidWorks...");

            // Inform listeners
            DisconnectedFromSolidWorks();

            // And plug-in domain listeners
            PlugInIntegration.DisconnectedFromSolidWorks(this);

            // Log it
            Logger.LogDebugSource($"Tearing down...");

            // Remove it from the list. Do this before calling PlugInIntegration.Teardown
            AddInIntegration.ActiveAddIns.Remove(this);

            // Clean up plug-in app domain
            PlugInIntegration.Teardown();

            // Return ok
            return true;
        }

        #endregion

        #region Assembly Resolve Methods

        /// <summary>
        /// Adds any reference assemblies to the assemblies that get resolved when loading assemblies
        /// based on the reference type. To add all references from a project, pass in any type that is
        /// contained in the project as the reference type
        /// </summary>
        /// <typeparam name="ReferenceType">The type contained in the assembly where the references are</typeparam>
        public void AddReferenceAssemblies<ReferenceType>()
        {
            // Find all reference assemblies from the type
            var referencedAssemblies = typeof(ReferenceType).Assembly.GetReferencedAssemblies();

            // If there are any references
            if (referencedAssemblies?.Length > 0)
                // Add them
                mReferencedAssemblies.AddRange(referencedAssemblies);
        }

        /// <summary>
        /// Attempts to resolve missing assemblies based on a list of known references
        /// primarily from SolidDna and the Add-in project itself
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="args"></param>
        /// <returns></returns>
        private Assembly CurrentDomain_AssemblyResolve(object sender, ResolveEventArgs args)
        {
            // Try and find a reference assembly that matches...
            var resolvedAssembly = mReferencedAssemblies.FirstOrDefault(f => string.Equals(f.FullName, args.Name, StringComparison.InvariantCultureIgnoreCase));

            // If we didn't find any assembly
            if (resolvedAssembly == null)
                // Return null
                return null;

            // If we found a match...
            try
            {
                // Try and load the assembly
                var assembly = Assembly.Load(resolvedAssembly.Name);

                // If it loaded...
                if (assembly != null)
                    // Return it
                    return assembly;

                // Otherwise, throw file not found
                throw new FileNotFoundException();
            }
            catch
            {
                //
                // Try to load by filename - split out the filename of the full assembly name
                // and append the base path of the original assembly (i.e. look in the same directory)
                //
                // NOTE: this doesn't account for special search paths but then that never
                //       worked before either
                //
                var parts = resolvedAssembly.Name.Split(',');
                var filePath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + "\\" + parts[0].Trim() + ".dll";

                // Try and load assembly and let it throw FileNotFound if not there 
                // as it's an expected failure if not found
                return Assembly.LoadFrom(filePath);
            }
        }

        #endregion

        #region Connected to SolidWorks Event Calls

        /// <summary>
        /// When the add-in has connected to SolidWorks
        /// </summary>
        public void OnConnectedToSolidWorks()
        {
            // Log it
            Logger.LogDebugSource($"Firing ConnectedToSolidWorks event...");

            ConnectedToSolidWorks();
        }

        /// <summary>
        /// When the add-in has disconnected to SolidWorks
        /// </summary>
        public void OnDisconnectedFromSolidWorks()
        {
            // Log it
            Logger.LogDebugSource($"Firing DisconnectedFromSolidWorks event...");

            DisconnectedFromSolidWorks();
        }

        #endregion

        #region Stand alone methods

        ///// <summary>
        ///// Attempts to set the SolidWorks property to the active SolidWorks instance.
        ///// Remember to call <see cref="TearDown"/> once done.
        ///// </summary>
        ///// <returns></returns>
        //public static bool ConnectToActiveSolidWorks()
        //{
        //    // Create new blank add-in
        //    var addin = new BlankAddInIntegration();

        //    // Return if we successfully got an instance
        //    return AddInIntegration.ConnectToActiveSolidWorks();
        //}

        #endregion
    }
}
