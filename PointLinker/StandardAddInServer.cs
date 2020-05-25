//''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
// Copyright (c) Autodesk, Inc. All rights reserved 
// Written by Xiaodong Liang 2011 - ADN/Developer Technical Services
//
// This software is provided as is, without any warranty that 
// it will work. You choose to use this tool at your own risk.
// Neither Autodesk nor the author Xiaodong Liang can be taken 
// as responsible for any damage this tool can cause to your data. 
// Please always make a back up of your data prior to use this tool,
// as it will modify the documents involved in the feature transformation.
//'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


//INSTANT C# NOTE: Formerly VB project-level imports:

using System;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using Microsoft.Win32;
using PointLinker.Commands;
using PointLinker.Utilities;

namespace PointLinker
{
	namespace PointLinker
	{
		[ProgId("PointLinker.StandardAddInServer"), Guid("0ab6a38a-1619-43e5-a505-ec7497bd2d4b"),ComVisible(true)]
		public class StandardAddInServer : Inventor.ApplicationAddInServer
		{
            // Create a logger for use in this class
            private static log4net.ILog log = log4net.LogManager.GetLogger(MethodBase.GetCurrentMethod().DeclaringType);

            public StandardAddInServer()
            {
                Assembly thisAssembly = Assembly.GetExecutingAssembly();
                var fi = new FileInfo(thisAssembly.Location + ".log4net");
                log4net.GlobalContext.Properties["LogFileName"] = fi.DirectoryName + "\\Log\\pointLinker";
                log4net.Config.XmlConfigurator.Configure(fi);
                log = log4net.LogManager.GetLogger(MethodBase.GetCurrentMethod().DeclaringType);
            }
			// Inventor application object.
			private Inventor.Application m_inventorApplication;


#region ApplicationAddInServer Members

			public void Activate(Inventor.ApplicationAddInSite addInSiteObject, bool firstTime)
			{
                log.Debug("Attempting to Load pointLinker addin");
                Type addinType = this.GetType();
                AdnInventorUtilities.Initialize(m_inventorApplication, addinType);
				// Initialize AddIn members.
				m_inventorApplication = addInSiteObject.Application;
                AdnCommand.AddCommand(new PointLinkerCtrlCmd(m_inventorApplication));
                AdnCommand.AddCommand(new PointLinkerSketchCtrlCmd(m_inventorApplication));
                AdnCommand.AddCommand(new AboutCtrlCmd(m_inventorApplication));
                AdnCommand.AddCommand(new HelpCtrlCmd(m_inventorApplication));

                // Only after all commands have been added,
                // load Ribbon UI from customized xml file.
                // Make sure "InternalName" of above commands is matching 
                // "internalName" tag described in xml of corresponding command.
                AdnRibbonBuilder.CreateRibbon(m_inventorApplication,addinType,
                   "PointLinker.resources.ribbons.xml");

                log.Debug("pointLinker loaded successfully");

			}


			public void Deactivate()
			{
                
				// Release objects.         
				Marshal.ReleaseComObject(m_inventorApplication);
				m_inventorApplication = null;


			}

			public object Automation
			{

				// This property is provided to allow the AddIn to expose an API 
				// of its own to other programs. Typically, this  would be done by
				// implementing the AddIn's API interface in a class and returning 
				// that class object through this property.

				get
				{
					return null;
				}

			}

			public void ExecuteCommand(int commandID)
			{

				// Note:this method is now obsolete, you should use the 
				// ControlDefinition functionality for implementing commands.

			}

#endregion


			// This property uses reflection to get the value for the GuidAttribute attached to the class.
//ORIGINAL LINE: Public Shared ReadOnly Property AddInGuid(ByVal t As Type) As String
//INSTANT C# NOTE: C# does not support parameterized properties - the following property has been rewritten as a function:
			public static string AddInGuid(Type t)
			{
				string tempAddInGuid = null;
				string guid = "";
				try
				{
					object[] customAttributes = t.GetCustomAttributes(typeof(GuidAttribute), false);
					GuidAttribute guidAttribute = (GuidAttribute)(customAttributes[0]);
					guid = "{" + guidAttribute.Value.ToString() + "}";
				}
				finally
				{
					tempAddInGuid = guid;
				}
				return tempAddInGuid;
			}



		}

	}


}