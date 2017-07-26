using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OutlookTools
{
    // Alteryx UI tool controls must derive from AlteryxGuiToolkit.Plugins.IPlugin.
    // Your class must then implement the five methods below:
    // GetConfigurationGui - called by Alteryx to construct a new configuration control.
    // GetEngineEntryPoint - called by Alteryx to identify the proper entry point for the Engine
    //                       implementation for this tool.
    // GetIcon - called by Alteryx to retrieve a bitmap to display in the flowchart view.
    // GetInputConnections - called by Alteryx to determine what input connections are supported.
    // GetOuputConnections - called by Alteryx to determine what output connections are supported.
    public class OutlookInputTool : AlteryxGuiToolkit.Plugins.IPlugin
    {
        // This method is called by Alteryx to construct a new configuration control.
        public AlteryxGuiToolkit.Plugins.IPluginConfiguration GetConfigurationGui()
        {
            // Return a new instance of our GUI control.
            return new OutlookInputToolGui();
        }

        // This method is called by Alteryx to identify the proper entry point for the Engine
        // implementation for this tool.  The two parameters are the name of the dll containing
        // the entry point, and the name of the class that implements INetPlugin.  If this entry
        // point is not found, Alteryx will still allow the tool to be added to a module and
        // configured, but the module will not run.
        public AlteryxGuiToolkit.Plugins.EntryPoint GetEngineEntryPoint()
        {
            return new AlteryxGuiToolkit.Plugins.EntryPoint("OutlookTools.dll", "OutlookTools.OutlookInputToolEngine", true);
        }

        // This method is called by Alteryx to retrieve a bitmap to display in the flowchart
        // view.  For performance reasons, we cache the bitmap.
        private System.Drawing.Bitmap m_icon;
        public System.Drawing.Image GetIcon()
        {
            // NOTE: If you follow this example and include the bitmap in the dll, be sure to specify
            // its "Build Action" as "Embedded Resource", or this code will not be able to locate it!
            if (m_icon == null)
            {
                // Get the assembly we are built into.
                System.IO.Stream s =
                    typeof(OutlookInputTool).Assembly.GetManifestResourceStream("OutlookTools.OutlookInputIcon.png");

                // Load the bitmap from the stream.
                m_icon = (System.Drawing.Bitmap)System.Drawing.Bitmap.FromStream(s);
                m_icon.MakeTransparent();
            }

            return m_icon;
        }

        // Each distinct input connection this tool will (potentially) allow must be given
        // a name.  This name is used by the engine to distinguish between different input streams.
        public AlteryxGuiToolkit.Plugins.Connection[] GetInputConnections()
        {
            // Since this example is an input tool and has no input connections, we will return an empty array.
            return new AlteryxGuiToolkit.Plugins.Connection[] { };
        }

        // Each distinct output connection this tool will (potentially) allow must be given
        // a name.  This name is used by the engine to distinguish between different output streams.
        public AlteryxGuiToolkit.Plugins.Connection[] GetOutputConnections()
        {
            // In this example, we only have one output connection called "Output".
            return new AlteryxGuiToolkit.Plugins.Connection[] { new AlteryxGuiToolkit.Plugins.Connection("Message Output") { Label = 'M' }, new AlteryxGuiToolkit.Plugins.Connection("Attachment Output") { Label = 'A' } };
        }
    }
}
