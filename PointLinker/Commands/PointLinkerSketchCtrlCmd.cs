using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Windows.Forms;
using Inventor;
using PointLinker.PointLinker;
using PointLinker.Utilities;
using Application = Inventor.Application;

namespace PointLinker.Commands
{
    class PointLinkerSketchCtrlCmd : AdnButtonCommandBase
    {
        // Create a logger for use in this class
        private static log4net.ILog log = log4net.LogManager.GetLogger(MethodBase.GetCurrentMethod().DeclaringType);
        private MainForm _mainForm;
        private const int MDelaytime = 500;

        public PointLinkerSketchCtrlCmd(Application application)
            : base(application)
        {
        }

        public override string DisplayName
        {
            get { return "pointLinker"; }
        }

        public override string InternalName
        {
            get { return "coolOrange.PointLinker.PointLinkerSketchCtrlCmd"; }
        }

        public override CommandTypesEnum Classification
        {
            get { return CommandTypesEnum.kEditMaskCmdType; }
        }

        public override string ClientId
        {
            get
            {
                Type t = typeof(StandardAddInServer);
                return t.GUID.ToString("B");
            }
        }

        public override string Description
        {
            get { return "Associative Point Linker for Inventor"; }
        }

        public override string ToolTipText
        {
            get { return "Associative Point Linker for Inventor"; }
        }

        public override ButtonDisplayEnum ButtonDisplay
        {
            get { return ButtonDisplayEnum.kDisplayTextInLearningMode; }
        }

        public override string StandardIconName
        {
            get
            {
                return "PointLinker.resources.pointLinker.ico";
            }
        }

        public override string LargeIconName
        {
            get
            {
                return "PointLinker.resources.pointLinker.ico";
            }
        }

        protected override void OnExecute(NameValueMap context)
        {
            log.Info("pointLinker started...");

            if (!(Application.ActiveEditObject is Sketch) && !(Application.ActiveEditObject is Sketch3D))
            {

                MessageBox.Show(Properties.Settings.Default.notSketch, Properties.Settings.Default.messageCaption, MessageBoxButtons.OK, MessageBoxIcon.Information);
                Terminate();
                return;

            }

            RegisterCommandForm(new MainForm(){FromSketch = true, InvApp = Application}, true);

            
            log.Info("pointLinker completed.");
        }

        protected override void OnHelp(NameValueMap context)
        {
        }
    }
}
