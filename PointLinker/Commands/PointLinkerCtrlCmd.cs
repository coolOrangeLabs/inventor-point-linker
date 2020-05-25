using System;
using System.Reflection;
using Inventor;
using PointLinker.PointLinker;
using PointLinker.Utilities;

namespace PointLinker.Commands
{
    class PointLinkerCtrlCmd : AdnButtonCommandBase
    {
        // Create a logger for use in this class
        private static log4net.ILog log = log4net.LogManager.GetLogger(MethodBase.GetCurrentMethod().DeclaringType);
        private MainForm _mainForm;
        private const int MDelaytime = 500;

        public PointLinkerCtrlCmd(Application application)
            : base(application)
        {
        }

        public override string DisplayName
        {
            get { return "pointLinker"; }
        }

        public override string InternalName
        {
            get { return "coolOrange.PointLinker.PointLinkerCtrlCmd"; }
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
            log.Info("pointLinker Command started...");

            RegisterCommandForm(new MainForm(){FromSketch = false, InvApp = Application}, true);

            log.Info("pointLinker Command completed.");
        }

        protected override void OnHelp(NameValueMap context)
        {
        }


    }
}
