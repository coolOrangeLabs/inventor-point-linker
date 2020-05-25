using System;
using Inventor;
using PointLinker.PointLinker;
using PointLinker.Utilities;

namespace PointLinker.Commands
{
    class HelpCtrlCmd : AdnButtonCommandBase
    {
        public HelpCtrlCmd(Application application) : base(application)
        {
        }

        public override string DisplayName
        {
            get { return "Help"; }
        }

        public override string InternalName
        {
            get { return "coolOrange.PointLinker.HelpCtrlCmd"; }
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
            get { return "coolOrange pointLinker help"; }
        }

        public override string ToolTipText
        {
            get { return "pointLinker help"; }
        }

        public override ButtonDisplayEnum ButtonDisplay
        {
            get { return ButtonDisplayEnum.kDisplayTextInLearningMode; }
        }

        public override string StandardIconName
        {
            get
            {
                return "PointLinker.resources.Help.ico";
            }
        }

        public override string LargeIconName
        {
            get
            {
                return "PointLinker.resources.Help.ico";
            }
        }

        protected override void OnExecute(NameValueMap context)
        {
            System.Diagnostics.Process.Start("http://wiki.coolorange.com/display/POIN/");
            Terminate();
        }

        protected override void OnHelp(NameValueMap context)
        {
        }
    }
}
