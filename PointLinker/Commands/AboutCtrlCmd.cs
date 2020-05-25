using System;
using System.Reflection;
using Inventor;
using PointLinker.PointLinker;
using PointLinker.Properties;
using PointLinker.Utilities;

namespace PointLinker.Commands
{
    class AboutCtrlCmd : AdnButtonCommandBase
    {
        public AboutCtrlCmd(Application application) : base(application)
        {
        }

        public override string DisplayName
        {
            get { return "About"; }
        }

        public override string InternalName
        {
            get { return "coolOrange.PointLinker.AboutCtrlCmd"; }
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
            get { return "coolOrange pointLinker version"; }
        }

        public override string ToolTipText
        {
            get { return "About pointLinker"; }
        }

        public override ButtonDisplayEnum ButtonDisplay
        {
            get { return ButtonDisplayEnum.kDisplayTextInLearningMode; }
        }

        public override string StandardIconName
        {
            get
            {
                return "PointLinker.resources.about.ico";
            }
        }

        public override string LargeIconName
        {
            get
            {
                return "PointLinker.resources.INV_APPs_PointLinker_32x32.ico";
            }
        }

        protected override void OnExecute(NameValueMap context)
        {
            var frmSplashAbout = new FrmSplash()
            {
                lblInfo = { Text = "Free License" },
                versionlbl = { Text = "2019" },
                buildlbl = { Text = Assembly.GetExecutingAssembly().GetName().Version.ToString() },
                BackgroundImage = Resources.pointLinker1
            };
           RegisterCommandForm(frmSplashAbout, true);
        }

        protected override void OnHelp(NameValueMap context)
        {
        }
    }
}
