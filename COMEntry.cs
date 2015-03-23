using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using NetOffice.OfficeApi.Enums;
using NetOffice.OutlookApi.Tools;
using NetOffice.Tools;
using OutLook = NetOffice.OutlookApi;
using Office = NetOffice.OfficeApi;

namespace TestOutlookAddin
{
    [COMAddin("Test Addin For Outlook", "", 3), CustomUI("TestOutlookAddin.RibbonUI.xml"), RegistryLocation(RegistrySaveLocation.CurrentUser)]
    [Guid("AFE67651-951D-4A42-8CAB-E9BF7E219DDF"), ProgId("TestAddinForOutlook")]
    public class COMEntry : COMAddin
    {

        NetOffice.OutlookApi.Application _outlookApplication;
        private NetOffice.OfficeApi.IRibbonUI _ribbon;

        NetOffice.OfficeApi.CommandBarButton LogonBtn;

        public COMEntry()
        {
            OnStartupComplete += Addin_OnStartupComplete;
            OnConnection += Addin_OnConnection;
            OnDisconnection += Addin_OnDisconnection;
        }

        /// <summary>
        /// Override get custom ui method, only show ribbon at start explorer
        /// </summary>
        /// <param name="RibbonID"></param>
        /// <returns></returns>
        public override string GetCustomUI(string RibbonID)
        {
            if (RibbonID != "Microsoft.Outlook.Explorer") return "";
            var ui = base.GetCustomUI(RibbonID);
            return ui;
        }

        private void Addin_OnDisconnection(ext_DisconnectMode RemoveMode, ref Array custom)
        {
            try
            {
                if (null != _outlookApplication)
                    _outlookApplication.Dispose();
            }
            catch (Exception exception)
            {
                // 处理
            }
        }

        private void Addin_OnConnection(object app, ext_ConnectMode ConnectMode, object AddInInst, ref Array custom)
        {
            try
            {
                _outlookApplication = new OutLook.Application(null, app);
                TaskPanes.Add(typeof(TaskPaneContainerControl), "侧边栏标题");
                TaskPanes[0].DockPosition = MsoCTPDockPosition.msoCTPDockPositionBottom;
                TaskPanes[0].DockPositionRestrict = MsoCTPDockPositionRestrict.msoCTPDockPositionRestrictNoChange;
                TaskPanes[0].Height = 100;
                TaskPanes[0].Visible = true;
                TaskPanes[0].Arguments = new object[] { this };
                // register add signature after new mail
                var i = _outlookApplication.Inspectors as OutLook.Inspectors;
                if (i != null)
                {
                    i.NewInspectorEvent += i_NewInspectorEvent;
                }
            }
            catch (Exception exception)
            {
                // 处理
            }
        }

        private OutLook.MailItem _newMail;

        void i_NewInspectorEvent(OutLook._Inspector Inspector)
        {
            _newMail = Inspector.CurrentItem as OutLook.MailItem;
            if (_newMail == null || !String.IsNullOrEmpty(_newMail.EntryID)) return;
            _newMail.SendEvent += newMail_SendEvent;
        }

        void newMail_SendEvent(ref bool Cancel)
        {
            _newMail.HTMLBody = _newMail.HTMLBody.Replace("</body>", "<p class=MsoNormal><span lang=EN-US><a href=\"http://www.camcard.com\">Click Me</a></span></p></body>");
        }

        private void Addin_OnStartupComplete(ref Array custom)
        {
            if (!_outlookApplication.Version.StartsWith("15.0") && !_outlookApplication.Version.StartsWith("14.0"))
            {
                try
                {
                    SetupGui();
                }
                catch (Exception exception)
                {
                    // 处理
                }
            }
        }

        private void SetupGui()
        {
            /* create commandbar */
            Office.CommandBar commandBar = _outlookApplication.ActiveExplorer().CommandBars.Add("工具栏名称", MsoBarPosition.msoBarTop, System.Type.Missing, true);
            commandBar.Visible = true;

            // add popup to commandbar
            //Office.CommandBarPopup commandBarPop = (Office.CommandBarPopup)commandBar.Controls.Add(MsoControlType.msoControlPopup, System.Type.Missing, System.Type.Missing, System.Type.Missing, true);
            //commandBarPop.Caption = CultureRes.ProductTitle;
            //commandBarPop.Tag = CultureRes.ProductTitle;

            // add a button to the popup
            LogonBtn = (Office.CommandBarButton)commandBar.Controls.Add(MsoControlType.msoControlButton, Type.Missing, Type.Missing, Type.Missing, true);
            LogonBtn.Style = MsoButtonStyle.msoButtonIconAndCaption;
            LogonBtn.Picture = PictureConverter.IconToPicture(Properties.Resources.SampleIcon2);
            LogonBtn.Mask = PictureConverter.ImageToPicture(Properties.Resources.sampleicon2Mask);
            LogonBtn.Caption = "按钮";
            //LogonBtn.ClickEvent += new NetOffice.OfficeApi.CommandBarButton_ClickEventHandler(LoginBtn_ClickEvent);
        }

        public void LoadAction(Office.IRibbonUI control)
        {
            _ribbon = control;
        }

        public string GetButtonLabel(NetOffice.OfficeApi.IRibbonControl control)
        {
            return "自定义\n";
        }

        public void ButtonAction(NetOffice.OfficeApi.IRibbonControl control)
        {
            MessageBox.Show("Hello World");
        }

        public stdole.IPictureDisp GetButtonImage(NetOffice.OfficeApi.IRibbonControl control)
        {
            return PictureConverter.IconToPictureDisp(Properties.Resources.SampleIcon2);
        }
    }
}
