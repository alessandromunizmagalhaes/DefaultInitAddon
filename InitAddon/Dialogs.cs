namespace InitAddon
{
    public static class Dialogs
    {
        public static SAPbouiCOM.Application SBOApplication;

        public static void Info(string msg)
        {
            BarMessage(msg, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
        }

        public static void PopupInfo(string msg)
        {
            Info(msg);
            SBOApplication.MessageBox(msg);
        }

        public static void Success(string msg)
        {
            BarMessage(msg, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
        }

        public static void PopupSuccess(string msg)
        {
            Success(msg);
            SBOApplication.MessageBox(msg);
        }

        public static void Error(string msg)
        {
            BarMessage(msg, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
        }

        public static void PopupError(string msg)
        {
            Error(msg);
            SBOApplication.MessageBox(msg);
        }

        public static bool Confirm(string msg)
        {
            return SBOApplication.MessageBox(msg, 1, "Sim", "Não") == 1;
        }

        private static void BarMessage(string msg, SAPbouiCOM.BoStatusBarMessageType type)
        {
            SBOApplication.StatusBar.SetText(msg, SAPbouiCOM.BoMessageTime.bmt_Short, type);
        }

        public static void RecebeSBOApplication(SAPbouiCOM.Application application)
        {
            SBOApplication = application;
        }
    }
}
