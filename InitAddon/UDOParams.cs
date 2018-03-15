namespace InitAddon
{
    public class UDOParams
    {
        public SAPbobsCOM.BoYesNoEnum CanCancel { get; set; }
        public SAPbobsCOM.BoYesNoEnum CanClose { get; set; }
        public SAPbobsCOM.BoYesNoEnum CanCreateDefaultForm { get; set; }
        public SAPbobsCOM.BoYesNoEnum CanDelete { get; set; }
        public SAPbobsCOM.BoYesNoEnum CanFind { get; set; }
        public SAPbobsCOM.BoYesNoEnum CanLog { get; set; }
        public SAPbobsCOM.BoYesNoEnum CanYearTransfer { get; set; }
        public SAPbobsCOM.BoYesNoEnum ManageSeries { get; set; }

        public UDOParams(
                SAPbobsCOM.BoYesNoEnum canCancel = SAPbobsCOM.BoYesNoEnum.tYES
                , SAPbobsCOM.BoYesNoEnum canClose = SAPbobsCOM.BoYesNoEnum.tYES
                , SAPbobsCOM.BoYesNoEnum canCreateDefaultForm = SAPbobsCOM.BoYesNoEnum.tYES
                , SAPbobsCOM.BoYesNoEnum canDelete = SAPbobsCOM.BoYesNoEnum.tYES
                , SAPbobsCOM.BoYesNoEnum canFind = SAPbobsCOM.BoYesNoEnum.tYES
                , SAPbobsCOM.BoYesNoEnum canLog = SAPbobsCOM.BoYesNoEnum.tYES
                , SAPbobsCOM.BoYesNoEnum canYearTransfer = SAPbobsCOM.BoYesNoEnum.tYES
                , SAPbobsCOM.BoYesNoEnum manageSeries = SAPbobsCOM.BoYesNoEnum.tYES)
        {
            CanCancel = canCancel;
            CanClose = canClose;
            CanCreateDefaultForm = canCreateDefaultForm;
            CanDelete = canDelete;
            CanFind = canFind;
            CanLog = canLog;
            CanYearTransfer = canYearTransfer;
            ManageSeries = manageSeries;
        }
    }
}
