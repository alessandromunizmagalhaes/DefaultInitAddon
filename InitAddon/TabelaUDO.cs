using SAPbobsCOM;
using System.Collections.Generic;

namespace InitAddon
{
    public class TabelaUDO : Tabela
    {
        public TabelaUDO(string nome, string descricao, BoUTBTableType tipo, List<Coluna> colunas, UDOParams udoParams) : base(nome, descricao, tipo, colunas)
        {
            if (tipo == BoUTBTableType.bott_NoObject || tipo == BoUTBTableType.bott_NoObjectAutoIncrement)
                throw new CustomException($"Erro ao instanciar tabela UDO. O tipo {tipo} não pode ser utilizado em tabelas UDO.");

            CanCancel = udoParams.CanCancel;
            CanClose = udoParams.CanClose;
            CanCreateDefaultForm = udoParams.CanCreateDefaultForm;
            CanDelete = udoParams.CanDelete;
            CanFind = udoParams.CanFind;
            CanLog = udoParams.CanLog;
            CanYearTransfer = udoParams.CanYearTransfer;
            ManageSeries = udoParams.ManageSeries;
        }

        public SAPbobsCOM.BoYesNoEnum CanCancel { get; set; }
        public SAPbobsCOM.BoYesNoEnum CanClose { get; set; }
        public SAPbobsCOM.BoYesNoEnum CanCreateDefaultForm { get; set; }
        public SAPbobsCOM.BoYesNoEnum CanDelete { get; set; }
        public SAPbobsCOM.BoYesNoEnum CanFind { get; set; }
        public SAPbobsCOM.BoYesNoEnum CanLog { get; set; }
        public SAPbobsCOM.BoYesNoEnum CanYearTransfer { get; set; }
        public SAPbobsCOM.BoYesNoEnum ManageSeries { get; set; }
        public SAPbobsCOM.BoUDOObjType ObjectType
        {
            get
            {
                if (Tipo == BoUTBTableType.bott_Document || Tipo == BoUTBTableType.bott_DocumentLines)
                {
                    return BoUDOObjType.boud_Document;
                }
                else if (Tipo == BoUTBTableType.bott_MasterData || Tipo == BoUTBTableType.bott_MasterDataLines)
                {
                    return BoUDOObjType.boud_MasterData;
                }
                else
                {
                    return BoUDOObjType.boud_MasterData;
                }
            }
        }
    }
}
