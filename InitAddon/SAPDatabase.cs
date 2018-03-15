using System;

namespace InitAddon
{
    public static class SAPDatabase
    {
        public static SAPbobsCOM.Company oCompany;

        #region :: Criação de Tabelas e campos

        public static void CriarTabela(Tabela tabela)
        {
            if (!ExisteTabela(tabela.NomeSemArroba))
            {
                CriarUserTable(tabela);

                bool tabela_is_UDO = tabela is TabelaUDO;
                TabelaUDO tabelaUDO = tabela_is_UDO ? (TabelaUDO)tabela : null;

                if (tabela_is_UDO)
                {
                    CriarTabelaComoUDO(tabelaUDO);
                }

                int i = 0;
                foreach (var coluna in tabela.Colunas)
                {
                    if (!ExisteColuna(tabela.NomeComArroba, coluna.Nome))
                    {
                        CriarColuna(tabela.NomeComArroba, coluna);

                        if (tabela_is_UDO)
                        {
                            // tem que fazer essa gambiarra horrível, porque o primeiro elemento a ser colocado como UDO,
                            // tem que obrigatóriamente ser o campo CODE.
                            // como eu não passo ele como um campo que eu quero usar na Lista de Colunas
                            // sempre que for o primeiro, passa o Code.
                            // horroroso mas é o que tem pra hoje.
                            if (i == 0)
                            {
                                CriarColunaComoUDO(tabelaUDO, new ColunaVarchar("Code", "Código", false));
                            }

                            CriarColunaComoUDO(tabelaUDO, coluna);
                        }
                    }
                    i++;
                }
            }
        }

        private static void CriarColunaComoUDO(TabelaUDO tabela, Coluna coluna)
        {
            GC.Collect();
            SAPbobsCOM.UserObjectsMD objUserObjectMD = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD);
            if (objUserObjectMD.GetByKey(tabela.NomeSemArroba))
            {
                objUserObjectMD.FormColumns.FormColumnAlias = coluna.Nome;
                objUserObjectMD.FormColumns.FormColumnDescription = coluna.Descricao;
                objUserObjectMD.FormColumns.Editable = SAPbobsCOM.BoYesNoEnum.tYES;
                objUserObjectMD.FormColumns.Add();

                if (objUserObjectMD.Update() != 0)
                {
                    throw new CustomException($"Erro ao tentar criar o campo {coluna.Nome} na tabela {tabela.NomeSemArroba} como UDO.\nErro: {oCompany.GetLastErrorDescription()}");
                }
            }

            System.Runtime.InteropServices.Marshal.ReleaseComObject(objUserObjectMD);
            objUserObjectMD = null;
            GC.Collect();
        }

        private static void CriarTabelaComoUDO(TabelaUDO tabela)
        {
            GC.Collect();
            SAPbobsCOM.UserObjectsMD objUserObjectMD = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD);

            objUserObjectMD.TableName = tabela.NomeSemArroba;
            objUserObjectMD.Name = tabela.NomeSemArroba;
            objUserObjectMD.Code = tabela.NomeSemArroba;

            objUserObjectMD.CanCancel = tabela.CanCancel;
            objUserObjectMD.CanClose = tabela.CanClose;
            objUserObjectMD.CanCreateDefaultForm = tabela.CanCreateDefaultForm;
            objUserObjectMD.CanDelete = tabela.CanDelete;
            objUserObjectMD.CanFind = tabela.CanFind;
            objUserObjectMD.CanLog = tabela.CanLog;
            objUserObjectMD.CanYearTransfer = tabela.CanYearTransfer;
            objUserObjectMD.ManageSeries = tabela.ManageSeries;
            objUserObjectMD.ObjectType = tabela.ObjectType;

            if (objUserObjectMD.Add() != 0)
            {
                throw new CustomException($"Erro ao tentar criar a tabela {tabela.NomeSemArroba} como UDO.\nErro: {oCompany.GetLastErrorDescription()}");
            }

            System.Runtime.InteropServices.Marshal.ReleaseComObject(objUserObjectMD);
            objUserObjectMD = null;
            GC.Collect();
        }

        public static void CriarColuna(string nome_tabela, Coluna coluna)
        {
            CriarUserField(nome_tabela, coluna);

            if (coluna.Obrigatorio)
            {
                SetaCampoObrigatorio(nome_tabela, coluna.Nome);
            }

            foreach (var valor_valido in coluna.ValoresValidos)
            {
                AdicionaValorValido(nome_tabela, coluna.Nome, valor_valido.Valor, valor_valido.Descricao);
            }

            if (!String.IsNullOrEmpty(coluna.ValorPadrao))
            {
                SetaValorPadrao(nome_tabela, coluna.Nome, coluna.ValorPadrao);
            }
        }

        public static void CriarUserTable(Tabela tabela)
        {
            GC.Collect();
            SAPbobsCOM.UserTablesMD oUserTablesMD = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserTables);

            oUserTablesMD.TableName = tabela.NomeSemArroba;
            oUserTablesMD.TableDescription = tabela.Descricao;
            oUserTablesMD.TableType = tabela.Tipo;

            var error_code = oUserTablesMD.Add();
            if (error_code != 0)
            {
                throw new CustomException($"Erro ao tentar criar a tabela {tabela.NomeComArroba}.\nErro: {oCompany.GetLastErrorDescription()}");
            }

            oUserTablesMD = null;
            GC.Collect(); // Release the handle to the table
        }

        public static void CriarUserField(string nome_tabela, Coluna coluna)
        {
            GC.Collect();
            SAPbobsCOM.UserFieldsMD objUserFieldsMD = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields);
            objUserFieldsMD.TableName = nome_tabela;
            objUserFieldsMD.Name = coluna.Nome;
            objUserFieldsMD.Description = coluna.Descricao;

            if (coluna is ColunaVarchar)
            {
                objUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha;
            }
            else if (coluna is ColunaText)
            {
                objUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Memo;
            }
            else if (coluna is ColunaDate)
            {
                objUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Date;
            }
            else if (coluna is ColunaInt)
            {
                objUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Numeric;
            }
            else if (coluna is ColunaTime)
            {
                objUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Date;
                objUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_Time;
            }
            else if (coluna is ColunaPercent)
            {
                objUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Float;
                objUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_Percentage;
            }
            else if (coluna is ColunaSum)
            {
                objUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Float;
                objUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_Sum;
            }
            else if (coluna is ColunaQuantity)
            {
                objUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Float;
                objUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_Percentage;
            }
            else if (coluna is ColunaPrice)
            {
                objUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Float;
                objUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_Price;
            }
            else if (coluna is ColunaLink)
            {
                objUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Memo;
                objUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_Link;
            }
            else if (coluna is ColunaLink)
            {
                objUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha;
                objUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_Image;
            }

            if (coluna.Tamanho > 0)
            {
                objUserFieldsMD.EditSize = coluna.Tamanho;
            }

            var error_code = objUserFieldsMD.Add();

            if (error_code != 0)
            {
                throw new CustomException($"Erro ao tentar criar o campo {coluna.Nome}.\nErro: {oCompany.GetLastErrorDescription()}");
            }

            System.Runtime.InteropServices.Marshal.ReleaseComObject(objUserFieldsMD);
            objUserFieldsMD = null;
        }

        #endregion


        #region :: Exclusão de tabelas e campos

        public static bool ExcluirTabela(string nome_tabela)
        {
            GC.Collect();
            SAPbobsCOM.UserTablesMD objUserTablesMD = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserTables);
            objUserTablesMD.TableName = nome_tabela;
            objUserTablesMD.GetByKey(nome_tabela);

            var error_code = objUserTablesMD.Remove();

            System.Runtime.InteropServices.Marshal.ReleaseComObject(objUserTablesMD);
            objUserTablesMD = null;

            if (error_code == 0)
                return true;
            else
                throw new CustomException($"Erro ao tentar remover a tabela {nome_tabela}.\nErro: {oCompany.GetLastErrorDescription()}");

        }

        public static bool ExcluirColuna(string nome_tabela, string nome_campo)
        {
            int FieldId = GetFieldId(nome_tabela, nome_campo);

            GC.Collect();
            SAPbobsCOM.UserFieldsMD oUserFieldsMD = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields);
            if (oUserFieldsMD.GetByKey(nome_tabela, FieldId))
            {
                //removendo campo de tabela
                var error_code = oUserFieldsMD.Remove();
                if (error_code != 0)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserFieldsMD);
                    throw new CustomException($"Erro ao tentar remover o campo {nome_campo} da tabela {nome_tabela}.\nErro: {oCompany.GetLastErrorDescription()}");
                }
            }

            System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserFieldsMD);
            return true;
        }

        #endregion


        #region :: Valores válidos, Valores Padrão e Campo obrigatório

        public static void AdicionaValorValido(string nome_tabela, string nome_campo, string valor, string descricao)
        {
            bool valorExiste = false;
            int campoID = GetFieldId(nome_tabela, nome_campo);

            if (ExisteValorValido(nome_tabela, campoID, valor))
            {
                valorExiste = true;
            }
            else
            {
                GC.Collect();
                SAPbobsCOM.UserFieldsMD objUserFieldsMD = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields);
                objUserFieldsMD.GetByKey(nome_tabela, campoID);
                SAPbobsCOM.ValidValuesMD oValidValues;
                oValidValues = objUserFieldsMD.ValidValues;

                var error_code = 0;
                if (!valorExiste)
                {
                    if (oValidValues.Value != "")
                    {
                        oValidValues.Add();
                        oValidValues.SetCurrentLine(oValidValues.Count - 1);
                        oValidValues.Value = valor;
                        oValidValues.Description = descricao;
                        error_code = objUserFieldsMD.Update();
                    }
                    else
                    {
                        oValidValues.SetCurrentLine(oValidValues.Count - 1);
                        oValidValues.Value = valor;
                        oValidValues.Description = descricao;
                        error_code = objUserFieldsMD.Update();
                    }

                    if (error_code != 0)
                    {
                        throw new CustomException($"Erro ao tentar adicionar valor válido {valor} na coluna {nome_campo} na tabela {nome_tabela}.\nErro: {oCompany.GetLastErrorDescription()}");
                    }
                }
                else
                {
                    error_code = objUserFieldsMD.Update();
                    if (error_code != 0)
                    {
                        throw new CustomException($"Erro ao tentar atualizar valor válido {valor} na coluna {nome_campo} na tabela {nome_tabela}.\nErro: {oCompany.GetLastErrorDescription()}");
                    }
                }
            }



        }

        public static bool SetaCampoObrigatorio(string nome_tabela, string nome_campo)
        {
            int campoID = GetFieldId(nome_tabela, nome_campo);
            SAPbobsCOM.UserFieldsMD objUserFieldsMD;
            GC.Collect();
            objUserFieldsMD = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields);

            if (objUserFieldsMD.GetByKey(nome_tabela, campoID))
            {
                objUserFieldsMD.Mandatory = SAPbobsCOM.BoYesNoEnum.tYES;
                var CodErroDB = objUserFieldsMD.Update();

                if (CodErroDB != 0)
                {
                    throw new CustomException($"Erro ao tentar tornar o campo {nome_campo} da tabela {nome_tabela} obrigatório.\nErro: {oCompany.GetLastErrorDescription()}");
                }
            }
            return true;
        }

        public static bool SetaValorPadrao(string nome_tabela, string nome_campo, string valor)
        {
            bool valorExiste = false;
            int campoID = GetFieldId(nome_tabela, nome_campo);

            if (ExisteValorValido(nome_tabela, campoID, valor))
            {
                valorExiste = true;
            }

            //se existe esse valor válido
            if (valorExiste && (ExisteValorPadraoSetado(nome_tabela, campoID, valor)) == false)
            {
                GC.Collect();
                SAPbobsCOM.UserFieldsMD objUserFieldsMD;
                objUserFieldsMD = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields);

                if (objUserFieldsMD.GetByKey(nome_tabela, campoID))
                    objUserFieldsMD.DefaultValue = valor;

                var error_code = objUserFieldsMD.Update();
                if (error_code != 0)
                {
                    objUserFieldsMD = null;
                    throw new CustomException($"Erro ao tentar setar o valor padrão {valor} para o campo {nome_campo} da tabela {nome_tabela}.\nErro: {oCompany.GetLastErrorDescription()}");
                }
                else
                {
                    objUserFieldsMD = null;
                    return true;
                }
            }
            else
            {
                return false;
            }
        }

        #endregion


        #region :: Helpers


        public static bool ExisteTabela(string nome_tabela)
        {
            SAPbobsCOM.Recordset rs = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            string sql = "SELECT COUNT(*) FROM OUTB WHERE TableName = '" + nome_tabela + "'";
            rs.DoQuery(sql);

            if (rs.Fields.Item(0).Value <= 0)
            {
                rs = null;
                sql = null;
                return false;
            }
            else
            {
                rs = null;
                sql = null;
                return true;
            }
        }

        public static bool ExisteColuna(string nome_tabela, string nome_campo)
        {
            SAPbobsCOM.Recordset rs = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            var sql = $"SELECT COUNT(*) FROM CUFD (NOLOCK) WHERE TableID='{nome_tabela}' and AliasID='{nome_campo}'";

            rs.DoQuery(sql);
            if (rs.Fields.Item(0).Value <= 0)
            {
                rs = null;
                return false;
            }
            else
            {
                rs = null;
                return true;
            }
        }

        public static bool ExisteValorValido(string nome_tabela, int campoID, string valor)
        {
            SAPbobsCOM.Recordset rs = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            string sql =
                $@"SELECT COUNT(*) FROM UFD1 (NOLOCK) 
                    WHERE TableID='{nome_tabela}' AND
                        FieldID='{campoID}' AND
                        FldValue='{valor}'";

            rs.DoQuery(sql);
            if (rs.Fields.Item(0).Value > 0)
            {
                rs = null;
                return true;
            }

            rs = null;
            return false;
        }

        public static bool ExisteValorPadraoSetado(string nome_tabela, int campoID, string valor)
        {
            SAPbobsCOM.Recordset rs = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            string sql = $@"SELECT COUNT(*) FROM CUFD (NOLOCK) 
            Where TableID='{nome_tabela}' AND
                  FieldID='{campoID}' AND
                  dflt='{valor}'";
            rs.DoQuery(sql);

            if ((rs.Fields.Item(0).Value) <= 0)
            {
                rs = null;
                return false;
            }
            else
            {
                rs = null;
                return true;
            }
        }

        public static int GetFieldId(string nome_tabela, string nome_campo)
        {
            SAPbobsCOM.Recordset rs = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            string sql = $@" SELECT FieldID FROM CUFD (NOLOCK)  WHERE TableID='{nome_tabela}' AND AliasID='{nome_campo}'";
            rs.DoQuery(sql);
            if (rs.Fields.Item(0).Value >= 0)
            {
                return rs.Fields.Item(0).Value;
            }
            else
            {
                rs = null;
                return -1;
            }
        }

        #endregion


        #region :: Outros

        public static void RecebeCompany(SAPbobsCOM.Company company)
        {
            oCompany = company;
        }

        #endregion
    }
}