using System;
using System.Collections.Generic;

namespace InitAddon
{
    public static class SAPDatabase
    {
        public static SAPbobsCOM.Company _company;

        #region :: Gestão de Tabelas

        public static void CriarTabela(Tabela tabela)
        {
            bool tabela_is_UDO = tabela is TabelaUDO;
            TabelaUDO tabelaUDO = tabela_is_UDO ? (TabelaUDO)tabela : null;

            if (tabela_is_UDO)
            {
                foreach (var tabelaFilha in tabelaUDO.TabelasFilhas)
                {
                    // atenção para a recursividade aqui
                    CriarTabela(tabelaFilha);
                }
            }

            if (!ExisteTabela(tabela.NomeSemArroba))
            {
                CriarUserTable(tabela);

                foreach (var coluna in tabela.Colunas)
                {
                    if (!ExisteColuna(tabela.NomeComArroba, coluna.Nome))
                    {
                        CriarColuna(tabela.NomeComArroba, coluna);
                    }
                }

                // não tem como ao mesmo tempo que criar uma coluna, já marcar ela como udo
                // tem que criar todas as colunas e depois iterar denovo só pra adicionar o UDO
                // regras SAP, senão vc não consegue adicionar o UDO via DI.
                if (tabela_is_UDO)
                {
                    DefinirTabelaComoUDO(tabelaUDO);

                    DefinirColunasComoUDO(tabela.NomeSemArroba, tabela.Colunas, true);

                    DefinirTabelasComoFilhasDoUDO(tabela.NomeSemArroba, tabelaUDO.TabelasFilhas);
                }
            }
        }

        private static void DefinirTabelasComoFilhasDoUDO(string nomeTabelaPai, List<Tabela> tabelasFilhas)
        {
            GC.Collect();
            SAPbobsCOM.UserObjectsMD objUserObjectMD = _company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD);
            if (objUserObjectMD.GetByKey(nomeTabelaPai))
            {
                // tem que percorrer denovo todas as tabelas filhas, para seta-las como filhas.
                // não dá pra aproveitar o for lá encima, no momento que o for lá emcima é executado, o UDO pai não existe,
                // então quando vai fazer GETBYKEY() com o nome do pai, sempre retorna false.
                foreach (var tabelaFilha in tabelasFilhas)
                {
                    objUserObjectMD.ChildTables.TableName = tabelaFilha.NomeSemArroba;
                    objUserObjectMD.ChildTables.Add();
                }

                if (objUserObjectMD.Update() != 0)
                {
                    throw new CustomException($"Erro ao tentar adicionar as tabela filhas da tabela {nomeTabelaPai} que é UDO.\nErro: {_company.GetLastErrorDescription()}");
                }
            }

            System.Runtime.InteropServices.Marshal.ReleaseComObject(objUserObjectMD);
            objUserObjectMD = null;
            GC.Collect();
        }

        private static void DefinirTabelaComoUDO(TabelaUDO tabela)
        {
            GC.Collect();
            SAPbobsCOM.UserObjectsMD objUserObjectMD = _company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD);

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
                throw new CustomException($"Erro ao tentar criar a tabela {tabela.NomeSemArroba} como UDO.\nErro: {_company.GetLastErrorDescription()}");
            }

            System.Runtime.InteropServices.Marshal.ReleaseComObject(objUserObjectMD);
            objUserObjectMD = null;
            GC.Collect();
        }

        private static void CriarUserTable(Tabela tabela)
        {
            GC.Collect();
            SAPbobsCOM.UserTablesMD oUserTablesMD = _company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserTables);

            oUserTablesMD.TableName = tabela.NomeSemArroba;
            oUserTablesMD.TableDescription = tabela.Descricao;
            oUserTablesMD.TableType = tabela.Tipo;

            if (oUserTablesMD.Add() != 0)
            {
                throw new CustomException($"Erro ao tentar criar a tabela {tabela.NomeComArroba}.\nErro: {_company.GetLastErrorDescription()}");
            }

            oUserTablesMD = null;
            GC.Collect(); // Release the handle to the table
        }

        public static void ExcluirTabela(string nomeSemArroba)
        {
            GC.Collect();
            SAPbobsCOM.UserObjectsMD oUDO = _company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD);

            if (oUDO.GetByKey(nomeSemArroba))
            {
                if (oUDO.Remove() != 0)
                    throw new CustomException($"Erro ao tentar remover a definição de UDO da tabela {nomeSemArroba}.\nErro: {_company.GetLastErrorDescription()}");
            }

            System.Runtime.InteropServices.Marshal.ReleaseComObject(oUDO);
            oUDO = null;

            SAPbobsCOM.UserTablesMD objUserTablesMD = _company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserTables);
            if (objUserTablesMD.GetByKey(nomeSemArroba))
            {
                objUserTablesMD.TableName = nomeSemArroba;

                if (objUserTablesMD.Remove() != 0)
                    throw new CustomException($"Erro ao tentar remover a tabela {nomeSemArroba}.\nErro: {_company.GetLastErrorDescription()}");
            }
            else
            {
                throw new CustomException($"tabela {nomeSemArroba} não encontrada para realizar remoção.");
            }

            System.Runtime.InteropServices.Marshal.ReleaseComObject(objUserTablesMD);
            objUserTablesMD = null;
        }

        public static bool ExisteTabela(string nome_tabela)
        {
            SAPbobsCOM.Recordset rs = _company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
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

        #endregion


        #region :: Gestão de Campos

        public static void CriarColuna(string nome_tabela, Coluna coluna)
        {
            CriarUserField(nome_tabela, coluna);

            if (coluna.Obrigatoria)
            {
                SetarCampoComoObrigatorio(nome_tabela, coluna.Nome);
            }

            foreach (var valor_valido in coluna.ValoresValidos)
            {
                AdicionarValorValido(nome_tabela, coluna.Nome, valor_valido.Valor, valor_valido.Descricao);
            }

            if (!String.IsNullOrEmpty(coluna.ValorPadrao))
            {
                SetarValorPadrao(nome_tabela, coluna.Nome, coluna.ValorPadrao);
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="nome_tabela"></param>
        /// <param name="colunas"></param>
        /// <param name="criar_campo_code_antes">
        /// quando se está criando uma nova tabela
        /// tem que fazer essa gambiarra horrível, porque o primeiro elemento a ser colocado como UDO,
        /// tem que obrigatóriamente ser o campo CODE.
        /// como eu não passo ele como um campo que eu quero usar na definição de Lista de Colunas da minha Tabela
        /// sempre que for o primeiro, inventa uma coluna ficticia e passa o Code.
        /// horroroso mas é o jeito.
        /// </param>
        public static void DefinirColunasComoUDO(string nome_tabela, List<Coluna> colunas, bool criar_campo_code_antes = false)
        {
            GC.Collect();
            SAPbobsCOM.UserObjectsMD objUserObjectMD = _company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD);
            if (objUserObjectMD.GetByKey(nome_tabela))
            {
                int i = 1;
                foreach (var coluna in colunas)
                {
                    // quando se está criando uma nova tabela
                    // tem que fazer essa gambiarra horrível, porque o primeiro elemento a ser colocado como UDO,
                    // tem que obrigatóriamente ser o campo CODE.
                    // como eu não passo ele como um campo que eu quero usar na definição de Lista de Colunas da minha Tabela
                    // sempre que for o primeiro, inventa uma coluna ficticia e passa o Code.
                    // horroroso mas é o jeito.
                    if (criar_campo_code_antes && i == 1)
                    {
                        AdicionarFindColumnsAoObjeto(objUserObjectMD, new ColunaVarchar("Code", "Código", 0, false));
                    }

                    AdicionarFindColumnsAoObjeto(objUserObjectMD, coluna);

                    i++;
                }

                if (objUserObjectMD.Update() != 0)
                {
                    throw new CustomException($"Erro ao tentar criar as colunas da tabela {nome_tabela} como UDO.\nErro: {_company.GetLastErrorDescription()}");
                }
            }

            System.Runtime.InteropServices.Marshal.ReleaseComObject(objUserObjectMD);
            objUserObjectMD = null;
            GC.Collect();
        }

        private static void CriarUserField(string nome_tabela, Coluna coluna)
        {
            GC.Collect();
            SAPbobsCOM.UserFieldsMD objUserFieldsMD = _company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields);
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


            if (objUserFieldsMD.Add() != 0)
            {
                throw new CustomException($"Erro ao tentar criar o campo {coluna.Nome}.\nErro: {_company.GetLastErrorDescription()}");
            }

            System.Runtime.InteropServices.Marshal.ReleaseComObject(objUserFieldsMD);
            objUserFieldsMD = null;
        }

        private static void AdicionarFindColumnsAoObjeto(SAPbobsCOM.UserObjectsMD objUserObjectMD, Coluna coluna)
        {
            objUserObjectMD.FindColumns.ColumnAlias = coluna.Nome;
            objUserObjectMD.FindColumns.ColumnDescription = coluna.Descricao;
            objUserObjectMD.FindColumns.Add();
        }

        public static void ExcluirColuna(string nome_tabela, string nome_campo)
        {
            int FieldId = GetFieldId(nome_tabela, nome_campo);

            GC.Collect();
            SAPbobsCOM.UserFieldsMD oUserFieldsMD = _company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields);
            if (oUserFieldsMD.GetByKey(nome_tabela, FieldId))
            {
                if (oUserFieldsMD.Remove() != 0)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserFieldsMD);
                    throw new CustomException($"Erro ao tentar remover o campo {nome_campo} da tabela {nome_tabela}.\nErro: {_company.GetLastErrorDescription()}");
                }
            }
            else
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserFieldsMD);
                throw new CustomException($"Campo {nome_campo} não encontrado na tabela {nome_tabela} para realizar a exclusão.");
            }
        }

        public static bool ExisteColuna(string nome_tabela, string nome_campo)
        {
            SAPbobsCOM.Recordset rs = _company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
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

        #endregion


        #region :: Valores válidos, Valores Padrão e Obrigatoriedade

        public static void AdicionarValorValido(string nome_tabela, string nome_campo, string valor, string descricao)
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
                SAPbobsCOM.UserFieldsMD objUserFieldsMD = _company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields);
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
                        throw new CustomException($"Erro ao tentar adicionar valor válido {valor} na coluna {nome_campo} na tabela {nome_tabela}.\nErro: {_company.GetLastErrorDescription()}");
                    }
                }
                else
                {
                    error_code = objUserFieldsMD.Update();
                    if (error_code != 0)
                    {
                        throw new CustomException($"Erro ao tentar atualizar valor válido {valor} na coluna {nome_campo} na tabela {nome_tabela}.\nErro: {_company.GetLastErrorDescription()}");
                    }
                }
            }
        }

        public static bool SetarCampoComoObrigatorio(string nome_tabela, string nome_campo)
        {
            int campoID = GetFieldId(nome_tabela, nome_campo);
            SAPbobsCOM.UserFieldsMD objUserFieldsMD;
            GC.Collect();
            objUserFieldsMD = _company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields);

            if (objUserFieldsMD.GetByKey(nome_tabela, campoID))
            {
                objUserFieldsMD.Mandatory = SAPbobsCOM.BoYesNoEnum.tYES;

                if (objUserFieldsMD.Update() != 0)
                {
                    throw new CustomException($"Erro ao tentar tornar o campo {nome_campo} da tabela {nome_tabela} obrigatório.\nErro: {_company.GetLastErrorDescription()}");
                }
            }
            return true;
        }

        public static bool SetarValorPadrao(string nome_tabela, string nome_campo, string valor)
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
                objUserFieldsMD = _company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields);

                if (objUserFieldsMD.GetByKey(nome_tabela, campoID))
                    objUserFieldsMD.DefaultValue = valor;

                if (objUserFieldsMD.Update() != 0)
                {
                    objUserFieldsMD = null;
                    throw new CustomException($"Erro ao tentar setar o valor padrão {valor} para o campo {nome_campo} da tabela {nome_tabela}.\nErro: {_company.GetLastErrorDescription()}");
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

        public static bool ExisteValorValido(string nome_tabela, int campoID, string valor)
        {
            SAPbobsCOM.Recordset rs = _company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
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
            SAPbobsCOM.Recordset rs = _company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
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

        #endregion


        #region :: Helpers

        public static int GetFieldId(string nome_tabela, string nome_campo)
        {
            SAPbobsCOM.Recordset rs = _company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
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
            _company = company;
        }

        #endregion
    }
}