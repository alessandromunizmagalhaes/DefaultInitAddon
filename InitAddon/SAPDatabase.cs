using System;

namespace InitAddon
{
    public static class SAPDatabase
    {
        public static SAPbobsCOM.Company oCompany;

        public static void CriarTabela(Tabela tabela)
        {
            if (!ExisteTabela(tabela.NomeSemArroba))
            {
                CriarUserTable(tabela);

                foreach (var coluna in tabela.Colunas)
                {
                    if (!ExisteColuna(tabela.NomeComArroba, coluna.Nome))
                    {
                        CriarUserField(tabela.NomeComArroba, coluna);
                    }
                }
            }
        }

        //Verifica a existencia de tabela de usuário 
        public static bool ExisteTabela(string cTable)
        {
            SAPbobsCOM.Recordset rs = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            string sql = "SELECT COUNT(*) FROM OUTB WHERE TableName = '" + cTable + "'";
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

        private static void CriarUserField(string nome_tabela, Coluna coluna)
        {
            GC.Collect();
            SAPbobsCOM.UserFieldsMD objUserFieldsMD = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields);
            objUserFieldsMD.TableName = nome_tabela;
            objUserFieldsMD.Name = coluna.Nome;
            objUserFieldsMD.Description = coluna.Descricao;

            if (coluna.Tipo == ColunaTipo.Varchar)
            {
                objUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha;
            }
            else if (coluna.Tipo == ColunaTipo.Text)
            {
                objUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Memo;
            }
            else if (coluna.Tipo == ColunaTipo.Date)
            {
                objUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Date;
            }
            else if (coluna.Tipo == ColunaTipo.Int)
            {
                objUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Numeric;
            }
            else if (coluna.Tipo == ColunaTipo.Time)
            {
                objUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Date;
                objUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_Time;
            }
            else if (coluna.Tipo == ColunaTipo.Percent)
            {
                objUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Float;
                objUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_Percentage;
            }
            else if (coluna.Tipo == ColunaTipo.Sum)
            {
                objUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Float;
                objUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_Sum;
            }
            else if (coluna.Tipo == ColunaTipo.Quantity)
            {
                objUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Float;
                objUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_Percentage;
            }
            else if (coluna.Tipo == ColunaTipo.Price)
            {
                objUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Float;
                objUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_Price;
            }
            else if (coluna.Tipo == ColunaTipo.Link)
            {
                objUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Memo;
                objUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_Link;
            }
            else if (coluna.Tipo == ColunaTipo.Image)
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

        private static void CriarUserTable(Tabela tabela)
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

        private static bool ExisteColuna(string nome_tabela, string nome_campo)
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

        public static bool ExcluiCampo(string nome_tabela, string nome_campo)
        {
            int FieldId = FieldID(nome_tabela, nome_campo);

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

        public static int FieldID(string nome_tabela, string nome_campo)
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

        public static void RecebeCompany(SAPbobsCOM.Company company)
        {
            oCompany = company;
        }
    }
}