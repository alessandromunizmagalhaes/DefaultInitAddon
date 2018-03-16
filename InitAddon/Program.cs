using System;
using System.Collections.Generic;

namespace InitAddon
{


    static class Program
    {
        private static string AddonName = "Addon Example";
        public static SAPbouiCOM.Application SBOApplication;
        public static SAPbobsCOM.Company oCompany;

        [STAThread]
        static void Main()
        {
            ConectarComSAP();

            CriarEstruturaDeDados();

            CriarMenus();

            DeclararEventos();

            Dialogs.Success(":: " + AddonName + " :: Iniciado.");

            // deixa a aplicação ativa
            System.Windows.Forms.Application.Run();
        }

        private static void ConectarComSAP()
        {
            SAPConnection.SBOApplicationHandler applicationHandler = null;
            applicationHandler += Dialogs.RecebeSBOApplication;
            applicationHandler += SAPMenus.RecebeSBOApplication;
            applicationHandler += applicationParam => SBOApplication = applicationParam;

            SAPConnection.CompanyHandler companyHandler = null;
            companyHandler += companyParam => oCompany = companyParam;
            companyHandler += SAPDatabase.RecebeCompany;

            SAPConnection.Connect(applicationHandler, companyHandler);
        }

        private static void CriarEstruturaDeDados()
        {
            Dialogs.Info(":: " + AddonName + " :: Criando tabelas e estruturas de dados ...");

            try
            {
                oCompany.StartTransaction();

                var tabela_detalhe_item = new Tabela("U_UPD_CCD1", "Detalhes do item Previsto"
                    , SAPbobsCOM.BoUTBTableType.bott_DocumentLines
                    , new List<Coluna>() {
                        new ColunaVarchar("ItemCode","Código do Item", 30, true),
                        new ColunaVarchar("ItemName","Descrição do Item", 120, true),
                        new ColunaPercent("PercItem","Percentagem Classe",true),
                        new ColunaInt("teste","teste",true),
                });

                var tabela_contratos = new TabelaUDO("U_UPD_OCCD", "Definições Gerais do Contrato"
                    , SAPbobsCOM.BoUTBTableType.bott_Document
                    , new List<Coluna>() {
                        new ColunaVarchar("CardCode","Código Fornecedor", 15,true, ""),
                        new ColunaVarchar("CardName","Descrição Fornecedor", 100,true, ""),
                        new ColunaVarchar("CtName","Pessoa de Contato", 50,true, ""),
                        new ColunaVarchar("Tel1","Pessoa de Contato", 15,true, ""),
                        new ColunaVarchar("EMail","E-mail", 50,true, ""),
                        new ColunaDate("DtPrEnt","Data Previsão Entrega",true),
                        new ColunaDate("DtPrPgt","Data Programa Entrega",true),
                        new ColunaInt("ModCtto","Modalidade Contrato",true),
                    }
                    , new UDOParams() { CanDelete = SAPbobsCOM.BoYesNoEnum.tNO }
                    , new List<Tabela>() { tabela_detalhe_item }
                );

                SAPDatabase.CriarTabela(tabela_contratos);

                oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
            }
            catch (CustomException e)
            {
                Dialogs.PopupError(e.Message);
            }
            catch (Exception e)
            {
                oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                Dialogs.PopupError("Erro interno.\nErro: " + e.Message);
            }
        }

        private static void CriarMenus()
        {
            Dialogs.Info(":: " + AddonName + " :: Criando menus ...");

            try
            {
                SAPMenus.RemoverMenus();

                SAPMenus.CriarMenus();
            }
            catch (Exception e)
            {
                Dialogs.PopupError("Erro ao inserir menus.\nErro: " + e.Message);
            }
        }

        private static void DeclararEventos()
        {
            SBOApplication.AppEvent += SBOApplication_AppEvent;
        }

        #region :: Declaração Eventos

        private static void SBOApplication_AppEvent(SAPbouiCOM.BoAppEventTypes EventType)
        {
            if (EventType == SAPbouiCOM.BoAppEventTypes.aet_ShutDown)
            {
                SAPMenus.RemoverMenus();
            }
        }

        #endregion
    }
}
