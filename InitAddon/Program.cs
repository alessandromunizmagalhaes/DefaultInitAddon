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
                var tabela_upd_teste = new Tabela(
                        "UPD_PCK_TESTE"
                        , "Apenas uma tabela de teste"
                        , SAPbobsCOM.BoUTBTableType.bott_MasterData
                        , new List<Coluna>() {
                              new Coluna("varchar", "um teste", ColunaTipo.Varchar, 10)
                            , new Coluna("date", "dois teste", ColunaTipo.Date)
                            , new Coluna("int", "três teste", ColunaTipo.Int)
                            , new Coluna("perc", "quatro teste", ColunaTipo.Percent, 10)
                        }
                );

                SAPDatabase.ExcluirTabela(tabela_upd_teste.NomeSemArroba);
                SAPDatabase.CriarTabela(tabela_upd_teste);
            }
            catch (CustomException e)
            {
                Dialogs.PopupError(e.Message);
            }
            catch (Exception e)
            {
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
