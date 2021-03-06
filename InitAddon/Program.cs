﻿using SAPHelper;
using System;
using System.Collections.Generic;

namespace InitAddon
{


    static class Program
    {
        private static string _addonName = "Addon Example";
        private static string _startupath = System.Windows.Forms.Application.StartupPath;
        public static SAPbouiCOM.Application _sBOApplication;
        public static SAPbobsCOM.Company _company;

        [STAThread]
        static void Main()
        {
            ConectarComSAP();

            CriarEstruturaDeDados();

            CriarMenus();

            DeclararEventos();

            Dialogs.Success(":: " + _addonName + " :: Iniciado.");

            // deixa a aplicação ativa
            System.Windows.Forms.Application.Run();
        }

        private static void ConectarComSAP()
        {
            SAPConnection.SBOApplicationHandler applicationHandler = null;
            applicationHandler += Dialogs.RecebeSBOApplication;
            applicationHandler += Menu.RecebeSBOApplication;
            applicationHandler += applicationParam => _sBOApplication = applicationParam;

            SAPConnection.CompanyHandler companyHandler = null;
            companyHandler += companyParam => _company = companyParam;
            companyHandler += Database.RecebeCompany;

            try
            {
                SAPConnection.Connect(applicationHandler, companyHandler);
            }
            catch (Exception e)
            {
                System.Windows.Forms.MessageBox.Show(e.Message);
                System.Windows.Forms.Application.Exit();
            }
        }

        private static void CriarEstruturaDeDados()
        {
            Dialogs.Info(":: " + _addonName + " :: Criando tabelas e estruturas de dados ...");

            try
            {
                _company.StartTransaction();

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
                        new ColunaInt("teste","teste",true),
                    }
                    , new UDOParams() { CanDelete = SAPbobsCOM.BoYesNoEnum.tNO }
                    , new List<Tabela>() { tabela_detalhe_item }
                );

                //Database.ExcluirColuna(tabela_contratos.NomeComArroba, "teste");


                //var coluna_teste = new ColunaInt("testex", "xtestex", true);
                //Database.CriarColuna(tabela_contratos.NomeComArroba, coluna_teste);
                //Database.DefinirColunasComoUDO(tabela_contratos.NomeComArroba, new List<Coluna>() { coluna_teste });

                //Database.ExcluirTabela(tabela_contratos.NomeSemArroba);

                _company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
            }
            catch (DatabaseException e)
            {
                Dialogs.PopupError(e.Message);
            }
            catch (Exception e)
            {
                _company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                Dialogs.PopupError("Erro interno.\nErro: " + e.Message);
            }
        }

        private static void CriarMenus()
        {
            Dialogs.Info(":: " + _addonName + " :: Criando menus ...");

            try
            {
                Menu.RemoverMenus(_startupath + @"/remover_menus.xml");

                Menu.CriarMenus(_startupath + @"/criar_menus.xml");
            }
            catch (Exception e)
            {
                Dialogs.PopupError("Erro ao inserir menus.\nErro: " + e.Message);
            }
        }

        private static void DeclararEventos()
        {
            _sBOApplication.AppEvent += SBOApplication_AppEvent;
        }

        #region :: Declaração Eventos

        private static void SBOApplication_AppEvent(SAPbouiCOM.BoAppEventTypes EventType)
        {
            if (EventType == SAPbouiCOM.BoAppEventTypes.aet_ShutDown)
            {
                Menu.RemoverMenus(_startupath + @"/remover_menus.xml");
            }
        }

        #endregion
    }
}
