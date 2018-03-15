using System;
using System.Windows.Forms;
using System.Xml;

namespace InitAddon
{
    public static class SAPMenus
    {
        public static SAPbouiCOM.Application SBOApplication;

        public static void CriarMenus(string xmlpath = "")
        {
            LoadBatches(!String.IsNullOrEmpty(xmlpath) ? xmlpath : Application.StartupPath + "/criar_menus.xml");
        }

        public static void RemoverMenus(string xmlpath = "")
        {
            LoadBatches(!String.IsNullOrEmpty(xmlpath) ? xmlpath : Application.StartupPath + "/remover_menus.xml");
        }

        public static void LoadBatches(string xmlpath)
        {
            XmlDocument xml = new XmlDocument();
            xml.Load(xmlpath);
            SBOApplication.LoadBatchActions(xml.InnerXml);
        }

        public static void RecebeSBOApplication(SAPbouiCOM.Application application)
        {
            SBOApplication = application;
        }
    }
}