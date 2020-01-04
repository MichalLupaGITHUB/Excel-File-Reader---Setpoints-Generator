using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Linq;


namespace ConvTool.Classes
{
    public class SaveXMLFile
    {
        #region Private Members
        
        private string XMLFileString;

        #endregion Private Members

        #region Public Members

        public XmlDocument XMLFile = new XmlDocument();

        #endregion Public Members

        #region Constructors

        public SaveXMLFile()
        {

        }

        #endregion Constructors

        #region Private Methods
        /// <summary>
        /// Set appropriate format for displaying on the TextBox
        /// </summary>
        /// <param name="xmlString"></param>
        /// <returns></returns>
        private string FormatXML(string xmlString)
        {
            XmlDocument TempDocument = new XmlDocument();

            TempDocument.LoadXml(xmlString); // Loading XML String

            StringBuilder TempString = new StringBuilder();

            System.IO.TextWriter TempTextW = new System.IO.StringWriter(TempString);

            XmlTextWriter TempXML = new XmlTextWriter(TempTextW);

            TempXML.Formatting = Formatting.Indented; // Set Format for XML String

            TempDocument.Save(TempXML);

            TempXML.Close();

            return TempString.ToString();
        }

        #endregion Private Methods

        #region Public Methods

        /// <summary>
        /// Creating XML File from loaded EXCEL data
        /// </summary>
        /// <param name="General"></param>
        /// <param name="CondCase"></param>
        /// <param name="INCA"></param>
        public void CreateXMLFile(XmlNode General, XmlNode CondCase, XmlNode INCA)
        {
            try
            {
                XmlDeclaration DeclarationNode = XMLFile.CreateXmlDeclaration("1.0", "UTF-8", null);
                XmlElement Root = XMLFile.DocumentElement;
                XMLFile.InsertBefore(DeclarationNode, Root);

                XmlNode RootNode = XMLFile.CreateElement("Root");
                XMLFile.AppendChild(RootNode);

                XmlNode GeneralNode = XMLFile.ImportNode(General.LastChild, true);
                RootNode.AppendChild(GeneralNode);

                XmlNode CondCaseNode = XMLFile.ImportNode(CondCase.LastChild, true);
                RootNode.AppendChild(CondCaseNode);

                XmlNode INCANode = XMLFile.ImportNode(INCA.LastChild, true);
                RootNode.AppendChild(INCANode);               
            }
            catch (Exception)
            {
                System.Windows.Forms.MessageBox.Show("Saving XML File Error!", "ConvTool - Error", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
                Environment.Exit(11);
            }
        }

        /// <summary>
        /// Save XML File
        /// </summary>
        /// <param name="DestinationPath"></param>
        public void SaveXMLNodesFile(string DestinationPath)
        {
            XMLFile.Save(DestinationPath + "\\" + "XMLFile.xml");
        }

        /// <summary>
        /// Converting XML Data To String
        /// </summary>
        /// <returns></returns>
        public string ConvertXMLDocumentToString()
        {
            using (var stringWriter = new StringWriter()) // using IDisposable interface
            using (var xmlTextWriter = XmlWriter.Create(stringWriter)) // using IDisposable interface
            {
                XMLFile.WriteTo(xmlTextWriter);
                xmlTextWriter.Flush();
                XMLFileString = stringWriter.GetStringBuilder().ToString();
            }

            XMLFileString = FormatXML(XMLFileString);

            return XMLFileString;
        }

        #endregion Public Methods
    }
}
