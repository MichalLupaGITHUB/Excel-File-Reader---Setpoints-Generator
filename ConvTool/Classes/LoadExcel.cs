using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelDataReader;
using System.Windows;

namespace ConvTool.Classes
{
    public class LoadExcel
    {
        #region Private members

        private FileStream ExcelFile;
       
        #endregion Private members

        #region Constructors
        public LoadExcel()
        {
            
        }
        #endregion Constructors

        #region Public Methods

        /// <summary>
        /// Open Excel File - Opening Excel File
        /// </summary>
        /// <param name="ExcelFilePath"></param>
        /// <returns></returns>
        public FileStream OpenExcelFile(string ExcelFilePath)
        {
            try
            {
                ExcelFile = File.Open(ExcelFilePath, FileMode.Open, FileAccess.Read);
            }
            catch (Exception)
            {
                System.Windows.Forms.MessageBox.Show("Loading Excel File Error!\nExcel File must be closed!", "ConvTool - Error", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
                Environment.Exit(5);
            }

            return ExcelFile;    
        }

        /// <summary>
        /// Close Excel File - Closing Excel File
        /// </summary>
        public void CloseExcelFile()
        {
            ExcelFile.Close();
        }
        #endregion Public Methods
    }
}
