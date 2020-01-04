using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConvTool.Classes
{
    public class ExcelPath
    {
        #region Private Members

        private string ExcelPathRetVal;

        #endregion Private Members

        #region Constructors
        public ExcelPath()
        {

        }
        #endregion Constructors

        #region Public Methods
        /// <summary>
        /// Get Excel Path - Get path to the Excel File
        /// </summary>
        /// <param name="ExcelPath"></param>
        /// <returns></returns>
        public string GetExcelPath(string ExcelPath)
        {
            ExcelPathRetVal = ExcelPath;
            return ExcelPathRetVal;
        }

        #endregion Public Methods
    }
}
