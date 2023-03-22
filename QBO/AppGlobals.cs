using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace QBOXlAddIn
{
    internal sealed partial class AppGlobals
    {
        private AppGlobals()
        {
        }

        private static ThisAddIn _ThisAddIn;

        private static global::Microsoft.Office.Tools.Excel.ApplicationFactory _factory;

        public static string ACCESS_TOKEN;
        public static string REFRESH_TOKEN;
        public static long ACCESS_TOKEN_EXPIRES_IN;
        public static string REALM_ID;
        public static string CODE;
        public static string STATE;
        public static string REPORT_START_DATE;
        public static string REPORT_END_DATE;

        internal static ThisAddIn ThisAddIn
        {
            get
            {
                return _ThisAddIn;
            }
            set
            {
                if ((_ThisAddIn == null))
                {
                    _ThisAddIn = value;
                }
                else
                {
                    throw new System.NotSupportedException();
                }
            }
        }

        internal static global::Microsoft.Office.Tools.Excel.ApplicationFactory Factory
        {
            get
            {
                return _factory;
            }
            set
            {
                if ((_factory == null))
                {
                    _factory = value;
                }
                else
                {
                    throw new System.NotSupportedException();
                }
            }
        }

    }


}
