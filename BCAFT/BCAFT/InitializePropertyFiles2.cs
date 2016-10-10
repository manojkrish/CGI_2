using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;

namespace BCAFT
{
    interface InitializePropertyFiles2
    {        
        public Object initializeCustomerJourneyPropertyFile()
        {
            string strResultsPath = Directory.GetParent(Directory.GetCurrentDirectory()).Parent.Parent.FullName + "\\ObjectProperties";
            Object CJProperties = new Properties(strResultsPath + "\\CustomerJourney.txt");
            return CJProperties;
        }

        public Object initializeAdminPropertyFile()
        {
            string strResultsPath = Directory.GetParent(Directory.GetCurrentDirectory()).Parent.Parent.FullName + "\\ObjectProperties";
            Object CJProperties = new Properties(strResultsPath + "\\Admin.txt");
            return CJProperties;
        }
    }
}
