using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using OfficeOpenXml;
using ExcelDataReader;
using Z.Dapper.Plus;
using System.IO;

namespace InsertDataFromExcel
{
    class Customer
    {
        public string id { get; set; }
        public string tk { get; set; }
        public string mk { get; set; }
    }
}
