﻿using SAPbouiCOM;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace REDFARM.Addons
{
    public class Utils
    {
        public static string[] ReadTxt(string path) => File.ReadAllLines(path);
    }
}
