﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SharePointCSOMRestSeries
{
    class Program
    {
        static void Main(string[] args)
        {
            CSOMDay1 objCSOM = new CSOMDay1();
            objCSOM.ConnectSharePointOnline();
            Console.ReadLine();
        }

    }
}
