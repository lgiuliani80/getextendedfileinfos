using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WinProps;

namespace EnumColumns
{
    class Program
    {
        static void Main(string[] args)
        {
            var props = PropertyKey.Keys;

            foreach (var p in props)
            {
                Console.WriteLine($"{p.Key} = {p.Value.FmtId} {p.Value.Pid}");
            }
        }
    }
}
