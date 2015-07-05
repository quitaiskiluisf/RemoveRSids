using System;
using System.Text;
using System.IO;
using System.Xml;
using RemoveRSids;

class Program
{
    public static void Main(string[] argsv)
    {
        for (int i = 0; i < argsv.Length; i++)
            using (RemovedorRevisionSessionIdentifiers r = new RemovedorRevisionSessionIdentifiers(argsv[i], Console.OpenStandardOutput()))
                r.Processa();
    }
}
