using System;
using System.Linq;
using NeoFileMagic;


class Program
{
    static void Main()
    {
        var doc = Ods.Load("./sample.ods");
        var sheet = doc.Sheets[0];
        if (sheet.RowCount == 0)
        {
            Console.WriteLine("Sheet 為空");
            return;
        }

        var row0 = sheet.Rows[0];
        var values = row0.Cells.Select(c => Ods.OneLine(c, TextHandling.CollapseToSpace)).ToArray();
        for (int i = 0; i < values.Length; i++)
        {
            Console.WriteLine($"[{i}] {values[i]}");
        }
    }

   
}
