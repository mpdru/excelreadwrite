// See https://aka.ms/new-console-template for more information

using System;
using System.Runtime.InteropServices;

namespace MyApp // Note: actual namespace depends on the project name.
{
    internal class Program
    {
        static void Main(string[] args)
        {

            var arquivo = "C:\\Users\\marco\\Desktop\\teste.xlsx";

            try
            {

                using (var package = new OfficeOpenXml.ExcelPackage(new FileInfo(arquivo)))
                {

                    package.Save();

                    var firstSheet = package.Workbook.Worksheets[0];

                    var celA2 = firstSheet.Cells["A2"].Text;
                    var celB2 = firstSheet.Cells["B2"].Text;

                    var colACount = firstSheet.Cells["A:A"].Count();
                    var colBCount = firstSheet.Cells["B:B"].Count();

                    Console.WriteLine("Sheet 1 Data");
                    Console.WriteLine($"Cell A2 Value   : {celA2}");
                    Console.WriteLine($"Cell B2 Value   : {celB2}");
                    Console.WriteLine($"Col A Elements  : {colACount}");
                    Console.WriteLine($"Col B Elements  : {colBCount}");

                    var status = "";
                    var agora = "";

                    for(int i = 2; i <= colACount; i++) {

                        status = firstSheet.Cells["B" + i.ToString()].Text;
                        Console.WriteLine($"Status atual  : {status}");

                        if(!status.Equals("ok")) {
                            //processar
                                //se sucesso
                                firstSheet.Cells["B" + i.ToString()].Value = "ok";
                                agora = DateTime.Now.ToString("dd'/'MM'/'yyyy' 'HH':'mm':'ss");
                                firstSheet.Cells["C" + i.ToString()].Value = agora;
                                Console.WriteLine($"Agora    :    {agora}");
                        }

                    }

                    package.Save();

                    Console.WriteLine("");
                }

            } catch {

                Console.WriteLine("**********************************************************************************");
                Console.WriteLine(" ALERTA");
                Console.WriteLine(" O arquivo " + arquivo + " está aberto / em uso.");
                Console.WriteLine(" Feche o arquivo e reinicie o programa!");
                Console.WriteLine("**********************************************************************************");
                
            }

        }
     
    }
}

