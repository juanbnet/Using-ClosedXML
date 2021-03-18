using InfDiario.Data;
using System;
using System.Data;

namespace InfDiario
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Generando Informe Diario...");
            DataTable dt = new DataTable();
            Informe informe = new Informe();
            dt = informe.getInforme();

            ReporteXlsx reporteXlsx = new ReporteXlsx();

            Console.WriteLine("Generando Archivo Excel...");
            reporteXlsx.createXlsx(dt);

            Console.WriteLine("Proceso Finalizado");

            Console.ReadKey();
        }
    }
}
