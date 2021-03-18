using System;
using System.Data;
using ClosedXML.Excel;

namespace InfDiario
{
    class ReporteXlsx
    {
        public void createXlsx(DataTable dataTable)
        {

            XLWorkbook wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Datos");

            ws.Range("A1").Value = "Factura";
            ws.Range("B1").Value = "Fecha Factura";
            ws.Range("C1").Value = "Historia Clínica";
            ws.Range("D1").Value = "Ingreso";
            ws.Range("E1").Value = "Mes";
            ws.Range("F1").Value = "Etapa";
            ws.Range("G1").Value = "Dias";
            ws.Range("H1").Value = "Fecha de Entrada";
            ws.Range("I1").Value = "Servicio";
            ws.Range("J1").Value = "Valor";
            ws.Range("K1").Value = "Convenio";
            ws.Range("L1").Value = "Plan";           
            ws.Range("M1").Value = "Marca No Pos";

            int row = 2;         

            for (int i=0;i<= dataTable.Rows.Count-1;i++)
            {

                ws.Cell(row, "A").Value = dataTable.Rows[i]["Factura"].ToString().Trim();
                ws.Cell(row, "B").Value = dataTable.Rows[i]["Fecha Factura"].ToString().Trim();
                ws.Cell(row, "C").Value = dataTable.Rows[i]["Historia Clínica"].ToString().Trim();
                ws.Cell(row, "D").Value = dataTable.Rows[i]["Ingreso"].ToString().Trim();
                ws.Cell(row, "E").Value = dataTable.Rows[i]["Mes"].ToString().Trim();
                ws.Cell(row, "F").Value = dataTable.Rows[i]["Etapa"].ToString().Trim();
                ws.Cell(row, "G").Value = dataTable.Rows[i]["Dias"].ToString().Trim();
                ws.Cell(row, "H").Value = dataTable.Rows[i]["Fecha Entrada"].ToString().Trim();
                ws.Cell(row, "I").Value = dataTable.Rows[i]["Servicio"].ToString().Trim();
                ws.Cell(row, "J").Value = dataTable.Rows[i]["Valor"].ToString().Trim();
                ws.Cell(row, "K").Value = dataTable.Rows[i]["Convenio"].ToString().Trim();
                ws.Cell(row, "L").Value = dataTable.Rows[i]["Plan"].ToString().Trim();
                ws.Cell(row, "M").Value = dataTable.Rows[i]["Marca No Pos"].ToString().Trim();

                row += 1;               
            }


            var lastRow = ws.LastRowUsed().RowNumber();
            var rng = ws.Range("A1:M" + lastRow);
            var rngTitle = ws.Range("A1:M1");

            //Formato del encabezado
            rngTitle.Style.Font.Bold = true;
            rngTitle.Style.Font.FontColor = XLColor.White;
            rngTitle.Style.Fill.BackgroundColor = XLColor.CadmiumGreen;  //XLColor.BlueGreen;   //FromArgb(0, 128, 0);

            rngTitle.SetAutoFilter();

            ws.Columns().AdjustToContents(); //Ajusta columnas de la hoja Datos                     

            //--- PIVOT TABLE ---

            // Add a new sheet for our pivot table
            var ptSheet = wb.Worksheets.Add("Tabla");
            // Create the pivot table, using the data from the "Datos" table
            var pt = ptSheet.PivotTables.Add("Tabla", ptSheet.Cell(1, 1), rng);
                       
            //Filters
            pt.ReportFilters.Add("Convenio");
            pt.ReportFilters.Add("Servicio");
            pt.ReportFilters.Add("Marca No Pos");

            //Rows
            var etapa=pt.RowLabels.Add("Etapa");
            etapa.SetSort(XLPivotSortType.Ascending);          

            //Values
            //Cuenta de Valor
            var cuentaValor=pt.Values.Add("Valor", "Cuenta de Valor");
            cuentaValor.SummaryFormula = XLPivotSummary.Count;

            //Porcentaje Cuenta de Valor
            var porcCuentaValor = pt.Values.Add("Valor", "% Facturas");
            porcCuentaValor.SummaryFormula = XLPivotSummary.Count;
            porcCuentaValor.ShowAsPercentageOfTotal();
            porcCuentaValor.NumberFormat.Format = "0.00%";

            //Promedio de días
            var dias = pt.Values.Add("Dias", "Promedio de Días");
            dias.SummaryFormula = XLPivotSummary.Average;
            dias.NumberFormat.Format = "0";

            //Suma de valor
            var sumValor = pt.Values.Add("Valor", "Suma de Valor");
            sumValor.SummaryFormula = XLPivotSummary.Sum;
            sumValor.NumberFormat.Format = "$###,###,###";

            //Porcentaje Suma de Valor
            var porcValor = pt.Values.Add("Valor", "%Valor");
            porcValor.ShowAsPercentageOfTotal();
            porcValor.NumberFormat.Format = "0.00%";

            pt.ClassicPivotTableLayout = true;            

            ptSheet.SetTabActive();

            ptSheet.ColumnWidth = 20;
            ptSheet.Columns("A").Width=50;

            var xlsFileName= AppDomain.CurrentDomain.BaseDirectory + @"Informe\InforDiario-" + DateTime.Now.ToString("yyyyMMddHHmm");
            wb.SaveAs(xlsFileName + ".xlsx");

            
            
        }
    }
}
