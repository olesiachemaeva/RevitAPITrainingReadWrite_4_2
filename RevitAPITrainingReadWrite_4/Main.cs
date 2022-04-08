using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.UI;
using Newtonsoft.Json;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using RevitAPITrainingReadWrite;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
//using System.Windows.Forms;


//Выведите в файл Excel следующие значения всех труб: имя типа трубы, наружный
//диаметр трубы, внутренний диаметр трубы, длина трубы.




namespace RevitAPITrainingReadWrite_4
{
    [Transaction(TransactionMode.Manual)]
    public class Main : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
          
                UIApplication uiapp = commandData.Application;
                UIDocument uidoc = uiapp.ActiveUIDocument;
                Document doc = uidoc.Document;


            string pipeInfo = string.Empty;

            var pipes = new FilteredElementCollector(doc)  //элемент трубы
                .OfClass(typeof(Pipe))
                .Cast<Pipe>()
                .ToList();




            string excelPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "pipes.xlsx");
            using (FileStream stream = new FileStream(excelPath, FileMode.Create, FileAccess.Write))
            {
                IWorkbook workbook = new XSSFWorkbook();
                ISheet sheet = workbook.CreateSheet("pipeInfo");
                int rowIndex = 0;
                foreach (var pipe in pipes)
                {
                    sheet.SetCellValue(rowIndex, columnIndex: 0, pipe.get_Parameter(BuiltInParameter.ELEM_TYPE_PARAM)
                        .AsValueString());
                    sheet.SetCellValue(rowIndex, columnIndex: 1, pipe.get_Parameter(BuiltInParameter.RBS_PIPE_OUTER_DIAMETER)
                        .AsValueString());
                    sheet.SetCellValue(rowIndex, columnIndex: 2, pipe.get_Parameter(BuiltInParameter.RBS_PIPE_INNER_DIAM_PARAM)
                        .AsValueString());
                    sheet.SetCellValue(rowIndex, columnIndex: 3, pipe.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH)
                        .AsValueString());

                    rowIndex++;

                }
                workbook.Write(stream);
                workbook.Close();
            }
            return Result.Succeeded;
        }
    }
}
