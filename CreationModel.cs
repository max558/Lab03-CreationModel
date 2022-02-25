using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using RevitAPITreningLibrary;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CreationModel
{
    [Transaction(TransactionMode.Manual)]
    public class CreationModel : IExternalCommand
    {
        public List<Level> LeveList { get; set; } = new List<Level>();

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            Document doc = commandData.Application.ActiveUIDocument.Document;
            LeveList = SelectionUtils.SelectAllElement<Level>(commandData);
            Level level1 = GetLevelofName("Уровень 1");
            Level level2 = GetLevelofName("Уровень 2");

            //Создание коробки из стен
            List<Wall> wallListCreat = CreateRectangleWall(doc, 10000, 5000, level1, level2);

 
            return Result.Succeeded;
        }

        /*
         *** === Получение уровня по его имени === ***
         * Входные данные: 
         * name - строковое имя уровня
         * Вывод: уровень
         */
        private  Level GetLevelofName(string name)
        {
            if (name == string.Empty)
            {
                return null;
            }
            Level res = LeveList
                .Where(x => x.Name.Equals(name))
                .FirstOrDefault();
            return res;
        }

        /*
         *** === Создание прямоугольника из стен === ***
         * Входные данные:
         * doc - активный документ
         * width, depth - ширина и глубина прямоугольника в мм
         * levelMin - уровнь, на котором создаются стены
         * levelMax - уровнь, до которого создаются стены
         * Выходные данные:
         * В случае успеха - список созданных стен, в противном случае - пустой список
         */
        private  List<Wall> CreateRectangleWall(Document doc,
                                               double width,
                                               double depth,
                                               Level levelMin,
                                               Level levelMax)
        {
            List<Wall> resWall = new List<Wall>();

            if (width == 0
                && depth == 0)
            {
                return resWall;
            }

            if (levelMin == null
                && levelMax == null)
            {
                return resWall;
            }

            double inWidth = UnitUtils.ConvertToInternalUnits(width, UnitTypeId.Millimeters);
            double inLength = UnitUtils.ConvertToInternalUnits(depth, UnitTypeId.Millimeters);
            double dx = inWidth / 2;
            double dy = inLength / 2;
            List<XYZ> points = new List<XYZ>();
            points.Add(new XYZ(-dx, -dy, 0));
            points.Add(new XYZ(dx, -dy, 0));
            points.Add(new XYZ(dx, dy, 0));
            points.Add(new XYZ(-dx, dy, 0));
            points.Add(new XYZ(-dx, -dy, 0));

            using (Transaction ts = new Transaction(doc, "Создание стен"))
            {
                ts.Start();
                for (int i = 0; i < points.Count - 1; i++)
                {
                    Line line = Line.CreateBound(points[i], points[i + 1]);
                    Wall wall = Wall.Create(doc, line, levelMin.Id, false);
                    wall.get_Parameter(BuiltInParameter.WALL_HEIGHT_TYPE).Set(levelMax.Id);
                    resWall.Add(wall);
                }
                ts.Commit();
            }
            return resWall;
        }
    }
}
