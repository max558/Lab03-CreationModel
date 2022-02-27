﻿using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Structure;
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

            //using (Transaction ts = new Transaction(doc, "Создание модели"))
            //{
            //    ts.Start();
            //Создание коробки из стен
            List<Wall> wallListCreat = CreateRectangleWall(doc, 10000, 5000, level1, level2);

            //вставка двери
            List<XYZ> pointListDoor = GetPointInstance(doc, wallListCreat[0], null, 0, 0, 0);
            XYZ pointDooor = pointListDoor[0];
            FamilyInstance door = AddInstance(commandData, level1, wallListCreat[0], pointDooor, BuiltInCategory.OST_Doors, "Одиночные-Щитовые", "0915 x 2134 мм");

            //Вставка окон

            List<FamilyInstance> windowList = new List<FamilyInstance>();
            double widthWindow = UnitUtils.ConvertToInternalUnits(915, UnitTypeId.Millimeters);
            double lengthWindowCorner = UnitUtils.ConvertToInternalUnits(1500, UnitTypeId.Millimeters);
            double lengthWindowWindow = UnitUtils.ConvertToInternalUnits(1200, UnitTypeId.Millimeters);

            foreach (var wall in wallListCreat)
            {
                List<XYZ> pointWindowList = GetPointInstance(doc, wall, door, widthWindow, lengthWindowCorner, lengthWindowWindow);
                foreach (var pointWind in pointWindowList)
                {
                    FamilyInstance window = AddInstance(commandData, level1, wall, pointWind, BuiltInCategory.OST_Windows, "Фиксированные", "0915 x 1220 мм");
                    double heightWindow = UnitUtils.ConvertToInternalUnits(934, UnitTypeId.Millimeters);
                    ParametrUtils.SetParametrFamilyInstance<double>(doc, window, "Высота нижнего бруса", heightWindow);
                    windowList.Add(window);
                }
            }


            //    ts.Commit();
            //}

            return Result.Succeeded;
        }

        /*
         *** === Получение уровня по его имени === ***
         * Входные данные: 
         * name - строковое имя уровня
         * Вывод: уровень
         */
        private Level GetLevelofName(string name)
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
        private List<Wall> CreateRectangleWall(Document doc,
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
            points.Add(new XYZ(-dx, dy, 0));
            points.Add(new XYZ(dx, dy, 0));
            points.Add(new XYZ(dx, -dy, 0));
            points.Add(new XYZ(-dx, -dy, 0));
            points.Add(new XYZ(-dx, dy, 0));

            using (Transaction ts = new Transaction(doc, "Создание коробки из стен"))
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

        /*
         *** === Создание экземпляра семейства === ***
         * Входные данные:
         * commandData - 
         * level - уровень
         * wall - стена основа
         * builtInCategory - категория семейства
         * nameFamaly - строковое имя семейства
         * nameInstance - строковое имя типа
         * Выходные данные:
         * создание экземпляра семейства на основе стены
         */
        private FamilyInstance AddInstance(ExternalCommandData commandData,
                                           Level level, Wall wall, XYZ pointInst,
                                           BuiltInCategory builtInCategory,
                                           string nameFamaly, string nameInstance)
        {
            Document doc = commandData.Application.ActiveUIDocument.Document;
            List<FamilySymbol> familySymbols = SelectionUtils.SelectAllElementCategoryType<FamilySymbol>(commandData, builtInCategory);
            FamilySymbol familySymbType = familySymbols
                .Where(x => x.Name.Equals(nameInstance))
                .Where(x => x.FamilyName.Equals(nameFamaly))
                .FirstOrDefault();

            FamilyInstance familyInst;
            using (Transaction tsa = new Transaction(doc, "Вставка экземпляра семейства"))
            {
                tsa.Start();

                if (!familySymbType.IsActive)
                {
                    familySymbType.Activate();
                }
                familyInst = doc.Create.NewFamilyInstance(pointInst, familySymbType, wall, level, StructuralType.NonStructural);

                tsa.Commit();
            }

            return familyInst;
        }

        /*
         *** === Определение точек вставки экземпляров === ***
         * Входные данные:
         * wall - выбранная стена
         * doorInst - существующий экземпляр двери
         * widthWindow - ширина окна
         * lengthWindowCorner - расстояние между углом здания и окном
         * lengthWindowWindow - расстояние между окнами (между границами окон)
         * Выходные данные:
         * При наличии двери список равномерно распределенных точек вставки
         * Если экземпляр двери null - то список центральной (относительно стены) точки вставки
         */
        private List<XYZ> GetPointInstance(Document doc, Wall wall, FamilyInstance doorInst,
                                            double widthWindow,
                                            double lengthWindowCorner, double lengthWindowWindow)
        {
            List<XYZ> resPointList = new List<XYZ>();

            LocationCurve hostCurve = wall.Location as LocationCurve;
            double lengthWall = hostCurve.Curve.Length;
            Curve curveWall = hostCurve.Curve;

            XYZ pointStart = hostCurve.Curve.GetEndPoint(0);
            XYZ pointEnd = hostCurve.Curve.GetEndPoint(1);

            if (doorInst == null)
            {
                //Определение центральной точки стены
                XYZ point = (pointStart + pointEnd) / 2;

                resPointList.Add(point);
                return resPointList;
            }

            //Определяем параметры двери, ширину и точку вставки
            FamilySymbol doorSymb = doorInst.Symbol;
            Element doorElement = doc.GetElement(doorSymb.Id);
            Parameter widthDoorParam = doorElement.LookupParameter("Ширина");
            double widthDoor = 0;
            if (widthDoorParam.StorageType == StorageType.Double)
            {
                widthDoor = widthDoorParam.AsDouble();
            }
            ElementId wallHostId = doorInst.Host.Id;

            //Проверка двери на наличие в стене
            if (wall.Id == wallHostId)
            {
                XYZ pointDoorInst = (doorInst.Location as LocationPoint).Point;

                //Поиск точек начала и конца
                Curve dorCurv1 = Line.CreateBound(pointStart, pointDoorInst);
                XYZ doorStart = dorCurv1.Evaluate((dorCurv1.Length - widthDoor / 2), false);
                Curve dorCurv2 = Line.CreateBound(pointDoorInst, pointEnd);
                XYZ doorEnd = dorCurv2.Evaluate(widthDoor / 2, false);
                List<XYZ> pointList1 = FoundPointInsert(pointStart, doorStart, widthWindow, lengthWindowCorner, lengthWindowWindow);
                List<XYZ> pointList2 = FoundPointInsert(doorEnd, pointEnd, widthWindow, lengthWindowCorner, lengthWindowWindow);
                foreach (var item in pointList2)
                {
                    pointList1.Add(item);
                }
                resPointList = pointList1;
            }
            else
            {
                resPointList = FoundPointInsert(pointStart, pointEnd, widthWindow, lengthWindowCorner, lengthWindowWindow);
            }



            return resPointList;
        }

        /*
         *** === Формирование списка точек на линии между 2-мя крайними точками === ***
         * Входные данные:
         * widthWindow - ширина окна
         * lengthWindowCorner - расстояние между углом здания и окном
         * lengthWindowWindow - расстояние между окнами (между границами окон)
         * Выходные данные:
         * Список точек на прямой
         */
        private List<XYZ> FoundPointInsert(XYZ pointStart, XYZ pointEnd,
                                           double widthWindow, double lengthWindowCorner,
                                           double lengthWindowWindow)
        {
            List<XYZ> resPointList = new List<XYZ>();
            Curve curve = Line.CreateBound(pointStart, pointEnd);
            //определяем расстояние между вставляемыми экземплярами
            double nCalc = (curve.Length - 2 * lengthWindowCorner + lengthWindowWindow) / (widthWindow + lengthWindowWindow);
            int nInt = Convert.ToInt32(nCalc);

            double distanceWindow = widthWindow + lengthWindowWindow;

            if (
                (nCalc - nInt) > 0
                && nInt > 1)
            {
                distanceWindow = widthWindow + ((curve.Length - (2 * lengthWindowCorner + nInt * widthWindow)) / (nInt - 1));
            }
            else
            {
                if (nInt == 1)
                {
                    distanceWindow = curve.Length / 2;
                }
            }

            // определяем точки на прямой по расчитанному расстоянию
            for (int i = 0; i < nInt; i++)
            {
                XYZ point = curve.Evaluate(lengthWindowCorner + widthWindow / 2 + (i * distanceWindow), false);
                resPointList.Add(point);
            }
            return resPointList;
        }
    }
}
