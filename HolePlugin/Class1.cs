using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HolePlugin
{
    [Transaction(TransactionMode.Manual)]

    public class AddHole : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            Document arDoc = commandData.Application.ActiveUIDocument.Document; //обращаемся к файлу АР
            Document ovDoc = arDoc.Application.Documents.OfType<Document>().Where(x => x.Title.Contains("ОВ")).FirstOrDefault(); //в файле АР находим вязанный файл ОВ
            if (ovDoc == null) //проверяем, что ОВ файл найден
            {
                TaskDialog.Show("Ошибка", "Не найден ОВ файл");
                return Result.Cancelled;
            }

            FamilySymbol familySymbol = new FilteredElementCollector(arDoc) //находим семейство отверстия
                .OfClass(typeof(FamilySymbol))
                .OfCategory(BuiltInCategory.OST_GenericModel)
                .OfType<FamilySymbol>()
                .Where(x => x.FamilyName.Equals("Отверстие"))
                .FirstOrDefault();
            if (familySymbol == null) //проверяем, что семейство отверстия загружено в модель
            {
                TaskDialog.Show("Ошибка", "Не найдено семейство \"Отверстия\"");
                return Result.Cancelled;
            }

            List<Duct> ducts = new FilteredElementCollector(ovDoc) //находим список воздуховодов
                .OfClass(typeof(Duct))
                .OfType<Duct>()
                .ToList();

            List<Pipe> pipes = new FilteredElementCollector(ovDoc) //находим список труб
                .OfClass(typeof(Pipe))
                .OfType<Pipe>()
                .ToList();

            View3D view3D = new FilteredElementCollector(arDoc) //находим 3d вид
                .OfClass(typeof(View3D))
                .OfType<View3D>()
                .Where(x => !x.IsTemplate)
                .FirstOrDefault();
            if (view3D == null) //проверяем, что зd вид найден
            {
                TaskDialog.Show("Ошибка", "Не найден 3D вид");
                return Result.Cancelled;
            }
            //создаем объект ReferenceIntersector при помощи конструктора
            ReferenceIntersector referenceIntersector = new ReferenceIntersector(new ElementClassFilter(typeof(Wall)), FindReferenceTarget.Element, view3D);

            Transaction transaction0 = new Transaction(arDoc);
            transaction0.Start("Расстановка отверстий");

            if (!familySymbol.IsActive)
            {
                familySymbol.Activate();
            }
            transaction0.Commit();

            Transaction transaction = new Transaction(arDoc);
            transaction.Start("Расстановка отверстий");

            //далее находим пересечения каждого из воздуховодов из коллекции ducts со стенами при помощи метода Find
            foreach (Duct d in ducts)
            {
                Line curve = (d.Location as LocationCurve).Curve as Line; //находим кривую из воздуховода и приводим ее к типу Линия
                XYZ point = curve.GetEndPoint(0); //получаем исходную точку
                XYZ direction = curve.Direction;

                //методу Find нужно передать 2 аргумента, исходную точку и вектор(направление)
                List<ReferenceWithContext> intersections = referenceIntersector.Find(point, direction)
                    .Where(x => x.Proximity <= curve.Length)
                    .Distinct(new ReferenceWithContextElementEqualityComparer()) //Distinct - метод расширения, который из всех объект подходящих по заданному критерию
                                                                                 //оставляет только один
                    .ToList(); //получаем набор всехх пересечений

                foreach (ReferenceWithContext refer in intersections) //теперь переберем все пересечения и вставим в точки экземпляры семейства отверстий
                {
                    double proximity = refer.Proximity; //расстояние до объекта
                    Reference reference = refer.GetReference(); //ссылка на объект
                    Wall wall = arDoc.GetElement(reference.ElementId) as Wall;
                    Level level = arDoc.GetElement(wall.LevelId) as Level;
                    XYZ pointHole = point + (direction * proximity); //чтобы найти точку вставки к стартовой точке прибавляем направление умноженное на расстояние

                    FamilyInstance hole = arDoc.Create.NewFamilyInstance(pointHole, familySymbol, wall, level, StructuralType.NonStructural);
                    Parameter width = hole.LookupParameter("Ширина");
                    Parameter height = hole.LookupParameter("Высота");
                    width.Set(d.Diameter);
                    height.Set(d.Diameter);
                }
            }

            foreach (Pipe p in pipes)
            {
                Line curve = (p.Location as LocationCurve).Curve as Line; //находим кривую из трубы и приводим ее к типу Линия
                XYZ point = curve.GetEndPoint(0); //получаем исходную точку
                XYZ direction = curve.Direction;

                //методу Find нужно передать 2 аргумента, исходную точку и вектор(направление)
                List<ReferenceWithContext> intersections = referenceIntersector.Find(point, direction)
                    .Where(x => x.Proximity <= curve.Length)
                    .Distinct(new ReferenceWithContextElementEqualityComparer()) //Distinct - метод расширения, который из всех объект подходящих по заданному критерию
                                                                                 //оставляет только один
                    .ToList(); //получаем набор всех пересечений

                foreach (ReferenceWithContext refer in intersections) //теперь переберем все пересечения и вставим в точки экземпляры семейства отверстий
                {
                    double proximity = refer.Proximity; //расстояние до объекта
                    Reference reference = refer.GetReference(); //ссылка на объект
                    Wall wall = arDoc.GetElement(reference.ElementId) as Wall;
                    Level level = arDoc.GetElement(wall.LevelId) as Level;
                    XYZ pointHole = point + (direction * proximity); //чтобы найти точку вставки к стартовой точке прибавляем направление умноженное на расстояние

                    FamilyInstance hole = arDoc.Create.NewFamilyInstance(pointHole, familySymbol, wall, level, StructuralType.NonStructural);
                    Parameter width = hole.LookupParameter("Ширина");
                    Parameter height = hole.LookupParameter("Высота");
                    width.Set(p.Diameter);
                    height.Set(p.Diameter);
                }
            }
            transaction.Commit();
            return Result.Succeeded;

        }
        public class ReferenceWithContextElementEqualityComparer : IEqualityComparer<ReferenceWithContext>
        {
            public bool Equals(ReferenceWithContext x, ReferenceWithContext y) //метод определяет будут ли 2 заданных объекта одинаковыми, возвращает true или false
            {
                if (ReferenceEquals(x, y)) return true;
                if (ReferenceEquals(null, x)) return false;
                if (ReferenceEquals(null, y)) return false;

                var xReference = x.GetReference();
                var yReference = y.GetReference();
                //проверяем если у двух элементов одинаковый ElementId - мы получим точки на одной стене
                return xReference.LinkedElementId == yReference.LinkedElementId //возвращает Id выбранных элементов из связанного файла
                    && xReference.ElementId == yReference.ElementId; //возвращает Id выбранных элементов из файла
            }

            public int GetHashCode(ReferenceWithContext obj) //возвращает хэшкод объекта
            {
                var reference = obj.GetReference();
                unchecked
                {
                    return (reference.LinkedElementId.GetHashCode() * 397) ^ reference.ElementId.GetHashCode();
                }
            }
        }
    }
}
