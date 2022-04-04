using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HolePlugin
{
    [TransactionAttribute(TransactionMode.Manual)]
    public class AddHole : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            // основной документ-файл
            Document arDoc = commandData.Application.ActiveUIDocument.Document;
            // связанный документ-файл
            Document ovDoc = arDoc.Application.Documents.OfType<Document>().Where(x => x.Title.Contains("ОВ")).FirstOrDefault();
            //если вдруг файл не будет обнаружен, то
            if (ovDoc == null)
            {
                TaskDialog.Show("Ошибка", "Не найден файл ОВ");
                return Result.Cancelled;
            }

            // убедимся, что в файл АР загружено семейство с отверстиями
            FamilySymbol familySymbol = new FilteredElementCollector(arDoc)
                .OfClass(typeof(FamilySymbol))
                .OfCategory(BuiltInCategory.OST_GenericModel)
                .OfType<FamilySymbol>()
                .Where(x => x.Name.Equals("Тип 1"))
                .FirstOrDefault();

            // в случае, когда семейство не найдено
            if (familySymbol == null)
            {
                TaskDialog.Show("Ошибка", "Не найдено данное семейство \"Отверстия\"");
                return Result.Cancelled;
            }

            // найдём все воздуховоды
            List<Duct> ducts = new FilteredElementCollector(ovDoc)
                .OfClass(typeof(Duct))
                .OfType<Duct>()
                .ToList();

            // найдём 3D вид
            View3D view3D = new FilteredElementCollector(arDoc)
                .OfClass(typeof(View3D))
                .OfType<View3D>()
                // важно, чтобы 3D вид не являлся шаблоном вида, поэтому доавим усл-е, что св-во IsTemplate не установлено
                .Where(x => !x.IsTemplate)
                .FirstOrDefault();

            // в случае, когда 3D вид не найден
            if (view3D == null)
            {
                TaskDialog.Show("Ошибка", "Не найден 3D вид");
                return Result.Cancelled;
            }

            // создаём объект ReferenceIntersector
            ReferenceIntersector referenceIntersector = new ReferenceIntersector(
                // мы хотим фильтровать по классу
                new ElementClassFilter(typeof(Wall)),
                // мы хотим найти элемент
                FindReferenceTarget.Element,
                view3D);

            Transaction tr = new Transaction(arDoc);
            tr.Start("Расстановка отверстий");

            if (!familySymbol.IsActive)
                familySymbol.Activate();

            // чтобы получить результат, т.е. набор пересекаемых стен. А применять будет с каждым воздуховодом, найденным в коллекции ducts
            foreach (Duct duct in ducts)
            {
                Line curve = (duct.Location as LocationCurve).Curve as Line;

                // теперь точку, можно получить из curve
                XYZ point = curve.GetEndPoint(0);

                // вектор направления
                XYZ direction = curve.Direction;

                // методу Find нужно передать 2 аргумента: исходная точка и направление из этой точки, а точку - из кривой
                List<ReferenceWithContext> intersections = referenceIntersector.Find(point, direction)
                    // ограничим нашу коллекцию. нужны объекты, у кот-х Proximity(расстояние) не превышает длину воздуховода
                    .Where(x => x.Proximity <= curve.Length)
                    // добавляем метод расширения Distinct - из всех объектов, кот. совпадают по зад. критерию, оставляет только один. в кач-ве аргумента указ-м экземпляр класса ReferenceWithContextElementEqualityComparer
                    .Distinct(new ReferenceWithContextElementEqualityComparer())
                    .ToList();

                // в итоге получим набор всех пересечений в виде объектов ReferenceWithContext
                foreach (ReferenceWithContext rwc in intersections)
                {
                    // определим точку вставки и добавим туда экземпляр семейства "Отверстия"
                    double proximity = rwc.Proximity;
                    Reference reference = rwc.GetReference();

                    // зная ElementId можно получить в док-те непосредственно стену, на кот-ю идёт эта сссылка
                    Wall wall = arDoc.GetElement(reference.ElementId) as Wall; // на выходе получим элемент
                    // получим уровень из стены
                    Level level = arDoc.GetElement(wall.LevelId) as Level;

                    XYZ pointHole = point + (direction * proximity);

                    // вставка экземпляра семейства
                    FamilyInstance holes = arDoc.Create.NewFamilyInstance(pointHole, familySymbol, wall, level, StructuralType.NonStructural);

                    // нам нужны длина и ширина, определим из диаметра трубы
                    Parameter width = holes.LookupParameter("Ширина");
                    Parameter height = holes.LookupParameter("Высота");

                    // установим значение Set, кот-й возьмём у воздуховода
                    width.Set(duct.Diameter);
                    height.Set(duct.Diameter);
                }
            }
            tr.Commit();

            return Result.Succeeded;
        }
    }

    public class ReferenceWithContextElementEqualityComparer : IEqualityComparer<ReferenceWithContext>
    {
        // определяет, убдут ли два заданных объекта (аргументы) определяемыми
        public bool Equals(ReferenceWithContext x, ReferenceWithContext y)
        {
            if (ReferenceEquals(x, y)) return true;
            if (ReferenceEquals(null, x)) return false;
            if (ReferenceEquals(null, y)) return false;

            var xReference = x.GetReference();

            var yReference = y.GetReference();

            // если у обоих элементов одинаковые ElementId, то получим точки на одной стене, то вернётся true
            // LinkedElementId - ElementId из связанного файла, они так же д.б. одинаковы
            return xReference.LinkedElementId == yReference.LinkedElementId
                       && xReference.ElementId == yReference.ElementId;
        }

        // должен возвращать HashCode объекта
        public int GetHashCode(ReferenceWithContext obj)
        {
            var reference = obj.GetReference();

            unchecked
            {
                return (reference.LinkedElementId.GetHashCode() * 397) ^ reference.ElementId.GetHashCode();
            }
        }
    }
}
