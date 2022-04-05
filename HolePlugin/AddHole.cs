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
    [TransactionAttribute(TransactionMode.Manual)]
    public class AddHole : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            // основной документ-файл
            Document arDoc = commandData.Application.ActiveUIDocument.Document;
            Document ovDoc = arDoc.Application.Documents.OfType<Document>().Where(x => x.Title.Contains("ОВ")).FirstOrDefault();

            if (ovDoc == null)
            {
                TaskDialog.Show("Ошибка", "Не найден файл ОВ");
                return Result.Cancelled;
            }

            List<Duct> ducts = GetDucts(ovDoc);
            List<Pipe> pipes = GetPipes(ovDoc);
            View3D view3D = GetView3D(arDoc);
            FamilySymbol familySymbolRectangleHole = GetFamilySymbol1(arDoc);
            FamilySymbol familySymbolRoundHole = GetFamilySymbol2(arDoc);

            // в случае, когда семейство не найдено
            if (familySymbolRectangleHole == null &&
                familySymbolRoundHole == null)
            {
                TaskDialog.Show("Ошибка", $"Не найдено данное семейство:{Environment.NewLine}\"Отверстие\"{Environment.NewLine}\"Круглое отверстие\""); ;
                return Result.Cancelled;
            }

            try
            {
                CreateRectangleHole(arDoc, ducts, view3D, familySymbolRectangleHole);
                CreateRoundHole(arDoc, pipes, view3D, familySymbolRoundHole);
            }
            catch (OperationCanceledException)
            {
                return Result.Cancelled;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                return Result.Cancelled;
            }

            return Result.Succeeded;
        }

        private static List<Pipe> GetPipes(Document _ovDoc)
        {
            // найдём все трубы
            List<Pipe> pipes = new FilteredElementCollector(_ovDoc)
                .OfClass(typeof(Pipe))
                .OfType<Pipe>()
                .ToList();

            return pipes;
        }

        private static List<Duct> GetDucts(Document ovDoc)
        {
            // найдём все воздуховоды
            List<Duct> ducts = new FilteredElementCollector(ovDoc)
                .OfClass(typeof(Duct))
                .OfType<Duct>()
                .ToList();

            return ducts;
        }

        private static View3D GetView3D(Document arDoc)
        {
            // найдём 3D вид
            View3D view3D = new FilteredElementCollector(arDoc)
                .OfClass(typeof(View3D))
                .OfType<View3D>()
                .Where(x => !x.IsTemplate)
                .FirstOrDefault();

            if (view3D == null)
            {
                TaskDialog.Show("Ошибка", "Не найден 3D вид");
                return null;
            }

            return view3D;
        }

        private static FamilySymbol GetFamilySymbol1(Document arDoc)
        {
            FamilySymbol familySymbolRectangleHole = new FilteredElementCollector(arDoc)
                .OfClass(typeof(FamilySymbol))
                .OfCategory(BuiltInCategory.OST_GenericModel)
                .OfType<FamilySymbol>()
                .Where(x => x.FamilyName.Equals("Отверстие"))
                .FirstOrDefault();

            return familySymbolRectangleHole;
        }

        private static FamilySymbol GetFamilySymbol2(Document arDoc)
        {
            // убедимся, что в файл АР загружено семейство с отверстиями
            FamilySymbol familySymbolRoundHole = new FilteredElementCollector(arDoc)
                .OfClass(typeof(FamilySymbol))
                .OfCategory(BuiltInCategory.OST_GenericModel)
                .OfType<FamilySymbol>()
                .Where(x => x.FamilyName.Equals("Круглое отверстие"))
                .FirstOrDefault();

            return familySymbolRoundHole;
        }

        private static void CreateRectangleHole(Document arDoc, List<Duct> ducts, View3D view3D, FamilySymbol familySymbolRectangleHole)
        {
            ReferenceIntersector referenceIntersector = new ReferenceIntersector(
                new ElementClassFilter(typeof(Wall)),
                FindReferenceTarget.Element,
                view3D);

            Transaction ts = new Transaction(arDoc);
            ts.Start("Активация отверстий");
            if (!familySymbolRectangleHole.IsActive)
            {
                familySymbolRectangleHole.Activate();
            }
            ts.Commit();

            Transaction tr = new Transaction(arDoc);
            tr.Start("Расстановка прямоугольных отверстий");

            foreach (Duct duct in ducts)
            {
                Line curve = (duct.Location as LocationCurve).Curve as Line;
                XYZ point = curve.GetEndPoint(0);
                XYZ direction = curve.Direction;

                List<ReferenceWithContext> intersections = referenceIntersector.Find(point, direction)
                    .Where(x => x.Proximity <= curve.Length)
                    .Distinct(new ReferenceWithContextElementEqualityComparer())
                    .ToList();

                foreach (ReferenceWithContext rwc in intersections)
                {
                    double proximity = rwc.Proximity;
                    Reference reference = rwc.GetReference();
                    Wall wall = arDoc.GetElement(reference.ElementId) as Wall;
                    Level level = arDoc.GetElement(wall.LevelId) as Level;
                    XYZ pointHole = point + (direction * proximity);

                    FamilyInstance holes = arDoc.Create.NewFamilyInstance(pointHole, familySymbolRectangleHole, wall, level, StructuralType.NonStructural);
                    Parameter width = holes.LookupParameter("Ширина");
                    Parameter height = holes.LookupParameter("Высота");

                    width.Set(duct.Diameter);
                    height.Set(duct.Diameter);
                }
            }

            tr.Commit();
        }

        private static void CreateRoundHole(Document arDoc, List<Pipe> pipes, View3D view3D, FamilySymbol familySymbolRoundHole)
        {
            ReferenceIntersector referenceIntersector = new ReferenceIntersector(
                new ElementClassFilter(typeof(Wall)),
                FindReferenceTarget.Element,
                view3D);

            Transaction tr = new Transaction(arDoc);
            tr.Start("Активация круглых отверстий");
            if (!familySymbolRoundHole.IsActive)
            {
                familySymbolRoundHole.Activate();
            }

            tr.Commit();

            Transaction ts = new Transaction(arDoc);
            ts.Start("Расстановка круглых отверстий");

            foreach (Pipe pipe in pipes)
            {
                Line curve = (pipe.Location as LocationCurve).Curve as Line;
                XYZ point = curve.GetEndPoint(0);
                XYZ direction = curve.Direction;

                List<ReferenceWithContext> intersections = referenceIntersector.Find(point, direction)
                    .Where(x => x.Proximity <= curve.Length)
                    .Distinct(new ReferenceWithContextElementEqualityComparer())
                    .ToList();

                foreach (ReferenceWithContext rwc in intersections)
                {
                    double proximity = rwc.Proximity;
                    Reference reference = rwc.GetReference();
                    Wall wall = arDoc.GetElement(reference.ElementId) as Wall;
                    Level level = arDoc.GetElement(wall.LevelId) as Level;
                    XYZ pointHole = point + (direction * proximity);

                    FamilyInstance holes = arDoc.Create.NewFamilyInstance(pointHole, familySymbolRoundHole, wall, level, StructuralType.NonStructural);
                    Parameter rad = holes.LookupParameter("Радиус");
                    double diam = pipe.get_Parameter(BuiltInParameter.RBS_PIPE_OUTER_DIAMETER).AsDouble();

                    rad.Set(diam / 2);
                }
            }

            ts.Commit();
        }
    }

    public class ReferenceWithContextElementEqualityComparer : IEqualityComparer<ReferenceWithContext>
    {
        public bool Equals(ReferenceWithContext x, ReferenceWithContext y)
        {
            if (ReferenceEquals(x, y)) return true;
            if (ReferenceEquals(null, x)) return false;
            if (ReferenceEquals(null, y)) return false;

            var xReference = x.GetReference();
            var yReference = y.GetReference();

            return xReference.LinkedElementId == yReference.LinkedElementId
                       && xReference.ElementId == yReference.ElementId;
        }

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
