using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using Autodesk.Civil.DatabaseServices;


namespace Materialy
{
    public static class MaterialSectionExtension
    {
       public static double Area2(this MaterialSection materialSection)
        {
            double area = 0.0;
            SectionPointCollection sectionPts = materialSection.SectionPoints; //wierzchołki materiału

            using (Polyline boundary = new Polyline())
            {
                // zamknięta polilinia z punktów przekroju materiału
                for(int i = 0; i < sectionPts.Count; i++)
                {
                    Point2d pt = new Point2d(sectionPts[i].Location.X, sectionPts[i].Location.Y);
                    boundary.AddVertexAt(i, pt, 0, 0, 0);
                }
                boundary.Closed = true;
                //powierzhnia za pomocą Autodesk.AutoCAD.Geometry
                area = boundary.Area;
            }

            return area;
        }
    }
}