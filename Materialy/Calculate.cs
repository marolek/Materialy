using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;


namespace Materialy
{
    class Calculate
    {
        public static double Area(Point2dCollection p2)
        {
            double result = 0.0;

            // Create a polyline with 5 points
            using (Polyline acPoly = new Polyline())
            {
                for(int i = 0; i < p2.Count; i++)
                {
                    acPoly.AddVertexAt(i, p2[i], 0, 0, 0);
                }
                acPoly.Closed = true;
                result = acPoly.Area;
            }

            return result;
        }
    }
}