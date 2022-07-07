using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;


namespace Materialy
{
    class Calculate
    {
        //Obliczanie powierzchni z polilinii utworzone z kolekcji punktów
        public static double Area(Point2dCollection p2)
        {
            double result = 0.0;

            // Tworzenie zamkniętej polilinii
            using (Polyline acPoly = new Polyline())
            {
                for(int i = 0; i < p2.Count; i++)
                {
                    acPoly.AddVertexAt(i, p2[i], 0, 0, 0);
                }
                acPoly.Closed = true;
                //powierzhnia za pomocą Autodesk.AutoCAD.Geometry
                result = acPoly.Area;
            }

            return result;
        }
    }
}