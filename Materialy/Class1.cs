using System.Runtime;
using System;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.GraphicsInterface;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD;
using Autodesk.Civil.ApplicationServices;
using Autodesk.Civil.DatabaseServices;

using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
//using System.Reflection;
//using System.Runtime.CompilerServices;
//using System.Security;
//using System.Text;
//using System.Threading.Tasks;
//using Autodesk.AutoCAD.Interop;
//using Autodesk.AECC.Interop.UiLand;

namespace Materialy
{
    public class Class1 : IExtensionApplication
    {
        public int polecenie = 0;

        [Autodesk.AutoCAD.Runtime.CommandMethod("ListaMaterialowSEP")]
        public void ListaMaterialow()
        {
            CivilDocument Civdoc = Autodesk.Civil.ApplicationServices.CivilApplication.ActiveDocument;
            Document acDoc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = acDoc.Editor;
            Database acCurDb = acDoc.Database;
            ObjectIdCollection alignments = Civdoc.GetAlignmentIds();

            ed.WriteMessage("\nWywołano polecenie: {0}", polecenie);

            using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
            {
                try
                {
                    string RaportPath = String.Format("{0}\\raporty\\", Path.GetDirectoryName(acCurDb.Filename));
                    System.IO.Directory.CreateDirectory(RaportPath);
                    foreach (ObjectId AlignID in alignments) 
                    {
                        Autodesk.AutoCAD.DatabaseServices.DBObject dbObj = acTrans.GetObject(AlignID, OpenMode.ForRead);
                        Alignment LiniaTrasowania = (Alignment)dbObj;
                        ObjectIdCollection samplelinesGroupID = LiniaTrasowania.GetSampleLineGroupIds();
                        
                        foreach (ObjectId SampleLineGroupID in samplelinesGroupID) 
                        {
                            SampleLineGroup GrupaLiniiSamplowania = (SampleLineGroup)acTrans.GetObject(SampleLineGroupID, OpenMode.ForRead);

                            QTOMaterialListCollection ListyMaterialow = GrupaLiniiSamplowania.MaterialLists;
                            foreach (QTOMaterialList ListaMaterialow in ListyMaterialow) 
                            {

                                string OutputFileName = String.Format(ListyMaterialow.Count>1 ? "{0}{1}_{2}.txt" : "{0}{1}.txt", RaportPath, LiniaTrasowania.Name, ListaMaterialow.Name);
                                ed.WriteMessage("\nPlik raportu: {0}", OutputFileName);
                                System.IO.StreamWriter OutputFile = new System.IO.StreamWriter(OutputFileName);
                           
                                //Wiersz nagłówków
                                OutputFile.Write("Pikieta\t");
                                foreach (QTOMaterial Material in ListaMaterialow)
                                {
                                    string MaterialInfo = String.Format(Material.Guid == ListaMaterialow.Last().Guid ? "{0}\n" : "{0}\t", Material.Name);
                                    OutputFile.Write(MaterialInfo);
                                }

                                //Wiersz jednostek
                                OutputFile.Write("Jednostka\t");
                                for (int i = 1; i < ListaMaterialow.Count; i++) 
                                {
                                    OutputFile.Write("[ m² ]\t");
                                }
                                OutputFile.Write("[ m² ]\n");

                                foreach (ObjectId SampleLineID in GrupaLiniiSamplowania.GetSampleLineIds()) 
                                {
                                    SampleLine LiniaSamplowania = (SampleLine)acTrans.GetObject(SampleLineID, OpenMode.ForRead);
                                    OutputFile.Write("{0}\t",LiniaSamplowania.Station);
                                    foreach (QTOMaterial Material in ListaMaterialow)
                                    {
                                        ObjectId MaterialSectionID = LiniaSamplowania.GetMaterialSectionId(ListaMaterialow.Guid, Material.Guid);
                                        MaterialSection MaterialSectionObject = (MaterialSection)acTrans.GetObject(MaterialSectionID, OpenMode.ForRead);
                                        
                                        double area = 0.0;
                                        Point2dCollection p2 = new Point2dCollection();
                                        SectionPointCollection sPts = MaterialSectionObject.SectionPoints;
                                        foreach(SectionPoint pt in sPts)
                                        {
                                            p2.Add(new Point2d(pt.Location.X, pt.Location.Y));
                                        }
                                        area = Calculate.Area(p2);
                                        OutputFile.Write(Material.Guid == ListaMaterialow.Last().Guid ? "{0}\n" : "{0}\t",area);
                                    }
                                }

                                OutputFile.Close();
                                string OldRaportPath = String.Format("{0}old\\", RaportPath);
                                System.IO.Directory.CreateDirectory(OldRaportPath);

                                String czas = DateTime.Now.ToString("yyyyddMM");
                                string destFile = String.Format(ListyMaterialow.Count>1 ? "{0}{1}_{2}_{3}.txt" : "{0}{1}_{3}.txt", OldRaportPath, LiniaTrasowania.Name, ListaMaterialow.Name, czas);
                                System.IO.File.Copy(OutputFileName, destFile, true);
                            }
                        }
                    }
                }
                catch (System.Exception ex) { }
                acTrans.Commit();
            }

        }

        [Autodesk.AutoCAD.Runtime.CommandMethod("ListaMaterialow")] 
        public void ListaMaterialow2()
        {
            CivilDocument Civdoc = Autodesk.Civil.ApplicationServices.CivilApplication.ActiveDocument;
            Document acDoc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = acDoc.Editor;
            Database acCurDb = acDoc.Database;
            ObjectIdCollection alignments = Civdoc.GetAlignmentIds();

            using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
            {
                try
                {
                    string RaportPath = String.Format("{0}\\raporty\\", Path.GetDirectoryName(acCurDb.Filename));
                    System.IO.Directory.CreateDirectory(RaportPath);

                    foreach (ObjectId AlignID in alignments) 
                    {
                        Autodesk.AutoCAD.DatabaseServices.DBObject dbObj = acTrans.GetObject(AlignID, OpenMode.ForRead);
                        Alignment LiniaTrasowania = (Alignment)dbObj;
                        ObjectIdCollection samplelinesGroupID = LiniaTrasowania.GetSampleLineGroupIds();

                        string OutputFileName = String.Format("{0}{1}.txt", RaportPath, LiniaTrasowania.Name);
                        ed.WriteMessage("\nPlik raportu: {0}", OutputFileName);
                        System.IO.StreamWriter OutputFile = new System.IO.StreamWriter(OutputFileName);

                        foreach (ObjectId SampleLineGroupID in samplelinesGroupID) 
                        {
                            SampleLineGroup GrupaLiniiSamplowania = (SampleLineGroup)acTrans.GetObject(SampleLineGroupID, OpenMode.ForRead);

                            QTOMaterialListCollection ListyMaterialow = GrupaLiniiSamplowania.MaterialLists;

                            int j=0;

                            //Wiersz nagłówków
                            OutputFile.Write("Pikieta\t");
                            foreach (QTOMaterialList ListaMaterialow in ListyMaterialow) 
                            {
                                foreach (QTOMaterial Material in ListaMaterialow)
                                {
                                    string MaterialInfo = "";
                                    if (ListaMaterialow.Guid == ListyMaterialow.Last().Guid)
                                    {
                                        MaterialInfo = String.Format(Material.Guid == ListaMaterialow.Last().Guid ? "{0}\n" : "{0}\t", Material.Name);
                                    }
                                    else
                                    {
                                        MaterialInfo = String.Format("{0}\t", Material.Name);
                                    }
                                    OutputFile.Write(MaterialInfo);
                                    j++;
                                }
                            }
                            
                            // //Wiersz jednostek
                            // OutputFile.Write("Jednostka\t");
                            // for (int i = 1; i < j; i++) 
                            // {
                            //     OutputFile.Write("[ m² ]\t");
                            // }
                            // OutputFile.Write("[ m² ]\n");

                            foreach (ObjectId SampleLineID in GrupaLiniiSamplowania.GetSampleLineIds()) 
                            {
                                SampleLine LiniaSamplowania = (SampleLine)acTrans.GetObject(SampleLineID, OpenMode.ForRead);
                                OutputFile.Write("{0}\t",LiniaSamplowania.Station);
                                                            
                                foreach (QTOMaterialList ListaMaterialow in ListyMaterialow) 
                                {
                                    foreach (QTOMaterial Material in ListaMaterialow)
                                    {
                                        ObjectId MaterialSectionID = LiniaSamplowania.GetMaterialSectionId(ListaMaterialow.Guid, Material.Guid);
                                        MaterialSection MaterialSectionObject = (MaterialSection)acTrans.GetObject(MaterialSectionID, OpenMode.ForRead);
                                        
                                        double area = 0.0;
                                        Point2dCollection p2 = new Point2dCollection();
                                        SectionPointCollection sPts = MaterialSectionObject.SectionPoints;
                                        foreach(SectionPoint pt in sPts)
                                        {
                                            p2.Add(new Point2d(pt.Location.X, pt.Location.Y));
                                        }
                                        
                                        area = Calculate.Area(p2);
                                        if (ListaMaterialow.Guid == ListyMaterialow.Last().Guid)
                                        {
                                            OutputFile.Write(Material.Guid == ListaMaterialow.Last().Guid ? "{0}\n" : "{0}\t",area);
                                        }
                                        else
                                        {
                                            OutputFile.Write("{0}\t",area);
                                        }
                                    }
                                }
                            }
                        }
                        OutputFile.Close();
                        string OldRaportPath = String.Format("{0}old\\", RaportPath);
                        System.IO.Directory.CreateDirectory(OldRaportPath);

                        String czas = DateTime.Now.ToString("yyyyMMdd");
                        string destFile = String.Format("{0}{1}_{2}.txt", OldRaportPath, LiniaTrasowania.Name, czas);
                        System.IO.File.Copy(OutputFileName, destFile, true);
                    }
                }
                catch (System.Exception ex) { }
                acTrans.Commit();
            }
        }

        #region IExtensionApplication Members

        public void Initialize()
        {
            Document acDoc = Application.DocumentManager.MdiActiveDocument;
            acDoc.Editor.WriteMessage("\nWczytano dodatek generujący raport materiałów - Użyj polecenia: ListaMaterialow");
        }

        public void Terminate()
        {
        }
        
        #endregion
    }
}
