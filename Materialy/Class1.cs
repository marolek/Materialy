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
        //do identyfikacji polecenia ListaMaterialowSEP lub ListaMaterialow - ToDo
        //public int polecenie = 0;

        //ListaMaterialowSEP - Lista materiałów w podziale na grupy materiałowe
        [Autodesk.AutoCAD.Runtime.CommandMethod("ListaMaterialowSEP")]
        public void ListaMaterialowSEP()
        {
            CivilDocument Civdoc = Autodesk.Civil.ApplicationServices.CivilApplication.ActiveDocument;
            Document acDoc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = acDoc.Editor;
            Database acCurDb = acDoc.Database;
            ObjectIdCollection alignments = Civdoc.GetAlignmentIds();

            //ed.WriteMessage("\nWywołano polecenie: {0}", polecenie);

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

                                string OutputFileName = String.Format(ListyMaterialow.Count>1 ? "{0}{1}_{2}_{3}.txt" : "{0}{1}_{3}.txt", RaportPath, LiniaTrasowania.Name, ListaMaterialow.Name, GrupaLiniiSamplowania.Name);
                                ed.WriteMessage("\nPlik raportu: {0}", OutputFileName);
                                System.IO.StreamWriter OutputFile = new System.IO.StreamWriter(OutputFileName);
                           
                                //Wiersz nagłówków
                                string naglowekTxt = "Pikieta";
                                foreach (QTOMaterial Material in ListaMaterialow)
                                    naglowekTxt += String.Format("\t{0}", Material.Name);

                                OutputFile.WriteLine(naglowekTxt);

                                // // Wiersz jednostek
                                // string jednostkiTxt = "Jednostka";
                                // jednostkiTxt += String.Concat(Enumerable.Repeat("\t[ m² ]", ListaMaterialow.Count));
                                // OutputFile.WriteLine(jednostkiTxt);

                                foreach (ObjectId SampleLineID in GrupaLiniiSamplowania.GetSampleLineIds()) 
                                {
                                    SampleLine LiniaSamplowania = (SampleLine)acTrans.GetObject(SampleLineID, OpenMode.ForRead);
                                    
                                    string wierszDanych = LiniaSamplowania.Station.ToString();
                                    foreach (QTOMaterial Material in ListaMaterialow)
                                    {
                                        ObjectId MaterialSectionID = LiniaSamplowania.GetMaterialSectionId(ListaMaterialow.Guid, Material.Guid);
                                        MaterialSection MaterialSectionObject = (MaterialSection)acTrans.GetObject(MaterialSectionID, OpenMode.ForRead);
                                        
                                        double area = MaterialSectionObject.Area2();
                                        wierszDanych += String.Format("\t{0}", area);
                                    }
                                    OutputFile.WriteLine(wierszDanych);
                                }

                                OutputFile.Close();
                                string OldRaportPath = String.Format("{0}old\\", RaportPath);
                                System.IO.Directory.CreateDirectory(OldRaportPath);

                                String czas = DateTime.Now.ToString("yyyyddMM");
                                string destFile = String.Format(ListyMaterialow.Count>1 ? "{0}{1}_{2}_{3}_{4}.txt" : "{0}{1}_{3}_{4}.txt", OldRaportPath, LiniaTrasowania.Name, ListaMaterialow.Name, GrupaLiniiSamplowania.Name, czas);
                                System.IO.File.Copy(OutputFileName, destFile, true);
                            }
                        }
                    }
                }
                catch (System.Exception ex) {
                    ed.WriteMessage("Błąd transakcji: {0}", ex);
                 }
                acTrans.Commit();
            }

        }

        //ListaMaterialow - Lista materiałów scalone wszystkie grupy materiałowe
        [Autodesk.AutoCAD.Runtime.CommandMethod("ListaMaterialow")] 
        public void ListaMaterialow()
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

                            // Wiersz nagłówków
                            string naglowekTxt = "Pikieta";
                            foreach (QTOMaterialList ListaMaterialow in ListyMaterialow) 
                                foreach (QTOMaterial Material in ListaMaterialow)
                                    naglowekTxt += String.Format("\t{0}", Material.Name);

                            OutputFile.WriteLine(naglowekTxt);

                            // // Wiersz jednostek
                            // string jednostkiTxt = "Jednostka";
                            // jednostkiTxt += String.Concat(Enumerable.Repeat("\t[ m² ]", ListaMaterialow.Count));
                            // OutputFile.WriteLine(jednostkiTxt);

                            
                            foreach (ObjectId SampleLineID in GrupaLiniiSamplowania.GetSampleLineIds()) 
                            {
                                SampleLine LiniaSamplowania = (SampleLine)acTrans.GetObject(SampleLineID, OpenMode.ForRead);
                                string wierszDanych = LiniaSamplowania.Station.ToString();
                                                            
                                foreach (QTOMaterialList ListaMaterialow in ListyMaterialow) 
                                {
                                    foreach (QTOMaterial Material in ListaMaterialow)
                                    {
                                        ObjectId MaterialSectionID = LiniaSamplowania.GetMaterialSectionId(ListaMaterialow.Guid, Material.Guid);
                                        MaterialSection MaterialSectionObject = (MaterialSection)acTrans.GetObject(MaterialSectionID, OpenMode.ForRead);

                                        double area = MaterialSectionObject.Area2();
                                        wierszDanych += String.Format("\t{0}", area);
                                    }
                                }
                                OutputFile.WriteLine(wierszDanych);
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
                catch (System.Exception ex) { 
                    ed.WriteMessage("Błąd transakcji: {0}", ex);
                }
                acTrans.Commit();
            }
        }

        #region IExtensionApplication Members

        public void Initialize()
        {
            Document acDoc = Application.DocumentManager.MdiActiveDocument;
            acDoc.Editor.WriteMessage("\nWczytano dodatek generujący raport materiałów - Użyj polecenia: ListaMaterialow lub ListaMaterialowSEP");
        }

        public void Terminate()
        {
        }
        
        #endregion
    }
}
