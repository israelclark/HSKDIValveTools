// (C) Copyright 2013 by Microsoft 
//
using System;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.EditorInput;
using acadApp = Autodesk.AutoCAD.ApplicationServices.Application;
using Autodesk.AutoCAD;
using Autodesk.AutoCAD.Internal;
using Autodesk.AutoCAD.Windows;
using System.Text;
using System.Collections.Generic;
using HSKDICommon;

// This line is not mandatory, but improves loading performances
[assembly: CommandClass(typeof(HSKDIValveTools.MyCommands))]

namespace HSKDIValveTools
{

    // This class is instantiated by AutoCAD for each document when
    // a command is called by the user the first time in the context
    // of a given document. In other words, non static data in this class
    // is implicitly per-document!
    public class MyCommands
    {
        [CommandMethod("NumberValves")]
        static public void NumberValves()
        {
            Document doc = acadApp.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;
            DBObject obj;            
            ObjectIdCollection objIds = new ObjectIdCollection();
            string controllerLetter = "";
            int prevControllerNumber = 0;

            // Ask the user to select Controller
            PromptEntityOptions peo = new PromptEntityOptions("\nSelect Controller Block or to Resume Previous, Select Designator to Resume From: ");
            peo.SetRejectMessage("\nObject must be a Block.");
            peo.AddAllowedClass(typeof(BlockReference), false);
            peo.AddAllowedClass(typeof(MLeader), false);
            peo.AllowObjectOnLockedLayer = true;            

            PromptEntityResult res = ed.GetEntity(peo);
            if (res.Status == PromptStatus.OK)
            {
                Transaction tr = doc.TransactionManager.StartTransaction();
                using (tr)
                {
                    //Get the Controller letter &| previous controller number
                    obj = tr.GetObject(res.ObjectId, OpenMode.ForRead);
                    if (obj.GetType().ToString() == "Autodesk.AutoCAD.DatabaseServices.MLeader")
                    {
                        MLeader ml = (MLeader)obj;
                        if (ml != null)
                        {
                            ObjectId designatorId = ml.BlockContentId;
                            BlockTableRecord btr = (BlockTableRecord)tr.GetObject(designatorId, OpenMode.ForRead);
                            foreach (ObjectId atid in btr)
                            {
                                obj = tr.GetObject(atid, OpenMode.ForRead, false) as DBObject;

                                if (atid.ObjectClass.IsDerivedFrom(RXClass.GetClass(typeof(AttributeDefinition))))
                                {
                                    AttributeDefinition attdef = (AttributeDefinition)obj as AttributeDefinition;

                                    if (attdef.Tag.ToUpper() == "ZONE")
                                    {
                                        string tempstring = "";
                                        string tempdouble = "";
                                        AttributeReference ar = ml.GetBlockAttribute(attdef.ObjectId);

                                        if (ar.Tag != null)
                                        {
                                            foreach (char c in ar.TextString)
                                            {
                                                tempdouble += HSKDICommon.Commands.SetNumber(c.ToString()) != -1 ? c.ToString() : "";
                                                if (tempdouble == "") tempstring += c.ToString();
                                            }
                                            controllerLetter = tempstring;
                                            prevControllerNumber = (int)HSKDICommon.Commands.SetNumber(tempdouble);
                                        }
                                    }
                                }
                            }
                            if ((prevControllerNumber == 0) || (controllerLetter == ""))
                            {
                                //was not a previously numbered designator
                            }
                        }
                    }

                    if (obj.GetType().ToString() == "Autodesk.AutoCAD.DatabaseServices.BlockReference")
                    {
                        BlockReference br = obj as BlockReference;
                        string blockName = br.Name;
                        if (br.IsDynamicBlock)
                        {
                            BlockTableRecord btr = (BlockTableRecord)tr.GetObject(br.DynamicBlockTableRecord, OpenMode.ForRead);
                            blockName = btr.Name;
                        }
                                               
                        if (br != null)
                        {
                            if (blockName.ToUpper().Contains("CONTROLLER"))
                            {
                                foreach (ObjectId atId in br.AttributeCollection)
                                {
                                    obj = tr.GetObject(atId, OpenMode.ForRead);
                                    AttributeReference ar = obj as AttributeReference;

                                    if (ar.Tag != null)
                                    {
                                        if (ar.Tag.ToUpper() == "CONTROLLER LETTER") controllerLetter = ar.TextString;
                                    }

                                }
                            }
                            else if (blockName.ToUpper().Contains("ZONE"))
                            {                               
                                foreach (ObjectId atId in br.AttributeCollection)
                                {
                                    obj = tr.GetObject(atId, OpenMode.ForRead);
                                    AttributeReference ar = obj as AttributeReference;
                                    
                                            string tempstring = "";
                                            string tempdouble = "";
                                           

                                            if (ar.Tag.ToUpper() == "ZONE")
                                            {
                                                foreach (char c in ar.TextString)
                                                {
                                                    tempdouble += HSKDICommon.Commands.SetNumber(c.ToString()) != -1 ? c.ToString() : "";
                                                    if (tempdouble == "") tempstring += c.ToString();
                                                }
                                                controllerLetter = tempstring;
                                                prevControllerNumber = (int)HSKDICommon.Commands.SetNumber(tempdouble);
                                            }
                                        
                                }
                                
                            }
                        }
                    }
                    else
                    {
                        //Not a controller block                        
                    }
                    tr.Commit();
                }

                ed.WriteMessage("\n");
                while (res.Status == PromptStatus.OK)
                {
                    //Number the zones
                    tr = doc.TransactionManager.StartTransaction();
                    using (tr)
                    {
                        // Ask the user to select valve designator &| valve            
                        peo.Message = "\nSelect Designators in order: ";
                        peo.SetRejectMessage("\nObject must be a Block or Multileader.");
                        peo.AddAllowedClass(typeof(BlockReference), false);
                        peo.AddAllowedClass(typeof(MLeader), false);
                        peo.AllowObjectOnLockedLayer = true;

                        res = ed.GetEntity(peo);
                        if (res.Status == PromptStatus.OK)
                        {
                            if (res.ObjectId != null)
                            {
                                obj = tr.GetObject(res.ObjectId, OpenMode.ForRead);
                                if (obj.GetType().ToString() == "Autodesk.AutoCAD.DatabaseServices.MLeader")
                                {
                                    MLeader ml = obj as MLeader;
                                    if (ml != null)
                                    {
                                        ObjectId designatorBlockId = ml.BlockContentId;
                                        ObjectId designatorMultileaderId = ml.Id;
                                        HSKDICommon.ListPair targetAttrib = new HSKDICommon.ListPair("ZONE", "");
                                        BlockTableRecord mbtr = (BlockTableRecord)tr.GetObject(designatorBlockId, OpenMode.ForRead);
                                        foreach (ObjectId atid in mbtr)
                                        {
                                            obj = tr.GetObject(atid, OpenMode.ForRead, false) as DBObject;

                                            if (atid.ObjectClass.IsDerivedFrom(RXClass.GetClass(typeof(AttributeDefinition))))
                                            {
                                                AttributeDefinition attdef = (AttributeDefinition)obj as AttributeDefinition;

                                                if (attdef.Tag.ToUpper() == targetAttrib.tag && !objIds.Contains(designatorMultileaderId)) // check your existing tag
                                                {
                                                    objIds.Add(designatorMultileaderId);
                                                    AttributeReference attRef = ml.GetBlockAttribute(attdef.ObjectId);
                                                    attRef = ml.GetBlockAttribute(attdef.ObjectId);
                                                    if (!attRef.IsWriteEnabled) attRef.UpgradeOpen();
                                                    attRef.TextString = controllerLetter + (prevControllerNumber + objIds.Count).ToString();
                                                    if (!ml.IsWriteEnabled) ml.UpgradeOpen();
                                                    ml.SetBlockAttribute(attdef.ObjectId, attRef);
                                                    ed.WriteMessage("\nTotal Designators selected: {0}.", objIds.Count);
                                                }
                                            }
                                        }
                                    }
                                }
                                if (obj.GetType().ToString() == "Autodesk.AutoCAD.DatabaseServices.BlockReference")
                                {
                                    BlockReference br = obj as BlockReference;
                                    if (br != null)
                                    {
                                        foreach (ObjectId arId in br.AttributeCollection)
                                        {
                                            obj = tr.GetObject(arId, OpenMode.ForRead);
                                            AttributeReference ar = obj as AttributeReference;
                                            BlockTableRecord btr = (BlockTableRecord)tr.GetObject(br.IsDynamicBlock ? br.DynamicBlockTableRecord : br.BlockTableRecord, OpenMode.ForRead);

                                            if (ar.Tag != null)
                                            {
                                                if (ar.Tag.ToUpper() == "ZONE" && !objIds.Contains(br.ObjectId))
                                                {
                                                    objIds.Add(br.ObjectId);
                                                    ar.UpgradeOpen();
                                                    ar.TextString = controllerLetter + (prevControllerNumber + objIds.Count).ToString();
                                                    ar.DowngradeOpen();
                                                    ed.WriteMessage("\nTotal Designators selected: {0}.", objIds.Count);
                                                }
                                            }
                                        }
                                    }
                                }
                            }                            
                        }
                        tr.Commit();
                    }
                }
            }
        }


        [CommandMethod("SizeValves", CommandFlags.UsePickSet | CommandFlags.Redraw | CommandFlags.Modal)]
        public void SizeValves()
        {
            Editor ed = acadApp.DocumentManager.MdiActiveDocument.Editor;
            Database db = HostApplicationServices.WorkingDatabase;
            Transaction tr = db.TransactionManager.StartTransaction();

            // Start the transaction
            try
            {
                PromptSelectionOptions opts = new PromptSelectionOptions();
                opts.MessageForAdding = "Select blocks or multileaders: ";
                PromptSelectionResult res = ed.GetSelection(opts);

                // Do nothing if selection is unsuccessful
                if (res.Status != PromptStatus.OK) return;

                SelectionSet selSet = res.Value;
                ObjectId[] idArray = selSet.GetObjectIds();
                foreach (ObjectId Id in idArray)
                {
                    ObjectId designatorId = Id;
                    DBObject selectedObj = tr.GetObject(designatorId, OpenMode.ForRead);
                    //ed.WriteMessage("\nSelected: {0}", selectedObj.GetType());

                    if (selectedObj.GetType().ToString() == "Autodesk.AutoCAD.DatabaseServices.MLeader")
                    {
                        MLeader mlRef = (MLeader)tr.GetObject(designatorId, OpenMode.ForRead);

                        if (mlRef != null)
                        {
                            designatorId = mlRef.BlockContentId;
                            HSKDICommon.ListPair targetAttrib = new HSKDICommon.ListPair("DIAMETER", "");
                            BlockTableRecord mbtr = (BlockTableRecord)tr.GetObject(designatorId, OpenMode.ForRead);

                            foreach (ObjectId atid in mbtr)
                            {
                                DBObject obj = tr.GetObject(atid, OpenMode.ForRead, false) as DBObject;

                                if (atid.ObjectClass.IsDerivedFrom(RXClass.GetClass(typeof(AttributeDefinition))))
                                {
                                    AttributeDefinition attdef = (AttributeDefinition)obj as AttributeDefinition;

                                    if (attdef.Tag.ToUpper() == "FLOW") // check your existing tag
                                    {
                                        AttributeReference attRef = mlRef.GetBlockAttribute(attdef.ObjectId);            
                                        if (!attRef.IsWriteEnabled) attRef.UpgradeOpen();
                                        List<HSKDICommon.ListPair> myAtts = new List<HSKDICommon.ListPair>();
                                        myAtts.Add(new HSKDICommon.ListPair(attRef.Tag, attRef.TextString));
                                        int flow = 0;
                                        if (myAtts.Count > 0 && attRef.Tag.ToUpper() == "FLOW")
                                        {
                                            if (!attRef.IsWriteEnabled) attRef.UpgradeOpen();
                                            flow = (int) ReadAttribute(tr, myAtts, "FLOW");
                                            targetAttrib = new HSKDICommon.ListPair("DIAMETER", "");
                                            ValveSize(targetAttrib, flow);                                         
                                            //UpdateAttribute(tr, mlRef.BlockContentId, targetAttrib);
                                            //ed.WriteMessage("\nValve Flow = {0}, sizing valve to {1}.", flow, targetAttrib.textString);
                                        }
                                        foreach (ObjectId targetAtid in mbtr)
                                        {
                                            obj = tr.GetObject(targetAtid, OpenMode.ForRead, false);

                                            if (atid.ObjectClass.IsDerivedFrom(RXClass.GetClass(typeof(AttributeDefinition))))
                                            {
                                                try
                                                {
                                                    attdef = (AttributeDefinition)obj as AttributeDefinition;
                                                    if (attdef.Tag.ToUpper() == "DIAMETER") // check your existing tag
                                                    {
                                                        
                                                        
                                                        attRef = mlRef.GetBlockAttribute(attdef.ObjectId);
                                                        if (!attRef.IsWriteEnabled) attRef.UpgradeOpen();
                                                        attRef.TextString = targetAttrib.textString;
                                                        if (!mlRef.IsWriteEnabled) mlRef.UpgradeOpen();
                                                        mlRef.SetBlockAttribute(attdef.ObjectId, attRef);
                                                        ed.WriteMessage("\nValve Flow = {0}, sizing valve to {1}.", flow, targetAttrib.textString);
                                                    }
                                                }
                                                catch
                                                {
                                                    
                                                }
                                            }
                                        }
                                    }
                                }
                            }                            
                            mlRef.RecordGraphicsModified(true);
                        }
                    }

                    if (selectedObj.GetType().ToString() == "Autodesk.AutoCAD.DatabaseServices.BlockReference")
                    {

                        BlockReference blkRef = (BlockReference)tr.GetObject(designatorId, OpenMode.ForRead);

                        if (blkRef != null)
                        {
                            AttributeCollection attCol = blkRef.AttributeCollection;

                            foreach (ObjectId attId in attCol)
                            {
                                AttributeReference attRef = (AttributeReference)tr.GetObject(attId, OpenMode.ForRead);
                                List<HSKDICommon.ListPair> myAtts = new List<HSKDICommon.ListPair>();
                                myAtts.Add(new HSKDICommon.ListPair(attRef.Tag, attRef.TextString));

                                if (myAtts.Count > 0 && attRef.Tag.ToUpper() == "FLOW")
                                {
                                    double flow = ReadAttribute(tr, myAtts, "FLOW");
                                    HSKDICommon.ListPair targetAttrib = new HSKDICommon.ListPair("DIAMETER", "");
                                    ValveSize(targetAttrib, flow);  
                                    UpdateAttribute(tr, blkRef, targetAttrib);
                                    ed.WriteMessage("\nValve Flow = {0}, sizing valve to {1}.", flow, targetAttrib.textString);
                                }
                            }
                        }
                    }
                }
                ed.WriteMessage("\nDone Sizing.");
                tr.Commit();
            }
            catch (Autodesk.AutoCAD.Runtime.Exception ex)
            {
                ed.WriteMessage(("Exception: " + ex.Message));
            }
            finally
            {
                tr.Dispose();
            }
        }

        private static void ValveSize(HSKDICommon.ListPair targetAttrib, double flow)
        {
            if (flow <= 25) targetAttrib.textString = "1\'\'";
            else if (flow <= 50) targetAttrib.textString = "1.5\'\'";
            else if (flow <= 100) targetAttrib.textString = "2\'\'";            
        }

        public static double ReadAttribute(Transaction tr, List<HSKDICommon.ListPair> myAtts, string tag)
        {
            double value = 0;
            myAtts.ForEach(delegate(HSKDICommon.ListPair a)
            {
                if (a.tag.ToUpper() == tag.ToUpper())
                {
                    try
                    {
                        //numerical values
                        double myDouble = HSKDICommon.Commands.SetNumber(a.textString);


                        if (myDouble.ToString() == a.textString)
                        {
                            value = (Math.Round(myDouble, 0));
                        }
                    }
                    catch
                    {
                        value = HSKDICommon.Commands.SetNumber(a.textString);
                    }
                }
            });

            return value;
        }

        public static void UpdateAttribute(Transaction tr, BlockReference blkRef, HSKDICommon.ListPair attribs)
        {
            AttributeCollection attCol = blkRef.AttributeCollection;

            foreach (ObjectId attId in attCol)
            {
                AttributeReference attRef = (AttributeReference)tr.GetObject(attId, OpenMode.ForRead);
                if (attRef.Tag.ToUpper() == attribs.tag)
                {
                    attRef.UpgradeOpen();
                    attRef.TextString = attribs.textString;
                    attRef.DowngradeOpen();
                }
            }
        }

    }

}
