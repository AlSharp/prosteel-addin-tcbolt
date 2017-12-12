using System;
using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.Runtime.CompilerServices;
using System.Windows.Forms;
using System.Collections;
using SRI=System.Runtime.InteropServices;

using Bentley.ProStructures;
using ProStructuresOE;
using Bentley.ProStructures.CadSystem;
using Bentley.ProStructures.Configuration;
using Bentley.ProStructures.Drawing;
using Bentley.ProStructures.Geometry.Data;
using Bentley.ProStructures.Geometry.Utilities;
using Bentley.ProStructures.Miscellaneous;
using Bentley.ProStructures.Property;
using Bentley.ProStructures.Entity;
using Bentley.ProStructures.Steel.Plate;
using Bentley.ProStructures.Steel.Primitive;
using Bentley.ProStructures.Steel.Bolt;
using Bentley.ProStructures.Modification.Edit;
using Bentley.ProStructures.Modification.ObjectData;
using Microsoft.VisualBasic;
using PlugInBase;

using BM=Bentley.MicroStation;
using BMW=Bentley.MicroStation.WinForms;
using BMI=Bentley.MicroStation.InteropServices;
using BCOM=Bentley.Interop.MicroStationDGN;

namespace GSFBolt
{

public class TcBolt
{
	public double Diameter; 
	public double Length; 
	public double KlemmLength;
	public double LengthAdd; 
	public double HeadHeight; 
	public double HeadDiameter; 
	public double WasherOuterDiameter; 
	public double WasherThickness; 
	public double NutKeySize; 
	public double NutHeight; 
	public int ObjectId;
	public string Name;
	public string Type;
	public string MaterialName;
	public int MaterialIndex;
	public int Color;
	public bool PartListFlag;
	public bool ModifyFlag;
	public int PartFamilyClassIndex;
	public PsPoint InsertPoint;
	public PsVector XAxis;
	public PsVector YAxis;
	public PsVector ZAxis;

	public TcBolt()
	{

	}
	public void SetParams(double TcBoltDiameter, string TcBoltType, double BoltAddLength)
	{
		PsUnits psUnits = new PsUnits();
		Diameter = TcBoltDiameter;
		LengthAdd = BoltAddLength;
		if (TcBoltDiameter == 0.5D)
        {
        	HeadHeight = 0.323D;
    		HeadDiameter = 1.126D;
    		WasherOuterDiameter = 1.063D;
    		WasherThickness = 0.097D;
    		NutKeySize = 0.875D;
    		NutHeight = 0.504D;
        }
        else if (TcBoltDiameter == 0.625D)
        {
        	HeadHeight = 0.403D;
    		HeadDiameter = 1.313D;
    		WasherOuterDiameter = 1.313D;
    		WasherThickness = 0.122D;
    		NutKeySize = 1.062D;
    		NutHeight = 0.631D;
        }
        else if (TcBoltDiameter == 0.75D)
	    {
			HeadHeight = 0.483D;
			HeadDiameter = 1.580D;
	   		WasherOuterDiameter = 1.468D;
	   		WasherThickness = 0.122D;
	   		NutKeySize = 1.250D;
	   		NutHeight = 0.758D;
	    }
		else if (TcBoltDiameter == 0.875D)
		{
			HeadHeight = 0.563D;
			HeadDiameter = 1.880D;
 			WasherOuterDiameter = 1.750D;
			WasherThickness = 0.136D;
			NutKeySize = 1.438D;
			NutHeight = 0.885D;
		}
		else if (TcBoltDiameter == 1D)
		{
			HeadHeight = 0.627D;
			HeadDiameter = 2.158D;
			WasherOuterDiameter = 2.000D;
			WasherThickness = 0.136D;
			NutKeySize = 1.625D;
			NutHeight = 1.012D;
		}
		else
		{
			HeadHeight = 0.718D;
			HeadDiameter = 2.375D;
			WasherOuterDiameter = 2.250D;
			WasherThickness = 0.177D;
			NutKeySize = 1.812D;
			NutHeight = 1.139D;
		}

		if (TcBoltType == "SHOP TC A325")
		{
			Type = "SHOP TC Bolt";
			MaterialName = "ASTM A325 Gr.81";
			Color = 20; 
		}
		else if (TcBoltType == "FIELD TC A325")
		{
			Type = "FIELD TC Bolt";
			MaterialName = "ASTM A325 Gr.81";
			Color = 7; 
		}
		else if (TcBoltType == "SHOP TC A490")
		{
			Type = "SHOP TC Bolt";
			MaterialName = "ASTM A490";
			Color = 202; 
		}
		else
		{
			Type = "FIELD TC Bolt";
			MaterialName = "ASTM A490";
			Color = 94; 
		}
	}

	public void SetStandardBoltParams(double StandardBoltDiameter, string StandardBoltStyle, double BoltAddLength)
	{
		PsUnits psUnits = new PsUnits();
		Diameter = StandardBoltDiameter;
		LengthAdd = BoltAddLength;
		if (StandardBoltDiameter == 0.5D)
        {
        	HeadHeight = 0.323D;
    		HeadDiameter = 1.126D;
    		WasherOuterDiameter = 1.063D;
    		WasherThickness = 0.097D;
    		NutKeySize = 0.875D;
    		NutHeight = 0.504D;
        }
        else if (StandardBoltDiameter == 0.625D)
        {
        	HeadHeight = 0.403D;
    		HeadDiameter = 1.313D;
    		WasherOuterDiameter = 1.313D;
    		WasherThickness = 0.122D;
    		NutKeySize = 1.062D;
    		NutHeight = 0.631D;
        }
        else if (StandardBoltDiameter == 0.75D)
	    {
			HeadHeight = 0.483D;
			HeadDiameter = 1.580D;
	   		WasherOuterDiameter = 1.468D;
	   		WasherThickness = 0.122D;
	   		NutKeySize = 1.250D;
	   		NutHeight = 0.758D;
	    }
		else if (StandardBoltDiameter == 0.875D)
		{
			HeadHeight = 0.563D;
			HeadDiameter = 1.880D;
 			WasherOuterDiameter = 1.750D;
			WasherThickness = 0.136D;
			NutKeySize = 1.438D;
			NutHeight = 0.885D;
		}
		else if (StandardBoltDiameter == 1D)
		{
			HeadHeight = 0.627D;
			HeadDiameter = 2.158D;
			WasherOuterDiameter = 2.000D;
			WasherThickness = 0.136D;
			NutKeySize = 1.625D;
			NutHeight = 1.012D;
		}
		else
		{
			HeadHeight = 0.718D;
			HeadDiameter = 2.375D;
			WasherOuterDiameter = 2.250D;
			WasherThickness = 0.177D;
			NutKeySize = 1.812D;
			NutHeight = 1.139D;
		}

		if (StandardBoltStyle == "A325 SHOP")
		{
			Type = "SHOP TC Bolt";
			MaterialName = "ASTM A325 Gr.81";
			Color = 20; 
		}
		else if (StandardBoltStyle == "A325 FIELD")
		{
			Type = "FIELD TC Bolt";
			MaterialName = "ASTM A325 Gr.81";
			Color = 7; 
		}
		else if (StandardBoltStyle == "A490 SHOP")
		{
			Type = "SHOP TC Bolt";
			MaterialName = "ASTM A490";
			Color = 202; 
		}
		else
		{
			Type = "FIELD TC Bolt";
			MaterialName = "ASTM A490";
			Color = 94; 
		}
	}

	public void CreateByTwoPoints(PsPoint StartPoint, PsPoint EndPoint)
	{
		PsCreatePrimitive psCylinder = new PsCreatePrimitive();
		PsCreatePrimitive psHead1 = new PsCreatePrimitive();
		PsCreatePrimitive psHead2 = new PsCreatePrimitive();
		PsCreatePrimitive psWasher = new PsCreatePrimitive();
		PsCreatePrimitive psNut = new PsCreatePrimitive();
		PsPolygon psNutPolygon = new PsPolygon();

		PsVector psVector1 = new PsVector();
		PsVector psVector2 = new PsVector();
		PsVector psVector3 = new PsVector();
		PsVector psVector = new PsVector();

		psVector1.SetFromPoints(StartPoint, EndPoint);
		psVector2 = psVector.GetPerpendicularVector(psVector1);
		psVector3.SetFromCrossProduct(psVector1, psVector2);

		InsertPoint = StartPoint;
		KlemmLength = psVector1.Length;

		psVector1.Normalize();
		psVector2.Normalize();
		psVector3.Normalize();

		ZAxis = psVector1;
		XAxis = psVector2;
		YAxis = psVector3;
      
		psCylinder.SetXYPlane(XAxis, YAxis);
		psCylinder.SetInsertPoint(InsertPoint);
		Length = Math.Round((KlemmLength + WasherThickness + NutHeight + LengthAdd + 0.375D)*4, MidpointRounding.ToEven)/4;
		psCylinder.CreateCylinder(Diameter/2, Length);

                // Head
		PsVector XAxisNegative = new PsVector();
		XAxisNegative = XAxis.Clone();
		XAxisNegative.Invert();

		psHead1.SetXYPlane(XAxisNegative,YAxis);
		psHead1.SetInsertPoint(InsertPoint);
		psHead1.CreateCylinder(HeadDiameter/2, HeadHeight/4);

		PsPoint head2InsertPoint = new PsPoint();
		head2InsertPoint = InsertPoint.Clone();
		head2InsertPoint.AddScaled(ZAxis, -1*HeadHeight/4);

		psHead2.SetXYPlane(XAxisNegative, YAxis);
		psHead2.SetInsertPoint(head2InsertPoint);
		psHead2.CreateCone(HeadDiameter/2, HeadDiameter/6, 0.75*HeadHeight);

                // Washer
		psWasher.SetXYPlane(XAxis, YAxis);
		psWasher.SetInsertPoint(EndPoint);
		psWasher.CreateCylinder(WasherOuterDiameter/2, WasherThickness);

                // Nut
		PsPoint nutInsertPoint = new PsPoint();
		nutInsertPoint = EndPoint.Clone();
		nutInsertPoint.AddScaled(ZAxis, WasherThickness);
		psNutPolygon.createPolygon(6, NutKeySize/2, false);
		psNut.SetXYPlane(XAxis, YAxis);
		psNut.SetPolygon(psNutPolygon);
		psNut.SetInsertPoint(nutInsertPoint);
		psNut.CreateExtrusion(NutHeight, 0D, 0D);

                // Unite all
		int cylinderId = psCylinder.ObjectId;
		int head1Id = psHead1.ObjectId;
		int head2Id = psHead2.ObjectId;
		int washerId = psWasher.ObjectId;
		int nutId = psNut.ObjectId;

		int[] ids = {head1Id, head2Id, washerId, nutId};

		PsCutObjects psUniteObjects = new PsCutObjects();
		psUniteObjects.SetObjectId(cylinderId);
		PsTransaction psTransaction = new PsTransaction();
		foreach (int i in ids)
		{
			psUniteObjects.SetAsBooleanCut(i);
			psUniteObjects.SetSubBodyType(SubBodyType.kAddBody);
			psUniteObjects.CreateLogicalLink = false;
			psUniteObjects.Apply();
			psTransaction.EraseLongId((long)i);
			psTransaction.Close();
		}

		PsUnits psUnits = new PsUnits();
		PsMaterialTable psMaterialTable = new PsMaterialTable();
		PsObjectProperties psObjectProperties = new PsObjectProperties();
		psObjectProperties.Name = Type + " " + psUnits.ConvertToText(Diameter) + "x" + psUnits.ConvertToText(Length) + 
								" " + MaterialName;
		psObjectProperties.Material = psMaterialTable.get_MaterialIndexFromName(MaterialName);
		psObjectProperties.ColorIndex = Color;
		PartListFlag = true;
		psObjectProperties.PartListFlag = PartListFlag;
		PartFamilyClassIndex = -1;
		psObjectProperties.FamilyClass = PartFamilyClassIndex;
		psObjectProperties.Length = Length;
		psObjectProperties.writeTo(cylinderId);

		PsPrimitive psPrimitive = new PsPrimitive();
		psTransaction.GetObject((long)cylinderId, PsOpenMode.kForWrite, ref psPrimitive);
		psPrimitive.writeProps(psObjectProperties);
		psTransaction.Close();
	}
	public void ConvertToTCBolt(int ObjectId)
	{
		PsObjectProperties psObjectProperties = new PsObjectProperties();
		PsTransaction psTransaction = new PsTransaction();
		psObjectProperties.readFrom(ObjectId);
		if (psObjectProperties.ObjectType == ObjectType.kBolt)
		{
			PsBolt standardBolt = new PsBolt();
			psTransaction.GetObject((long)ObjectId, PsOpenMode.kForWrite, ref standardBolt);
			if ((new ArrayList(new string[] {"A325 SHOP", "A325 FIELD", "A490 SHOP", "A490 FIELD"})).Contains(standardBolt.BoltStyleName))
			{
				SetStandardBoltParams(standardBolt.Diameter, standardBolt.BoltStyleName, standardBolt.LengthAddition);
				Length = standardBolt.Length;
				Diameter = standardBolt.Diameter;
				KlemmLength = standardBolt.KlemmLength;
				InsertPoint = standardBolt.InsertPoint; 
				XAxis = standardBolt.XAxis;
				YAxis = standardBolt.YAxis;
				PsPoint psEndPoint = new PsPoint();
				PsVector tempVector = new PsVector();
				tempVector.SetFromCrossProduct(XAxis, YAxis);
				tempVector.Normalize();
				ZAxis = tempVector.Clone();
				psEndPoint = InsertPoint.Clone();
				psEndPoint.AddScaled(ZAxis, KlemmLength);
				CreateByTwoPoints(InsertPoint, psEndPoint);
				psTransaction.Close();
				psTransaction.EraseLongId((long)ObjectId);
			}
			psTransaction.Close();
		}
	}

}
public class AnchorBolt
{
	public double Diameter; 
	public double Length;
	public double Embedment;
	public double PartThickness;
	public double WasherOuterDiameter; 
	public double WasherThickness; 
	public double NutKeySize; 
	public double NutHeight; 
	public int ObjectId;
	public string Type;
	public string Name;
	public string MaterialName;
	public int MaterialIndex;
	public int Color;
	public bool PartListFlag;
	public bool ModifyFlag;
	public int PartFamilyClassIndex;
	public PsPoint InsertPoint;
	public PsVector XAxis;
	public PsVector YAxis;
	public PsVector ZAxis;

	public AnchorBolt()
	{

	}
	public void SetParams(string AnchorName, double AnchorDiameter, double AnchorEmbedment, double AnchorPartThickness, double AnchorLength)
	{
		Type = AnchorName;
		Diameter = AnchorDiameter;
		Embedment = AnchorEmbedment;
		PartThickness = AnchorPartThickness;
		Length = AnchorLength;
		Color = 7;
		PartListFlag = true;
		PartFamilyClassIndex = -1;
		if (AnchorDiameter == 0.25D)
        {
    		WasherOuterDiameter = 0.734D;
    		WasherThickness = 0.071D;
    		NutKeySize = 0.500D;
    		NutHeight = 0.220D;
    	}
		else if (AnchorDiameter == 0.375D)
        {
    		WasherOuterDiameter = 1.000D;
    		WasherThickness = 0.071D;
    		NutKeySize = 0.688D;
    		NutHeight = 0.377D;
    	}
		else if (AnchorDiameter == 0.5D)
        {
    		WasherOuterDiameter = 1.250D;
    		WasherThickness = 0.112D;
    		NutKeySize = 0.875D;
    		NutHeight = 0.504D;
        }
        else if (AnchorDiameter == 0.625D)
        {
    		WasherOuterDiameter = 1.75D;
    		WasherThickness = 0.112D;
    		NutKeySize = 1.062D;
    		NutHeight = 0.631D;
        }
        else if (AnchorDiameter == 0.75D)
	    {
	   		WasherOuterDiameter = 2.000D;
	   		WasherThickness = 0.112D;
	   		NutKeySize = 1.250D;
	   		NutHeight = 0.758D;
	    }
		else if (AnchorDiameter == 0.875D)
		{
 			WasherOuterDiameter = 2.250D;
			WasherThickness = 0.174D;
			NutKeySize = 1.438D;
			NutHeight = 0.885D;
		}
		else if (AnchorDiameter == 1D)
		{
			WasherOuterDiameter = 2.500D;
			WasherThickness = 0.174D;
			NutKeySize = 1.625D;
			NutHeight = 1.012D;
		}
		else if (AnchorDiameter == 1.125D)
		{
			WasherOuterDiameter = 2.750D;
			WasherThickness = 0.174D;
			NutKeySize = 1.812D;
			NutHeight = 1.139D;
		}
		else
		{
			WasherOuterDiameter = 3.000D;
			WasherThickness = 0.174D;
			NutKeySize = 2.000D;
			NutHeight = 1.251D;
		}
	}
	public void GetNutAndWasherSizesByDiameter(double AnchorDiameter)
	{
		if (AnchorDiameter == 0.25D)
        {
    		WasherThickness = 0.071D;
    		NutHeight = 0.220D;
    	}
		else if (AnchorDiameter == 0.375D)
        {
    		WasherThickness = 0.071D;
    		NutHeight = 0.377D;
    	}
		else if (AnchorDiameter == 0.5D)
        {
    		WasherThickness = 0.112D;
    		NutHeight = 0.504D;
        }
        else if (AnchorDiameter == 0.625D)
        {
    		WasherThickness = 0.112D;
    		NutHeight = 0.631D;
        }
        else if (AnchorDiameter == 0.75D)
	    {
	   		WasherThickness = 0.112D;
	   		NutHeight = 0.758D;
	    }
		else if (AnchorDiameter == 0.875D)
		{
			WasherThickness = 0.174D;
			NutHeight = 0.885D;
		}
		else if (AnchorDiameter == 1D)
		{
			WasherThickness = 0.174D;
			NutHeight = 1.012D;
		}
		else if (AnchorDiameter == 1.125D)
		{
			WasherThickness = 0.174D;
			NutHeight = 1.139D;
		}
		else
		{
			WasherThickness = 0.174D;
			NutHeight = 1.251D;
		}
	}
	public int CreateAnchorCode111(PsPoint Point1, PsPoint Point2)
	{
		PsCreatePrimitive psCylinder = new PsCreatePrimitive();
		PsCreatePrimitive psWasher = new PsCreatePrimitive();
		PsCreatePrimitive psNut = new PsCreatePrimitive();
		PsPolygon psNutPolygon = new PsPolygon();

		PsVector psVector1 = new PsVector();
		PsVector psVector2 = new PsVector();
		PsVector psVector3 = new PsVector();
		PsVector psVector = new PsVector();

		psVector1.SetFromPoints(Point1, Point2);
		psVector2 = psVector.GetPerpendicularVector(psVector1);
		psVector3.SetFromCrossProduct(psVector1, psVector2);

		InsertPoint = Point1;

		psVector1.Normalize();
		psVector2.Normalize();
		psVector3.Normalize();

		ZAxis = psVector1;
		XAxis = psVector2;
		YAxis = psVector3;

		PsVector XAxisNegative = new PsVector();
		XAxisNegative = XAxis.Clone();
		XAxisNegative.Invert();

				//Cylinder
		PsPoint CylinderInsertPoint = new PsPoint();
		CylinderInsertPoint = InsertPoint.Clone();
		CylinderInsertPoint.AddScaled(ZAxis, -1*(Length - Embedment - PartThickness));
      
		psCylinder.SetXYPlane(XAxis, YAxis);
		psCylinder.SetInsertPoint(CylinderInsertPoint);
		psCylinder.CreateCylinder(Diameter/2, Length);
                // Washer
		psWasher.SetXYPlane(XAxisNegative, YAxis);
		psWasher.SetInsertPoint(InsertPoint);
		psWasher.CreateCylinder(WasherOuterDiameter/2, WasherThickness);

                // Nut
		PsPoint nutInsertPoint = new PsPoint();
		nutInsertPoint = InsertPoint.Clone();
		nutInsertPoint.AddScaled(ZAxis, -1*WasherThickness);
		psNutPolygon.createPolygon(6, NutKeySize/2, false);
		psNut.SetXYPlane(XAxisNegative, YAxis);
		psNut.SetPolygon(psNutPolygon);
		psNut.SetInsertPoint(nutInsertPoint);
		psNut.CreateExtrusion(NutHeight, 0D, 0D);

                // Unite all
		int cylinderId = psCylinder.ObjectId;
		int washerId = psWasher.ObjectId;
		int nutId = psNut.ObjectId;

		int[] ids = {washerId, nutId};

		PsCutObjects psUniteObjects = new PsCutObjects();
		psUniteObjects.SetObjectId(cylinderId);
		PsTransaction psTransaction = new PsTransaction();
		foreach (int i in ids)
		{
			psUniteObjects.SetAsBooleanCut(i);
			psUniteObjects.SetSubBodyType(SubBodyType.kAddBody);
			psUniteObjects.CreateLogicalLink = false;
			psUniteObjects.Apply();
			psTransaction.EraseLongId((long)i);
			psTransaction.Close();
		}

		PsUnits psUnits = new PsUnits();
		PsObjectProperties psObjectProperties = new PsObjectProperties();
		psObjectProperties.Name = Type + " " + psUnits.ConvertToText(Diameter) + "x" + psUnits.ConvertToText(Length);
		psObjectProperties.ColorIndex = Color;
		psObjectProperties.PartListFlag = PartListFlag;
		psObjectProperties.FamilyClass = PartFamilyClassIndex;
		psObjectProperties.Length = Length;
		psObjectProperties.writeTo(cylinderId);
		PsPrimitive psPrimitive = new PsPrimitive();
		psTransaction.GetObject((long)cylinderId, PsOpenMode.kForWrite, ref psPrimitive);
		psPrimitive.writeProps(psObjectProperties);
		psTransaction.Close();
		ObjectId = cylinderId;
        return ObjectId;
	}
}
}