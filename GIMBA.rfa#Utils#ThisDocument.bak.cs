// REVIT DOCUMENT MACRO MODULE "Utils"
// Text dump of Revit materials properties
//
// Author    : Matthias Wolff
// Modified  : 2021-06-01
// Latest    : https://github.com/matthias-wolff/Revit-Generic-ISO-Metric-Bolt-Assembly/blob/main/GIMBA.rfa%23Utils%23ThisDocument.bak.cs
// References:
// [1] About Macro Manager and the Revit Macro IDE
//     https://help.autodesk.com/view/RVT/2022/ENU/?guid=GUID-071913D8-214A-45AB-A798-A81653E77F88
// [2] Revit API 2022
//     https://www.revitapidocs.com/2022/
// [3] C# Language Reference
//     https://docs.microsoft.com/en-us/dotnet/csharp/language-reference
// [4] .NET API Documentation, Method PropertyInfo.GetValue
//     https://docs.microsoft.com/de-de/dotnet/api/system.reflection.propertyinfo.getvalue?view=net-5.0
// [5] The Building Coder. Material, Physical and Thermal Assets. Nov. 2019. Retrieved May 2021.
//     https://thebuildingcoder.typepad.com/blog/2019/11/material-physical-and-thermal-assets.html

using System;                           // For C# EventArgs
using System.Collections.Generic;       // For C# enumerable (lists, etc.), etc.
using System.IO;                        // For file input/output
using System.Reflection;                // For accessing C# properties
using System.Diagnostics;               // For Process.Start
using System.Runtime.CompilerServices;  // For determining path of this source file
using System.Linq;                      // For .Cast, .ToList, etc
using Autodesk.Revit.UI;                // For TaskDialog, UIApplication, etc.
using Autodesk.Revit.DB;                // For acessing Revit properties
using Autodesk.Revit.DB.Visual;         // For acessing Revit materials, appearance assets, etc.

/// <summary>Provides text dumps of Revit materials and assets</summary>
public static class Dump
{
  /// <summary>Writes one line of text.</summary>
  /// <param name="line">The text line to write</param>
  private static string WriteLine(string line)
  {
    return line+"\n";
  }
  
  /// <summary>Text dump of a Revit parameter</summary>
  /// <param name="aprop">The parameter</param>
  /// <param name="prefix">Output line prefix, default is an empty string</param>
  public static string ToString(Parameter param, string prefix="")
  {
    if (param==null)
      return Dump.WriteLine("<null> (Parameter)");
    
    string dump = "";
    string n = param.Definition.Name;
    string t = param.StorageType.ToString();
    string v = param.AsValueString();
    dump += Dump.WriteLine(prefix+String.Format("{0} ({1}): {2}",n,t,v));
    
    return dump;
  }
  
  /// <summary></summary>Text dump of a C# property</summary>
  /// <param name="pi">The property</param>
  /// <param name="obj">The instance for which tp display the property</param> 
  /// <param name="prefix">Output line prefix, default is an empty string</param>
  /// <param name="pn">Alternative name of property, default is <c>pi.Name</c></param>
  /// <param name="pt">Alternative type name of property, default is <c>pi.PropertyType.Name</c></param>
  public static string ToString(PropertyInfo pi, Object obj, string prefix="", string pn=null, string pt=null)
  {
    if (pi==null)
      return Dump.WriteLine("<null> (PropertyInfo)");

    string d = "";
    int    l = pi.GetIndexParameters()!=null ? pi.GetIndexParameters().Length : 0;
    string n = String.IsNullOrEmpty(pn) ? pi.Name : pn;
    string t = String.IsNullOrEmpty(pt) ? pi.PropertyType+"" : pt;
    object v = null;
    if (l==0)
    {
      v  = pi.GetValue(obj);
      d += Dump.WriteLine(prefix+String.Format("{0} ({1}): {2}",n,t,v));
    }
    else
      foreach (ParameterInfo pari in pi.GetIndexParameters())
      {
        int i = pari.Position;
        n = pari.Name;
        t = pari.ParameterType.Name;
        v = "<value ???>"; //TODO This does not work: pi.GetValue(obj, new object[] { i });
        d += Dump.WriteLine(prefix+String.Format("{0}[{1}] ({2}): {3}",n,i,t,v)); 
      }
    
    return d;
  }

  /// <summary>Simplified text dump of a Revit asset property</summary>
  /// <param name="aprop">The asset property</param>
  /// <param name="prefix">Output line prefix, default is an empty string</param>
  public static string ToString(AssetProperty ap, string prefix="")
  {
    if (ap==null)
      return Dump.WriteLine("<null> (AssetProperty)");
    
    string d = "";
    
    // Dump (ordinary) properties if asset property object
    PropertyInfo[] pis = ap.GetType().GetProperties();      // Retrieve all properties of asset propery object
    PropertyInfo   piv = ap.GetType().GetProperty("Value"); // "Value" property of asset property object (if any)

    if (piv!=null)
      // Skip all properties except "Value"
      d += Dump.ToString(piv,ap,prefix,ap.Name+".Value",ap.Type.ToString());
    else
    {
      // Dump all properties
      string t = ap.GetType().Name;
      string n = ap.Name;

      if (ap is AssetPropertyList)
      {
        // Special case: asset property list -> dump properties
        d += Dump.WriteLine(prefix+String.Format("{0} ({1})",n,t));
        AssetPropertyList apl = ap as AssetPropertyList;
        foreach (AssetProperty ap2 in apl.GetValue())
          d += Dump.ToString(ap2,"  "+prefix);
      }
      else if (ap is AssetPropertyDoubleArray2d)
      {
        // Special case: 2x double array -> pretty-print
        IList<Double> l = (ap as AssetPropertyDoubleArray4d).GetValueAsDoubles();
        d += Dump.WriteLine(prefix+String.Format(
          "{0} ({1}): [{2}, {3}]",n,t,
          l.ElementAt(0), l.ElementAt(1)
         ));
      }
      else if (ap is AssetPropertyDoubleArray3d)
      {
        // Special case: 2x double array -> pretty-print
        IList<Double> l = (ap as AssetPropertyDoubleArray4d).GetValueAsDoubles();
        d += Dump.WriteLine(prefix+String.Format(
          "{0} ({1}): [{2}, {3}, {4}]",n,t,
          l.ElementAt(0), l.ElementAt(1), l.ElementAt(1)
         ));
      }
      else if (ap is AssetPropertyDoubleArray4d)
      {
        // Special case: 4x double array -> pretty-print
        IList<Double> l = (ap as AssetPropertyDoubleArray4d).GetValueAsDoubles();
        d += Dump.WriteLine(prefix+String.Format(
          "{0} ({1}): [{2}, {3}, {4}, {5}]",n,t,
          l.ElementAt(0), l.ElementAt(1), l.ElementAt(2), l.ElementAt(3)
         ));
      }
      else
      {
        // Gernel case: Dump C# properties
        d += Dump.WriteLine(prefix+String.Format("{0} ({1})",n,t));
        foreach (PropertyInfo pi in pis)
          d += Dump.ToString(pi,ap,"  "+prefix);
      }
    }
    
    // Rest will be indented one level more further
    prefix = "  "+prefix;

    // Dump connected properties
    if (ap.NumberOfConnectedProperties>0)
    {
      d += Dump.WriteLine(prefix+"<Connected Properties>");
      foreach (AssetProperty apc in ap.GetAllConnectedProperties())
        d += Dump.ToString(apc,"  "+prefix);
    }

    // Dump single connected asset
    Asset sca = ap.GetSingleConnectedAsset();
    if (sca!=null)
    {
      d += Dump.WriteLine(prefix+"<Single Connected Asset>");
      d += Dump.ToString(sca,"  "+prefix);
    }
    
    return d;
  }

  /// <summary>Text dump of a Revit asset</summary>
  /// <param name="asset">The asset</param>
  /// <param name="prefix">Output line prefix, default is an empty string</param>
  public static string ToString(Asset asset, string prefix="")
  {
    if (asset==null)
      return Dump.WriteLine("<null> (Asset)");

    string d = "";

    // Dump ordinary properties
    d += Dump.WriteLine(prefix+"<Ordinary Properties>");
    PropertyInfo[] pis = asset.GetType().GetProperties();
    foreach (PropertyInfo pi in pis)       
      d += Dump.ToString(pi,asset,"  "+prefix);

    // Dump asset properties
    d += Dump.WriteLine(prefix+"<Asset Properties>");
    for (int i=0; i<asset.Size; i++)
      d += Dump.ToString(asset[i],"  "+prefix);

    // Dump connected properties
    if (asset.NumberOfConnectedProperties>0)
    {
      d += Dump.WriteLine(prefix+"<Connected Properties>");
      foreach (AssetProperty apc in asset.GetAllConnectedProperties())
        d += Dump.ToString(apc,"  "+prefix);
    }

    // Dump single connected asset
    Asset sca = asset.GetSingleConnectedAsset();
    if (sca!=null)
    {
      d += Dump.WriteLine(prefix+"<Single Connected Asset>");
      d += Dump.ToString(sca,"  "+prefix);
    }

    return d;
  }

  /// <summary>Text dump of a Revit structural (phyical) asset</summary>
  /// <param name="asset">The structural asset</param>
  /// <param name="physicalPropSet">The physical property set element</param>
  /// <param name="prefix">Output line prefix, default is an empty string</param>
  public static string ToString(StructuralAsset asset, PropertySetElement physicalPropSet, string prefix="")
  {
    if (asset==null)
      return Dump.WriteLine("<null> (Asset)");

    string d = "";

    // Dump ordinary properties
    d += Dump.WriteLine(prefix+"<Ordinary Properties>");
    PropertyInfo[] pis = asset.GetType().GetProperties();
    foreach (PropertyInfo pi in pis)       
      d += Dump.ToString(pi,asset,"  "+prefix);

    // Dump thermal asset parameters
    d += Dump.WriteLine(prefix+"<Parameters>");
    ICollection<Parameter> parameters = physicalPropSet.GetOrderedParameters();
    foreach (Parameter p in parameters)
      d += Dump.ToString(p,"  "+prefix);

    return d;
  }

  /// <summary>Text dump of a Revit thermal asset</summary>
  /// <param name="asset">The thermal asset</param>
  /// <param name="thermalPropSet">The thermal property set element</param>
  /// <param name="prefix">Output line prefix, default is an empty string</param>
  public static string ToString(ThermalAsset asset, PropertySetElement thermalPropSet, string prefix="")
  {
    if (asset==null)
      return Dump.WriteLine("<null> (Asset)");

    string d = "";

    // Dump ordinary properties
    d += Dump.WriteLine(prefix+"<Ordinary Properties>");
    PropertyInfo[] pis = asset.GetType().GetProperties();
    foreach (PropertyInfo pi in pis)       
      d += Dump.ToString(pi,asset,"  "+prefix);

    // Dump thermal asset parameters
    d += Dump.WriteLine(prefix+"<Parameters>");
    ICollection<Parameter> parameters = thermalPropSet.GetOrderedParameters();
    foreach (Parameter p in parameters)
      d += Dump.ToString(p,"  "+prefix);

    return d;
  }

  /// <summary>Text dump of a Revit material</summary>
  /// <param name="material">The material</param>
  /// <param name="prefix">Output line prefix, default is an empty string</param>
  public static string ToString(Material material, string prefix="")
  {
    if (material==null)
      return Dump.WriteLine("<null> (Material)");

    string d = Dump.WriteLine(prefix+String.Format("Material \"{0}\"",material.Name));
    prefix = "  "+prefix;

    Document doc = material.Document;
    d += Dump.WriteLine(prefix+String.Format("<Document PathName>: {0}",doc.PathName));
    
    // Dump ordinary properties
    d += Dump.WriteLine(prefix+"<Ordinary Properties>");
    PropertyInfo[] pis = material.GetType().GetProperties();
    foreach (PropertyInfo pi in pis)       
      d += Dump.ToString(pi,material,"  "+prefix);

    // Dump built-in paramters
    d += Dump.WriteLine(prefix+"<Built-in Parameters>");
    foreach (BuiltInParameter bip in Enum.GetValues(typeof(BuiltInParameter)))
    {
      Parameter param = material.get_Parameter(bip);
      if (param!=null)
        d += Dump.ToString(param,"  "+prefix+"["+bip.ToString()+"] ");
    }

    // Dump other parameters
    d += Dump.WriteLine(prefix+"<Other Parameters>");
    ParameterMap parammap = material.ParametersMap;
    if (parammap.Size>0)
      foreach (Parameter param in parammap)
        Dump.ToString(param,"  "+prefix);
    else
        Dump.WriteLine("    "+prefix+"<none>");
 
    // Appearance (Rendering) Asset
    string s = "<Appearance Asset>";
    AppearanceAssetElement appearanceElement = doc.GetElement(material.AppearanceAssetId) as AppearanceAssetElement;
    if (appearanceElement!=null)
    {
      d += Dump.WriteLine(prefix+s);
      Asset appearanceAsset = appearanceElement.GetRenderingAsset();
      d += Dump.ToString(appearanceAsset,"  "+prefix);
    }
    else
      d += Dump.WriteLine(prefix+s+": -none-");

    // Physical (Structural) Asset
    s = "<Physical Asset>";
    PropertySetElement physicalPropSet = doc.GetElement(material.StructuralAssetId) as PropertySetElement;
    if (physicalPropSet!=null)
    {
      d += Dump.WriteLine(prefix+s);
      StructuralAsset physicalAsset = physicalPropSet.GetStructuralAsset();
      d += Dump.ToString(physicalAsset,physicalPropSet,"  "+prefix);
    }
    else
      d += Dump.WriteLine(prefix+s+": -none-");

    // Thermal Asset
    s = "<Thermal Asset>";
    PropertySetElement thermalPropSet = doc.GetElement(material.ThermalAssetId) as PropertySetElement;
    if (thermalPropSet!=null)
    {
      d += Dump.WriteLine(prefix+s);
      ThermalAsset thermalAsset = thermalPropSet.GetThermalAsset();
      d += Dump.ToString(thermalAsset,thermalPropSet,"  "+prefix);
    }
    else
      d += Dump.WriteLine(prefix+s+": -none-");

    return d;
  }
  
  /// <summary>Dumps the properties of all materials in a Revit document to a string</summary>
  /// <param name="document">The Revit document</param>
  /// <returns>The dump</returns>
  public static string AllMerials(Document document)
  {
    if (document==null)
      return "Document is <null>. Cannot dump anything.";

    // Get list of all materials in document
    IList<Material> materials =                                
      (new FilteredElementCollector(document))
      .OfClass(typeof(Material))
      .ToElements()
      .Cast<Material>()
      .ToList();
    
    // - Dump material properties to string
    string d = "";
    foreach (Material material in materials)
      d += Dump.ToString(material,"- ");

    return d;
  }
}

namespace Utils
{
  [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
  [Autodesk.Revit.DB.Macros.AddInId("BFB2B587-B93B-4025-9397-5BC1FDF07D39")]
	public partial class ThisDocument
	{
	  
	  #region Module startup and shutdown code

		private void Module_Startup(object sender, EventArgs e)
		{
		}

		private void Module_Shutdown(object sender, EventArgs e)
		{
		}

		#endregion

		#region Revit Macros generated code

		private void InternalStartup()
		{
			this.Startup += new System.EventHandler(Module_Startup);
			this.Shutdown += new System.EventHandler(Module_Shutdown);
		}

		#endregion

		#region Utilities
		
		/// <summary>Returns the path if the current C# source file</summary>
    /// <seealso href="https://stackoverflow.com/questions/47841441/how-do-i-get-the-path-to-the-current-c-sharp-source-code-file"
    ///   >Stackoverflow. How do I get the path to the current C# source code file?</seealso>
    public static string GetThisFilePath([CallerFilePath] string path = null)
    {
      return path;
    }

    #endregion

    public void Dump_Materials()
    {
      // Collect some info
      Document actDoc      = this.Application.ActiveUIDocument.Document;
      string   actDocFile  = Path.GetFileName(actDoc.PathName);
      string   thisDocPath = Path.GetDirectoryName(this.Document.PathName);
      string   thisDocFile = Path.GetFileName(this.Document.PathName);

      // Backup this source file in folder of GIMBA.rfa
      // NOTE: Revit 2022 is tending to lose the most recent changes to this source file!
      try
      {
        string srcFn = GetThisFilePath();
        string dstFn = Path.Combine(thisDocPath,thisDocFile+"#Utils#ThisDocument.bak.cs");
        File.Copy(srcFn,dstFn,true);
        // NOTE: Will only work as long project is open in SharpDevelop...
      }
      catch {/*...hence ignore any exceptions*/}
      
      // Dump all materials contained in active Revit document to a text file
      string       fn = Path.Combine(thisDocPath,"Materials_in_"+actDocFile+".txt");
      FileStream   fs = new FileStream(fn,FileMode.Create,FileAccess.Write);  
      StreamWriter fw = new StreamWriter(fs);  
      fw.WriteLine(Dump.AllMerials(actDoc)); // <-- See class Dump above!
      fw.Close(); 

      // Show wrap-up dialog
      TaskDialog td = new TaskDialog("Dump Materials Properties");
      td.MainInstruction = "A text dump of all materials residing in "
                         + actDocFile + " has been written.";
      td.MainContent     = "Dump file: "+ fn;
      td.CommonButtons   = TaskDialogCommonButtons.Close;
      td.AddCommandLink(
        TaskDialogCommandLinkId.CommandLink1,
        "Open dump file"
      );
      TaskDialogResult tdr = td.Show();
      switch (tdr)
      {
        case TaskDialogResult.CommandLink1:
          Process.Start(fn);
          break;
        default:
          // Nothing to be done
          break;
      }
    }
	}
}

// EOF