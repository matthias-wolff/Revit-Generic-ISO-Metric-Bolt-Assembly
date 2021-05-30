// Text dump of Revit materials properties
//
// Author    : Matthias Wolff
// Modified  : 2021-05-30
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
    
    // TODO: Implement Dump.Parameter(...)!
    return "";
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

    string dump = "";
    string n = String.IsNullOrEmpty(pn) ? pi.Name : pn;
    string t = String.IsNullOrEmpty(pt) ? pi.PropertyType.Name : pt;
    if (pi.GetIndexParameters().Length==0)
    {
      Object v = pi.GetValue(obj);
      dump += Dump.WriteLine(prefix+String.Format("{0} ({1}): {2}",n,t,v));
    }
    else
      // TODO: Implement indexed properties in Dump.Property()
      dump += Dump.WriteLine(prefix+String.Format("{0} ({1}): -indexed-",n,"?")); 
    
    return dump;
  }

  /// <summary>Simplified text dump of a Revit asset property</summary>
  /// <param name="aprop">The asset property</param>
  /// <param name="prefix">Output line prefix, default is an empty string</param>
  public static string ToString(AssetProperty ap, string prefix="")
  {
    if (ap==null)
      return Dump.WriteLine("<null> (AssetProperty)");
    
    string dump = "";
    
    // Dump (ordinary) properties if asset property object
    PropertyInfo[] pis = ap.GetType().GetProperties();      // Retrieve all properties of asset propery object
    PropertyInfo   piv = ap.GetType().GetProperty("Value"); // "Value" property of asset property object (if any)

    if (piv!=null)
      // Skip all properties except "Value"
      dump += Dump.ToString(piv,ap,prefix,ap.Name+".Value",ap.Type.ToString());
    else
    {
      // Dump all properties
      string t = ap.GetType().Name;
      string n = ap.Name;
      dump += Dump.WriteLine(prefix+String.Format("{0} ({1})",n,t));
      foreach (PropertyInfo pi in pis)
        dump += Dump.ToString(pi,ap,"  "+prefix);
    }
    
    // Rest will be indented one level more further
    prefix = "  "+prefix;

    // Dump connected properties
    if (ap.NumberOfConnectedProperties>0)
    {
      dump += Dump.WriteLine(prefix+"<Connected Properties>");
      foreach (AssetProperty apc in ap.GetAllConnectedProperties())
        dump += Dump.ToString(apc,"  "+prefix);
    }

    // Dump single connected asset
    Asset sca = ap.GetSingleConnectedAsset();
    if (sca!=null)
    {
      dump += Dump.WriteLine(prefix+"<Single Connected Asset>");
      dump += Dump.ToString(sca,"  "+prefix);
    }
    
    return dump;
  }

  /// <summary>Text dump of a Revit asset</summary>
  /// <param name="asset">The asset</param>
  /// <param name="prefix">Output line prefix, default is an empty string</param>
  public static string ToString(Asset asset, string prefix="")
  {
    if (asset==null)
      return Dump.WriteLine("<null> (Asset)");

    string dump = "";

    // Dump ordinary properties
    dump += Dump.WriteLine(prefix+"<Ordinary Properties>");
    PropertyInfo[] pis = asset.GetType().GetProperties();
    foreach (PropertyInfo pi in pis)       
      dump += Dump.ToString(pi,asset,"  "+prefix);

    // Dump asset properties
    dump += Dump.WriteLine(prefix+"<Asset Properties>");
    for (int i=0; i<asset.Size; i++)
      dump += Dump.ToString(asset[i],"  "+prefix);

    // Dump connected properties
    if (asset.NumberOfConnectedProperties>0)
    {
      dump += Dump.WriteLine(prefix+"<Connected Properties>");
      foreach (AssetProperty apc in asset.GetAllConnectedProperties())
        dump += Dump.ToString(apc,"  "+prefix);
    }

    // Dump single connected asset
    Asset sca = asset.GetSingleConnectedAsset();
    if (sca!=null)
    {
      dump += Dump.WriteLine(prefix+"<Single Connected Asset>");
      dump += Dump.ToString(sca,"  "+prefix);
    }
    
    return dump;
  }

  /// <summary>Text dump of a Revit material</summary>
  /// <param name="material">The material</param>
  /// <param name="prefix">Output line prefix, default is an empty string</param>
  public static string ToString(Material material, string prefix="")
  {
    if (material==null)
      return Dump.WriteLine("<null> (Material)");

    string dump = "";
    
    try
    {
      dump += Dump.WriteLine(prefix+String.Format("Material \"{0}\"",material.Name));
      prefix = "  "+prefix;

      Document doc = material.Document;
      dump += Dump.WriteLine(prefix+String.Format("<Document PathName>: {0}",doc.PathName));
      
      // Dump ordinary properties
      dump += Dump.WriteLine(prefix+"<Ordinary Properties>");
      PropertyInfo[] pis = material.GetType().GetProperties();
      foreach (PropertyInfo pi in pis)       
        dump += Dump.ToString(pi,material,"  "+prefix);
   
      // Appearance (Rendering) Asset
      string s = "<Appearance Asset>";
      AppearanceAssetElement appearanceElement = doc.GetElement(material.AppearanceAssetId) as AppearanceAssetElement;
      if (appearanceElement!=null)
      {
        dump += Dump.WriteLine(prefix+s);
        Asset appearanceAsset = appearanceElement.GetRenderingAsset();
        dump += Dump.ToString(appearanceAsset,"  "+prefix);
      }
      else
        dump += Dump.WriteLine(prefix+s+": -none-");
  
      // Physical (Structural) Asset
      s = "<Physical Asset>";
      PropertySetElement physicalPropSet = doc.GetElement(material.StructuralAssetId) as PropertySetElement;
      if (physicalPropSet!=null)
      {
        StructuralAsset physicalAsset = physicalPropSet.GetStructuralAsset();
        dump += Dump.WriteLine("  "+prefix+"Name: "+physicalAsset.Name);
        ICollection<Parameter> physicalParameters = physicalPropSet.GetOrderedParameters();
        foreach (Parameter p in physicalParameters)
          dump += Dump.ToString(p,"  "+prefix);
      }
      else
        dump += Dump.WriteLine(prefix+s+": -none-");
  
      // Thermal Asset
      s = "<Thermal Asset>";
      PropertySetElement thermalPropSet = doc.GetElement(material.ThermalAssetId) as PropertySetElement;
      if (thermalPropSet!=null)
      {
        ThermalAsset thermalAsset = thermalPropSet.GetThermalAsset();
        dump += Dump.WriteLine("  "+prefix+"Name: "+thermalAsset.Name);
        ICollection<Parameter> thermalParameters = thermalPropSet.GetOrderedParameters();
        foreach (Parameter p in thermalParameters)
          dump += Dump.ToString(p,"  "+prefix);
      }
      else
        dump += Dump.WriteLine(prefix+s+": -none-");
    }
    catch (Exception e)
    {
      dump += Dump.WriteLine(e.ToString());
    }

    return dump;
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
    string dump = "";
    foreach (Material material in materials)
      dump += Dump.ToString(material,"- ");

    return dump;
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
      // Backup this source file in folder of GIMBA.rfa
      // NOTE: Revit 2022 is tending to lose the most recent changes to this source file!
      string thisDocPath = Path.GetDirectoryName(this.Document.PathName);
      string thisDocFile = Path.GetFileName(this.Document.PathName);
      File.Copy(GetThisFilePath(),Path.Combine(thisDocPath,"GIMBA.rfa#Utils#ThisDocument.bak.cs"),true);
      
      // Dump all materials contained in active Revit document
      Document document   = this.Application.ActiveUIDocument.Document;
      string   actDocFile = Path.GetFileName(document.PathName);
      string   dump       = Dump.AllMerials(Document);
      
      // Write dump to a text file
      string       fn = Path.Combine(thisDocPath,"Materials_in_"+actDocFile+".txt");
      FileStream   fs = new FileStream(fn,FileMode.Create,FileAccess.Write);  
      StreamWriter fw = new StreamWriter(fs);  
      fw.WriteLine(dump);  
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