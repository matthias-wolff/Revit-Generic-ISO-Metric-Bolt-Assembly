// GIMBA - Generic ISO Metric Bolt Assembly
// Service macros
//
// Author    : Matthias Wolff
// Modified  : 20210526-1702
// References:
// [1] About Macro Manager and the Revit Macro IDE
//     https://help.autodesk.com/view/RVT/2022/ENU/?guid=GUID-071913D8-214A-45AB-A798-A81653E77F88
// [2] Revit API 2022
//     https://www.revitapidocs.com/2022/
// [3] C# Language Reference
//     https://docs.microsoft.com/en-us/dotnet/csharp/language-reference

using System;                           // For Object, Exception, etc.
using System.Diagnostics;               // For Process.Start
using System.Text.RegularExpressions;   // For regular expressions
using System.IO;                        // For file input/output
using System.Reflection;                // For listing asset and other properties
using System.Runtime.CompilerServices;  // For determining path of this source file
using Autodesk.Revit.UI;                // For TaskDialog, UIApplication, etc.
using Autodesk.Revit.DB;                // For Document, Material, etc.
using Autodesk.Revit.DB.Visual;         // For appearance asset
using System.Collections.Generic;       // For Dictionary, IEnumerable, etc.
using System.Linq;                      // For .Cast, .ToList, etc

/// <summary>Information on tasks completed by a Revit macro and on problems</summary>
public class Log
{
  
  #region Log File
  
  /// <summary>The fully qualified log file name (default is in temp folder)</summary>
  protected static string LogFileName = Path.Combine(Path.GetTempPath(),"GIMBA.log");
  
  /// <summary>Returns the log file name</summary>
  public static string GetLogFileName()
  {
    return Log.LogFileName;
  }
  
  /// <summary>Deletes the log file.</summary>
  protected static void DeleteLogFile()
  { 
    try { File.Delete(LogFileName); } catch {/* Ignore errors*/}
  }

  #endregion

  #region Log keeping
  
  /// <summary>Creates a new macro completion information object</summary>
  /// <param name="name">The macro name</param>
  /// <param name="doc">The Revit document</param>
  public static void Begin(string name, Document doc)
  {
    // Try to make log file path in Revit document's folder, use temp folder on errors
    try
    {
      string path = Path.GetDirectoryName(doc.PathName);
      Log.LogFileName = Path.Combine(path,"GIMBA.log");
      
      // Backup this source file (Revit cost me 10 hours of wirking yesterday... :((
      string ts = DateTime.Now.ToString("yyyyMMdd-HHmmss");
      File.Copy(Log.GetThisFilePath(),Path.Combine(path,"GIMBA.rfa#ThisDocument.bak.cs"),true);
      File.Copy(Log.GetThisFilePath(),Path.Combine(path,"GIMBA.rfa#ThisDocument.bak-"+ts+".cs"),true);
    }
    catch (Exception e)
    {
      Log.WL(e.ToString());
    }

    // Initialize log file
    DeleteLogFile();
    Log.WL("-------------------------------------------------------------------------------");
    Log.WL(String.Format(
      "Pass of {0}, timestamp {1}",
      name,
      DateTime.Now.ToString("yyyyMMdd-HHmmss")
    ));
  }

  /// <summary>Writes a line into the log file.</summary>
  /// <param name="line">The line</param>
  /// <seealso cref="https://github.com/jeremytammik/rvtmetaprop/blob/master/rvtmetaprop/Command.cs#L426-L437"/>
  public static void WL(string line="")
  {
    // Slow but safe...
    // TODO: Surround by try-catch? But how to display the error?
    FileStream   fs = new FileStream(LogFileName, FileMode.Append|FileMode.Create, FileAccess.Write);  
    StreamWriter fw = new StreamWriter(fs);  
    fw.WriteLine(line);  
    fw.Flush();  
    fw.Close(); 
  }

  /// <summary>Finishes with the log and shows a dialog informing the user on the task completion
  /// status including any errors.</summary>
  /// <param name="e">Top level exception to report on (optional)</param>
  public static void End(Exception e=null)
  {
    // Finish the log file
    if (e!=null)
      WL(e.ToString());
    WL();
    WL("Pass complete");
  }

  #endregion
  
  #region Miscellaneous Helpers
  
  /// <summary>Returns the path if the current C# source file</summary>
  /// <seealso href="https://stackoverflow.com/questions/47841441/how-do-i-get-the-path-to-the-current-c-sharp-source-code-file"
  ///   >Stackoverflow. How do I get the path to the current C# source code file?</seealso>
  public static string GetThisFilePath([CallerFilePath] string path = null)
  {
    return path;
  }

  /// <summary>Formats a human-readable message including a count</summary>
  /// <param name="count">The count</param>
  /// <param name="format">The message format string (see <see cref="String.Format"/>). Field {0} will 
  /// be filled with the <paramref name="count"/> or "no" if the <paramref name="count"/> is zero.
  /// Field {1} will be filled with the appropriate suffix.</param>
  /// <param name="pluralSuffix">The plural suffix (optional, default is "s")</param>
  /// <param name="singularSuffix">The singular suffix (optional, default is "")</param>
  /// <returns></returns>
  public static string MakeCountMsg(
    int    count, 
    string format, 
    string pluralSuffix   = "s", 
    string singularSuffix = ""
  )
  {
    string sCount  = count!=0 ? count+"" : "no";
    string sSuffix = count!=1 ? pluralSuffix : singularSuffix;
    return String.Format(format,sCount,sSuffix);
  }

  /// <summary>Formats a human-readable message including a count and appends "--> ok" or "--> 
  /// NOT OK" depending on a condition</summary>
  /// <param name="count">The count</param>
  /// <param name="condition">The condition</param>
  /// <param name="format">The message format string (see <see cref="String.Format"/>). Field {0} will 
  /// be filled with the <paramref name="count"/> or "no" if the <paramref name="count"/> is zero.
  /// Field {1} will be filled with the appropriate suffix.</param>
  /// <param name="pluralSuffix">The plural suffix (optional, default is "s")</param>
  /// <param name="singularSuffix">The singular suffix (optional, default is "")</param>
  /// <returns></returns>
  public static string MakeCountOkMsg(
    int    count, 
    bool   condition, 
    string format, 
    string pluralSuffix="s", 
    string singularSuffix=""
  )
  {
    string s = Log.MakeCountMsg(count,format,pluralSuffix,singularSuffix);
    return s + (condition ? " --> ok" : " --> NOT OK");
  }
 
  #endregion

}

/// <summary>Provides text dumps of Revit materials and assets</summary>
/// <seealso href="https://thebuildingcoder.typepad.com/blog/2019/11/material-physical-and-thermal-assets.html"
///   The Building Coder. Material, Physical and Thermal Assets. Nov. 2019. Retrieved May 2021.></seealso>
/// <seealso href="https://docs.microsoft.com/de-de/dotnet/api/system.reflection.propertyinfo.getvalue?view=net-5.0"
///   Microsoft. .NET API Documentation, Method PropertyInfo.GetValue. Retrieved May 2021.></seealso>
public static class Dump
{
  /// <summary>Writes one line of text. Modify this method to write to a file or whatever</summary>
  /// <param name="line">The text line to write</param>
  public static void WriteLine(string line="")
  {
    //Console.WriteLine(line);
    Log.WL(line);
  }
  
  /// <summary>Text dump of a Revit parameter</summary>
  /// <param name="aprop">The parameter</param>
  /// <param name="prefix">Output line prefix, default is an empty string</param>
  public static void Parameter(Parameter param, string prefix="")
  {
    // TODO: Implement Dump.Parameter(...)!
  }
  
  /// <summary></summary>Text dump of a C# property</summary>
  /// <param name="pi">The property</param>
  /// <param name="obj">The instance for which tp display the property</param> 
  /// <param name="prefix">Output line prefix, default is an empty string</param>
  /// <param name="pn">Alternative name of property, default is <c>pi.Name</c></param>
  /// <param name="pt">Alternative type name of property, default is <c>pi.PropertyType.Name</c></param>
  public static void Property(PropertyInfo pi, Object obj, string prefix="", string pn=null, string pt=null)
  {
    Debug.Assert(pi!=null);

    string n = String.IsNullOrEmpty(pn) ? pi.Name : pn;
    string t = String.IsNullOrEmpty(pt) ? pi.PropertyType.Name : pt;
    if (pi.GetIndexParameters().Length==0)
    {
      Object v = pi.GetValue(obj);
      Dump.WriteLine(prefix+String.Format("{0} ({1}): {2}",n,t,v));
    }
    else
      // TODO: Listing?
      Dump.WriteLine(prefix+String.Format("{0} ({1}): -indexed-",n,"?")); 
  }

  /// <summary>Simplified text dump of a Revit asset property</summary>
  /// <param name="aprop">The asset property</param>
  /// <param name="prefix">Output line prefix, default is an empty string</param>
  public static void AssetProperty(AssetProperty ap, string prefix="")
  {
    Debug.Assert(ap!=null);
    
    // Dump (ordinary) properties if asset property object
    PropertyInfo[] pis = ap.GetType().GetProperties();      // Retrieve all properties of asset propery object
    PropertyInfo   piv = ap.GetType().GetProperty("Value"); // "Value" property of asset property object (if any)

    if (piv!=null)
      // Skip all properties except "Value"
      Dump.Property(piv,ap,prefix,ap.Name+".Value",ap.Type.ToString());
    else
    {
      // Dump all properties
      string t = ap.GetType().Name;
      string n = ap.Name;
      Dump.WriteLine(prefix+String.Format("{0} ({1})",n,t));
      foreach (PropertyInfo pi in pis)
        Dump.Property(pi,ap,"  "+prefix);
    }
    
    // Rest will be indented one level more further
    prefix = "  "+prefix;

    // Dump connected properties
    if (ap.NumberOfConnectedProperties>0)
    {
      Dump.WriteLine(prefix+"<Connected Properties>");
      foreach (AssetProperty apc in ap.GetAllConnectedProperties())
        Dump.AssetProperty(apc,"  "+prefix);
    }

    // Dump single connected asset
    Asset sca = ap.GetSingleConnectedAsset();
    if (sca!=null)
    {
      Dump.WriteLine(prefix+"<Single Connected Asset>");
      Dump.Asset(sca,"  "+prefix);
    }
  }

  /// <summary>Text dump of a Revit asset</summary>
  /// <param name="asset">The asset</param>
  /// <param name="prefix">Output line prefix, default is an empty string</param>
  public static void Asset(Asset asset, string prefix="")
  {
    Debug.Assert(asset!=null);
    
    // Dump ordinary properties
    Dump.WriteLine(prefix+"<Ordinary Properties>");
    PropertyInfo[] pis = asset.GetType().GetProperties();
    foreach (PropertyInfo pi in pis)       
      Dump.Property(pi,asset,"  "+prefix);

    // Dump asset properties
      Dump.WriteLine(prefix+"<Asset Properties>");
    for (int i=0; i<asset.Size; i++)
      Dump.AssetProperty(asset[i],"  "+prefix);

    // Dump connected properties
    if (asset.NumberOfConnectedProperties>0)
    {
      Dump.WriteLine(prefix+"<Connected Properties>");
      foreach (AssetProperty apc in asset.GetAllConnectedProperties())
        Dump.AssetProperty(apc,"  "+prefix);
    }

    // Dump single connected asset
    Asset sca = asset.GetSingleConnectedAsset();
    if (sca!=null)
    {
      Dump.WriteLine(prefix+"<Single Connected Asset>");
      Dump.Asset(sca,"  "+prefix);
    }
  }

  /// <summary>Text dump of a Revit material</summary>
  /// <param name="material">The material</param>
  /// <param name="prefix">Output line prefix, default is an empty string</param>
  public static void Material(Material material, string prefix="")
  {
    Debug.Assert(material!=null);

    try
    {
      Dump.WriteLine(prefix+String.Format("Material \"{0}\"",material.Name));
      prefix = "  "+prefix;

      Document doc = material.Document;
      Dump.WriteLine(prefix+String.Format("<Document PathName>: {0}",doc.PathName));
      
          // Dump ordinary properties
      Dump.WriteLine(prefix+"<Ordinary Properties>");
      PropertyInfo[] pis = material.GetType().GetProperties();
      foreach (PropertyInfo pi in pis)       
        Dump.Property(pi,material,"  "+prefix);
   
      // Appearance (Rendering) Asset
      string s = "<Appearance Asset>";
      AppearanceAssetElement appearanceElement = doc.GetElement(material.AppearanceAssetId) as AppearanceAssetElement;
      if (appearanceElement!=null)
      {
        Dump.WriteLine(prefix+s);
        Asset appearanceAsset = appearanceElement.GetRenderingAsset();
        Dump.Asset(appearanceAsset,"  "+prefix);
      }
      else
        Dump.WriteLine(prefix+s+": -none-");
  
      // Physical (Structural) Asset
      s = "<Physical Asset>";
      PropertySetElement physicalPropSet = doc.GetElement(material.StructuralAssetId) as PropertySetElement;
      if (physicalPropSet!=null)
      {
        StructuralAsset physicalAsset = physicalPropSet.GetStructuralAsset();
        Dump.WriteLine("  "+prefix+"Name: "+physicalAsset.Name);
        ICollection<Parameter> physicalParameters = physicalPropSet.GetOrderedParameters();
        foreach (Parameter p in physicalParameters)
          Dump.Parameter(p,"  "+prefix);
      }
      else
        Dump.WriteLine(prefix+s+": -none-");
  
      // Thermal Asset
      s = "<Thermal Asset>";
      PropertySetElement thermalPropSet = doc.GetElement(material.ThermalAssetId) as PropertySetElement;
      if (thermalPropSet!=null)
      {
        ThermalAsset thermalAsset = thermalPropSet.GetThermalAsset();
        Dump.WriteLine("  "+prefix+"Name: "+thermalAsset.Name);
        ICollection<Parameter> thermalParameters = thermalPropSet.GetOrderedParameters();
        foreach (Parameter p in thermalParameters)
          Dump.Parameter(p,"  "+prefix);
      }
      else
        Dump.WriteLine(prefix+s+": -none-");
    }
    catch (Exception e)
    {
      Dump.WriteLine(e.ToString());
    }
  }
}

/// <summary>ISO metric thread properties object</summary>
/// <seealso href="http://www.iso-gewinde.at">Thread Calculator (German)</seealso>
public class ThreadGeometry
{

  #region Static API

  /// <summary>ISO metric threads dictionary</summary>
  private static Dictionary<string,ThreadGeometry> Threads;

  
  /// <summary>Returns the ISO metric threads dictionary.</summary>
  /// <seealso href="http://www.iso-gewinde.at">Thread Calculator (German)</seealso>
  protected internal static IList<ThreadGeometry> getList()
  {
    if (ThreadGeometry.Threads==null)
    {
      ThreadGeometry.Threads = new Dictionary<string,ThreadGeometry>();
//      new ThreadGeometry( 3,0.5 ); // NOTE: Newly created objects register themselves with dictionary
//      new ThreadGeometry( 4,0.7 );
      new ThreadGeometry( 5,0.8 );
//      new ThreadGeometry( 6,1   );
//      new ThreadGeometry( 8,1.25);
//      new ThreadGeometry(10,1.5 );
//      new ThreadGeometry(12,1.75);
      // TODO: Add threads M14..M64
    }
    return ThreadGeometry.Threads.Values.ToList();
  }
  
  /// <summary>Get one thread properties object</summary>
  /// <param name="D">The nominal thread diameter</param>
  /// <returns>The thread properties object or <c>null</c> if no object was found</returns>
  protected internal static ThreadGeometry Get(int D)
  {
    string key = String.Format("M{0}",D);
    try
    {
      return Threads[key];
    }
    catch (Exception e)
    {
      Log.WL(e.ToString());
    }
    return null;                     
  }

  #endregion
  
  #region Instance API

  /// <summary>Nominal diameter in millimeters</summary>
  protected internal int D; 

  /// <summary>Thread pitch in millimeters</summary>
  protected internal double P;

  /// <summary>Nominal circumference in millimeters</summary>
  protected internal double u;

  /// <summary>Thread pitch angle in degrees</summary>
  protected internal double beta;

  /// <summary>Creates a new ISO metric thread properties object.</summary>
  /// <param name="D">Nominal diameter in millimeters</param>
  /// <param name="P">Thread pitch in millimeters</param>
  private ThreadGeometry(int D, double P)
  {
    this.D    = D;
    this.P    = P;
    this.u    = Math.PI*this.D;
    this.beta = Math.Atan2(this.P,this.u)*180/Math.PI;

    ThreadGeometry.Threads.Add(String.Format("M{0}",this.D),this);
  }

  public override string ToString()
  {
    return string.Format("[IMThread D={0}, P={1}, u={2}, beta={3}]", D, P, u, beta);
  }

  #endregion
  
}

/// <summary>Manages GIMBA ISO metric screw thread materials</summary>
public static class ThreadMaterials
{

  #region Static Properties
  
  /// <summary>Plain materials</summary>
  public static string sPMATS = "GIMBA - {0}";
  public static Regex  PMATS  = new Regex("^"+String.Format(sPMATS,"([^-]+)")+"$");
 
  /// <summary>Thread materials</summary>
  public static string sTMATS = "GIMBA - {0} - M{1} thread";
  public static Regex  TMATS = new Regex("^"+String.Format(sTMATS,"(.+)",@"(\d+)")+"$");
 
  /// <summary>Thread template materials</summary>
  public static string sVMATS = @"^GIMBA - {0} - Thread template$";
  public static Regex  VMATS = new Regex("^"+String.Format(sVMATS,"(.+)")+"$");

  #endregion

  #region User Inferface
  
  /// <summary>Shows the welcome dialog</summary>
  /// <param name="td">The task dialog to use</param>
  /// <param name="doc">The Revit document to work on</param>
  /// <param name="tgeos">List of thread geometries</param>
  /// <param name="tMats">List of existing thread materials</param>
  /// <param name="vMats">List of thread template materials (valid or invalid)</param>
  /// <param name="vMatsValid">List of valid thread template materials</param>
  /// <returns></returns>
  public static TaskDialogResult DoWelcomeDialog(
    TaskDialog            td,
    Document              doc,
    IList<ThreadGeometry> tgeos,
    IList<Material>       tMats,
    IList<Material>       vMats, 
    IList<Material>       vMatsValid
  )
  {
    // Big decision: If ok, macro operations can be performed
    bool   ready   = vMatsValid.Count>0 && tgeos.Count>0;
    string docName = doc.Title + (doc.IsFamilyDocument ? ".rfa" : ".rvt");

    // Prepare texts
    string stgeos  = "* Found {0} thread geometr{1} --> ";
    string stMats  = "* Found {0} existing thread material{1} --> ok";
    string svvMats = "* Found {0} valid template material{1} --> ";
    string sivMats = "* Found {0} invalid template material{1} --> ";
    stgeos  = Log.MakeCountMsg(tgeos.Count,stgeos,"ies","y");
    stMats  = Log.MakeCountMsg(tMats.Count,stMats);
    svvMats = Log.MakeCountMsg(vMats.Count,svvMats);
    sivMats = Log.MakeCountMsg(vMats.Count-vMatsValid.Count,sivMats);
    stgeos  += tgeos.Count>0 ? "ok" : "NOT OK" ;
    svvMats += vMats.Count>0 ? "ok" : "NOT OK" ;
    sivMats += (vMats.Count-vMatsValid.Count)>0 ? "ignore" : "ok" ;
    td.ExpandedContent = "Summary of pre-check results:\n"
                       + stgeos  + "\n"
                       + stMats  + "\n"
                       + svvMats + "\n"
                       + sivMats;
    if (!ready)
      td.ExpandedContent += "\n\n"
                         +  "Issues marked with \"NOT OK\" obstruct operation.";
   
    // Configure task dialog    
    td.MainInstruction   = "This macro performs batch operations on ISO metric "
                         + "screw thread materials. ";
    td.MainContent       = "Working document is: "+docName+"\n";
    // FIXME: This displays GIMBA.html in editor, not in browser
    td.FooterText        = Path.Combine(Path.GetDirectoryName(doc.PathName),"GIMBA.html");
    td.FooterText        = "<a href=\""+td.FooterText+"\">Help</a>";
    td.AllowCancellation = true;
    td.TitleAutoPrefix   = false;
    if (ready)
    {
      // - Good to go: operations can be performed
      td.Title            = "Good to Go...";
      td.MainInstruction += "Please select an option!";
      td.MainContent     += "See details for results of pre-checks.";
      td.CommonButtons    = TaskDialogCommonButtons.Cancel;
      string sCmdSupp1    = "Will create {0} thread material{1}.";
      string sCmdSupp2    = "Please choose below whether {0} existing thread material{1} "
                          + "shall be overwritten.";
      sCmdSupp1           = Log.MakeCountMsg(tgeos.Count*vMats.Count,sCmdSupp1);
      sCmdSupp2           = Log.MakeCountMsg(tMats.Count,sCmdSupp2);
      td.AddCommandLink(
        TaskDialogCommandLinkId.CommandLink1,
        "Create Thread Materials",
        sCmdSupp1 + (tMats.Count>0 ? " "+sCmdSupp2 : "")
      );
      if (tMats.Count>0)
      {
        td.VerificationText = "Create shall overwrite existing thread material{1}";
        td.VerificationText = Log.MakeCountMsg(tMats.Count,td.VerificationText);
        sCmdSupp1           = "Will delete {0} existing thread material{1}. "
                            + "Template or other materials will not be deleted!";
        td.AddCommandLink(
          TaskDialogCommandLinkId.CommandLink2,
          "Delete Thread Materials",
          Log.MakeCountMsg(tMats.Count,sCmdSupp1)
        );
      }
    }
    else
    {
      // - Pre-checks failed: nothing can be done
      td.Title            = "Pre-Checks Failed";
      td.MainInstruction += "Pre-checks failed. No operation is possible on document.";
      td.FooterText      += " - <a href=\""+Log.GetLogFileName()+"\">View log file</a>";
      td.MainContent     += "See details and log file for further information.";
      td.CommonButtons    = TaskDialogCommonButtons.Close;
    }
    
    // Show task dialog
    TaskDialogResult tdr = td.Show();
    return tdr;
//    if (TaskDialogResult.CommandLink1 == tdr)
//    try { Process.Start(LogFileName); } catch {/*Ignore errors*/}
  }
  
  #endregion

  #region Asset Propeties Helpers
  
  /// <summary>Retries an asset property through which the property's value van be accessed</summary>
  /// <param name="asset">The asset to retrieve the property from</param>
  /// <param name="name">The name of the asset property to retrieve</param>
  /// <returns>The asset property</returns>
  /// <exception cref="Exception">If either argument is <c>null</c> or no property with the 
  /// specified name exists.</exception>
  public static T GetAP<T>(Asset asset, string name) where T : AssetProperty
  {
    return asset.FindByName(name) as T;
  }

  #endregion
  
  #region Workers on Materials

  /// <summary>Finds thread materials in a Revit document</summary>
  /// <param name="doc">Revit document to find materials in</param>
  /// <param name="filter">Filter for material names to find, a string (to match name exactly) or a Regex</param>
  /// <returns></returns>
  public static IEnumerable<Material> Find(Document doc, Object filter=null)
  {
    Regex rx = new Regex(".*");
    if (filter!=null)
      if (filter is string)
      {
        string s = filter as string;
        s  = s.Replace("(","\\(").Replace(")","\\)");
        rx = new Regex("^"+s+"$");
      }
      else if (filter is Regex)
        rx = filter as Regex;
      else
        throw new ArgumentException("Argument <filter> must be null, a string or a Regex");

    return new FilteredElementCollector(doc)
               .OfClass(typeof(Material))
               .ToElements()
               .Cast<Material>()
               .Where(mat => rx.IsMatch(mat.Name));
  }

  /// <summary>Checks whether a material is suitable as a template for thread materials.</summary>
  /// <param name="mat">The material</param>
  /// <exception cref="Exception">if the material is not suitable (exception message will contain 
  /// details)</exception>
  public static void CheckTemplate(Material vMat)
  {
    string prefix = "  - ";
    
    // Material cannot be null
    if (vMat==null)
      throw new Exception("Material is <null>"); // TODO: Add solution

    // Long name to display in messages
    string lname = String.Format("Material \"{0}\"",vMat.Name);
    Log.WL(String.Format(prefix+"Material \"{0}\"",vMat.Name));
    prefix = "  "+prefix;
    
    // Assert material residing in a document
    Document doc = vMat.Document;
    if (doc==null)
    {
      string problem = " does not reside in a Revit document";
      Log.WL(prefix+"Material"+problem+" -> FAILED");
      throw new Exception(lname+problem);
    }
    Log.WL(String.Format(prefix+"Material resides in document \"{0}\" -> OK",doc.Title));
    
    // Assert name is "GIMBA - <plain material name> - Thread template";
    string pattern = String.Format(sVMATS,"<plain material name>");
    if (!ThreadMaterials.VMATS.IsMatch(vMat.Name))
    {
      string problem = " has an invalid name, should be \""+pattern+"\"";
      Log.WL(prefix+"Material"+problem+" -> FAILED");
      throw new Exception(lname+problem);
    }
    Log.WL(String.Format(prefix+"Material name matches \"{0}\" -> OK",pattern));

    // Assert material has appearance asset
    try
    {
      ElementId              vAppAssElemId = vMat.AppearanceAssetId;
      AppearanceAssetElement vAppAssElem   = doc.GetElement(vAppAssElemId) as AppearanceAssetElement;
      Asset                  vAppAss        = vAppAssElem.GetRenderingAsset();
      Log.WL(String.Format(prefix+"Material has an appearance asset -> OK"));
 
      // Assert appearance asset has a gradient metal pattern shader
      try
      {
        AssetPropertyReference tMpsRef = GetAP<AssetPropertyReference>(vAppAss,"metal_pattern_shader");
        Asset tMpsApp = tMpsRef.GetSingleConnectedAsset();
        if (tMpsApp==null)
          throw new NullReferenceException();
        Log.WL(prefix+"Appearance asset has a metal pattern shader -> OK");
        
        // Assert metal pattern shader has gradient procedure map
        string baseSchema = GetAP<AssetPropertyString>(tMpsApp,"BaseSchema").Value;
        if ("GradientSchema".Equals(baseSchema))
          Log.WL(String.Format(prefix+"Metal pattern shader is a gradient procedure map -> OK"));
        else
        {
          Log.WL(String.Format(prefix+"Metal pattern shader is not a gradient procedure map -> FAILED"));
          throw new Exception("Base schema of metal pattern shader is \""+baseSchema+"\", should be \"GradientSchema\"");
        }
      }
      catch (Exception e)
      {
        string problem = "Appearance asset has not gradient metal pattern shader";
        Log.WL(prefix+problem+" -> FAILED");
        Log.WL(e.ToString());
        throw new Exception(problem,e);
      }
    }
    catch (Exception e)
    {
      string problem = " has no appearance asset or the appearance asset cannot be accessed";
      Log.WL(prefix+"Material"+problem+" -> FAILED");
      Log.WL(e.ToString());
      throw new Exception(lname+problem,e);
    }

  }

  /// <summary>Deletes a GIMBA thread material. For sefaety, the method will not delete any materials
  ///   whose name does not match <c>ThreadMaterials.TMATS</c>.</summary>
  /// <param name="material">The GIMBA thread material to be deleted</param>
  /// <exception cref="ArgumentException">If the material is not a GIMBA thread material</exception>
  private static void Delete(Material tMat)
  {
    Debug.Assert(tMat!=null);
    Debug.Assert(tMat.Document!=null);

    Log.WL(String.Format("- Deleting thread material \"{0}\"",tMat.Name));

    Document doc            = tMat.Document;
    ElementId tMatElemId    = tMat.Id;
    ElementId tAppAssElemId = tMat.AppearanceAssetId;

    doc.Delete(tMatElemId   );
    doc.Delete(tAppAssElemId);
  }

  /// <summary>Creates a new thread material by duplicating and editing a template material</summary>
  /// <param name="vMat">The thread template material to base the new thread material on</param>
  /// <param name="tgeo">The thread properties</param>
  /// <returns>The newly created material</returns>
  /// <seealso cref="https://www.revitapidocs.com/2019/96d557aa-e446-49c5-11cd-59fda2459e82.htm"/>
  private static Material Create(Material vMat, ThreadGeometry tgeo)
  {
    string prefix = "- ";
    //Dump.Material(vMat,prefix);
    Log.WL(prefix+String.Format(
      "Creating M{0} thread material from template \"{1}\"",
      tgeo.D,
      vMat.Name
     ));

    // Initialize
    Document doc         = vMat.Document;
    string   category    = VMATS.Match(vMat.Name).Groups[1]+"";   // E.g. "Steel galvanized"
    string   title       = String.Format(sTMATS,category,tgeo.D); // E.g. "GIMBA - Steel galvanized - M12 thread"
    string   description = String.Format(                         // E.g. "Generic ISO metric bolt assembly, steel galvanized with M12 tread"
                             "Generic ISO metric bolt assembly: {0} with M{1} tread", 
                             category,tgeo.D
                           );

    // Duplicate and edit template material
    ElementId              tAppAssElemId = ElementId.InvalidElementId;
    ElementId              vAppAssElemId = vMat.AppearanceAssetId;
    AppearanceAssetElement vAppAssElem   = doc.GetElement(vAppAssElemId) as AppearanceAssetElement;
  
    // - Duplicate the material
    Material tMat = vMat.Duplicate(title);

    // - Duplicate the appearance asset and the asset(s) in it
    AppearanceAssetElement tAppAssElem = vAppAssElem.Duplicate(title);

    // - Assign the asset element to the material
    tMat.AppearanceAssetId = tAppAssElem.Id;
    tAppAssElemId          = tAppAssElem.Id;

    // Make changes to the duplicate appearance asset
    using (AppearanceAssetEditScope editScope = new AppearanceAssetEditScope(vAppAssElem.Document /*TOOO: use doc*/))
    {
      // - Get editable copy of appearance asset
      Asset tAppAss = editScope.Start(tAppAssElemId);

      // - Edit appearance asset properties
      string keyword = GetAP<AssetPropertyString>(tAppAss,"keyword").Value;
      //GetAP<AssetPropertyString>(tAppAss,"category"   ).Value = category;
      GetAP<AssetPropertyString>(tAppAss,"description").Value = description;
      GetAP<AssetPropertyString>(tAppAss,"keyword"    ).Value = keyword+String.Format(":M{0}",tgeo.D);

      // - Modify single connected asset of metal_pattern_shader
      AssetPropertyReference tMpsRef = GetAP<AssetPropertyReference>(tAppAss,"metal_pattern_shader");
      Asset                  tMpsApp = tMpsRef.GetSingleConnectedAsset();
      GetAP<AssetPropertyDistance>(tMpsApp,"texture_RealWorldScaleX" ).Value = tgeo.P / 25.4;
      GetAP<AssetPropertyDistance>(tMpsApp,"texture_RealWorldScaleY" ).Value = tgeo.u / 25.4;
      GetAP<AssetPropertyDouble  >(tMpsApp,"texture_WAngle"          ).Value = 90 - tgeo.beta;
      GetAP<AssetPropertyBoolean >(tMpsApp,"texture_ScaleLock"       ).Value = false;
      GetAP<AssetPropertyBoolean >(tMpsApp,"texture_URepeat"         ).Value = true;
      GetAP<AssetPropertyBoolean >(tMpsApp,"texture_VRepeat"         ).Value = true;

      // - Commit edit
      editScope.Commit(true);
    }

    // Fertsch
    return tMat;
  }

  #endregion

  #region Main Function of Thread Materials Macro

  /// <summary>Main method of macro "Thread_Materials"</summary>
  /// <param name="doc">Revit document to load the thread materials template from and 
  /// to create the thread materials in</param>
	public static void MacroMain(Document doc)
	{
    Debug.Assert(doc!=null);

    // Pre-Checks: Try to make sure that actual operation will not fail
    Log.WL();
    Log.WL("Pre-Checks");

    // - Find and check thread geometries
    Log.WL("- Searching thread geometries");
    IList<ThreadGeometry> tgeos = ThreadGeometry.getList();
    foreach (ThreadGeometry tgeo in tgeos)
      Log.WL("   - "+tgeo);

    // - Find exisiting thread materials
    Log.WL("- Searching existing thread materials");
    IList<Material> tMats = ThreadMaterials.Find(doc,ThreadMaterials.TMATS).ToList();
    foreach (Material tMat in tMats)
      Log.WL(String.Format("   - Material \"{0}\"",tMat.Name));

    // - Find and check template materials
    Log.WL("- Searching and checking thread template materials");
    IList<Material> vMats = ThreadMaterials.Find(doc,ThreadMaterials.VMATS).ToList();
    List<Material> vMatsValid = new List<Material>(vMats);
    foreach (Material vMat in vMats)
      try
      {
        ThreadMaterials.CheckTemplate(vMat);
      }
      catch (Exception e)
      {
        string vMatName = "<null> material";
        if (vMat!=null) vMatName = vMat.Name;
        Log.WL("Check failed on template material \""+vMatName+"\"");
        Log.WL(e.ToString());
        vMatsValid.Remove(vMat);
      }
      
    // - Final checks
    string stgeos = "  - Found {0} thread geometr{1}";
    string stMats = "  - Found {0} existing thread material{1}";
    string svMats = "  - Found {0} valid template material{1}";
    stgeos = Log.MakeCountOkMsg(tgeos.Count,tgeos.Count>0,stgeos,"ies","y");
    stMats = Log.MakeCountOkMsg(tMats.Count,true,stMats);
    svMats = Log.MakeCountOkMsg(vMatsValid.Count,vMatsValid.Count>0,svMats);
    if (vMatsValid.Count>0 && tgeos.Count>0)
      Log.WL("- PRE-CHECK OK");
    else
      Log.WL("- PRE-CHECK FAILED");
    Log.WL(stgeos);
    Log.WL(stMats);
    Log.WL(svMats);

    // Do welcome dialog
    Log.WL();
    Log.WL("Welcome dialog...");
    bool createVmats = false;
    bool deleteVmats = false;
    TaskDialog td = new TaskDialog("?");
    switch (ThreadMaterials.DoWelcomeDialog(td,doc,tgeos,tMats,vMats,vMatsValid))
    {
      case TaskDialogResult.CommandLink1:
        Log.WL("- Create Thread Materials operation selected by user");
        createVmats = true;
        try { deleteVmats = td.WasVerificationChecked(); } catch (Exception e) {}
        break;
      case TaskDialogResult.CommandLink2:
        Log.WL("- Delete Thread Materials operation selected by user");
        deleteVmats = true;
        break;
      default:
        Log.WL("- Cancelled by user");
        return;
    }

    // Main procedure
    string transactionName = (createVmats ? "Create" : "Delete") + " Thread Materials";
    Log.WL();
    Log.WL("Starting transaction \""+transactionName+"\"");
    List<string>   tMatsDeleted      = new List<string  >();
    List<Material> tMatsCreated      = new List<Material>();
    List<Material> tMatsDeleteFailed = new List<Material>();
    List<string>   tMatsCreateFailed = new List<string  >();
    using (Transaction t=new Transaction(doc,transactionName))
    {
      // - Start transaction
      t.Start();

      if (deleteVmats)
      {
        // - Delete materials
        Log.WL();
        Log.WL("Deleting thread materials");
        foreach (Material tMat in tMats)
        {
          string tMatName = tMat.Name;
          try
          {
            ThreadMaterials.Delete(tMat);
            tMatsDeleted.Add(tMatName);
          }
          catch (Exception e)
          {
            tMatsDeleteFailed.Add(tMat);
            Log.WL(String.Format("  Failed to delete thread material \"{0}\"",tMatName));
            Log.WL(e.ToString());
          }
        }
      }

      if (createVmats)
      {
        // - Create materials
        Log.WL();
        Log.WL("Creating thread materials");
        
        foreach (Material vMat in vMatsValid)
          foreach (ThreadGeometry tgeo in tgeos)
            try
            {
              Material tMat = ThreadMaterials.Create(vMat,tgeo);
              tMatsCreated.Add(tMat);
            }
            catch (Exception e)
            {
              string stMatName = String.Format(
                                   "thread material M{0} from template {1}",
                                   tgeo.D,
                                   vMat.Name
                                 );
              tMatsCreateFailed.Add(stMatName);
              Log.WL("  Failed to create "+stMatName);
              Log.WL(e.ToString());
            }
      }

      Log.WL();
      Log.WL("Committing transaction \""+transactionName+"\"");
      t.Commit();
    }
    
  	// Do wrap-up dialog
    Log.WL();
    Log.WL("Wrap-up");
//    DoWrapupDialog();
	}

	#endregion

}

namespace GIMBA
{
  [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
  [Autodesk.Revit.DB.Macros.AddInId("7EC07557-AB49-4F14-94AD-696D6670B347")]
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

    #region Revit documents
  
    /// <summary>Returns the database-level document currently active in Revit or <c>null</c>
    /// if there is no active document. The method is tailored to work in both, document and
    /// application macros.</summary>
    private Document GetActiveDocument()
    {
      Object thiz = (Object)this;
      if (thiz is UIDocument)
      {
        // We're a document macro
        UIDocument doc = thiz as UIDocument;
        return doc.Application.ActiveUIDocument.Document;
      }
      else if (thiz is UIApplication)
      {
        // We're an application macro
        UIApplication app = thiz as UIApplication;
        return app.ActiveUIDocument.Document;
      }
      else
        // Dunno dude...
        return null;
    }
  
    /// <summary>Checks whether a Revit document is suitable for working on.</summary>
    /// <exception cref="GIMBAexception">if <paramref name="document"/> is <c>null</c> or not suitable</exception>
    private void CheckDocument(Document doc)
    {
      // Check if there is an active document
      if (doc==null)
        throw new Exception
          (
            "There is no active document. See details for further instructions.\n\n"
            +"Open a suitable Revit family and make sure that it contains base materials named \"GIMBA - <Name>\"!"
          );
     }

    #endregion

    #region Revit macro handlers

    public void Thread_Materials()
    {
      Document doc = GetActiveDocument();
      Log.Begin("Thread_Materials",doc);
      try
      {
        CheckDocument(doc);
        ThreadMaterials.MacroMain(doc);
        Log.End();
      }
      catch (Exception e)
      {
        // TODO: Show error dialog!
        Log.End(e);
      }
    }

//    public void TEST()
//    {
//      // Out of service
//    }

    #endregion		
	}
}

// EOF