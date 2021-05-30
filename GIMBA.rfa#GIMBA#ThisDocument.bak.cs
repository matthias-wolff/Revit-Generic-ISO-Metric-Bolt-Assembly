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
// [4] .NET API Documentation, Method PropertyInfo.GetValue
//     https://docs.microsoft.com/de-de/dotnet/api/system.reflection.propertyinfo.getvalue?view=net-5.0
// [5] The Building Coder. Material, Physical and Thermal Assets. Nov. 2019. Retrieved May 2021.
//     https://thebuildingcoder.typepad.com/blog/2019/11/material-physical-and-thermal-assets.html

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

/// <summary>Basic static utility method</summary>
public class Utils
{

  #region C# Reflection
  
  /// <summary>Returns the path if the current C# source file</summary>
  /// <seealso href="https://stackoverflow.com/questions/47841441/how-do-i-get-the-path-to-the-current-c-sharp-source-code-file"
  ///   >Stackoverflow. How do I get the path to the current C# source code file?</seealso>
  public static string GetThisFilePath([CallerFilePath] string path = null)
  {
    return path;
  }

  #endregion

  #region String formatters
  
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
    string s = Utils.MakeCountMsg(count,format,pluralSuffix,singularSuffix);
    return s + (condition ? " --> ok" : " --> NOT OK");
  }

  /// <summary>Creates a hyperlink to the module help page</summary>
  /// <param name="doc">The Revit document this module is residing in</param>
  /// <param name="text">The link text</param>
  public static string MakeHelpLink(Document doc, string text)
  {
    string path = Path.Combine(Path.GetDirectoryName(doc.PathName),"GIMBA.html");
    return "<a href=\""+path+"\">"+text+"</a>";
  }

  #endregion

}

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

  /// <summary>Returns a hyperlink to the log file</summary>
  /// <param name="text">The link text</param>
  public static string MakeLogFileLink(string text)
  {
    return "<a href=\""+Log.GetLogFileName()+"\">"+text+"</a>";
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
      File.Copy(Utils.GetThisFilePath(),Path.Combine(path,"GIMBA.rfa#GIMBA#ThisDocument.bak.cs"),true);
      File.Copy(Utils.GetThisFilePath(),Path.Combine(path,"GIMBA.rfa#GIMBA#ThisDocument.bak-"+ts+".cs"),true);
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

  /// <summary>Finishes with the log</summary>
  public static void End()
  {
    // Finish the log file
    WL();
    WL("Pass complete");
  }

  #endregion

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
      new ThreadGeometry( 3,0.5 ); // NOTE: Newly created objects register themselves with dictionary
      new ThreadGeometry( 4,0.7 );
      new ThreadGeometry( 5,0.8 );
      new ThreadGeometry( 6,1   );
      new ThreadGeometry( 8,1.25);
      new ThreadGeometry(10,1.5 );
      new ThreadGeometry(12,1.75);
      new ThreadGeometry(14,2   );
      new ThreadGeometry(16,2   );
      new ThreadGeometry(18,2.5 );
      new ThreadGeometry(20,2.5 );
      new ThreadGeometry(22,2.5 );
      new ThreadGeometry(24,3   );
      new ThreadGeometry(27,3   );
      new ThreadGeometry(30,3.5 );
      new ThreadGeometry(33,3.5 );
      new ThreadGeometry(36,4   );
      new ThreadGeometry(39,4   );
      new ThreadGeometry(42,4.5 );
      new ThreadGeometry(45,4.5 );
      new ThreadGeometry(48,5   );
      new ThreadGeometry(52,5   );
      new ThreadGeometry(56,5.5 );
      new ThreadGeometry(64,6   );
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

  /// <summary>Counter index for thread geometries</summary>
  public static int CNT_TGEO = 0;

  /// <summary>Counter index for (valid) template materials</summary>
  public static int CNT_VMAT = 1;

  /// <summary>Counter index for invalid template materials</summary>
  public static int CNT_VMAT_INVAL = 2;

  /// <summary>Counter index for exisiting thread materials</summary>
  public static int CNT_TMAT = 3;

  /// <summary>Counter index for skipped thread materials</summary>
  public static int CNT_SKIP = 4;

  /// <summary>Counter index for deleted thread materials</summary>
  public static int CNT_DEL = 5;

  /// <summary>Counter index for overwritten thread materials</summary>
  public static int CNT_OVR = 6;

  /// <summary>Counter index for newly created thread materials</summary>
  public static int CNT_CRE = 7;

  /// <summary>Counter index for thread materials that could not be deleted</summary>
  public static int CNT_DELFAIL = 8;

  /// <summary>Counter index for thread materials that could not be overwritten</summary>
  public static int CNT_OVRFAIL = 9;

  /// <summary>Counter index for thread materials that could not be created</summary>
  public static int CNT_CREFAIL = 10;

  /// <summary>Counter array size</summary>
  public static int CNT_MAX = 11;
  
  #endregion

  #region User Interface
  
  /// <summary>Shows the welcome dialog</summary>
  /// <param name="td">The task dialog to use</param>
  /// <param name="doc">The Revit document to work on</param>
  /// <param name="counts">Operation and error counters</param>
  /// <returns>The task dialog result</returns>
  public static TaskDialogResult DoWelcomeDialog(
    TaskDialog td,
    Document   doc,
    int[]      counts
  )
  {
    // Big decision: If ok, macro operations can be performed
    bool ready = counts[CNT_VMAT]>0 && counts[CNT_TGEO]>0;

    // Prepare texts
    string docName = doc.Title + (doc.IsFamilyDocument ? ".rfa" : ".rvt");
    string stgeos  = "* Found {0} thread geometr{1} --> ";
    string stMats  = "* Found {0} existing thread material{1} --> ok";
    string svvMats = "* Found {0} valid template material{1} --> ";
    string sivMats = "* Found {0} invalid template material{1} --> ";
    stgeos  = Utils.MakeCountMsg(counts[CNT_TGEO],stgeos,"ies","y");
    stMats  = Utils.MakeCountMsg(counts[CNT_TMAT],stMats);
    svvMats = Utils.MakeCountMsg(counts[CNT_VMAT],svvMats);
    sivMats = Utils.MakeCountMsg(counts[CNT_VMAT_INVAL],sivMats);
    stgeos  += counts[CNT_TGEO]>0 ? "ok" : "NOT OK" ;
    svvMats += counts[CNT_VMAT]>0 ? "ok" : "NOT OK" ;
    sivMats += counts[CNT_VMAT_INVAL]>0 ? "ignore" : "ok" ;
    td.ExpandedContent = "Summary of pre-check results:\n"
                       + stgeos  + "\n"
                       + stMats  + "\n"
                       + svvMats + "\n"
                       + sivMats;
    if (!ready)
    {
      td.MainIcon        =  TaskDialogIcon.TaskDialogIconWarning;
      td.ExpandedContent += "\n\n"
                         +  "Issues marked with \"NOT OK\" obstruct operation.";
    }
   
    // Configure task dialog    
    td.MainInstruction   = "This macro performs batch operations on ISO metric "
                         + "screw thread materials. ";
    td.MainContent       = "Working document is: "+docName+"\n";
    td.FooterText        = Utils.MakeHelpLink(doc,"Help");
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
      sCmdSupp1           = Utils.MakeCountMsg(counts[CNT_TGEO]*counts[CNT_VMAT],sCmdSupp1);
      sCmdSupp2           = Utils.MakeCountMsg(counts[CNT_TMAT],sCmdSupp2);
      td.AddCommandLink(
        TaskDialogCommandLinkId.CommandLink1,
        "Create Thread Materials",
        sCmdSupp1 + (counts[CNT_TMAT]>0 ? " "+sCmdSupp2 : "")
      );
      if (counts[CNT_TMAT]>0)
      {
        td.VerificationText = "Overwrite existing thread material{1}";
        td.VerificationText = Utils.MakeCountMsg(counts[CNT_TMAT],td.VerificationText);
        sCmdSupp1           = "Will delete {0} existing thread material{1}. "
                            + "Template or other materials will not be deleted!";
        td.AddCommandLink(
          TaskDialogCommandLinkId.CommandLink2,
          "Delete Thread Materials",
          Utils.MakeCountMsg(counts[CNT_TMAT],sCmdSupp1)
        );
      }
    }
    else
    {
      // - Pre-checks failed: nothing can be done
      td.Title            = "Pre-Checks Failed";
      td.MainInstruction += "Pre-checks failed, however. No operation is possible on document.";
      td.FooterText      += " - "+Log.MakeLogFileLink("View log file");
      td.MainContent     += "See details and log file for further information.";
      td.CommonButtons    = TaskDialogCommonButtons.Close;
    }

    // Show task dialog
    return td.Show();
  }
  
  /// <summary>Shows the wrap-up dialog</summary>
  /// <param name="td">The task dialog to use</param>
  /// <param name="doc">The Revit document that has been worked on</param>
  /// <param name="cmd">Command executed on document</param>
  /// <param name="counts">Operation and error counters</param>
  /// <returns>The task dialog result</returns>
  public static TaskDialogResult DoWrapupDialog(
    TaskDialog       td,
    Document         doc,
    TaskDialogResult cmd,
    int[]            counts
  )
  {
    // Prepare texts
    bool   errors  = counts[CNT_DELFAIL]>0 || counts[CNT_OVRFAIL]>0 || counts[CNT_CREFAIL]>0;
    string docName = doc.Title + (doc.IsFamilyDocument ? ".rfa" : ".rvt");
    
    // Configure task dialog
    td.Title             = "Operation Completed" + (errors ? " with Errors" : "" );
    td.MainContent       = "See details and log file for further information.";
    td.ExpandedContent   = "Summary of operations performed:";
    // FIXME: This displays GIMBA.html in editor, not in browser
    td.FooterText        = Path.Combine(Path.GetDirectoryName(doc.PathName),"GIMBA.html");
    td.FooterText        = Utils.MakeHelpLink(doc,"Help")
                         + " - " + Log.MakeLogFileLink("View log file");
    td.CommonButtons     = TaskDialogCommonButtons.Close;
    td.AllowCancellation = true;
    td.TitleAutoPrefix   = false;
    if (errors)
      td.MainIcon = TaskDialogIcon.TaskDialogIconWarning;

    switch (cmd)
    {
      case TaskDialogResult.CommandLink1:
        if (!errors && counts[CNT_CRE]==0 && counts[CNT_OVR]==0)
          td.MainInstruction = "All thread materials were already present. "
                             + "Did not create new materials.";
        else
          td.MainInstruction = Utils.MakeCountMsg(counts[CNT_CRE]+counts[CNT_OVR],"Created {0} thread material{1}.");
        td.ExpandedContent += Utils.MakeCountMsg(counts[CNT_CRE ],"\n* Created {0} new material{1}");
        td.ExpandedContent += Utils.MakeCountMsg(counts[CNT_OVR ],"\n* Overwrote {0} material{1}");
        if (counts[CNT_SKIP]>0)
          td.ExpandedContent += Utils.MakeCountMsg(counts[CNT_SKIP],"\n* Skipped {0} existing material{1}");
        if (counts[CNT_CREFAIL]>0)
          td.ExpandedContent += Utils.MakeCountMsg(counts[CNT_CREFAIL],"\n* Failed to create {0} material{1}");
        if (counts[CNT_OVRFAIL]>0)
          td.ExpandedContent += Utils.MakeCountMsg(counts[CNT_OVRFAIL],"\n* Failed to overwrite {0} material{1}");
        break;

      case TaskDialogResult.CommandLink2:
        if (!errors && counts[CNT_DEL]==0)
          td.MainInstruction = "No thread materials were found. Did not delete any materials.";
        else
          td.MainInstruction = Utils.MakeCountMsg(counts[CNT_DEL],"Deleted {0} thread material{1}.");
        td.ExpandedContent += Utils.MakeCountMsg(counts[CNT_DEL],"\n* Deleted {0} material{1}");
        if (counts[CNT_DELFAIL]>0)
          td.ExpandedContent += Utils.MakeCountMsg(counts[CNT_DELFAIL],"\n* Failed to delete {0} material{1}");
        break;

      default:
        Debug.Assert(false); // SHould be unreachable 
        break;
    }
 
    // Show task dialog
    return td.Show();
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

  /// <summary>Returns a thread material name basing on a template material and a thread geometry</summary>
  /// <param name="vMat">The thread template material to base the name on</param>
  /// <param name="tgeo">The thread properties to base the name on</param>
  /// <returns>The name</returns>
  public static string MakeName(Material vMat, ThreadGeometry tgeo)
  {
    string category = VMATS.Match(vMat.Name).Groups[1]+"";   // E.g. "Steel galvanized"
    return            String.Format(sTMATS,category,tgeo.D); // E.g. "GIMBA - Steel galvanized - M12 thread"
  }
  
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
      throw new Exception("Material is <null>");

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
  private static Material Create(string name, Material vMat, ThreadGeometry tgeo)
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
    string   description = String.Format(                         // E.g. "Generic ISO metric bolt assembly, steel galvanized with M12 tread"
                             "Generic ISO metric bolt assembly: {0} with M{1} tread", 
                             category,tgeo.D
                           );

    // Duplicate and edit template material
    ElementId              tAppAssElemId = ElementId.InvalidElementId;
    ElementId              vAppAssElemId = vMat.AppearanceAssetId;
    AppearanceAssetElement vAppAssElem   = doc.GetElement(vAppAssElemId) as AppearanceAssetElement;
  
    // - Duplicate the material
    Material tMat = vMat.Duplicate(name);

    // - Duplicate the appearance asset and the asset(s) in it
    AppearanceAssetElement tAppAssElem = vAppAssElem.Duplicate(name);

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
    int[] counts = new int[ThreadMaterials.CNT_MAX];

    // Pre-Checks: Try to make sure that actual operation will not fail
    Log.WL();
    Log.WL("Pre-Checks");

    // - Find and check thread geometries
    Log.WL("- Searching thread geometries");
    IList<ThreadGeometry> tgeos = ThreadGeometry.getList();
    foreach (ThreadGeometry tgeo in tgeos)
      Log.WL("   - "+tgeo);
    counts[CNT_TGEO] = tgeos.Count;

    // - Find exisiting thread materials
    Log.WL("- Searching existing thread materials");
    IList<Material> tMats = ThreadMaterials.Find(doc,ThreadMaterials.TMATS).ToList();
    foreach (Material tMat in tMats)
      Log.WL(String.Format("   - Material \"{0}\"",tMat.Name));
    counts[CNT_TMAT] = tMats.Count;

    // - Find and check template materials
    Log.WL("- Searching and checking thread template materials");
    IList<Material> vMats = ThreadMaterials.Find(doc,ThreadMaterials.VMATS).ToList();
    foreach (Material vMat in vMats)
      try
      {
        ThreadMaterials.CheckTemplate(vMat);
        counts[CNT_VMAT]++;
      }
      catch (Exception e)
      {
        string vMatName = "<null> material";
        if (vMat!=null) vMatName = vMat.Name;
        Log.WL("Check failed on template material \""+vMatName+"\"");
        Log.WL(e.ToString());
        vMats.Remove(vMat);
        counts[CNT_VMAT_INVAL]++;
      }
      
    // - Final checks
    string stgeos = "  - Found {0} thread geometr{1}";
    string stMats = "  - Found {0} existing thread material{1}";
    string svMats = "  - Found {0} valid template material{1}";
    stgeos = Utils.MakeCountOkMsg(counts[CNT_TGEO],counts[CNT_TGEO]>0,stgeos,"ies","y");
    stMats = Utils.MakeCountOkMsg(counts[CNT_TMAT],true,stMats);
    svMats = Utils.MakeCountOkMsg(counts[CNT_VMAT],counts[CNT_VMAT]>0,svMats);
    if (counts[CNT_VMAT]>0 && counts[CNT_TGEO]>0)
      Log.WL("- PRE-CHECK OK");
    else
      Log.WL("- PRE-CHECK FAILED");
    Log.WL(stgeos);
    Log.WL(stMats);
    Log.WL(svMats);

    // Do welcome dialog
    Log.WL();
    Log.WL("Welcome dialog...");
    TaskDialog       td  = new TaskDialog("?");
    TaskDialogResult cmd = ThreadMaterials.DoWelcomeDialog(td,doc,counts);
    switch (cmd)
    {
      case TaskDialogResult.CommandLink1:
        Log.WL("- Create Thread Materials operation selected by user");
        break;
      case TaskDialogResult.CommandLink2:
        Log.WL("- Delete Thread Materials operation selected by user");
        break;
      default:
        Log.WL("- Cancelled by user");
        return;
    }

    // Main procedure
    counts = new int[ThreadMaterials.CNT_MAX];
    string transactionName = (cmd==TaskDialogResult.CommandLink1 ? "Create" : "Delete") 
                           + " Thread Materials";
    Log.WL();
    Log.WL("Starting transaction \""+transactionName+"\"");
    using (Transaction t=new Transaction(doc,transactionName))
    {
      // - Start transaction
      t.Start();

      switch (cmd)
      {
        case TaskDialogResult.CommandLink1:
          // - Create materials
          Log.WL();
          Log.WL("Creating thread materials");
          bool overwrite = false;
          try { overwrite = td.WasVerificationChecked(); } catch {}
          
          foreach (Material vMat in vMats)
            foreach (ThreadGeometry tgeo in tgeos)
            {
              string tMatName = ThreadMaterials.MakeName(vMat,tgeo);

              // - Find existing material of same name
              Material oMat = null;
              try { oMat = ThreadMaterials.Find(doc,tMatName).ToList()[0]; } catch {};

              // - Remove old material of same name
              if (overwrite && oMat!=null)
                try
                {
                  ThreadMaterials.Delete(oMat);
                  oMat = null;
                  counts[CNT_OVR]++;
                }
                catch (Exception e)
                {
                  Log.WL("  Failed to remove \""+tMatName+"\"");
                  Log.WL(e.ToString());
                  counts[CNT_OVRFAIL]++;
                }

              // - Create new material
              if (oMat==null)
                try
                {
                  Material tMat = ThreadMaterials.Create(tMatName,vMat,tgeo);
                  if (!overwrite) counts[CNT_CRE]++;
                }
                catch (Exception e)
                {
                  Log.WL("  Failed to create \""+tMatName+"\"");
                  Log.WL(e.ToString());
                  if (overwrite)
                    counts[CNT_OVRFAIL]++;
                  else
                    counts[CNT_CREFAIL]++;
                }
              else
              {
                Log.WL("- Skip existing material \""+oMat.Name+"\"");
                counts[CNT_SKIP]++;
              }
            }
          break;

        case TaskDialogResult.CommandLink2:
          // - Delete materials
          Log.WL();
          Log.WL("Deleting thread materials");
          foreach (Material tMat in tMats)
          {
            string tMatName = tMat.Name;
            try
            {
              ThreadMaterials.Delete(tMat);
              counts[CNT_DEL]++;
            }
            catch (Exception e)
            {
              Log.WL(String.Format("  Failed to delete thread material \"{0}\"",tMatName));
              Log.WL(e.ToString());
              counts[CNT_DELFAIL]++;
            }
          }
          break;
        
        default:
          // Should be unreachable!
          Debug.Assert(false);
          break;
      }

      Log.WL();
      Log.WL("Committing transaction \""+transactionName+"\"");
      t.Commit();
    }
    
  	// Do wrap-up dialog
    Log.WL();
    Log.WL("Wrap-up");
    td = new TaskDialog("?");
    ThreadMaterials.DoWrapupDialog(td,doc,cmd,counts);
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

    #region User Interface
    
    public void DoErrorWrapupDialog(Exception e)
    {
      TaskDialog td      = new TaskDialog("An Error Occurred");
      td.MainInstruction = "An unrecoverable error occured executing the macro.";
      td.MainContent     = "See details and log file for more infromation.";
      td.ExpandedContent = e.ToString();
      td.CommonButtons   = TaskDialogCommonButtons.Close;
      td.MainIcon        = TaskDialogIcon.TaskDialogIconError;
      td.Show();
    }
    
    #endregion
    
    #region Revit macro handlers

    public void Thread_Materials()
    {
      Document doc = GetActiveDocument();
      Log.Begin("Thread_Materials",doc); // TODO: Use this document!?
      try
      {
        CheckDocument(doc);
        ThreadMaterials.MacroMain(doc);
      }
      catch (Exception e)
      {
        Log.WL(e.ToString());
        DoErrorWrapupDialog(e);
      }
      Log.End();
    }

    #endregion		

	}
}

// EOF