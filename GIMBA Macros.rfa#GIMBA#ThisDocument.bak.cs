// GIMBA - Generic ISO Metric Bolt Assembly
// Service macros
//
// Author    : Matthias Wolff
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
// [6] Wikipedia. Metrisches ISO-Gewinde (German). Retrieved June 2021.
//     https://de.wikipedia.org/wiki/Metrisches_ISO-Gewinde

using System;                           // For Object, Exception, etc.
using System.IO;                        // For file input/output
using System.Threading;                 // For globally setting invariant culture
using System.Globalization;             // For globally setting invariant culture
using System.Diagnostics;               // For Process.Start
using System.Text.RegularExpressions;   // For regular expressions
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

  #region Static fields

  /// <summary>Author of package</summary>
  public static string Author = "Matthias Wolff";

  /// <summary>URL of repository</summary>
  public static string RepoUrl = "https://github.com/matthias-wolff/Revit-Generic-ISO-Metric-Bolt-Assembly";

  /// <summary>URL of help page</summary>
  public static string HelpUrl = "https://matthias-wolff.github.io/Revit-Generic-ISO-Metric-Bolt-Assembly/GIMBA.html";

  #endregion

  #region Paths and Files
  
  /// <summary>Returns the path if the current C# source file</summary>
  /// <seealso href="https://stackoverflow.com/questions/47841441/how-do-i-get-the-path-to-the-current-c-sharp-source-code-file"
  ///   >Stackoverflow. How do I get the path to the current C# source code file?</seealso>
  public static string GetThisFilePath([CallerFilePath] string path = null)
  {
    return path;
  }
  
  /// <summary>Writes text data to a file</summary>
  /// <param name="data">The data string to write</param>
  /// <param name="path">The fully qualified pathname of the file to write</param>
  /// <param name="overwrite">if <c>true</c>, an exisiting file be overwritten</param>
  /// <returns>1 if a new file was created, -1 if an existing file was overwritten, and 0 if an existing file was skipped</returns>
  public static int WriteTextFile(string data, string path, bool overwrite)
  {
    bool fex = File.Exists(path);
    if (fex && !overwrite) return 0;
    FileStream   fs = new FileStream(path,FileMode.Create,FileAccess.Write);  
    StreamWriter fw = new StreamWriter(fs);  
    fw.Write(data);  
    fw.Close();
    return fex ? -1 : 1;
  }

  /// <summary>Creates a backup of this source file. Revit cost me 10 hours of working 
  /// yesterday. That's why I implemented this method... :((</summary>
  /// <param name="doc">The document this module is residing in. The backup file is created
  /// in the document's folder.</param>
  public static void BackupThisSource(Document doc)
  {
    string path = Path.GetDirectoryName(doc.PathName);
    string time = DateTime.Now.ToString("yyyyMMdd-HHmmss");
    string fnam = "GIMBA Macros.rfa#GIMBA#ThisDocument.bak";
    File.Copy(Utils.GetThisFilePath(),Path.Combine(path,fnam+".cs"),true);
    File.Copy(Utils.GetThisFilePath(),Path.Combine(path,fnam+"-"+time+".cs"),true);
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
    return "<a href=\""+Utils.HelpUrl+"\">"+text+"</a>";
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
  
  /// <summary>Creates a new log</summary>
  /// <param name="name">The macro name</param>
  /// <param name="doc">The Revit document containing this macro</param>
  public static void Begin(string name, Document doc)
  {
    // Try to make log file path in Revit document's folder, use temp folder on errors
    try
    {
      string path = Path.GetDirectoryName(doc.PathName);
      Log.LogFileName = Path.Combine(path,"GIMBA.log");
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

/// <summary>ISO metric bolt geometry</summary>
/// <seealso href="http://www.iso-gewinde.at">Thread Calculator (German)</seealso>
public class BoltGeometry
{
  private static readonly string CsvTlmm = "##LENGTH##MILLIMETERS";
  private static readonly string CsvToth = "##OTHER##";

  private static readonly string[] materialNames = new string[]{
    // TODO: Automate
    // "Steel", 
    "Steel galvanized", 
    // "Steel galvanized chromated",
    // "Stainless steel"
  };

  #region Static Fields

  /// <summary>ISO metric threads dictionary</summary>
  private static Dictionary<string,BoltGeometry> Geometries;

  #endregion

  #region Static API

  /// <summary>Returns the ISO metric threads dictionary.</summary>
  /// <seealso href="http://www.iso-gewinde.at">Thread Calculator (German)</seealso>
  protected internal static IList<BoltGeometry> GetList()
  {
    // Create bolt gemoetries dictionary of not yet exisiting
    if (BoltGeometry.Geometries==null)
    {
      BoltGeometry.Geometries = new Dictionary<string,BoltGeometry>();
      // ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
      // Parameter:   |  D|  P  |  s  |  k  |  a   | du1 | du2|  u  | dh1 | dh2 | dh3 | dgl|     cls                                           
      new BoltGeometry(  3, 0.5 ,  5.5,  2  ,  1.5 ,  3.2,   7, 0.5 ,  3.2,  3.4,  3.6,  10, new double[]{3,4,5,6,8,10,12,16,18,20,22,25,30,35,40,50,60}); 
      new BoltGeometry(  4, 0.7 ,  7  ,  2.8,  2.1 ,  4.3,   9, 0.8 ,  4.3,  4.5,  4.8,  20, new double[]{4,6,8,10,12,14,16,18,20,22,25,30,35,40,45,50,55,60,65,70,75,80});
      new BoltGeometry(  5, 0.8 ,  8  ,  3.5,  2.4 ,  5.3,  10, 1   ,  5.3,  5.5,  5.8,  50, new double[]{6,8,10,12,14,16,18,20,22,25,30,35,40,45,50,55,60,65,70,80,90,100});
      new BoltGeometry(  6, 1   , 10  ,  4  ,  3   ,  6.4,  12, 1.6 ,  6.4,  6.6,  7  ,  50, new double[]{6,8,10,12,14,16,18,20,22,25,28,30,35,40,45,50,55,60,65,70,75,80,85,90,100,110,120,130,140,150});
      new BoltGeometry(  8, 1.25, 13  ,  5.5,  3.75,  8.4,  16, 1.6 ,  8.4,  9  , 10  ,  50, new double[]{8,10,12,14,16,18,20,22,25,30,35,40,45,50,55,60,65,70,75,80,85,90,95,100,110,120,130,140,150,160,170,180,190,200});
      new BoltGeometry( 10, 1.5 , 17  ,  6.4,  4.5 , 10.5,  20, 2   , 10.5, 11  , 12  , 100, new double[]{10,12,16,18,20,22,25,28,30,35,40,45,50,55,60,65,70,75,80,85,90,100,110,120,130,140,150,160,170,180,190,200,220,240,280,300});
      new BoltGeometry( 12, 1.75, 19  ,  8  ,  5.5 , 13  ,  24, 2.5 , 13  , 13.5, 14.5, 100, new double[]{10,12,16,18,20,22,25,28,30,35,40,45,50,55,60,65,70,75,80,85,90,100,110,120,130,140,150,160,170,180,190,200,220,240,300});
      new BoltGeometry( 14, 2   , 22  ,  9  ,  6   , 15  ,  28, 2.5 , 15  , 15.5, 16.5, 100, new double[]{16,20,25,30,35,40,45,50,55,60,65,70,75,80,90,100,110,120,130,140,150,160,170,180,200,220});
      new BoltGeometry( 16, 2   , 24  , 10  ,  6   , 17  ,  30, 3   , 17  , 17.5, 18.5, 150, new double[]{12,16,20,25,30,35,40,45,50,55,60,65,70,75,80,85,90,95,100,110,120,130,140,150,160,170,180,190,200,210,220,230,240,250,260,280,300,320,340,400,500});
      new BoltGeometry( 18, 2.5 , 27  , 11.5,  7.5 , 19  ,  34, 3   , 19  , 20  , 21  , 150, new double[]{20,25,30,35,40,45,50,55,60,65,70,75,80,85,90,100,110,120,130,140,150,160,170,180,190,200});
      new BoltGeometry( 20, 2.5 , 30  , 12.5,  7.5 , 21  ,  37, 3   , 21  , 22  , 24  , 150, new double[]{20,25,30,35,40,45,50,55,60,65,70,75,80,85,90,100,110,120,130,140,150,160,170,180,190,200,210,220,230,240,250,260,280,300,360});
      new BoltGeometry( 22, 2.5 , 32  , 14  ,  7.5 , 23  ,  39, 3   , 23  , 24  , 26  , 150, new double[]{30,35,40,45,50,55,60,65,70,75,80,90,100,110,120,130,140,150,160,170,180,190,200});
      new BoltGeometry( 24, 3   , 36  , 15  ,  9   , 25  ,  44, 4   , 25  , 26  , 28  , 150, new double[]{25,30,35,40,45,50,55,60,65,70,75,80,85,90,100,110,120,130,140,150,160,170,180,190,200,210,220,230,240,250,260,280,300,320,500});
      new BoltGeometry( 27, 3   , 41  , 17  ,  9   , 28  ,  50, 4   , 28  , 30  , 32  , 150, new double[]{30,40,45,50,55,60,65,70,75,80,85,90,100,110,120,130,140,150,160,170,180,190,200,300});
      new BoltGeometry( 30, 3.5 , 46  , 19  , 10.5 , 31  ,  56, 4   , 31  , 33  , 35  , 150, new double[]{35,40,45,50,55,60,65,70,75,80,85,90,100,110,120,130,140,150,160,170,180,190,200,210,220,230,240,250,260,280,300,320,340,360,380,400,500,600});
      new BoltGeometry( 33, 3.5 , 50  , 21  , 10.5 , 34  ,  60, 5   , 34  , 36  , 39  , 150, new double[]{40,50,60,65,70,75,80,90,100,110,120,130,140,150,160,170,180,190,200,300});
      new BoltGeometry( 36, 4   , 55  , 23  , 12   , 37  ,  66, 5   , 37  , 39  , 42  , 150, new double[]{40,45,50,55,60,65,70,75,80,85,90,100,110,120,130,140,150,160,170,180,190,200,220,260,280,300,320,340,400,600});
      new BoltGeometry( 39, 4   , 60  , 25  , 12   , 40  ,  72, 6   , 40  , 43  , 45  , 150, new double[]{80,90,100,110,120,130,140,150,160,180,190,200});
      new BoltGeometry( 42, 4.5 , 65  , 26  , 13.5 , 43  ,  78, 7   , 43  , 46  , 48  , 150, new double[]{50,55,60,70,75,80,85,90,100,110,120,130,140,150,160,170,180,190,200,220,250,260,300,360,400});
      new BoltGeometry( 45, 4.5 , 70  , 28  , 13.5 , 46  ,  85, 7   , 47  , 49  , 52  , 150, new double[]{90,100,110,120,130,140,150});
      new BoltGeometry( 48, 5   , 75  , 30  , 15   , 50  ,  92, 8   , 50  , 52  , 56  , 150, new double[]{60,70,80,90,100,110,120,130,140,150,160,170,180,190,200,420});
      new BoltGeometry( 52, 5   , 80  , 33  , 15   , 54  ,  98, 8   , 54  , 57  , 61  , 180, new double[]{150,200});
      new BoltGeometry( 56, 5.5 , 85  , 35  , 16.5 , 58  , 105, 9   , 58  , 62  , 66  , 180, new double[]{130,140,150,160,170,190,200,220,240,250,260,280,300,380});
      new BoltGeometry( 64, 6   , 95  , 40  , 18   , 66  , 120, 9   , 66  , 70  , 74  , 250, new double[]{300});
      // ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
      // NOTE: Newly created objects register themselves with dictionary
    }
    
    // Return dictionary values, i.e., the geometries, as a list
    return BoltGeometry.Geometries.Values.ToList();
  }

  /// <summary>Writes the bolt type catalog file</summary>
  /// <param name="path">Fully qualified path of type catalog file to write</param>
  /// <param name="overwrite">If <c>true</c>, a previously existing type catalog file will be overwritten</param>
  /// <returns>1 if a new file was created, -1 if an existing file was overwritten, and 0 if an existing file was skipped</returns>
  public static int WriteBoltTypeCatalog(string path, bool overwrite)
  {
    Log.WL();
    Log.WL(path);
    
    // Make CSV Header
    string data 
      = ",Nominal Diameter"+CsvTlmm
      + ",Length"          +CsvTlmm
      + ",Shank"           +CsvToth
      + ",Material"        +CsvToth
      + ",Thread Material" +CsvToth
      + "\n";

    // Create type definitions
    foreach (BoltGeometry bg in BoltGeometry.GetList())
      foreach (double l in bg.cls)
        foreach (bool s in new bool[]{false,true})
          if (!s || l>50)
            foreach (string m in materialNames)
              data 
                += String.Format("M{0} x {1}{2} {3}",bg.D,l,s?" w/shank":"",m) // Type name
                +  String.Format(",{0}",bg.D  )                                // Nominal diameter
                +  String.Format(",{0}",l     )                                // Length
                +  String.Format(",{0}",s?1:0 )                                // Shank
                +  String.Format(",GIMBA - {0}",m)                             // Plain material
                +  String.Format(",GIMBA - {0} - M{1} thread",m,bg.D)          // Thread material
                +  "\n";

    // Write type catalog file
    int result = Utils.WriteTextFile(data,path,overwrite);
    switch(result)
    {
      case -1: Log.WL("- overwritten"); break;
      case  1: Log.WL("- created"    ); break;
      default: Log.WL("- skipped"    ); break;
    }
    return result;
  }

  /// <summary>Writes the assembly type catalog file</summary>
  /// <param name="path">Fully qualified path of type catalog file to write</param>
  /// <param name="overwrite">If <c>true</c>, a previously existing type catalog file will be overwritten</param>
  /// <returns>1 if a new file was created, -1 if an existing file was overwritten, and 0 if an existing file was skipped</returns>
  public static int WriteAssemblyTypeCatalog(string path, bool overwrite)
  {
    Log.WL();
    Log.WL(path);
    
    // Make CSV Header
    string data 
      = ",Nominal Diameter"+CsvTlmm
      + ",Grip Length"     +CsvTlmm
      + ",Shank"           +CsvToth
      + ",Material"        +CsvToth
      + ",Thread Material" +CsvToth
      + "\n";

    // Create type definitions
    foreach (BoltGeometry bg in BoltGeometry.GetList())
      foreach (bool s in new bool[]{false,true})
        if (!s || bg.dgl>50)
          foreach (string m in materialNames)
            data 
              += String.Format("M{0}{1} {2}",bg.D,s?" w/shank":"",m) // Type name
              +  String.Format(",{0}",bg.D  )                        // Nominal diameter
              +  String.Format(",{0}",bg.dgl)                        // Grip length
              +  String.Format(",{0}",s?1:0 )                        // Shank
              +  String.Format(",GIMBA - {0}",m)                     // Plain material
              +  String.Format(",GIMBA - {0} - M{1} thread",m,bg.D)  // Thread material
              +  "\n";

    // Write type catalog file
    int result = Utils.WriteTextFile(data,path,overwrite);
    switch(result)
    {
      case -1: Log.WL("- overwritten"); break;
      case  1: Log.WL("- created"    ); break;
      default: Log.WL("- skipped"    ); break;
    }
    return result;
  }

  /// <summary>Writes the grip-to-length lookup table file</summary>
  /// <param name="path">Fully qualified path of lookup table file to write</param>
  /// <param name="overwrite">If <c>true</c>, a previously existing type lookup table file will be overwritten</param>
  /// <returns>1 if a new file was created, -1 if an existing file was overwritten, and 0 if an existing file was skipped</returns>
  public static int WriteG2LTable(string path, bool overwrite)
  {
    Log.WL();
    Log.WL(path);
    
    // Make CSV Header
    string data 
      = ",D" +CsvTlmm
      + ",LG"+CsvTlmm
      + ",l" +CsvTlmm
      + "\n";

    // Create lookup table lines
    foreach (BoltGeometry bg in BoltGeometry.GetList())
      for (int LG=0; LG<=600; LG++)
      {
        int m;
        if      (LG< 23) m=2;
        else if (LG<100) m=5;
        else             m=10;
        if (LG%m==0)
        {
          double l    = 0;
          double lmin = LG + 2*bg.k + 2*bg.u;
          foreach (double L in bg.cls)
            if (L>lmin)
            {
              l = L;
              break;
            }
          if (l>0)
            data 
              += String.Format("M{0} x ]{1}[",bg.D,LG) // Line name (irrelevant)
              +  String.Format(",{0}"        ,bg.D   ) // Nominal diameter
              +  String.Format(",{0}"        ,LG     ) // Grip length
              +  String.Format(",{0}"        ,l      ) // Bolt length
              +  "\n";
        }
      }

    // Write lookup table file
    int result = Utils.WriteTextFile(data,path,overwrite);
    switch(result)
    {
      case -1: Log.WL("- overwritten"); break;
      case  1: Log.WL("- created"    ); break;
      default: Log.WL("- skipped"    ); break;
    }
    return result;
  }

  /// <summary>Writes the geometry parameters lookup table file</summary>
  /// <param name="path">Fully qualified path of lookup table file to write</param>
  /// <param name="overwrite">If <c>true</c>, a previously existing type lookup table file will be overwritten</param>
  /// <returns>1 if a new file was created, -1 if an existing file was overwritten, and 0 if an existing file was skipped</returns>
  public static int WriteMGeoTable(string path, bool overwrite)
  {
    Log.WL();
    Log.WL(path);
    
    // Make CSV Header
    string data 
      = ",D"   +CsvTlmm
      + ",P"   +CsvTlmm
      + ",H"   +CsvTlmm
      + ",d2"  +CsvTlmm
      + ",s"   +CsvTlmm
      + ",k"   +CsvTlmm
      + ",a"   +CsvTlmm
      + ",b2"  +CsvTlmm
      + ",b3"  +CsvTlmm
      + ",b4"  +CsvTlmm
      + ",du1" +CsvTlmm
      + ",du2" +CsvTlmm
      + ",u"   +CsvTlmm
      + ",dh1" +CsvTlmm
      + ",dh2" +CsvTlmm
      + ",dh3" +CsvTlmm
      + "\n";

    // Create lookup table lines
    foreach (BoltGeometry bg in BoltGeometry.GetList())
      data 
        += String.Format("M{0}"   ,bg.D  ) // Line name (irrelevant)
        +  String.Format(",{0}"   ,bg.D  ) // Nominal diameter
        +  String.Format(",{0}"   ,bg.P  ) // Thread pitch
        +  String.Format(",{0:F2}",bg.H  ) // Thread height
        +  String.Format(",{0:F2}",bg.d2 ) // Effective pitch diameter
        +  String.Format(",{0}"   ,bg.s  ) // Wrench size
        +  String.Format(",{0}"   ,bg.k  ) // Height of bolt cap and nut
        +  String.Format(",{0}"   ,bg.a  ) // Minimum distance between cap and thread
        +  String.Format(",{0}"   ,bg.b2 ) // Minimum thread length for bolt lengths < 125 mm
        +  String.Format(",{0}"   ,bg.b3 ) // Minimum thread length for bolt lengths < 200 mm
        +  String.Format(",{0}"   ,bg.b4 ) // Minimum thread length for bolt lengths >= 200 mm
        +  String.Format(",{0}"   ,bg.du1) // Diameter of washer clearance hole
        +  String.Format(",{0}"   ,bg.du2) // Washer diameter
        +  String.Format(",{0}"   ,bg.u  ) // Washer thickness
        +  String.Format(",{0}"   ,bg.dh1) // Fine clearance hole diameter
        +  String.Format(",{0}"   ,bg.dh2) // Medium clearance hole diameter
        +  String.Format(",{0}"   ,bg.dh3) // Coarse clearance hole diameter
        +  "\n";

    // Write lookup table file
    int result = Utils.WriteTextFile(data,path,overwrite);
    switch(result)
    {
      case -1: Log.WL("- overwritten"); break;
      case  1: Log.WL("- created"    ); break;
      default: Log.WL("- skipped"    ); break;
    }
    return result;
  }

  /// <summary>Writes the supported nominal diameters lookup table file</summary>
  /// <param name="path">Fully qualified path of lookup table file to write</param>
  /// <param name="overwrite">If <c>true</c>, a previously existing type lookup table file will be overwritten</param>
  /// <returns>1 if a new file was created, -1 if an existing file was overwritten, and 0 if an existing file was skipped</returns>
  public static int WriteD2DTable(string path, bool overwrite)
  {
    Log.WL();
    Log.WL(path);
    
    // Make CSV Header
    string data 
      = ",ND" +CsvTlmm
      + ",D"  +CsvTlmm
      + "\n";

    // Create lookup table lines
    IList<BoltGeometry> bgl = BoltGeometry.GetList();
    int i  = 0;        // Index in bolt geometries list
    int D1 = 3;        // Lower permissible nominal diameter
    int D2 = bgl[i].D; // Upper permissible nominal diameter
    for (int D=3; D<=64; D++)
    {
      data += String.Format("D={0},{1},{2}\n",D,D,D-D1<D2-D?D1:D2);
      if (D==D2)
      {
        D1=D2;
        if (i<bgl.Count-1)
          D2 = bgl[++i].D;
      }
    }

    // Write lookup table file
    int result = Utils.WriteTextFile(data,path,overwrite);
    switch(result)
    {
      case -1: Log.WL("- overwritten"); break;
      case  1: Log.WL("- created"    ); break;
      default: Log.WL("- skipped"    ); break;
    }
    return result;
  }

  /// <summary>Dumps the geometry parameters to an HTML table</summary>
  /// <param name="path">Fully qualified path of HTML file to write</param>
  /// <param name="overwrite">If <c>true</c>, a previously existing type lookup table file will be overwritten</param>
  /// <returns>1 if a new file was created, -1 if an existing file was overwritten, and 0 if an existing file was skipped</returns>
  public static int WriteMGeoHtml(string path, bool overwrite)
  {
    Log.WL();
    Log.WL(path);

    // Make HTML table header
    string data
      = "<table>\n"
      + "  <tr>\n"
      +  "    <th>Name</th>\n"           // Name
      +  "    <th>D</td>\n"              // Nominal diameter
      +  "    <th>P</td>\n"              // Thread pitch
      +  "    <th>H</td>\n"              // Thread height
      +  "    <th>d<sub>2</sub></td>\n"  // Effective pitch diameter
      +  "    <th>s</td>\n"              // Wrench size
      +  "    <th>k</td>\n"              // Height of bolt cap and nut
      +  "    <th>a</td>\n"              // Minimum distance between cap and thread
      +  "    <th>b<sub>2</sub></td>\n"  // Minimum thread length for bolt lengths < 125 mm
      +  "    <th>b<sub>3</sub></td>\n"  // Minimum thread length for bolt lengths < 200 mm
      +  "    <th>b<sub>4</sub></td>\n"  // Minimum thread length for bolt lengths >= 200 mm
      +  "    <th>d<sub>u1</sub></td>\n" // Diameter of washer clearance hole
      +  "    <th>d<sub>u2</sub></td>\n" // Washer diameter
      +  "    <th>u</td>\n"              // Washer thickness
      +  "    <th>d<sub>h1</sub></td>\n" // Fine clearance hole diameter
      +  "    <th>d<sub>h2</sub></td>\n" // Medium clearance hole diameter
      +  "    <th>d<sub>h3</sub></td>\n" // Coarse clearance hole diameter
      + "  </tr>\n";

    // TODO: Create HTML table lines
    foreach (BoltGeometry bg in BoltGeometry.GetList())
      data 
        += "  <tr>\n"
        +  String.Format("    <td>M{0}</td>\n"  ,bg.D  ) // Name
        +  String.Format("    <td>{0}</td>\n"   ,bg.D  ) // Nominal diameter
        +  String.Format("    <td>{0}</td>\n"   ,bg.P  ) // Thread pitch
        +  String.Format("    <td>{0:F2}</td>\n",bg.H  ) // Thread height
        +  String.Format("    <td>{0:F2}</td>\n",bg.d2 ) // Effective pitch diameter
        +  String.Format("    <td>{0}</td>\n"   ,bg.s  ) // Wrench size
        +  String.Format("    <td>{0}</td>\n"   ,bg.k  ) // Height of bolt cap and nut
        +  String.Format("    <td>{0}</td>\n"   ,bg.a  ) // Minimum distance between cap and thread
        +  String.Format("    <td>{0}</td>\n"   ,bg.b2 ) // Minimum thread length for bolt lengths < 125 mm
        +  String.Format("    <td>{0}</td>\n"   ,bg.b3 ) // Minimum thread length for bolt lengths < 200 mm
        +  String.Format("    <td>{0}</td>\n"   ,bg.b4 ) // Minimum thread length for bolt lengths >= 200 mm
        +  String.Format("    <td>{0}</td>\n"   ,bg.du1) // Diameter of washer clearance hole
        +  String.Format("    <td>{0}</td>\n"   ,bg.du2) // Washer diameter
        +  String.Format("    <td>{0}</td>\n"   ,bg.u  ) // Washer thickness
        +  String.Format("    <td>{0}</td>\n"   ,bg.dh1) // Fine clearance hole diameter
        +  String.Format("    <td>{0}</td>\n"   ,bg.dh2) // Medium clearance hole diameter
        +  String.Format("    <td>{0}</td>\n"   ,bg.dh3) // Coarse clearance hole diameter
        +  "  </tr>\n";

    // Make HTML table footer
    data += "</table>\n";

    // Write type catalog file
    int result = Utils.WriteTextFile(data,path,overwrite);
    switch(result)
    {
      case -1: Log.WL("- overwritten"); break;
      case  1: Log.WL("- created"    ); break;
      default: Log.WL("- skipped"    ); break;
    }
    return result;
  }
  
  #endregion

  #region Main Function of Catalogs And Tables Macro
  
  /// <summary>Manages type catalog and lookup table files for bolts and bolt assemlies.</summary>
  /// <param name="document">The document this macro resides in</param>
  public static void CatalogsAndTables(Document document)
  {
    TaskDialog td = new TaskDialog("?");

    // Prepare paths
    string docDir = Path.GetDirectoryName(document.PathName);
    string btcPth = Path.Combine(docDir,"Generic ISO Metric Bolt.txt"         ); // Bolt type catalog file
    string atcPth = Path.Combine(docDir,"Generic ISO Metric Bolt Assembly.txt"); // Assembly type catalog file
    string g2lPth = Path.Combine(docDir,"GIMBA G2L.csv"                       ); // Grip to bolt length lookup table file
    string mgePth = Path.Combine(docDir,"GIMBA MGeo.csv"                      ); // Geometry parameters lookup table file
    string d2dPth = Path.Combine(docDir,"GIMBA D2D.csv"                       ); // Supported nominal diameters lookup table file
    string mghPth = Path.Combine(docDir,"GIMBA MGeo.html"                     ); // Geometry parameters HTML table file
    
    // Pre-checks: Find existing type catalog and lookup table files
    int tcn  = 2; // Number of type catalog files
    int tbn  = 3; // Number of lookup table files
    int thn  = 1; // Number of HTML files
    int tcex = 0; // Number of existing type catalog files
    int tbex = 0; // Number of existing lookup table files
    int thex = 0; // Number of existing HTML files
    td.ExpandedContent = "Status of type catalog and lookup table files:";
    foreach (string path in new string[]{btcPth,atcPth,g2lPth,mgePth,d2dPth})
    {
      td.ExpandedContent+="\n* "+Path.GetFileName(path);
      if (File.Exists(path))
      {
        td.ExpandedContent+=" (exists)";
        if (".txt".Equals(Path.GetExtension(path),StringComparison.OrdinalIgnoreCase))
          tcex++;
        else if (".html".Equals(Path.GetExtension(path),StringComparison.OrdinalIgnoreCase))
          thex++;
        else
          tbex++;
      }
      else
        td.ExpandedContent+=" (does not exist)";
    }

    // Do welcome dialog
    Log.WL();
    Log.WL("Welcome dialog...");
    td.Title             = "Good to Go...";
    td.MainInstruction   = "This macro creates type catalog and lookup table "
                         + "for the generic ISO metric bolt and bolt assembly "
                         + "families. Please select an option!";
    td.MainContent       = "Working directory is: "+docDir+"\n"
                         + "See details for results of pre-checks.";
    td.CommonButtons     = TaskDialogCommonButtons.Cancel;
    td.FooterText        = Utils.MakeHelpLink(document,"Help");
    td.TitleAutoPrefix   = false;
    td.AllowCancellation = true;
    string sCmdSupp1     = "Will create {0} file{1}. ";
    string sCmdSupp2     = "Please choose below whether {0} existing file{1} "
                         + "shall be overwritten.";
    td.AddCommandLink(
      TaskDialogCommandLinkId.CommandLink1,
      "Create type catalog files",
      Utils.MakeCountMsg(tcn,sCmdSupp1) + (tcex>0 ? Utils.MakeCountMsg(tcex,sCmdSupp2) : "")
    );
    td.AddCommandLink(
      TaskDialogCommandLinkId.CommandLink2,
      "Create lookup table files",
      Utils.MakeCountMsg(tbn,sCmdSupp1) + (tbex>0 ? Utils.MakeCountMsg(tbex,sCmdSupp2) : "")
    );
    td.AddCommandLink(
      TaskDialogCommandLinkId.CommandLink3,
      "Dump geometry parameters to an HTML table",
      Utils.MakeCountMsg(thn,sCmdSupp1) + (thex>0 ? Utils.MakeCountMsg(tbex,sCmdSupp2) : "")
    );
    if (tcex+tbex>0)
      td.VerificationText = Utils.MakeCountMsg(tcex+tbex+thex,"Overwrite existing file{1}");
    TaskDialogResult tdr = td.Show();
    
    // Main operation
    bool doOvr = (tcex+tbex)==0 || td.WasVerificationChecked();
    int cre = 0; // Number of files newly created
    int ovr = 0; // Number of files overwritten
    int skp = 0; // Number of files skipped
    int err = 0; // Number of errors (max. 1 per file)
    Tuple<Func<string,bool,int>,string>[] schedule;
    string details = "";
    switch (tdr)
    {
      case TaskDialogResult.CommandLink1:
        Log.WL("- Create type catalog files operation selected by user");
        schedule = new Tuple<Func<string,bool,int>,string>[]{
          Tuple.Create<Func<string,bool,int>,string>(WriteBoltTypeCatalog    ,btcPth),
          Tuple.Create<Func<string,bool,int>,string>(WriteAssemblyTypeCatalog,atcPth)
        };
        break;
      case TaskDialogResult.CommandLink2:
        Log.WL("- Create lookup table files operation selected by user");
        schedule = new Tuple<Func<string,bool,int>,string>[]{
          Tuple.Create<Func<string,bool,int>,string>(WriteG2LTable ,g2lPth),
          Tuple.Create<Func<string,bool,int>,string>(WriteMGeoTable,mgePth),
          Tuple.Create<Func<string,bool,int>,string>(WriteD2DTable ,d2dPth)
        };
        break;
      case TaskDialogResult.CommandLink3:
        Log.WL("- Create lookup table files operation selected by user");
        schedule = new Tuple<Func<string,bool,int>,string>[]{
          Tuple.Create<Func<string,bool,int>,string>(WriteMGeoHtml,mghPth)
        };
        break;
      default:
        Log.WL("- Cancelled by user");
        return;
    }
    foreach(Tuple<Func<string,bool,int>,string> item in schedule)
      try
      {
        Func<string,bool,int> fnc  = item.Item1;
        string                path = item.Item2;
        details += "\n* "+Path.GetFileName(path)+" ";
        switch (fnc(path,doOvr))
        {
          case -1:
            ovr++;
            details +="(overwritten)";
            break;
          case 1:
            cre++;
            details +="(created)";
            break;
          default:
            skp++;
            details +="(skipped)";
            break;
        }
      }
      catch (Exception e)
      {
        err++;
        td.ExpandedContent +="(ERROR)";
        Log.WL(e.ToString());
      }
    
    // TODO: Do Wrap-up dialog
    Log.WL();
    Log.WL("Wrap up");
    td = new TaskDialog("Operation Completed");
    if (err>0)
    {
      td.Title += " with Errors";
      td.MainIcon = TaskDialogIcon.TaskDialogIconWarning;
    }
    td.MainContent       = "See details and log file for further information.";
    td.ExpandedContent   = "Summary of file operations performed:" + details;
    // FIXME: This displays GIMBA.html in editor, not in browser
    td.FooterText        = Path.Combine(Path.GetDirectoryName(document.PathName),"GIMBA.html");
    td.FooterText        = Utils.MakeHelpLink(document,"Help")
                         + " - " + Log.MakeLogFileLink("View log file");
    td.CommonButtons     = TaskDialogCommonButtons.Close;
    td.AllowCancellation = true;
    td.TitleAutoPrefix   = false;
    if (cre+ovr+err==0 && skp>0)
    {
      td.MainInstruction = "All output files were present. Nothing to be done.";
      td.MainContent     = "If you want to recreate the files, re-run the macro and check "
                         + "\"Overwrite exisiting files\". " + td.MainContent;
    }
    else
    {
      td.MainInstruction = "";
      if (cre>0) td.MainInstruction += Utils.MakeCountMsg(cre,"{0} file{1} created. "    );
      if (ovr>0) td.MainInstruction += Utils.MakeCountMsg(ovr,"{0} file{1} overwritten. ");
      if (skp>0) td.MainInstruction += Utils.MakeCountMsg(skp,"{0} file{1} skipped. "    );
      if (err>0) td.MainInstruction += Utils.MakeCountMsg(err,"{0} error{1} occurred. "  );
    }
    td.Show();
  }
  
  #endregion
  
  #region Instance Fields

  /// <summary>Nominal diameter in millimeters</summary>
  public readonly int D; 

  /// <summary>Thread pitch in millimeters (DIN 13, DIN ISO 68-1)</summary>
  public readonly double P;

  /// <summary>Wrench size in millimeters (DIN 931, DIN 933)</summary>
  public readonly double s;

  /// <summary>Height of bolt head and nut in millimeters (DIN 931, DIN 933)</summary>
  public readonly double k;

  /// <summary>Maximum distance from bolt head to thread in millimeters (DIN 933)</summary>
  public readonly double a;

  /// <summary>Diameter of washer clearance hole in millimeters (DIN 125)</summary>
  public readonly double du1;

  /// <summary>Washer diameter in millimeters (DIN 125)</summary>
  public readonly double du2;

  /// <summary>Washer thickness in millimeters (DIN 125)</summary>
  public readonly double u;

  /// <summary>Fine clearance hole diameter (H12) in millimeters (EN 20273)</summary>
  public readonly double dh1;

  /// <summary>Medium clearance hole diameter (H13) in millimeters (EN 20273)</summary>
  public readonly double dh2;

  /// <summary>Coarse clearance hole diameter (H14) in millimeters (EN 20273)</summary>
  public readonly double dh3;

  /// <summary>Default grip length of bolt assembly in millimeters (own definition)</summary>
  public readonly double dgl;
  
  /// <summary>Customary bolt lengths in millimeters</summary>
  public readonly double[] cls;

  /// <summary>Effective pitch diameter in millimeters (DIN 13, DIN ISO 68-1; computed)</summary>
  public readonly double d2;

  /// <summary>Thread height in millimeters (DIN 13, DIN ISO 68-1; computed)</summary>
  public readonly double H;

  /// <summary>Nominal circumference in millimeters (computed)</summary>
  public readonly double C;

  /// <summary>Thread pitch angle in degrees (computed)</summary>
  public readonly double beta;

  /// <summary>Minimum thread length in millimeters for bolt lengths &lt; 125 mm (DIN 931, computed)</summary>
  public readonly double b2;

  /// <summary>Minimum thread length in millimeters for bolt lengths &lt; 200 mm (DIN 931, computed)</summary>
  public readonly double b3;

  /// <summary>Minimum thread length in millimeters for bolt lengths &gt; 200 mm (DIN 931, computed)</summary>
  public readonly double b4;

  #endregion

  #region Instance API

  /// <summary>Creates a new ISO metric bolt gemoetry.</summary>
  /// <param name="D">Nominal diameter in millimeters</param>
  /// <param name="P">Thread pitch in millimeters</param>
  /// <param name="s">Wrench size in millimeters</param>
  /// <param name="k">Height of bolt head and nut in millimeters</param>
  /// <param name="a">Maximum distance from bolt head to thread in millimeters</param>
  /// <param name="du1">Diameter of washer clearance hole in millimeters</param>
  /// <param name="du2">Washer diameter in millimeters</param>
  /// <param name="u">Washer thickness in millimeters</param>
  /// <param name="dh1">Fine clearance hole diameter (H12) in millimeters</param>
  /// <param name="dh2">Medium clearance hole diameter (H13) in millimeters</param>
  /// <param name="dh3">Coarse clearance hole diameter (H14) in millimeters</param>
  /// <param name="dgl">Default grip length of bolt assembly in millimeters</param>
  /// <param name="cls">Customary bolt lengths in millimeters</param>
  private BoltGeometry(
    int      D, 
    double   P, 
    double   s, 
    double   k, 
    double   a, 
    double   du1, 
    double   du2, 
    double   u,
    double   dh1,
    double   dh2,
    double   dh3,
    double   dgl, 
    double[] cls
  )
  {
    // Set field values from arguments
    this.D    = D;
    this.P    = P;
    this.s    = s;
    this.k    = k;
    this.a    = a;
    this.du1  = du1;
    this.du2  = du2;
    this.u    = u;
    this.dh1  = dh1;
    this.dh2  = dh2;
    this.dh3  = dh3;
    this.dgl  = dgl;
    this.cls  = cls;

    // Cumpute field values from arguments
    this.d2   = D - 3*Math.Sqrt(3)/8 * P;
    this.H    = Math.Sqrt(3)/2 * P;
    this.C    = Math.PI*this.D;
    this.beta = Math.Atan2(P,C)*180/Math.PI;
    this.b2   = 2*D + 6;
    this.b3   = 2*D + 12;
    this.b4   = 2*D + 25;

    // Register bolt geometry with dictionary
    BoltGeometry.Geometries.Add(String.Format("M{0}",this.D),this);
  }

  public override string ToString()
  {
    return string.Format("[IMThread D={0}, P={1}, C={2}, beta={3}]", D, P, C, beta);
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
      string sCmdSupp1    = "Will create {0} thread material{1}. "
                          + "Depending on the number of materials being created, "
                          + "the operation may take a few seconds.";
      string sCmdSupp2    = "Please choose below whether {0} existing thread material{1} "
                          + "shall be overwritten.";
      sCmdSupp1           = Utils.MakeCountMsg(counts[CNT_TGEO]*counts[CNT_VMAT],sCmdSupp1);
      sCmdSupp2           = Utils.MakeCountMsg(counts[CNT_TMAT],sCmdSupp2);
      td.AddCommandLink(
        TaskDialogCommandLinkId.CommandLink1,
        "Create thread materials",
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
          "Delete thread materials",
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
  public static string MakeName(Material vMat, BoltGeometry tgeo)
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

  /// <summary>Finds the bump gradient map, if any, in an appearance asset</summary>
  /// <param name="appAss">The appearance asset</param>
  /// <returns>The bump gradient map</returns>
  /// <exception cref="ArgumentException">if the argument is <c>null</c> or not an appearance asset</exception>
  /// <exception cref="Exception"> if no bump gradient map is found.</exception>
  public static Asset GetBumpGradientMap(Asset appAss)
  {
    // Check argument
    if (appAss==null)
      throw new ArgumentException("Argument appAss must not be null");
    if (!appAss.AssetType.Equals(AssetType.Appearance))
      throw new ArgumentException("Argument appAss must be an appearance asset");

    // Find bump map
    Asset bmapAss = null;    // Result
    string[] bmapProps = {   // Appearance asset properties to examine
      "generic_bump_map",    // in generic schema
      "metal_pattern_shader" // in metal schema
    };
    foreach (string bmapProp in bmapProps)
    {
      AssetProperty apr = appAss.FindByName(bmapProp);
      if (apr==null) continue;
      bmapAss = apr.GetSingleConnectedAsset();
      if (bmapAss==null) continue;
      string baseSchema = GetAP<AssetPropertyString>(bmapAss,"BaseSchema").Value;
      if ("GradientSchema".Equals(baseSchema))
        return bmapAss;
    }

    // No bump gradient map found
    throw new Exception("No bump gradient map found");
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
      Asset                  vAppAss       = vAppAssElem.GetRenderingAsset();
      Log.WL(String.Format(prefix+"Material has an appearance asset -> OK"));

      // Assert appearance asset has a bump gradient map
      try
      {
        GetBumpGradientMap(vAppAss);
        Log.WL(prefix+"Appearance asset has a bump gradient map -> OK");
      }
      catch (Exception e)
      {
        string problem = "Appearance asset has no bump gradient map";
        Log.WL(prefix+problem+" -> FAILED");
        Log.WL(e.ToString());
        throw new Exception(lname+": "+problem,e);
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
  private static Material Create(string name, Material vMat, BoltGeometry tgeo)
  {
    string prefix = "- ";
    //Dump.Material(vMat,prefix);
    Log.WL(prefix+String.Format(
      "Creating M{0} thread material from template \"{1}\"",
      tgeo.D,
      vMat.Name
     ));

    // Initialize
    Document doc       = vMat.Document;
    string category    = VMATS.Match(vMat.Name).Groups[1]+"";   // E.g. "Steel galvanized"
    string description = String.Format(                         // E.g. "Generic ISO metric bolt assembly, steel galvanized with M12 tread"
                           "Generic ISO metric bolt assembly: {0} with M{1} tread", 
                           category,tgeo.D
                         );
    string  comments   = String.Format(
                           "Rendering material for M{0} thread. Use macro Thread_Materials in GIMBA.rfa "
                           + " to manage thread materials. See {1} for further instructions.",
                           tgeo.D,Utils.HelpUrl
                         );

    // Duplicate and edit template material
    ElementId              tAppAssElemId = ElementId.InvalidElementId;
    ElementId              vAppAssElemId = vMat.AppearanceAssetId;
    AppearanceAssetElement vAppAssElem   = doc.GetElement(vAppAssElemId) as AppearanceAssetElement;
  
    // - Duplicate the material
    Material tMat = vMat.Duplicate(name);
    
    // - Make changes to the Material
    tMat.get_Parameter(BuiltInParameter.ALL_MODEL_MANUFACTURER).Set(Utils.Author);
    tMat.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).Set(comments);
    tMat.get_Parameter(BuiltInParameter.ALL_MODEL_URL).Set(Utils.RepoUrl);
    tMat.get_Parameter(BuiltInParameter.ALL_MODEL_DESCRIPTION).Set(description);

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

      // - Modify bump gradient map
      Asset tBmap = GetBumpGradientMap(tAppAss);
      GetAP<AssetPropertyDistance>(tBmap,"texture_RealWorldScaleX" ).Value = tgeo.P / 25.4;
      GetAP<AssetPropertyDistance>(tBmap,"texture_RealWorldScaleY" ).Value = tgeo.C / 25.4;
      GetAP<AssetPropertyDouble  >(tBmap,"texture_WAngle"          ).Value = 90 - tgeo.beta;
      GetAP<AssetPropertyBoolean >(tBmap,"texture_ScaleLock"       ).Value = false;
      GetAP<AssetPropertyBoolean >(tBmap,"texture_URepeat"         ).Value = true;
      GetAP<AssetPropertyBoolean >(tBmap,"texture_VRepeat"         ).Value = true;

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
    IList<BoltGeometry> tgeos = BoltGeometry.GetList();
    foreach (BoltGeometry tgeo in tgeos)
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
    foreach (Material vMat in new List<Material>(vMats)) // Iterate new list allows removing elements in vMat
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
        Log.WL("- Create thread materials operation selected by user");
        break;
      case TaskDialogResult.CommandLink2:
        Log.WL("- Delete thread materials operation selected by user");
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
            foreach (BoltGeometry tgeo in tgeos)
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
      td.FooterText      = Log.MakeLogFileLink("View log");
      td.Show();
    }
    
    #endregion
    
    #region Revit macro handlers

    public void Thread_Materials()
    {
      Thread.CurrentThread.CurrentCulture = new CultureInfo(""); // Set invariant culture
      Log.Begin("Thread_Materials",this.Document);
      Utils.BackupThisSource(this.Document);
      try
      {
        Document actDoc = GetActiveDocument();
        CheckDocument(actDoc);
        ThreadMaterials.MacroMain(actDoc);
      }
      catch (Exception e)
      {
        Log.WL(e.ToString());
        DoErrorWrapupDialog(e);
      }
      Log.End();
    }
    
    public void Catalogs_And_Tables()
    {
      Thread.CurrentThread.CurrentCulture = new CultureInfo(""); // Set invariant culture
      Log.Begin("Catalogs_And_Tables",this.Document);
      Utils.BackupThisSource(this.Document);
      try
      {
        BoltGeometry.CatalogsAndTables(this.Document);
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