# Revit-Generic-ISO-Metric-Bolt-Assembly
Autodesk Revit families and macros for generic ISO metric bolts and bolt assemblies

## ISO Metric Hex Bolt Assembly
* extensive type catalog M 5 ... M 64
* including optional nut and washers
* freely adjustable bolt length
* all other dimensions automated conforming to DIN 931/ISO 4014 and DIN 933/ISO 7089
* grip and bolt modes (see preview)
* galvanized steel material and thread texture included for rendering the assembly like seen in preview
* minimalist geometry, no junk parameters
* Revit version: 2022+

## Remarks
* Visibility of assembly is controlled by detail level
  * Coarse: Assembly not visible(!)
  * Medium: Bolt head, nut, washers, and simplified bolt end visible
  * Fine: Detailed assembly visible
<!--
* Included thread texture is installed as follows:
  * Copy texture file "M12 Metric Thread - displace.png" into your project folder
  * Open material "Steel galvanized, metric thread" in the materials editor and go to the appearance tab
  * Re-link the three references to the texture file
  * Add the path of the texture file to your rendering path
  * You're good to go...
-->

## Standards
* ISO metric thread: DIN 13, DIN 68-1 (see e.g. http://www.iso-gewinde.at)
* Metric hex bolts /w shaft: DIN 931/ISO 4014 (see e.g. https://www.wegertseder.com/Download/techdat/DIN-931-Sechskantschrauben-mit-Schaft.pdf)
* Metric hex bolts: DIN-933/ISO 7089 (see e.g. https://www.wegertseder.com/Download/techdat/DIN-933-Sechskantschrauben-Gewinde-bis-Kopf.pdf)
