This repository contains examples of web content, online strategies and marketing examples from working at Loofe's Clothing. 

Examples

Size Guide
The aim of this project was to allow colleagues (who were not versed in HTML and CSS) to easily develop stylised branded size guides for the website. All they would need to do it complete the table as shown in the example: an excel document for each brand, a 'sheet' per department, headers on rows 3 and 4 with values row 5 onwards further details on how to use this template is included in sizeguide-example.xlsm.

This example includes the following files:
sizeguide-example.xlsm: contains a detail explanation on how to use the size guide template including example tables for a handful of categories.
website-size-guide-style.css: a local version of the CSS file used to style all size guides. Inline CSS is used minimally to allow for changes to be made in future more easily without needing to modify every size guide individually. This CSS includes the unit selector styles, this allows switching between for example centimetres and inches without using JavaScript facilitating faster loads.
ResetFormatMacro.vb: a small VBA (visual basic for applications) script to reset the format of all sheets, this was implemented because formatting could become messy if copy and paste was used to complete the tables.
MakeSizeGuide.vb: the VBA code that converts the tables in excel to HTML documents for use on the Loofe's Website. The commenting is currently minimal but should still explain what each function is doing and commenting will be improved in future.
BaseCode.xlsm: the base Excel document the contains the above VBA scripts that all other size guide templates link to, this was done so that any changes to code only need to be made once (within this document) and it was automatically be translated to all the other size guide files. IMPORTANT: this file must be one level above any size guide file such as with the Under Armour Example, is the files are rearranged the template cannot locate the base code and it will not generate the HTML files.
Under Armour Example: This folder contains the Excel document outlining the size guides for the brand Under Armour and all the HTML files this template generated.