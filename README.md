# SimExcelBlocks
Simulink Excel Blocks is a custom Simulink library which will be useful to the refer the Excel files within Simulink model.



Usage:
If this SimExcelBlocks folder is in MATLAB path, then this library blocks will be available to use from Simulink Library Browser.

There are three blocks available in this library.
1. FromExcel Block - Helps to read the data from the Excel file and use it as a source block data for Simulink
2. ToExcel Block - Helps to write the Simulink simulation results to the Excel file. It can be used as a Sink block in the Simulink model.
3. LookupExcel Block - Helps to create a lookup table which can use the data from an Excel file.

All these blocks are explained with ''sldemo_autotrans'' demo model available in Simulink toolbox. The sample blocks and the corresponding Excel data files to use it in sldemo_autotrans are available in the testdata folder.
Detailed help can be found from the following document "docs/ExcelBlocks_doc.pdf".


MATLAB Release Compatibility: Created with R2015b, Compatible with R2015b and later releases


Developed by: Sysenso Systems, https://sysenso.com/

Contact: contactus@sysenso.com

Note: Please share your comments and contact us if you are interested in updating the features further.
