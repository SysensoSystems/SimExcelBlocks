%% *|Simulink Excel Blocks|*
%
% Simulink Excel Blocks is a custom Simulink library which will be useful to the refer the Excel files within Simulink model.
%
% Developed by: Sysenso Systems, <https://sysenso.com/>
%
% Contact: contactus@sysenso.com
%
% Version:
% 1.0 - Initial Version.
%
%
% *|Detailed Information|*
%%
% *Excel Blocks Library*
%
% * If this SimExcelBlocks folder is in MATLAB path, then this library will be available to use from Simulink Library Browser.
%
% <<\images\\LibBrowser.png>>
%
% If this library is not visible in Simulink Library Browser, then close the Simulink and type following command in MATLAB command window.
% >> sl_refresh_customizations
% Now, it will become available within the Simulink Library Browser.
%
%%
% * There are three blocks available in this library.
%
% # FromExcel Block - Helps to read the data from the Excel file and use it as a source block data for Simulink
% # ToExcel Block - Helps to write the Simulink simulation results to the Excel file. It can be used as a Sink block in the Simulink model.
% # LookupExcel Block - Helps to create a lookup table which can use the data from an Excel file.
%
%%
% *Demo*
%
% All these blocks are explained in the below sections using ''sldemo_autotrans'' demo model available in Simulink toolbox.
% The sample blocks and the corresponding Excel data files to use it in sldemo_autotrans are available in the testdata folder.
%
% <<\images\\ReferBlocks.png>>
%
%%
% *FromExcel Block*
%
% This block has following dialog settings.
%
% <<\images\\FromExcelBlock.png>>
%
% # File - Browse the Excel file from where the data has be referred.
% # Number of Signals  - Number of signals that has to be referred from the Excel file.
% # Table Information - Table should be populated with unique name for every signal and also information related to sheet name/number and cell ranges.
%
%
% * Refer the below image. Instead of using the ManeuversGUI block as a source block in the sldemo_autotrans model, the user can use the FromExcel block by referring the data from Excel file.
%
% <<\images\\sldemo_fromblock.png>>
%
% * The dialog settings for the FromExcel block and the respective Excel file can be as shown below.
%
% <<\images\\FromExcelData.png>>
%
%%
% *ToExcel Block*
%
% This block has following dialog settings.
%
% <<\images\\ToExcelBlock.png>>
%
% # File - Path of the Excel file to write the Simulation results.
% # Number of Signals  - Number of signals that has to be written to the Excel file.
% # Table Information - Table should be populated with unique name for every signal and also information related to sheet name/number and the starting cell to write the data.
% The data can be written as a column vector or as a row vector.
%
%
% * Refer the below image. All the signals connected to the Scope block(Plot Results) are connected to the ToExcel block in the sldemo_autotrans model.
%
% <<\images\\sldemo_toblock.png>>
%
% * The dialog settings for the ToExcel block and the corresponding ToExcel file generated are shown below.
%
% <<\images\\ToExcelData.png>>
%
%%
% *LookupExcel Block*
%
% This block has following dialog settings.
%
% <<\images\\LookupExcelBlock.png>>
%
% # File - Issue the Excel file to read the lookup table data.
% # Lookup table dimension  - As of now, this block supports only 1D 0r 2D lookup tables.
% # Table Information - Table should be populated with the information related to sheet name/number and cell ranges to refer.
%
%
% * Refer the below image. The 2D lookup table in sldemo_autotrans/Engine/EngineTorque path can also be created with LookupExcel block by referring the data from Excel file.
%
% <<\images\\sldemo_lookupblock.png>>
%
% * The dialog settings for the LookupExcel block and the corresponding Excel file data can be as shown below.
%
% <<\images\\LookupExcelData.png>>
%
%%
% *Note: Please share your comments and contact us if you are interested in updating the features further.*
%