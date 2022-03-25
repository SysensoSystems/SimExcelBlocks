function blkStruct = slblocks
% SLBLOCKS Defines the block library for a specific Toolbox or Blockset.

blkStruct.MaskDisplay = 'disp(''ExcelBlocks'')';
blkStruct.OpenFcn = 'simExcelLibrary';
blkStruct.Name = 'Excel Blocks';

% Define the library list for the Simulink Library browser.
% Return the name of the library model and the name for it.
if exist('simExcelLibrary') == 4
    Browser(1).Library = 'simExcelLibrary';
    Browser(1).Name    = 'Excel Blocks';
    Browser(1).IsFlat  = 1;
    blkStruct.Browser = Browser;
end

end