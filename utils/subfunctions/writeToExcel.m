function writeToExcel(blockUserData)
% Helps to write the simulation results to the required excel file.
% It will be called as a 'StopFcn' callback from the ToExcel block.
%

tableData = blockUserData.tableData;
% Storing the data in excel file
for signalInd = 1:size(tableData,1)
    signalName = tableData{signalInd,1};
    sheetName = tableData{signalInd,2};
    startingCell =  tableData{signalInd,3};
    timeStartingCell = tableData{signalInd,4};
    
    % Extracting the simulated data
    textData = ['ToExcel_' signalName];
    workSpaceData = evalin('base',textData);
    timeName = [signalName '-Time'];
    if strcmpi(tableData{signalInd,5},'Column Vector')
        excelWriteData = [{signalName}; num2cell(workSpaceData.Data)];
        timeData = [{timeName}; num2cell(workSpaceData.Time)];
    else
        excelWriteData = [{signalName}, num2cell(workSpaceData.Data')];
        timeData  = [{timeName}, num2cell(workSpaceData.Time')];
    end
    if ~isnan(str2double(sheetName))
        xlswrite(blockUserData.fileName,excelWriteData,str2double(sheetName),startingCell);
        xlswrite(blockUserData.fileName,timeData,str2double(sheetName),timeStartingCell);
    else
        xlswrite(blockUserData.fileName,excelWriteData,sheetName,startingCell);
        xlswrite(blockUserData.fileName,timeData,sheetName,timeStartingCell);
    end
end

end