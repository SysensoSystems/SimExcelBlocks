function LookupExcelGUI(varargin)
% Helps to create a lookup table which can use the data from an Excel file.
%
% Developed by: Sysenso Systems, https://sysenso.com/
% Contact: contactus@sysenso.com
%
% Version:
% 1.0 - Initial Version.
%

% Iunput assignment
if isempty(varargin)
    handles.blockHandle = [];
else
    blockPath = varargin{1};
    handles.blockHandle = get_param(blockPath,'Handle');
end

% Check if a GUI is already open for the given block and avoid opening
% another GUI instance for the same block.
figureHandle = findall(0,'Name','FromExcel Block Parameters','UserData',handles.blockHandle);
if ~isempty(figureHandle)
    figure(figureHandle);
    return;
end

% Creating GUI
handles.figure = figure('Tag','LookupExcel Block','Visible','off','Menubar','none','Toolbar','none','Name','LookupExcel Block Parameters','NumberTitle','off','UserData',handles.blockHandle);
screenSize = get(0,'screensize');
set(handles.figure,'Units','Pixels','Position',[0.3*screenSize(3) 0.35*screenSize(4) 0.4*screenSize(3) 0.3*screenSize(4)]);
movegui(handles.figure,'center');

mainPanel = uiflowcontainer('v0','Units','norm','Position',[.01,.01,.98,.98],'parent',handles.figure);
set(mainPanel,'FlowDirection','TopDown');
browsePanel = uiflowcontainer('v0','parent',mainPanel);
set(browsePanel,'HeightLimits',[30,30]);
handles.name = uicontrol('Parent',browsePanel,'Style','checkbox','CData',nan(1,1,3),'String','File','FontSize',10,'HorizontalAlignment','center');
set(handles.name,'WidthLimits',[40,40]);
handles.textPath = uicontrol('Parent',browsePanel,'Style','checkbox','CData',nan(1,1,3),'String','','BackgroundColor',[1 1 1]);
set(handles.textPath,'WidthLimits',[100,inf]);
handles.browseButton = uicontrol('Parent',browsePanel,'Style','pushbutton','String','Browse');
set(handles.browseButton,'WidthLimits',[80,80]);
dimensionPanel = uiflowcontainer('v0','parent',mainPanel);
set(dimensionPanel,'HeightLimits',[30,30]);
handles.tableDimensionName = uicontrol('Parent',dimensionPanel,'Style','text','String','Enter the table dimension','FontSize',10,'HorizontalAlignment','right');
set(handles.tableDimensionName,'WidthLimits',[100,inf]);
handles.tableDimensions = uicontrol('Parent',dimensionPanel,'Style','edit','String','','BackgroundColor',[1 1 1]);
set(handles.tableDimensions,'WidthLimits',[40,40]);
handles.createButton = uicontrol('Parent',dimensionPanel,'Style','pushbutton','String','Create');
set(handles.createButton,'WidthLimits',[80,80]);
handles.uiTable = uitable('Parent',mainPanel,'ColumnName',{'Parameters' 'Sheet Number/Name' 'Cell Range'},'ColumnEditable',[false true true]);
actionPanel = uiflowcontainer('v0','parent',mainPanel);
set(actionPanel,'HeightLimits',[30,30]);
uicontainer('Parent',actionPanel);
handles.updateButton = uicontrol('Parent',actionPanel,'Style','pushbutton','String','Update');
set(handles.updateButton,'WidthLimits',[80,80]);
handles.cancelButton = uicontrol('Parent',actionPanel,'Style','pushbutton','String','Cancel');
set(handles.cancelButton,'WidthLimits',[80,80]);
handles.helpButton = uicontrol('Parent',actionPanel,'Style','pushbutton','String','Help');
set(handles.helpButton,'WidthLimits',[80,80]);
uicontainer('Parent',actionPanel);

% Restoring the data entered previously when the GUI is accessed once again
try
    restoreData = get_param(handles.blockHandle,'UserData');
    if ~isempty(restoreData)
        set(handles.textPath,'String',restoreData.fileName);
        set(handles.tableDimensions,'String',restoreData.signalEdit);
        set(handles.uiTable,'Data',restoreData.tableData);
    end
catch
end

% Column width management of the uitable
handles.columnWidth = [0.35 0.2 0.35];
figureResizeCallback([],[],handles);
set(handles.figure,'Visible','on');

% Callback function for the buttons
set(handles.browseButton,'Callback',@(h,e)browseButtonCallback(h,e,handles));
set(handles.createButton,'Callback',@(h,e)createButtonCallback(h,e,handles));
set(handles.updateButton,'Callback',@(h,e)updateButtonCallback(h,e,handles));
set(handles.cancelButton,'Callback',@(h,e)cancelButtonCallback(h,e,handles));
set(handles.helpButton,'Callback',@(h,e)helpButtonCallback(h,e,handles));
set(handles.figure,'SizeChangedFcn',@(h,e)figureResizeCallback(h,e,handles));

end
%--------------------------------------------------------------------------
function browseButtonCallback(hObject,event,handles)
% Opening the browse window for the user to select the excelfile from which
% data has to be imported

[filename, pathname] = uigetfile({'*.xlsx';'*xls'},'Pick an excel file');
fullName = [pathname filename];
winopen(fullName);
set(handles.textPath,'String',fullName);
end
%--------------------------------------------------------------------------
function createButtonCallback(hObject,event,handles)
% Creating table for the user to enter the address of the data

signalLength = str2double(get(handles.tableDimensions,'String'));
if (signalLength > 0) && (signalLength <3)
    tableData = {};
    tableData{1,1} = 'Table Data';
    for length = 1:signalLength
        tableData{length+1,1} = ['Breakpoint ' num2str(length)];
        for width = 2:3
            tableData{length,width} = '';
        end
    end
    set(handles.uiTable,'Data',tableData);
else
    msgbox('Please enter valid no of dimension either 1 or 2','Error');
    set(handles.tableDimensions,'String','');
    return;
end

end
%--------------------------------------------------------------------------
function updateButtonCallback(hObject,event,handles)
% Extracting the data entered by the user in the UITable

if isempty(get(handles.uiTable,'Data'))
    msgbox('Please enter the details for importing the data from excel','Error');
    return;
end
if isempty(get(handles.textPath,'String'))
    msgbox('Please select the file from which data has to be imported','Error');
    return;
end
tableData = get(handles.uiTable,'Data');
% Check all the fields are available
if ~isempty(find(cellfun(@(x) isempty(x), tableData)))
    msgbox('Data should not be empty. Please check.','Error');
    return;
end
xlsFile = get(handles.textPath,'String');
% Checking the sheet.
[~,xlSheets] = xlsfinfo(xlsFile);
tableSheets = {};
for signalInd = 1:size(tableData,1)
    sheetName = tableData{signalInd,2};
    sheetNum = str2double(sheetName);
    if isnumeric(sheetNum)
        if ~any(isequal(sheetNum,1:length(xlSheets)))
            msgbox(['The sheet given ''' sheetName ''' is not available in the Excel file.'],'Error');
            return;
        else
            sheetName = xlSheets{sheetNum};
        end
    elseif ischar(sheetName)
        if ~any(strcmp(sheetName,xlSheets))
            msgbox(['The sheet given ''' sheetName ''' is not available in the Excel file.'],'Error');
            return;
        end
    end
    tableSheets = [tableSheets; sheetName];
end

% Importing data for breakpoints
for signalInd = 2:size(tableData,1)
    try
        breakpointData = xlsread(xlsFile,tableSheets{signalInd},tableData{signalInd,3});
    catch
        msgbox(['Error in reading the ' xlsFile ',in the sheet:' tableSheets{signalInd} ' for the cell:' tableData{signalInd,3} '.'],'Error');
        return;
    end
    if any(isnan(breakpointData))
        msgbox('The table data imported is not having valid numeric data.','Error');
        return;
    end
    spreadSheetData.breakpoint(signalInd-1).data = breakpointData;
end

% Importing the table data
try
    tableDataSignal = xlsread(xlsFile,tableSheets{1},tableData{1,3});
catch
    msgbox(['Error in reading the ' xlsFile ',in the sheet:' tableSheets{1} ' for the cell:' tableData{1,3} '.'],'Error');
    return;
end
if any(isnan(tableDataSignal))
    msgbox('The table data imported is not having valid numeric data.','Error');
    return;
end
spreadSheetData.tableDimension = str2double(get(handles.tableDimensions,'String'));
spreadSheetData.table.data = tableDataSignal;

% Storing the table data enetered by the user
blockUserData.fileName = xlsFile;
blockUserData.signalEdit = get(handles.tableDimensions,'String');
blockUserData.tableData = tableData;

% Sending data to create blocks and set block parameters for the block
updateToLookupTableBlock(spreadSheetData,blockUserData,handles.blockHandle);
close(handles.figure);

end
%--------------------------------------------------------------------------
function updateToLookupTableBlock(spreadSheetData,blockUserData,blockHandle)
% Creates the lookblock block with table data

% Creating dimensions for the blocks and width between the blocks
portWidth = 30;
portHeight = 14;
blockWidth = 50;
blockHeight = spreadSheetData.tableDimension*50;
xPos = 200;
yPos = 20;
xGap = 150;
yGap = 40;

% Clearing the existing blocks
deleteInputLine = get_param(blockHandle,'LineHandles');
noOfLines = length(deleteInputLine.Inport);
for lineCount = 1:noOfLines
    if deleteInputLine.Inport(lineCount) ~= -1
        delete_line(deleteInputLine.Inport(lineCount));
    end
end
Simulink.SubSystem.deleteContents(blockHandle);
set_param(blockHandle,'UserDataPersistent','on');
set_param(blockHandle,'UserData', blockUserData);

% Creating blocks and connecting ports and setting block parameters with
% the data enetered by the user
blockPath = [get_param(blockHandle,'Parent') '/' get_param(blockHandle,'Name')];
lookUpPath = [blockPath '/' 'lookUpTable'];
outPortPath = [blockPath '/' 'Result'];
add_block('simulink/Lookup Tables/n-D Lookup Table',lookUpPath,'Position',[xPos yPos-blockHeight/2 xPos+blockWidth yPos+blockHeight/2]);
set_param(lookUpPath,'NumberOfTableDimensions',mat2str(spreadSheetData.tableDimension));
set_param(lookUpPath,'Table',mat2str(spreadSheetData.table.data));
portHandles = get_param(lookUpPath,'PortHandles');
for signalInd = 1:length(spreadSheetData.breakpoint)
    name = ['data' num2str(signalInd)];
    inPortPath = [blockPath '/' name];
    add_block('simulink/Sources/In1',inPortPath);
    portPosition = get_param(portHandles.Inport(signalInd),'Position');
    set_param(inPortPath,'Position',[portPosition(1)-xGap portPosition(2)-portHeight/2 portPosition(1)-xGap+portWidth portPosition(2)+portHeight/2]);
    set_param(lookUpPath,['BreakpointsForDimension' num2str(signalInd)],mat2str(spreadSheetData.breakpoint(signalInd).data));
    add_line(blockPath,[name '/1'],['lookUpTable/' num2str(signalInd)],'autorouting','on');
    yPos = yPos + yGap;
end
add_block('simulink/Sinks/Out1',outPortPath);
portPosition = get_param(portHandles.Outport,'Position');
set_param(outPortPath,'Position',[portPosition(1)+xGap portPosition(2)-portHeight/2 portPosition(1)+xGap+portWidth portPosition(2)+portHeight/2]);
add_line(blockPath,'lookUpTable/1','Result/1','autorouting','on');

end
%--------------------------------------------------------------------------
function cancelButtonCallback(hObject,event,handles)
% Closes the GUI without any prompt.

close(handles.figure);
end
%--------------------------------------------------------------------------
function helpButtonCallback(hObject,event,handles)
% Launches help file.

open('ExcelBlocks_doc.pdf');
end
%--------------------------------------------------------------------------
function figureResizeCallback(hObject,event,handles)

figureSize = get(handles.figure,'Position');
set(handles.uiTable,'ColumnWidth',num2cell(handles.columnWidth.*figureSize(3)));
end