function ToExcelGUI(varargin)
% Helps to write the Simulink simulation results to the Excel file. It can
% be used as a Sink block in the Simulink model.
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
handles.figure = figure('Tag','ToExcel Block','UserData',[],'Visible','off','Menubar','none','Toolbar','none','Name','ToExcel Block Parameters','NumberTitle','off','UserData',handles.blockHandle);
screenSize = get(0,'screensize');
set(handles.figure,'Units','Pixels','Position',[0.3*screenSize(3) 0.3*screenSize(4) 0.4*screenSize(3) 0.4*screenSize(4)]);
movegui(handles.figure,'center');

mainPanel = uiflowcontainer('v0','Units','normalized','Position',[.01,.01,.98,.98],'parent',handles.figure);
set(mainPanel,'FlowDirection','TopDown');
savePanel = uiflowcontainer('v0','parent',mainPanel);
set(savePanel,'HeightLimits',[30,30]);
handles.name = uicontrol('Parent',savePanel,'Style','checkbox','CData',nan(1,1,3),'String','File','FontSize',10,'HorizontalAlignment','center');
set(handles.name,'WidthLimits',[40,40]);
handles.textPath = uicontrol('Parent',savePanel,'Style','checkbox','CData',nan(1,1,3),'String','','BackgroundColor',[1 1 1]);
set(handles.textPath,'WidthLimits',[100,inf]);
handles.saveButton = uicontrol('Parent',savePanel,'Style','pushbutton','String','Save');
set(handles.saveButton,'WidthLimits',[80,80]);
signalPanel = uiflowcontainer('v0','parent',mainPanel);
set(signalPanel,'HeightLimits',[30,30]);
emptySpace = uicontainer('Parent',signalPanel);
set(emptySpace,'WidthLimits',[inf,inf]);
handles.signalNumName = uicontrol('Parent',signalPanel,'Style','text','CData',nan(1,1,3),'String','Number of signals to be saved ','FontSize',10,'HorizontalAlignment','right');
set(handles.signalNumName,'WidthLimits',[100,inf]);
handles.signalNum = uicontrol('Parent',signalPanel,'Style','edit','String','','BackgroundColor',[1 1 1]);
set(handles.signalNum,'WidthLimits',[40,40]);
handles.createButton = uicontrol('Parent',signalPanel,'Style','pushbutton','String','Create');
set(handles.createButton,'WidthLimits',[80,80]);
handles.uiTable = uitable('Parent',mainPanel,'ColumnName',{'Signal Name','Sheet Number/Name','Signal Starting Cell','Time Starting Cell','Save Format'},'ColumnEditable',true,'ColumnFormat',{'','','','',{'Column Vector','Row Vector'}});
tableEditPanel = uiflowcontainer('v0','parent',mainPanel);
set(tableEditPanel,'HeightLimits',[30,30]);
uicontainer('Parent',tableEditPanel);
handles.removeSignalButton = uicontrol('Parent',tableEditPanel,'Style','pushbutton','String','Remove Signal');
set(handles.removeSignalButton,'WidthLimits',[80,80]);
handles.upButton = uicontrol('Parent',tableEditPanel,'Style','pushbutton','String','Up');
set(handles.upButton,'WidthLimits',[80,80]);
handles.downButton = uicontrol('Parent',tableEditPanel,'Style','pushbutton','String','Down');
set(handles.downButton,'WidthLimits',[80,80]);
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
        set(handles.signalNum,'String',restoreData.signalEdit);
        set(handles.uiTable,'Data',restoreData.tableData);
    end
catch
end

% Column width management of the uitable
handles.columnWidth = [0.24 0.16 0.19 0.19 0.15];
figureResizeCallback([],[],handles);
set(handles.figure,'Visible','on');

% Callback function for the buttons
set(handles.saveButton,'Callback',@(h,e)saveButtonCallback(h,e,handles));
set(handles.createButton,'Callback',@(h,e)createButtonCallback(h,e,handles));
set(handles.removeSignalButton,'Callback',@(h,e)removeSignalButtonCallback(h,e,handles));
set(handles.upButton,'Callback',@(h,e)upButtonCallback(h,e,handles));
set(handles.downButton,'Callback',@(h,e)downButtonCallback(h,e,handles));
set(handles.uiTable,'CellSelectionCallback',@(h,e)cellSelectionCallback(h,e,handles));
set(handles.updateButton,'Callback',@(h,e)updateButtonCallback(h,e,handles));
set(handles.cancelButton,'Callback',@(h,e)cancelButtonCallback(h,e,handles));
set(handles.helpButton,'Callback',@(h,e)helpButtonCallback(h,e,handles));
set(handles.figure,'SizeChangedFcn',@(h,e)figureResizeCallback(h,e,handles));

end
%--------------------------------------------------------------------------
function cellSelectionCallback(hObject,event,handles)
% Records the current cell indices.

indices = event.Indices;
set(handles.uiTable,'UserData',indices);
end
%--------------------------------------------------------------------------
function saveButtonCallback(hObject,event,handles)
% To select the excel file to export the data.

[filename,pathname] = uiputfile({'*.xlsx';'*.xls'},'Save to');
if isequal(filename,0)
    return;
end
fullName = [pathname filename];
set(handles.textPath,'String',fullName);

end
%--------------------------------------------------------------------------
function createButtonCallback(hObject,event,handles)
% Creates table rows.

signalLength = str2double(get(handles.signalNum,'String'));
if signalLength > 0
    tableData = {};
    for rowInd = 1:signalLength
        for columnInd = 1:4
            tableData{rowInd,columnInd} = '';
        end
        tableData{rowInd,5} = 'Column Vector';
    end
    set(handles.uiTable,'Data',tableData);
else
    msgbox('Please enter valid no of signals','Error');
    return;
end

end
%--------------------------------------------------------------------------
function removeSignalButtonCallback(hObject,event,handles)
% Removing the row once the user select the rows and click remove signal
% button.

if isempty(get(handles.uiTable,'Data'))
    return;
end
selectedCells = get(handles.uiTable,'UserData');
removeIndex = unique(selectedCells(:,1));
if isempty(removeIndex)
    return;
end
tableData = get(handles.uiTable,'Data');
tableData(removeIndex(1:end),:) = [];
set(handles.uiTable,'Data',tableData);
set(handles.signalNum,'String',num2str(size(tableData,1)));

end
%--------------------------------------------------------------------------
function upButtonCallback(hObject,event,handles)
% Moving the selected row signal in the downward direction once the user
% selects up button.

if isempty(get(handles.uiTable,'Data'))
    return;
end
selectedCells = get(handles.uiTable,'UserData');
upIndex = unique(selectedCells(:,1));
if isempty(upIndex)
    return;
end
tableData = get(handles.uiTable,'Data');
if upIndex(1) == 1
    return;
end
upTableData = [tableData(1:upIndex(1)-2,:); tableData(upIndex,:); tableData(upIndex(1)-1,:); tableData(upIndex(end)+1:size(tableData,1),:)];
set(handles.uiTable,'Data',upTableData);

end
%--------------------------------------------------------------------------
function downButtonCallback(hObject,event,handles)
% Moving the selected row signal in the downward direction once the user
% selects down button.

if isempty(get(handles.uiTable,'Data'))
    return;
end
selectedCells = get(handles.uiTable,'UserData');
downIndex = unique(selectedCells(:,1));
if isempty(downIndex)
    return;
end
tableData = get(handles.uiTable,'Data');
if downIndex(end) == size(tableData,1)
    return;
end
tableData = get(handles.uiTable,'Data');
downTableData = [tableData(1:downIndex(1)-1,:); tableData(downIndex(end)+1,:); tableData(downIndex,:);  tableData(downIndex(end)+2:size(tableData,1),:)];
set(handles.uiTable,'Data',downTableData);

end
%--------------------------------------------------------------------------
function updateButtonCallback(hObject,event,handles)
% Storing the data enetered by the user in the UITable

if isempty(get(handles.uiTable,'Data'))
    msgbox('Please enter the signal details for exporting the data to excel.','Error');
    return;
end
if isempty(get(handles.textPath,'String'))
    msgbox('Please select the file to which data has to be exported.','Error');
    return;
end
tableData = get(handles.uiTable,'Data');
% Check all the fields are available
if ~isempty(find(cellfun(@(x) isempty(x), tableData)))
    msgbox('Data should not be empty. Please check.','Error');
    return;
end

% Validate the table information.
% File name, signal name, sheet name are need not be validated. Only cell
% name have to be validated. Because xlswrite will handle all such issues.
for signalInd = 1:size(tableData,1)
    signalName = tableData{signalInd,1};
    signalCellName = tableData{signalInd,3};
    timeCellName = tableData{signalInd,4};
    if isempty(strcmpi(regexp(signalCellName,'\w+\d+','match'),signalCellName)) || ~strcmpi(regexp(signalCellName,'\w+\d+','match'),signalCellName)
        msgbox(['The cell details given for signal: ' signalName ' to store the signal data is not valid.'],'Error');
        return;
    end
    if isempty(strcmpi(regexp(timeCellName,'\w+\d+','match'),timeCellName)) || ~strcmpi(regexp(timeCellName,'\w+\d+','match'),timeCellName)
        msgbox(['The cell details given for signal: ' signalName ' to store the signal data is not valid.'],'Error');
        return;
    end
end

% Storing the table data enetered by the user in the GUI
blockUserData.fileName = get(handles.textPath,'String');
blockUserData.signalEdit = get(handles.signalNum,'String');
blockUserData.tableData = tableData;

% Sending the data for creating block and to set block parameters
updateToExcelBlock(blockUserData,handles.blockHandle);
close(handles.figure);

end
%--------------------------------------------------------------------------
function updateToExcelBlock(blockUserData,blockHandle)
% Populates the ToExcel block with port and signal data

tableData = blockUserData.tableData;
% Creating dimensions for the blocks and width between the blocks
portWidth = 30;
portHeight = 14;
fwsWidth = 100;
fwsHeight = 20;
xPos = 35;
yPos = 20;
xGap = 180;
yGap = 40;

% Delete only if there are extra lines previously
lineHandles = get_param(blockHandle,'LineHandles');
for lineCount = size(tableData,1)+1:length(lineHandles.Inport)
    if lineHandles.Inport(lineCount) ~= -1
        delete_line(lineHandles.Inport(lineCount));
    end
end
% Record the line points to redraw it again after updating the block
% contents.
linePoints = [];
for lineCount = 1:length(lineHandles.Inport)
    if lineHandles.Inport(lineCount) ~= -1
        linePoints(lineCount).data = get_param(lineHandles.Inport(lineCount),'Points');
    end
end
Simulink.SubSystem.deleteContents(blockHandle);
set_param(blockHandle,'UserDataPersistent','on');
set_param(blockHandle,'UserData',blockUserData);

% Creating blocks and connecting ports and setting block parameters with
% the data enetered by the user
blockPath = [get_param(blockHandle,'Parent') '/' get_param(blockHandle,'Name')];
for signalInd = 1:size(tableData,1)
    twsName = ['data' num2str(signalInd)];
    portName = tableData{signalInd,1};
    workspacePath = [blockPath '/' twsName];
    inPortPath = [blockPath '/' portName];
    add_block('simulink/Sources/In1',inPortPath);
    add_block('simulink/Sinks/To Workspace',workspacePath);
    set_param(inPortPath,'Position',[xPos yPos-portHeight/2 xPos+portWidth yPos+portHeight/2]);
    set_param(workspacePath,'VariableName',['ToExcel_' portName],'Position',[xPos+xGap yPos-fwsHeight/2 xPos+xGap+fwsWidth yPos+fwsHeight/2]);
    add_line(blockPath,[portName '/1'],[twsName '/1'],'autorouting','on');
    yPos = yPos + yGap;
end
% Restore the line connection, if it was connected already.
for lineCount = 1:length(lineHandles.Inport)
    if lineHandles.Inport(lineCount) ~= -1
        set_param(lineHandles.Inport(lineCount),'Points',linePoints(lineCount).data);
    end
end

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
