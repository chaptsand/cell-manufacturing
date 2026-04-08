function SaveExcelSheet

global run filename sheetname pathname dataSaveDir fileHeader Key KeyRange 
global Stock StockRange
global sourcepathSES targetpathSES sourcefile factorNameList allListRange

%   Open Excel application process.
e_process = actxserver('excel.application');
e_process.DisplayAlerts = false;
if run == 1
    [sourcefile,pathname] = uigetfile( ...
        {'DE_template_*.xlsx','HD-DE template (*.xlsx)'; '*.xlsx','Excel workbook (*.xlsx)'}, ...
        'Select template source file');
    sourcepathSES = strcat(pathname,sourcefile);
    disp(['Template file selected = ' sourcefile]);
    Key = xlsread(sourcepathSES,'Reagent Cf',KeyRange);
    Stock = xlsread(sourcepathSES,'Reagent Cf',StockRange);
    [~,factorNameList,~] = xlsread(sourcepathSES,'Reagent Cf',allListRange);
end

%   Open preformatted source file.
e_file_source = e_process.Workbooks.Open(sourcepathSES);

%   Designate file to save.
fHeader = fileHeader(1:15);
filename = [fHeader '_data.xlsx'];

if run == 1
    foldername = fileHeader;
    mkdir(foldername);
    dataSaveDir = fullfile(pathname,foldername);
end

targetpathSES = strcat(pathname,filename);
A = NaN(1,1);
xlswrite(filename,A);
e_file_target = e_process.Workbooks.Open(targetpathSES); % open target file

if run == 1
    %   Get source sheet/template from source file.
    sheet_source = e_file_source.Sheets.Item('Reagent Cf');

    %   Copy source sheet/template into target workbook. Selecting the
    %   worksheet first is unnecessary and can fail when the target workbook
    %   is the active workbook.
    sheet_source.Copy(e_file_target.Sheets.Item(1));
end

sheet_source = e_file_source.Sheets.Item('Template');
sheet_source.Copy(e_file_target.Sheets.Item(1));
sheetname = ['vectors_gen' num2str(run,'%02i')];
e_file_target.Sheets.Item(1).Name = sheetname; % rename 1st sheet
e_file_target.Save; % save to the same file

%   Delete placeholder worksheets from the new workbook. Avoid relying on
%   localized/default sheet names such as "Sheet1".
if run == 1
    keepSheets = {sheetname, 'Reagent Cf'};
    for i = e_file_target.Worksheets.Count:-1:1
        ws = e_file_target.Worksheets.Item(i);
        if ~any(strcmp(ws.Name, keepSheets))
            ws.Delete;
        end
    end
    e_file_target.Save; % save to the same file
end

e_file_target.Close(false);
e_process.Quit;
