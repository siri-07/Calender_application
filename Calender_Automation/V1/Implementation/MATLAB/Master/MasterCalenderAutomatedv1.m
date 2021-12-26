%% Import data from spreadsheet
% Script for importing data from the following spreadsheet:
%
%    Workbook: C:\Users\vv972\OneDrive\Documents\MATLAB\Excel case study\Excel Automation\Automation_Sample Calender_v0.1.xlsx
%    Worksheet: Simple
%
% Auto-generated by MATLAB on 19-Jul-2021 13:12:35
clear all;

%% Taking data from GUI
Initiative = "GENESIS";
OutMonth = "January";
%% Set up the Import Options and import the data

opts = spreadsheetImportOptions("NumVariables", 11);

% Specify sheet and range
opts.Sheet = "Sample_GENESIS";
opts.DataRange = "A4:K34";

% Specify column names and types
opts.VariableNames = ["Month", "Date", "Day", "CourseCode", "Module", "Leads", "Leads1", "Leads2", "SessionSlot", "SessionTime", "Comments"];
opts.VariableTypes = ["string", "int8", "string", "string", "string", "string", "string", "string", "string", "string", "string"];

% Specify variable properties
opts = setvaropts(opts, "Month", "WhitespaceRule", "preserve");
opts = setvaropts(opts, ["Month", "Day", "CourseCode", "Module", "Leads", "Leads1", "Leads2", "SessionSlot", "SessionTime", "Comments"], "EmptyFieldRule", "auto");

% Import the data
tbl = readtable("C:\Users\vv972\OneDrive\Documents\MATLAB\Excel case study\Excel_Automation_Test\Automation_Sample Calender_v0.6.xlsx", opts, "UseExcel", false);

%% Assigning data to matlab variables
Month = tbl.Month;
Date = tbl.Date;
Day = tbl.Day;
CourseCode = tbl.CourseCode;
Module = tbl.Module;
Leads = tbl.Leads;
Leads1 = tbl.Leads1;
Leads2 = tbl.Leads2;
SessionSlot = tbl.SessionSlot;
SessionTime = tbl.SessionTime;
Comments = tbl.Comments;

%% Key importing from Master.xlsx Key worksheet
% Auto-generated by MATLAB on 23-Jul-2021 11:34:12

%% Set up the Import Options and import the data
optsKey = spreadsheetImportOptions("NumVariables", 8);

% Specify sheet and range
optsKey.Sheet = "Key";
optsKey.DataRange = "A2:H102";

% Specify column names and types
optsKey.VariableNames = ["FixedInitiativeTitles", "FixedInitiativeCodes", "FixedInitiativeColourCodes", "VarName4", "VarName5", "VarName6", "FixedCourseCodes", "FixedCourseTitles"];
optsKey.VariableTypes = ["string", "double", "string", "string", "string", "string", "string", "string"];

% Specify variable properties
optsKey = setvaropts(optsKey, ["FixedInitiativeTitles", "FixedInitiativeColourCodes", "VarName4", "VarName5", "VarName6", "FixedCourseCodes", "FixedCourseTitles"], "WhitespaceRule", "preserve");
optsKey = setvaropts(optsKey, ["FixedInitiativeTitles", "FixedInitiativeColourCodes", "VarName4", "VarName5", "VarName6", "FixedCourseCodes", "FixedCourseTitles"], "EmptyFieldRule", "auto");

% Import the data
tbKey = readtable("C:\Users\vv972\OneDrive\Documents\MATLAB\Excel case study\Master.xlsx", optsKey, "UseExcel", false);

%% Convert to output type
FixedInitiativeTitles = tbKey.FixedInitiativeTitles;
FixedInitiativeTitles = FixedInitiativeTitles(FixedInitiativeTitles~="");
FixedInitiativeCodes = tbKey.FixedInitiativeCodes;
FixedInitiativeCodes = FixedInitiativeCodes(1:length(FixedInitiativeTitles));
FixedInitiativeColourCodes = tbKey.FixedInitiativeColourCodes;
VarName4 = tbKey.VarName4;
VarName5 = tbKey.VarName5;
VarName6 = tbKey.VarName6;
FixedCourseCodes = tbKey.FixedCourseCodes;
FixedCourseCodes = FixedCourseCodes(FixedCourseCodes~="");
FixedCourseTitles = tbKey.FixedCourseTitles;
FixedCourseTitles = FixedCourseTitles(FixedCourseTitles~="");

%% Fixed courses
FixedCourses = [FixedCourseCodes FixedCourseTitles];

%% Importing the existing OutMonth.xlsx sheetdata from Master.xlsx workbook of fixed size = length(FixedCourseCodes)x 69

[~,UniqueCourseCodeOutMonth,~] = xlsread("C:\Users\vv972\OneDrive\Documents\MATLAB\Excel case study\Master.xlsx", OutMonth, 'A4:A9');
[~,RespectiveCourseTitleOutMonth,~] = xlsread("C:\Users\vv972\OneDrive\Documents\MATLAB\Excel case study\Master.xlsx", OutMonth, 'B4:B9');
[~,RespectiveFacultyOutMonth,~] = xlsread("C:\Users\vv972\OneDrive\Documents\MATLAB\Excel case study\Master.xlsx", OutMonth, 'C4:G9');
[TimeTableOutMonth,~,~] = xlsread("C:\Users\vv972\OneDrive\Documents\MATLAB\Excel case study\Master.xlsx", OutMonth, 'H4:BQ9');
UniqueCourseCodeOutMonthLength = length(UniqueCourseCodeOutMonth);
[~,RespectiveFacultyWidthOutMonth] = size(RespectiveFacultyOutMonth);

%% Fixing the error course codes
for i=1:length(CourseCode)
    if CourseCode(i,1)~=""
        TempFlag=0;
        for j=1:length(FixedCourses)
            if CourseCode(i,1)==FixedCourses(j,1)
                TempFlag=1;
            end
        end
        if TempFlag==0
            TempFlagError = 1;
            for k=1:length(FixedCourses)
                if Module(i)==FixedCourses(k,2)
                    CourseCode(i)=FixedCourses(k,1);
                    TempFlagError = 0;
                end
            end
            if TempFlagError == 1
                CourseCode(i)="";
            end
        end
    end
end

%% Defining different initiatives
FixedInitiatives = [FixedInitiativeTitles FixedInitiativeCodes];

%% Selecting the particular initiative code 
InitiativeCode = 11;
for i=1:length(FixedInitiatives)
    if ismember(Initiative,FixedInitiatives(i,1))
        InitiativeCode = FixedInitiatives(i,2);
    end
end

InitiativeCode = str2double(InitiativeCode);
%% UniqueCourseCode containing unique data for CourseCode
UniqueCourseCode = unique(CourseCode);
UniqueCourseCode = UniqueCourseCode(UniqueCourseCode~="");

%% Data containing the data-wise module names and respective faculties
Data = [Module Leads Leads1 Leads2];

%% Declaring variable to hold respective CourseTitle for UniqueCourseCode
RespectiveCourseTitle = strings(length(UniqueCourseCode),1);

%% Declaring matrix to hold repeatitive list of faculties for respective UniqueCourseCode
Faculty=strings(length(UniqueCourseCode),length(CourseCode)*3);

%% Declaring matrix to hold respective list of faculties for respective UniqueCourseCode
RespectiveFaculty=strings(length(UniqueCourseCode),5);

%% Initialising a TimeTable of zeros for UniqueCourseCode for a month of 31 days
TimeTable=zeros(length(UniqueCourseCode),62);

%% Logically assigning a CourseTitle, Faculty for every UniqueCourseCode
for i=1:length(UniqueCourseCode)
    for j=1:length(CourseCode)        
        if ismember(UniqueCourseCode(i,1),CourseCode(j,1))
            RespectiveCourseTitle(i,1)=Data(j,1);
            Faculty(i,((j-1)*3)+1:((j-1)*3)+3)=Data(j,2:4);
            if SessionSlot(j)=='M'
                TimeTable(i,(2*Date(j)-1))=InitiativeCode;
            elseif SessionSlot(j)=='A'
                TimeTable(i,2*Date(j))=InitiativeCode;
            elseif SessionSlot(j)=='F'
                TimeTable(i,(2*Date(j)-1))=InitiativeCode;
                TimeTable(i,2*Date(j))=InitiativeCode;
            end
        end    
    end    
end


for i=1:length(UniqueCourseCode)
    UniqueFaculty = unique(Faculty(i,:));
    UniqueFaculty = UniqueFaculty(UniqueFaculty~="");
    RespectiveFaculty(i,1:length(UniqueFaculty))=UniqueFaculty(:);
end


%% Writing UniqueCourseCode, RespectiveCourseTitle, RespectiveFaculty and TimeTable onto the Output.xlsx

if isempty(UniqueCourseCodeOutMonth)
    xlswrite('Master.xlsx',UniqueCourseCode,OutMonth,'A4');
    xlswrite('Master.xlsx',RespectiveCourseTitle,OutMonth,'B4');
    xlswrite('Master.xlsx',RespectiveFaculty,OutMonth,'C4');
    xlswrite('Master.xlsx',TimeTable,OutMonth,'H4');
    
    
    %% Colouring the cells in FinalTimeTable
    Excel = actxserver('excel.application');
    % Get Workbook object
    WB = Excel.Workbooks.Open(fullfile(pwd, 'Master.xlsx'),0,false);
    ColumnIndex = [ 'H' , 'I' , 'J' , 'K' , 'L' ,'M' , 'N' , 'O' , 'P' , 'Q' , 'R' , 'S' , 'T' , 'U' , 'W' , 'X' , 'Y' , 'Z' , 'AA' ,'AB' , 'AC' , 'AD' , 'AE' , 'AF' , 'AG' , 'AH' , 'AI' , 'AJ' , 'AK' , 'AL' , 'AM' , 'AN' , 'AO' , 'AP' , 'AQ' , 'AR' , 'AS' , 'AT' , 'AU' , 'AW' , 'AX' , 'AY' , 'AZ' , 'BA' ,'BB' , 'BC' , 'BD' , 'BE' , 'BF' , 'BG' , 'BH' , 'BI' , 'BJ' , 'BK' , 'BL' ,'BM' , 'BN' , 'BO' , 'BP' , 'BQ' ];
    FinalRow = length(UniqueCourseCode) + 4;
    for i = 4:FinalRow
        for j = 1:length(ColumnIndex)
            TempCourseCode = xlsread('Master.xlsx',OutMonth,[ColumnIndex(j),num2str(i)]);
            if TempCourseCode == 1
                WB.Worksheets.Item(OutMonth).Range([ColumnIndex(j),num2str(i)]).Interior.ColorIndex = 3;
            elseif TempCourseCode == 2
                WB.Worksheets.Item(OutMonth).Range([ColumnIndex(j),num2str(i)]).Interior.ColorIndex = 3;
            elseif TempCourseCode == 3
                WB.Worksheets.Item(OutMonth).Range([ColumnIndex(j),num2str(i)]).Interior.ColorIndex = 3;
            elseif TempCourseCode == 4
                WB.Worksheets.Item(OutMonth).Range([ColumnIndex(j),num2str(i)]).Interior.ColorIndex = 3;
            elseif TempCourseCode == 5
                WB.Worksheets.Item(OutMonth).Range([ColumnIndex(j),num2str(i)]).Interior.ColorIndex = 3;
            elseif TempCourseCode == 6
                WB.Worksheets.Item(OutMonth).Range([ColumnIndex(j),num2str(i)]).Interior.ColorIndex = 3;
            end
        end
    end

    % Save Workbook
    WB.Save();
    % Close Workbook
    WB.Close();
    % Quit Excel
    Excel.Quit();
else
    %Initialising the final outputs of UniqueCourseCodesFinal, RespectiveCourseTitleFinal, RespectiveFacultyFinal
    TempCounterFinal = 0;
    for i=1:length(UniqueCourseCode)
        for j=1:length(UniqueCourseCodeOutMonth)
            if ismember(UniqueCourseCode(i,1),UniqueCourseCode(j,1))
                TempCounterFinal= TempCounterFinal + 1;
                break;
            end
        end
    end
    FinalLength = length(UniqueCourseCodeOutMonth) + length(UniqueCourseCode) - TempCounterFinal;    
    
    FinalUniqueCourseCode = strings(FinalLength,1);
    FinalUniqueCourseCode(1:length(UniqueCourseCodeOutMonth),1)= UniqueCourseCodeOutMonth(:,1);
    
    FinalRespectiveCourseTitle = strings(FinalLength,1);
    FinalRespectiveCourseTitle(1:length(UniqueCourseCodeOutMonth),1)= RespectiveCourseTitleOutMonth(:,1);
    
    FinalRespectiveFaculty = strings(FinalLength,5);
    FinalRespectiveFaculty(1:length(UniqueCourseCodeOutMonth),1:RespectiveFacultyWidthOutMonth) = RespectiveFacultyOutMonth;

    FinalTimeTable = zeros(FinalLength,62);
    FinalTimeTable(1:length(UniqueCourseCodeOutMonth),:) = TimeTableOutMonth;
    
    TempCounterCourse = 1;
    TempCounterFaculty = 1;
    for i=1:length(UniqueCourseCode)
        TempFlagCourse=0;
        for j=1:length(FinalUniqueCourseCode)
            if ismember(UniqueCourseCode(i,1),FinalUniqueCourseCode(j,1))
                TempFlagCourse = 1;
                TempRow = j;
            end
        end
        if TempFlagCourse == 1
            RespectiveFacultyLength = unique(RespectiveFaculty(i,:));
            RespectiveFacultyLength = length(RespectiveFacultyLength(RespectiveFacultyLength~=""));
            FinalRespectiveFacultyLength = unique(FinalRespectiveFaculty(TempRow,:));
            FinalRespectiveFacultyLength = length(FinalRespectiveFacultyLength(FinalRespectiveFacultyLength~=""));        
        
            for x=1:RespectiveFacultyLength
                TempFlagFaculty = 0;
                for y=1:FinalRespectiveFacultyLength
                    if ismember(RespectiveFaculty(i,x),FinalRespectiveFaculty(TempRow,y))
                        TempFlagFaculty = 1;
                    end
                end
                if TempFlagFaculty == 0
                    FinalRespectiveFaculty(TempRow,FinalRespectiveFacultyLength+TempCounterFaculty) = RespectiveFaculty(i,x);
                    TempCounterFaculty = TempCounterFaculty +1;
                end
                if (TempCounterFaculty + FinalRespectiveFacultyLength) > 5
                    TempCounterFaculty = 1;
                    break;
                end
            end
            for a=1:62
                if TimeTable(i,a) == InitiativeCode
                    FinalTimeTable(TempRow,a) = TimeTable(i,a);
                end
            end
        elseif TempFlagCourse == 0
            TempRowFinal = length(UniqueCourseCodeOutMonth)+ TempCounterCourse;
            FinalUniqueCourseCode(TempRowFinal,1) = UniqueCourseCode(i,1);
            FinalRespectiveCourseTitle(TempRowFinal,1) = RespectiveCourseTitle(i,1);
            FinalRespectiveFaculty(TempRowFinal,:) = RespectiveFaculty(i,:);
            FinalTimeTable(TempRowFinal,:)=TimeTable(i,:);
            TempCounterCourse = TempCounterCourse + 1;
            if TempCounterCourse > length(FixedCourseCodes)
                break;
            end
        end
    end
    xlswrite('Master.xlsx',FinalUniqueCourseCode,OutMonth,'A4');
    xlswrite('Master.xlsx',FinalRespectiveCourseTitle,OutMonth,'B4');
    xlswrite('Master.xlsx',FinalRespectiveFaculty,OutMonth,'C4');
    xlswrite('Master.xlsx',FinalTimeTable,OutMonth,'H4');
    
    
    %Colouring the cells in FinalTimeTable

    Excel = actxserver('excel.application');
    % Get Workbook object
    WB = Excel.Workbooks.Open(fullfile(pwd, 'Master.xlsx'),0,false);
    ColumnIndex = [ 'H' , 'I' , 'J' , 'K' , 'L' ,'M' , 'N' , 'O' , 'P' , 'Q' , 'R' , 'S' , 'T' , 'U' , 'W' , 'X' , 'Y' , 'Z' , 'AA' ,'AB' , 'AC' , 'AD' , 'AE' , 'AF' , 'AG' , 'AH' , 'AI' , 'AJ' , 'AK' , 'AL' , 'AM' , 'AN' , 'AO' , 'AP' , 'AQ' , 'AR' , 'AS' , 'AT' , 'AU' , 'AW' , 'AX' , 'AY' , 'AZ' , 'BA' ,'BB' , 'BC' , 'BD' , 'BE' , 'BF' , 'BG' , 'BH' , 'BI' , 'BJ' , 'BK' , 'BL' ,'BM' , 'BN' , 'BO' , 'BP' , 'BQ' ];
    FinalRow = FinalLength + 4;
    for i = 4:FinalRow
        for j = 1:length(ColumnIndex)
            TempCourseCode = xlsread('Master.xlsx',OutMonth,[ColumnIndex(j),num2str(i)]);
            if TempCourseCode == 1
                WB.Worksheets.Item(OutMonth).Range([ColumnIndex(j),num2str(i)]).Interior.ColorIndex = 3;
            elseif TempCourseCode == 2
                WB.Worksheets.Item(OutMonth).Range([ColumnIndex(j),num2str(i)]).Interior.ColorIndex = 3;
            elseif TempCourseCode == 3
                WB.Worksheets.Item(OutMonth).Range([ColumnIndex(j),num2str(i)]).Interior.ColorIndex = 3;
            elseif TempCourseCode == 4
                WB.Worksheets.Item(OutMonth).Range([ColumnIndex(j),num2str(i)]).Interior.ColorIndex = 3;
            elseif TempCourseCode == 5
                WB.Worksheets.Item(OutMonth).Range([ColumnIndex(j),num2str(i)]).Interior.ColorIndex = 3;
            elseif TempCourseCode == 6
                WB.Worksheets.Item(OutMonth).Range([ColumnIndex(j),num2str(i)]).Interior.ColorIndex = 3;
            end
        end
    end

    % Save Workbook
    WB.Save();
    % Close Workbook
    WB.Close();
    % Quit Excel
    Excel.Quit();

end
%% Rows not analysed

RowsNotAnalysed = zeros(length(CourseCode(CourseCode=="")));
TempRowsCounter = 1;
 for i=1:length(CourseCode)
     if CourseCode(i,1)==""
         RowsNotAnalysed(TempRowsCounter)=Date(i);
         TempRowsCounter = TempRowsCounter + 1;
     end
 end
 RowsNotAnalysed=RowsNotAnalysed(RowsNotAnalysed ~=0);
 
fig = uifigure;
uialert(fig,sprintf('Date %g was not analysed\n',RowsNotAnalysed),'Warning');
 
%% Clear temporary variables
clear opts tbl