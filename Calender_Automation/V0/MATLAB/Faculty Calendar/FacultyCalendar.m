opts = spreadsheetImportOptions("NumVariables", 11);

% Specify sheet and range 
opts.Sheet = "Sample_GENESIS";
opts.DataRange = "A4:K100";

% Specify column names and types
opts.VariableNames = ["VarName1", "Date", "Day", "CourseCode", "Module", "Leads", "Leads1", "Leads2", "SessionSlot", "SessionTime", "Comments"];
opts.VariableTypes = ["string", "double", "categorical", "categorical", "categorical", "categorical", "categorical", "categorical", "categorical", "categorical", "categorical"];

% Specify variable properties
opts = setvaropts(opts, "VarName1", "WhitespaceRule", "preserve");
opts = setvaropts(opts, ["VarName1", "Day", "CourseCode", "Module", "Leads", "Leads1", "Leads2", "SessionSlot", "SessionTime", "Comments"], "EmptyFieldRule", "auto");

% Import the data
Automation_SampleCalender_v0_1 = readtable('Automation_Sample Calender_v0.6.xlsx', opts, "UseExcel", false);
%% Convert to output type

VarName1 = Automation_SampleCalender_v0_1.VarName1;
Date = Automation_SampleCalender_v0_1.Date;
Day = Automation_SampleCalender_v0_1.Day;
CourseCode = Automation_SampleCalender_v0_1.CourseCode;
Module = Automation_SampleCalender_v0_1.Module;
Leads = Automation_SampleCalender_v0_1.Leads;
Leads1 = Automation_SampleCalender_v0_1.Leads1;
Leads2 = Automation_SampleCalender_v0_1.Leads2;
SessionSlot = Automation_SampleCalender_v0_1.SessionSlot;
SessionTime = Automation_SampleCalender_v0_1.SessionTime;
Comments = Automation_SampleCalender_v0_1.Comments;
%% Clear temporary variables

clear Automation_SampleCalender_v0_1 opts
TogetherArray = cat(1,Leads,Leads1,Leads2);
FacultyList = unique(TogetherArray);
FacultyList = rmmissing(FacultyList);
FacultyList = upper(FacultyList);
Month = "July";
Month = upper(Month);
Initiative = "GENESIS";
FacultyMSlots = zeros(length(FacultyList),31);
FacultyASlots = zeros(length(FacultyList),31);
for i=1:length(Leads)
    for j = 1:length(FacultyList)
        if(Leads(i)==FacultyList(j) || Leads1(i)==FacultyList(j) || Leads2(i)==FacultyList(j) )
            if(SessionSlot(i)=='M')
                FacultyMSlots(j,Date(i)) = FacultyMSlots(j,Date(i)) + 1;
            elseif(SessionSlot(i)=='A')
                FacultyASlots(j,Date(i)) = FacultyASlots(j,Date(i)) + 1;
            elseif(SessionSlot(i)=='F')
                FacultyMSlots(j,Date(i)) = FacultyMSlots(j,Date(i)) + 1;
                FacultyASlots(j,Date(i)) = FacultyASlots(j,Date(i)) + 1;
            end
        end
    end
end
for i=1:length(FacultyList)
    for j=1:31
        if(Initiative=="GENESIS PRO")&(FacultyMSlots(i,j)==1)
            FacultyMSlots(i,j) = 2;
        elseif(Initiative=="BUILD / STEP UP")&(FacultyMSlots(i,j)==1)
            FacultyMSlots(i,j) = 3;
        elseif(Initiative=="OPEN TRAININGS")&(FacultyMSlots(i,j)==1)
            FacultyMSlots(i,j) = 4;
        elseif(Initiative=="STEPin")&(FacultyMSlots(i,j)==1)
            FacultyMSlots(i,j) = 5;
        elseif(Initiative=="TEECH PRAAAVINYA")&(FacultyMSlots(i,j)==1)
            FacultyMSlots(i,j) = 6;
        elseif(Initiative=="OTHERS")&(FacultyMSlots(i,j)==1)
            FacultyMSlots(i,j) = 7;
        end
        if(Initiative=="GENESIS PRO")&(FacultyASlots(i,j)==1)
            FacultyASlots(i,j) = 2;
        elseif(Initiative=="BUILD / STEP UP")&(FacultyASlots(i,j)==1)
            FacultyASlots(i,j) = 3;
        elseif(Initiative=="OPEN TRAININGS")&(FacultyASlots(i,j)==1)
            FacultyASlots(i,j) = 4;
        elseif(Initiative=="STEPin")&(FacultyASlots(i,j)==1)
            FacultyASlots(i,j) = 5;
        elseif(Initiative=="TEECH PRAAAVINYA")&(FacultyASlots(i,j)==1)
            FacultyASlots(i,j) = 6;
        elseif(Initiative=="OTHERS")&(FacultyASlots(i,j)==1)
            FacultyASlots(i,j) = 7;
        end
    end
end

Months =            ["JANUARY","FEBRUARY","MARCH","APRIL","MAY","JUNE","JULY","AUGUST","SEPTEMBER","OCTOBER","NOVEMBER","DECEMBER"];
Range = [5,24,31,50,57,76,83,102,109,128,135,154,161,180,187,206,213,232,239,258,265,284,291,310]; 
FacultyRange = ["B5:B24","B31:50","B57:B76","B83:B102","B109:B128","B135:B154","B161:B180","B187:B206","B213:B232","B238:B258","B265:B284","B291:B310"];
MAlpha = ['A','C','E','G','I','K','M','O','Q','S','U','W','Y'];
AAlpha = ['B','D','F','H','J','L','N','P','R','T','V','X','Z'];
for i=1:12
    if(Month == Months(i))
        index=i;
    end
end
[~,FacultyDown] = xlsread('FacultyCalendar_Output.xlsx',1,FacultyRange(index));
FacLen = length(FacultyDown);
ExMSlots = zeros(FacLen,31);
ExASlots = zeros(FacLen,31);
for i=1:FacLen
    for j=1:12
        ExMSlots(i,j) = table2array(readtable('FacultyCalendar_Output.xlsx','Range',[MAlpha(j+1),num2str(Range(index*2-1)+i-1),':',MAlpha(j+1),num2str(Range(index*2-1)+i-1)]));
        ExASlots(i,j) = table2array(readtable('FacultyCalendar_Output.xlsx','Range',[AAlpha(j+1),num2str(Range(index*2-1)+i-1),':',AAlpha(j+1),num2str(Range(index*2-1)+i-1)]));
    end
    for j=13:25
        ExMSlots(i,j) = table2array(readtable('FacultyCalendar_Output.xlsx','Range',['A',MAlpha(j-12),num2str(Range(index*2-1)+i-1),':','A',MAlpha(j-12),num2str(Range(index*2-1)+i-1)]));
        ExASlots(i,j) = table2array(readtable('FacultyCalendar_Output.xlsx','Range',['A',AAlpha(j-12),num2str(Range(index*2-1)+i-1),':','A',AAlpha(j-12),num2str(Range(index*2-1)+i-1)]));
    end
    for j=26:31
        ExMSlots(i,j) = table2array(readtable('FacultyCalendar_Output.xlsx','Range',['B',MAlpha(j-25),num2str(Range(index*2-1)+i-1),':','B',MAlpha(j-25),num2str(Range(index*2-1)+i-1)]));
        ExASlots(i,j) = table2array(readtable('FacultyCalendar_Output.xlsx','Range',['B',AAlpha(j-25),num2str(Range(index*2-1)+i-1),':','B',AAlpha(j-25),num2str(Range(index*2-1)+i-1)]));
    end
end
ExMSlots(isnan(ExMSlots)) = 0;
ExASlots(isnan(ExASlots)) = 0;
UpFaculty = strings(FacLen+length(FacultyList),1);
UpMSlots = zeros(FacLen+length(FacultyList),31);
UpASlots = zeros(FacLen+length(FacultyList),31);
kc=1;
CheckList = zeros(length(FacultyList),1);
CheckList1 = zeros(FacLen,1);
for i=1:FacLen
    for j=1:length(FacultyList)
        if(FacultyDown(i)==FacultyList(j))
            CheckList(j)=1;
            CheckList1(i)=1;
            for k=1:31
                if(ExMSlots(i,k) ~= FacultyMSlots(j,k))
                    if(ExMSlots(i,k)==0)
                        ExMSlots(i,k) = FacultyMSlots(j,k);
                    else
                        ExMSlots(i,k) = (FacultyMSlots(j,k)*10) + (ExMSlots(i,k));
                    end
                end
                if(ExASlots(i,k) ~= FacultyASlots(j,k))
                    if(ExASlots(i,k)==0)
                        ExASlots(i,k) = FacultyASlots(j,k);
                    else
                        ExASlots(i,k) = (FacultyASlots(j,k)*10) + (ExASlots(i,k));
                    end
                end
            end
            UpFaculty(kc) = FacultyDown(i);
            UpMSlots(kc,:) = ExMSlots(i,:);
            UpASlots(kc,:) = ExASlots(i,:);
            kc = kc+1;
            i = i+1;
            break;
        end
    end
end
for i=1:FacLen
    if(CheckList1(i)==0)
        UpFaculty(kc) = FacultyDown(i);
        UpMSlots(kc,:) = ExMSlots(i,:);
        UpASlots(kc,:) = ExASlots(i,:);
        kc = kc+1;
    end
end
for i=1:length(FacultyList)
    if(CheckList(i)==0)
        UpFaculty(kc) = FacultyList(i);
        UpMSlots(kc,:) = FacultyMSlots(i,:);
        UpASlots(kc,:) = FacultyASlots(i,:);
        kc = kc+1;
    end
end
for i=1:kc-1
    writematrix(UpFaculty(i),'FacultyCalendar_Output.xlsx','Sheet',1,'Range',['B',num2str(Range(index*2-1)+i-1)])
    for j=1:12
        writematrix(UpMSlots(i,j),'FacultyCalendar_Output.xlsx','Sheet',1,'Range',[MAlpha(j+1),num2str(Range(index*2-1)+i-1)]);
        writematrix(UpASlots(i,j),'FacultyCalendar_Output.xlsx','Sheet',1,'Range',[AAlpha(j+1),num2str(Range(index*2-1)+i-1)]);
    end
    for j=13:25
        writematrix(UpMSlots(i,j),'FacultyCalendar_Output.xlsx','Sheet',1,'Range',['A',MAlpha(j-12),num2str(Range(index*2-1)+i-1)]);
        writematrix(UpASlots(i,j),'FacultyCalendar_Output.xlsx','Sheet',1,'Range',['A',AAlpha(j-12),num2str(Range(index*2-1)+i-1)]);
    end
    for j=26:31
        writematrix(UpMSlots(i,j),'FacultyCalendar_Output.xlsx','Sheet',1,'Range',['B',MAlpha(j-25),num2str(Range(index*2-1)+i-1)]);
        writematrix(UpMSlots(i,j),'FacultyCalendar_Output.xlsx','Sheet',1,'Range',['B',AAlpha(j-25),num2str(Range(index*2-1)+i-1)]);
    end
end
