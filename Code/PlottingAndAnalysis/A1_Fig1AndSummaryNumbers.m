clear; clc

%Ritwika VPS, May 2026
%This script plots Fig1 in the main text amd generates summary tables (total numbers by year and by type) + a .txt file that writes some more summary 
% info to file.

BasePath = '~/Desktop/GoogleDriveFiles/research/DARCLEPaper2025/DARCLEPaper2025_ExtractionCode/Data/';
MetadataPath = strcat(BasePath,'A0_MetadataFiles/'); %path to metadata files
OpPath = strcat(BasePath,'A3_AnalysisOutputFiles/'); %path to write processed data files (to save summary numbers etc)

%read in table with full extraction details: pre and posthoc extraction tables combined + all duplicates removed, all retyping done, all non-LF and 
% non-English studies id'd.
DataTab = readtable(strcat(MetadataPath,'SF3_DARCLE_ExtractionList_Unified.xlsx')); 
LFAndEngOnlyCombinedTab = DataTab(~contains(DataTab.Type,'Non-longform') & ...
    ~contains(DataTab.Type,'NonEnglishContent'),:);  %Remove non-LF and non-English papers

uYear = unique(LFAndEngOnlyCombinedTab.PublishedYear); %Get unique years
u_LFonlyType_org = {'Observational','Experimental','Validation','Computational','Review',...
'Commentaries','Best Practice','Other'}'; %organise types
%Check
if ~isequal(sort(u_LFonlyType_org),sort(unique(LFAndEngOnlyCombinedTab.Type)))
    error('User-defined type list and type list from table do not match')
end

%Get the number of papers of each type
clear NumOfType
for i = 1:numel(u_LFonlyType_org)
    NumOfType(i,1) = height(LFAndEngOnlyCombinedTab(contains(LFAndEngOnlyCombinedTab.Type,u_LFonlyType_org{i}),:));
end

%Get totals by year and by type (i.e., a 2d total), as well as totals by year.
for i = 1:numel(uYear)
    NumPerYear(i,1) = height(LFAndEngOnlyCombinedTab(LFAndEngOnlyCombinedTab.PublishedYear == uYear(i),:));
    for j = 1:numel(u_LFonlyType_org)
        NumYrType(i,j) = height(LFAndEngOnlyCombinedTab(LFAndEngOnlyCombinedTab.PublishedYear == uYear(i)...
        & contains(LFAndEngOnlyCombinedTab.Type,u_LFonlyType_org{j}),:));
    end
end

%OkabeItoMap_Rgb = hex2rgb(flip({'#E69F00', '#56B4E9','#009E73','#F0E442','#0072B2','#D55E00','#CC79A7','#000000'}));

%set colours
Clrs = [0 0 146; %DB
0 78 162; %LB
0 146 150; %GB
87 185 134; %G
181 211 105; %LG
254 222 93; %Y
255 160 39; %O
229 82 0]/256; %R

%% Plotting: by year and by type totals etc
figure1 = figure('PaperType','<custom>','PaperSize',[24.5 15.5],'WindowState','maximized','Color',[1 1 1]);


% Create axes: area plot
axes1 = axes('Position',[0.209421112372304 0.124930825826234 0.782633371169126 0.861878453038675]); hold(axes1,'on');
Aplot = area(uYear(3:end),NumYrType(3:end,:));
for i = 1:height(Clrs)
    Aplot(i).EdgeColor = 'none';
    Aplot(i).FaceColor = Clrs(i,:);
end
ylabel({''}); xlabel({'Year of publication'}); %labels
xlim(axes1,[2008.8 2026.2]); ylim(axes1,[0 55]); %axes limits
box(axes1,'on'); hold(axes1,'off');
set(axes1,'FontSize',24,'TickDir','none','XGrid','on','XMinorGrid','on','XMinorTick','on','XTick',...
    [2009 2010 2011 2012 2013 2014 2015 2016 2017 2018 2019 2020 2021 2022 2023 2024 2025 2026],'XTickLabel',...
    {'2009','2010','2011','2012','2013','2014','2015','2016','2017','2018','2019','2020','2021','2022','2023','2024','2025','2026'}, ...
    'XTickLabelRotation',90,'YGrid','on','YMinorGrid','on','YMinorTick','on','YTickLabel',{'','','','','',''}); % Set the remaining axes properties
legend(u_LFonlyType_org,'Direction','normal','Position', ... %legend
    [0.207480526578143 0.544575894587917 0.247354497354497 0.134852216748768],'IconColumnWidth',40,'NumColumns',2) %set direction to normal


% Create axes: bars for earlier years
axes2 = axes('Position',[0.0720771850170262 0.124930825826234 0.064131668558456 0.861878453038676]); hold(axes2,'on');
for i = 1:2 %only the first two years
    b = bar(uYear(i),NumYrType(i,:),'stacked');
    for j = 1:height(Clrs)
        b(j).EdgeColor = 'none';
        b(j).FaceColor = Clrs(j,:);
    end
end
ylabel({'Number of publications'}); %labels
xlim(axes2,[1978.2 1986.5]); ylim(axes2,[0 55]); %axes limits
box(axes2,'on'); hold(axes2,'off');
set(axes2,'FontSize',24,'XGrid','on','XLimitMethod','tight','XTick',[1979 1985],'XTickLabel',{'1979','1985'}, ...
    'XTickLabelRotation',90,'YGrid','on','YLimitMethod','tight','YMinorGrid','on','ZLimitMethod','tight'); % Set the remaining axes properties


% Create axes: grey box
axes3 = axes('Position',[0.130533484676504 0.124930825826234 0.0817253121452894 0.861878453038676]);
ylim(axes3,[0 55]); %limits
box(axes3,'on');
set(axes3,'BoxStyle','full','Color',[0.8 0.8 0.8],'FontSize',24,'XColor',[0 0 0], ...
    'XTick',zeros(1,0),'YColor','none','YTick',zeros(1,0)); % Set the remaining axes properties


% Create axes: inset
axes4 = axes('Position',[0.236095346197503 0.752002649030655 0.207718501702611 0.221915285451198]); hold(axes4,'on');
bar(uYear,NumPerYear,'FaceColor',[0 0 0],'BarWidth',0.6);
xlim(axes4,[1978 2027]); ylim(axes4,[0 56]); %limits
box(axes4,'on'); hold(axes4,'off');
set(axes4,'FontSize',22,'TickDir','none','XLimitMethod','tight','XTick',[1979 1985 2009 2014 2026], ...
    'XTickLabelRotation',90,'YLimitMethod','tight','YTick',[0 15 35 55],'ZLimitMethod','tight'); % Set the remaining axes properties


%Annotations and textboxes
annotation(figure1,'textbox',[0.118355933463056 0.127618036088509 0.105654761904762 0.0474137931034483],'VerticalAlignment','middle',...
    'String',{'(1986-2008)'},'HorizontalAlignment','center','FontSize',24,'EdgeColor','none');
annotation(figure1,'arrow',[0.218501702610669 0.21850170261067],[0.25292345934373 0.193070788993822],'LineWidth',2.5,'HeadStyle','cback2');
annotation(figure1,'arrow',[0.443813847900113 0.443813847900113],[0.330271525642073 0.270418855292165],'LineWidth',2.5,'HeadStyle','cback2');
annotation(figure1,'textbox',[0.206619180494746 0.247882440192693 0.125496031746032 0.083128078817734],'VerticalAlignment','middle',...
    'String',{'LENA released','(2009)'},'LineWidth',0.25,'HorizontalAlignment','center','FontSize',24,'BackgroundColor',[1 1 1]);
annotation(figure1,'textbox',[0.375302817715557 0.32615131680411 0.137731481481481 0.083128078817734],'VerticalAlignment','middle',...
    'String',{'DARCLE formed','(2014)'},'LineWidth',0.25,'HorizontalAlignment','center','FontSize',24,'BackgroundColor',[1 1 1]);



%% Write to output files:
cd(OpPath) 

%Write summary numbers to text file
fileID = fopen('SF4a_ExtractionSummaryNumbers.txt', 'w');  %open file

%Write to file
fprintf(fileID,'Total number of studies included: %i \n',height(LFAndEngOnlyCombinedTab)); %total number of included studies
fprintf(fileID,['Total number of studies excluded (note that these exclusions are only studies that ' ...
    'are identified as non-LF and/or non-English *after/during* manual review/extraction): %i \n'], ...
    height(DataTab)-height(LFAndEngOnlyCombinedTab)); %how many studies were excluded
fprintf(fileID,'Total number of studies excluded as non-LF during manual extraction/review: %i \n', ...
    height(DataTab(contains(DataTab.Type,'Non-longform'),:))); %how many studies were excluded as non-LF
fprintf(fileID,'Total number of studies excluded as non-English during manual extraction/review: %i \n', ...
    height(DataTab(contains(DataTab.Type,'NonEnglishContent'),:))); %how many studies were excluded as non-English
fprintf(fileID,'Earliest year represented is %i and latest year represented is %i \n', ...
    min(LFAndEngOnlyCombinedTab.PublishedYear),max(LFAndEngOnlyCombinedTab.PublishedYear));

%CLOSE!! file
fclose(fileID); 


%Table with included types, total number, and percetage of total
SummaryByTypeTab = table(u_LFonlyType_org,NumOfType,NumOfType/height(LFAndEngOnlyCombinedTab)*100);
SummaryByTypeTab.Properties.VariableNames = {'Type','NumByType','PercentOfTotalByType'};
writetable(SummaryByTypeTab,'SF4b_SummaryNumbersByType.csv')


%Table with by year totals and percetage of total
SummaryByYearTab = table(uYear,NumPerYear,NumPerYear/height(LFAndEngOnlyCombinedTab)*100);
SummaryByYearTab.Properties.VariableNames = {'Year','NumByYear','PercentOfTotalByYear'};
writetable(SummaryByYearTab,'SF4c_SummaryNumbersByYear.csv')

