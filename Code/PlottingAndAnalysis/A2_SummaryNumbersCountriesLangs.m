clear; clc

% Ritwika VPS, May 2026
% This script performs more checks on (atuomatically and manually, as necessary) extracted + reviewed files, gets summary numbers for countries 
% and languages, and writes associated outputs to file as appropriate.

%Paths
BasePath = '~/Desktop/GoogleDriveFiles/research/DARCLEPaper2025/DARCLEPaper2025_ExtractionCode/Data/';
MetadataPath = strcat(BasePath,'A0_MetadataFiles/'); %path to metadata files
ExtractedFilesPath = strcat(BasePath,'A2_DataForAnalysis_PostAutoAndManualExtraction_2026_04/'); %path to extracted files
OpPath = strcat(BasePath,'A3_AnalysisOutputFiles/'); %path to write processed data files (to save summary numbers etc)

%read in table with full extraction metadata: pre and posthoc extraction tables combined + all duplicates removed, all retyping done, all non-LF and 
% non-English studies id'd.
DataTab = readtable(strcat(MetadataPath,'SF3_DARCLE_ExtractionList_Unified.xlsx')); 
DataTab_ObsExpOnly = DataTab(contains(DataTab.Type,{'Observational','Experimental'},'IgnoreCase',true),:);  %Only obs and experimental

%% 1. Get extracted data 
cd(ExtractedFilesPath)

%---1a. Data from manually extracted sheets-----------------
% First we get the list of manually extracted tables (dir()). Then, the second posthoc table is excluded because there was no manual extraction
% associated with teh second posthoc round. The sheets in these tables correspond to an extraction category (e.g., recording device, software, etc.),
% lists of papers that were not automatically extracted for *all* categories except community/consortia (see extraction code for details), and
% lists of papers that were tagged to be excluded from extraction because they were identified as duplicates or not experimental/obs, etc during
% manual review. For the prehoc and first posthoc tables, each sheet is read in for the required subset of columns and combined into a single table.
NotAutoExtractedPapers_Dir = dir('ListOfNotExtractedPapers_ExceptCommsAndConsortia*'); %Reading manually extracted sheet: there should be three.
%CHECK!
if numel(NotAutoExtractedPapers_Dir) ~= 3
    error('There should be 3 sheets of not extracted papers (1 pre-hoc and 2 post-hoc)')
end

%Remove table from posthoc round2,'ListOfNotExtractedPapers_ExceptCommsAndConsortia_PosthocRd2.xlsx' (empty table). 
NotAutoExtractedPapers_Dir = NotAutoExtractedPapers_Dir(~contains({NotAutoExtractedPapers_Dir.name}, ...
                                                                   'ListOfNotExtractedPapers_ExceptCommsAndConsortia_PosthocRd2.xlsx'));
%Get sheets for excel files for manual review/extraction corresponding to prehoc and first round of posthoc extraction.
NotAutoExtractedTabs_SheetList = {'Papers_to_Exclude','Main_Generated','RecDevice','Software','CountriesAndLangs','ManualAnnot','Topics'}; %names of sheets
%For each sheet, get the list of cols to read in (to pass into the user-defined fn below)
NotAutoExtractedTabs_SheetsColsToRead = {{'PaperTitle_ToC','PaperTitle_NoToC','FileName'}, %Papers_to_Exclude
                                          {'PaperTitle_ToC','PaperTitle_NoToC','FileName'}, %Main_Generated
                                          {'FileName','LENA','Camera','RecOrMic'}, %RecDevice
                                          {'FileName','LENA','Praat','OpenSMILE','ALICE','SpeechDetAlg'}, %Software
                                          {'FileName','CountriesProcessed','LanguagesProcessed'}, %CountriesAndLangs
                                          {'FileName','ManualAnnot'}, %ManualAnnot
                                          {'FileName','CDSvsADS','VocMaturity','Multiling',...
                                                'CommDisorder','SES','SchoolProg','Multimodal','StressAnxEmo','CryNonCry','NICU'}}; %Topics
%For each item in the sheets list, read through the not-auto-extracted tables and combine. So, this would, for eg, read the 'Papers_To_Exclude' sheet from 
% all tables listing papers that weren't automatically extracted (i.e., the pre-hoc round and the first post-hoc round, cuz the second posthoc extraction
% had no papers that could not be automatically extracted), and combine those sheets (for only the specified columns) into a single table. And that table
% is then stored in a cell array.
for i = 1:numel(NotAutoExtractedTabs_SheetList)
    NotAutoExtractedTabs_FullTab{i} = CombineTabs(NotAutoExtractedPapers_Dir,NotAutoExtractedTabs_SheetList{i},NotAutoExtractedTabs_SheetsColsToRead{i});
end

%Rename the countries and languages table (so we can merge with the automatically extracted table)
if isequal(sort(NotAutoExtractedTabs_FullTab{5}.Properties.VariableNames), ...
        sort({'FileName','CountriesProcessed','LanguagesProcessed'})) %check that we have the correct table
    NotAutoExtractedTabs_FullTab{5}.Properties.VariableNames = {'FileName','Countries','Languages'};
else
    error('Attempting to rename the wrong table')
end

%---1b. Data from automatically extracted sheets-----------------
% Get the categories (each category is its own set of extracted sheets for each round (pre- and two post-hocs) of extraction) that were extracted. These
% strings will be used to dir() automatically extracted spreadsheets for that category and combine spreadsheets at the category level into a single table.
AutoExtractedPapers_Categories = {'RecDevice_ExtractedSents*','Software_ExtractedSents*','CountriesLangsStatesUS_ExtractedSents*',...
                                  'ManualAnnot_ExtractedSents*','Topics_ExtractedBinary*','CommunityConsortia_ExtractedBinary*'};
AutoExtractedPapers_ByCategCols = {{'FileName','LENA','Camera','RecOrMic'}, %RecDevice
                                    {'FileName','LENA','Praat','OpenSMILE','ALICE','SpeechDetAlg'}, %Software
                                    {'FileName','Countries','Languages'}, %CountriesAndLangs
                                    {'FileName','ManualAnnot'}, %ManualAnnnot
                                    {'FileName','CDSvsADS','VocMaturity','Multiling',...
                                                'CommDisorder','SES','SchoolProg','Multimodal','StressAnxEmo','CryNonCry','NICU'}, %Topics
                                    {'FileName','DARCLE','ACLEW','Homebank','Talkbank','ALICE','VTC','LangView','MPaL','ManyBabies'}}; %Communities/consortia
for i = 1:numel(AutoExtractedPapers_Categories)
    CurrCateg_Dir = dir(AutoExtractedPapers_Categories{i}); %get list of tables
    %CHECK
    if numel(CurrCateg_Dir) ~= 3
        error('There should be 3 spreadsheets (1 pre-hoc and 2 post-hoc)')
    end
    AutoExtractedTabs_FullTab{i} = CombineTabs(CurrCateg_Dir,'Sheet1',AutoExtractedPapers_ByCategCols{i});
end

%---1c. Checks Pt 1 ---
% Check that the total number of obs + experimental papers (from the final, re-typed, duplicate removed etc combined extraction sheet) matches the 
% number of papers listed in the automatic extraction table for that category (even if required manual review) + number of papers that could
% not be manually extracted - minus papers that were tagged as needing to be excluded during manual review/extraction (resulting in the final combined
% extraction sheet).

%Topics are only extracted automatically, so the number of entries in the combined topics spreadhseet minus the number of papers tagged to be excluded
% in the combined Papers_To_Exclude sheet from the manual extraction spreadsheets should be equal to the finalied number of observational and expt studies.
% This is because the final count is determined after id-ing duplicates and re-typed papers, and that info is present in the 'Papers_To_Exclude' combined
% sheet.
if (height(AutoExtractedTabs_FullTab{end}) - height(NotAutoExtractedTabs_FullTab{1})) ~= height(DataTab_ObsExpOnly)
    error('Topics numbers do not match (Checks Pt1')
end

%All other categories except topics
ManualInds = 3:7; AutoInds = 1:5; %indices for recdevice, software, countries and langs, manual annotations, and topics
for i = 1:numel(AutoInds)  %check for each category
    if ((height(AutoExtractedTabs_FullTab{AutoInds(i)}) + height(NotAutoExtractedTabs_FullTab{2}) - height(NotAutoExtractedTabs_FullTab{1})) ...
            ~= height(DataTab_ObsExpOnly))
        i
        error(['Final number of obs + expt papers not equal to number of automatically extracted (with or without needing further review) paper' ...
            '+ number of papers that could not be automatically extracted at all - number of papers that are excluded papers due to dupes and/or re-typing'])
        %NotAutoExtractedTabs_FullTab{2} is the list of papers that could not be automatically extracted at all
        %NotAutoExtractedTabs_FullTab{1} is the list of papers that are flagged for exclusion due to dupes id-d and/or re-typing.
    end
end


%---1d. Checks Pt 2 (more fine-grained checks) ---
% For each category (except topics, because topics are entirely extracted automatically), check that the final number of expt + obs papers is
% equal to the number of automatically extracted papers + number of manually extracted/reviewed studies - number of studies tagged for
% exclusion - number of studies in manually extracted/reviewed studies that replaces the corresponding automatically extracted entry (due to being
% manually reviewed and therefore, being present in the category-level manual review/extraction).
% Aside from studies that could not be automatically extracted at all, for some categories, some studies required manual review/extraction, even if
% they were present in the automatic extraction sheet. For these studies, the manual review/extraction entry is intended to replace the automatic 
% extraction one if it exists.



%SAVE fully processed combined tables with exclusions excluded (incl for
%topics)





%% Write outputs to file
cd(OpPath) 

%Write summary numbers to text file
fileID = fopen('S5_SummaryNumbersFromExtractedData.txt', 'w');  %open file

%Write to file:
%total number of included studies (after accounting for exclusions)
fprintf(fileID,'Total number of studies included (after accounting for exclusions; see next line): %i \n',height(DataTab_ObsExpOnly)); 
%Total number of studies excluded during/after manual extraction and review due to re-typing and id-ing dupes (including id-ing dupes through custom code).
fprintf(fileID,['The total number of studies excluded during/after the manual extraction and review stage' ... 
    'due to identifying dupes (including id-ing dupes through custom code) + study type ' ...
    'being changed after review (re-typing) is: %i \n'],height(NotAutoExtractedTabs_FullTab{1})); 
fprintf(fileID,['Topics were automatically extracted from ALL pdfs initially tagged as experimental or observational. ' ...
    'Accounting for exculded studies (see prev. line), number of studies topics were extracted from: %i \n'], ...
    height(AutoExtractedTabs_FullTab{end}) - height(NotAutoExtractedTabs_FullTab{1})); %Number of studies topics were extracted from (accounting for exclusions)


%CLOSE!! file
fclose(fileID); 




%------------------------------------------------------------------------------------------------------------------------------------
%% FUNCTIONS USED
%------------------------------------------------------------------------------------------------------------------------------------

%This function takes in a structure outputed from a dir() (IpDir; lists extracted spreadsheets for a given category or from manual extraction), 
% the sheet within the spreadsheet to extract data from (SheetName), and the column names to be read in (ColsToRead).
% Outputs a concatenated tab (FullTab) for all tables in IpDir.
%
% So, for instance, dir(ListOfNotExtractedPapers_ExceptCommsAndConsortia*) would return an IpDir with the 3 .xlsx files with info for manual extraction.
% Note that we exclude the excel file corresponding to the second posthoc round for 'ListOfNotExtractedPapers_ExceptCommsAndConsortia'.
% Now, these excel files have various sheets, which can be accessed using SheetName and the columns to be read in can be specified using ColsToRead.
% Then, this function will concatenate (vertically) tables corresponding to, say, SheetName = 'Papers_to_Exclude' and ColsToRead for the prehoc
% and the first posthoc extractions. 
% 
% For the automatically extracted tables, there is only one sheet (Sheet1), so the SheetName is fixed. So, say, for IpDir = dir(RecDevice_ExtractedSents*),
% thsi function will concatenate tables for the automatically extracted recdevice spreadsheets based on ColsToRead.
function [FullTab] = CombineTabs(IpDir,SheetName,ColsToRead)

    FullTab = readtable(IpDir(1).name,'Sheet',SheetName); %read in first spreadsheet in IpDir based on SheetName
    FullTab = FullTab(:,ColsToRead); %only keep ColsToRead

    for i = 2:numel(IpDir) %go through the rest of IpDir
        if ~isempty(readtable(IpDir(i).name)) %check to make sure that table is not empty
            tempTab = readtable(IpDir(i).name,'Sheet',SheetName); %read in table based on SheetName
            tempTab = tempTab(:,ColsToRead); %keep ColsToRead
            FullTab = [FullTab; tempTab]; %Add to output table
        end
    end
end

%-------------------------------------------------------------------------------------------------
function [] = CheckNums_ObsExpt_vs_ExtractedStudies()

end


