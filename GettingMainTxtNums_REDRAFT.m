clear; clc

%reading in extraction sheets and getting obs/expyt only + combining
PreHoc = readtable('DARCLE_extraction_2.16.26.xlsx','Sheet','Corrected_ReTyped_PostManualExt');
PostHoc = readtable('Post-hoc Paper Extractions_3.23.26.xlsx','Sheet','Corrected_ReTyped_PostManualExt');
PreHoc_ObsExp = PreHoc(contains(PreHoc.Type,{'Observational','Experimental'},'IgnoreCase',true),:);
PostHoc_ObsExp = PostHoc(contains(PostHoc.Type,{'Observational','Experimental'},'IgnoreCase',true),:);
ExtracSheet_ObsExp_Combined = [PreHoc_ObsExp; PostHoc_ObsExp];

%Reading manually extracted sheet: list of papers manually extracted + list
%of papers to be excluded
NotExtractedPapersList = dir('ListOfNotExtractedPapers_ExceptCommsAndConsortia*'); %there should be three
PapersToExclude_Full = readtable(NotExtractedPapersList(1).name,'Sheet','Papers_to_Exclude');
PapersToExclude_Full.PaperTitle_NoToC_Updated = []; %This column is empty in 1??
ManualExtracFull = readtable(NotExtractedPapersList(1).name,'Sheet','Main_Generated');
for i = 2:numel(NotExtractedPapersList)
    if ~isempty(readtable(NotExtractedPapersList(i).name)) %the third one is empty cuz it's posthoc round 2
        tempExcTab = readtable(NotExtractedPapersList(i).name,'Sheet','Papers_to_Exclude');
        tempExcTab.PaperTitle_NoToC_Updated = [];
        PapersToExclude_Full = [PapersToExclude_Full; tempExcTab];
        ManualExtracFull = [ManualExtracFull; readtable(NotExtractedPapersList(i).name,'Sheet','Main_Generated')];
    end
end

%file names to exclude from the full list of papers to exclude
FnamesToExclude = PapersToExclude_Full.FileName;

%Automatically extrcted countries and languages
CountryLangExtractedList = dir('CountriesLangsStatesUS_ExtractedSents*');
CountryLangAutoFull = readtable(CountryLangExtractedList(1).name);
CountryLangAutoFull.StatesUS_sents = [];
for i = 2:numel(CountryLangExtractedList)
    TempTab = readtable(CountryLangExtractedList(i).name);
    TempTab.StatesUS_sents = [];
    CountryLangAutoFull = [CountryLangAutoFull; TempTab];
end

%checks that numbers match
height(ExtracSheet_ObsExp_Combined)
height(CountryLangAutoFull) + height(ManualExtracFull) - height(PapersToExclude_Full)

%manually extracted countries and langs
CountryLangManualFull = readtable(NotExtractedPapersList(1).name,'Sheet','CountriesAndLangs');
CountryLangManualFull.Notes = [];
for i = 2:numel(NotExtractedPapersList)
    if ~isempty(readtable(NotExtractedPapersList(i).name)) %the third one is empty cuz it's posthoc round 2
        tempExcTab = readtable(NotExtractedPapersList(i).name,'Sheet','CountriesAndLangs');
        tempExcTab.Notes = [];
        CountryLangManualFull = [CountryLangManualFull; tempExcTab];
    end
end

%auto and manual extracted countries and langs after excluding papers that need exclusion
CountryLangAutoFull_w_ReqFilesExcluded = CountryLangAutoFull(~contains(CountryLangAutoFull.FileName,FnamesToExclude),:);
CountryLangManualFull_w_ReqFilesExcluded = CountryLangManualFull(~contains(CountryLangManualFull.FileName,FnamesToExclude),:);

%Remove file names in manual from auto (some papers were manually
%re-evaluated)
RepeatFnamesInManual = intersect(CountryLangAutoFull_w_ReqFilesExcluded.FileName,CountryLangManualFull_w_ReqFilesExcluded.FileName);
CountryLangAutoFull_w_ReqFilesExcluded_ManualDupesRemoved =...
        CountryLangAutoFull_w_ReqFilesExcluded(~contains(CountryLangAutoFull_w_ReqFilesExcluded.FileName,RepeatFnamesInManual),:);

%number check
height(CountryLangAutoFull_w_ReqFilesExcluded_ManualDupesRemoved) + height(CountryLangManualFull_w_ReqFilesExcluded)

%put together tables
CountryLangFinalAuto_SubTab = CountryLangAutoFull_w_ReqFilesExcluded_ManualDupesRemoved(:,{'FileName','Countries','Languages'});
CountryLangFinalManual_SubTab = CountryLangManualFull_w_ReqFilesExcluded(:,{'FileName','CountriesProcessed','LanguagesProcessed'});
CountryLangFinalManual_SubTab.Properties.VariableNames = {'FileName','Countries','Languages'};
CountryLang_PostProcessFull = [CountryLangFinalAuto_SubTab; CountryLangFinalManual_SubTab];

UnlistedCtry = GetUnlistedCtryLang('Countries',CountryLang_PostProcessFull);
UnlistedLang = GetUnlistedCtryLang('Languages',CountryLang_PostProcessFull);

%unique(UnlistedLang)
%unique(UnlistedCtry)

%countries presumed
CountryLang_PostProcessFull(contains(CountryLang_PostProcessFull.Countries,'presume','IgnoreCase',true),:);
%presumed united states
CountryLang_PostProcessFull(contains(CountryLang_PostProcessFull.Countries,'presumed united','IgnoreCase',true),:);
%languages presumed
CountryLang_PostProcessFull(contains(CountryLang_PostProcessFull.Languages,'presume','IgnoreCase',true),:);
%English presumed
CountryLang_PostProcessFull(contains(CountryLang_PostProcessFull.Languages,'presumed english','IgnoreCase',true),:);
%countries AND languages presumed
CountryLang_PostProcessFull(contains(CountryLang_PostProcessFull.Languages,'presume','IgnoreCase',true)...
& contains(CountryLang_PostProcessFull.Countries,'presume','IgnoreCase',true),:);
%presumed languages when study is US (presumed or otherwise)
CountryLang_PostProcessFull(contains(CountryLang_PostProcessFull.Languages,'presume','IgnoreCase',true)...
& contains(CountryLang_PostProcessFull.Countries,'United States','IgnoreCase',true),:);
%unspecified langs
CountryLang_PostProcessFull(contains(CountryLang_PostProcessFull.Languages,'unspecified','IgnoreCase',true),:);

%% LENA

%recdevice_auto
RecDevList = dir('RecDevice_ExtractedSents*');
RecDevAutoFull = readtable(RecDevList(1).name);
RecDevAutoFull = RecDevAutoFull(:,{'FileName','LENA','Camera','RecOrMic'});
for i = 2:numel(RecDevList)
    TempTab = readtable(RecDevList(i).name);
    TempTab = TempTab(:,{'FileName','LENA','Camera','RecOrMic'});
    %TempTab.StatesUS_sents = [];
    RecDevAutoFull = [RecDevAutoFull; TempTab];
end

%recdevice manual
RecDevManualFull = readtable(NotExtractedPapersList(1).name,'Sheet','RecDevice');
RecDevManualFull = RecDevManualFull(:,{'FileName','LENA','Camera','RecOrMic'});
for i = 2:numel(NotExtractedPapersList)
    if ~isempty(readtable(NotExtractedPapersList(i).name)) %the third one is empty cuz it's posthoc round 2
        tempExcTab = readtable(NotExtractedPapersList(i).name,'Sheet','RecDevice');
        tempExcTab = tempExcTab(:,{'FileName','LENA','Camera','RecOrMic'});
        RecDevManualFull = [RecDevManualFull; tempExcTab];
    end
end

%auto and manual extractedrecdev after excluding papers that need exclusion
RecDevAutoFull_w_ReqFilesExcluded = RecDevAutoFull(~contains(RecDevAutoFull.FileName,FnamesToExclude),:);
RecDevManualFull_w_ReqFilesExcluded = RecDevManualFull(~contains(RecDevManualFull.FileName,FnamesToExclude),:);

%Remove file names in manual from auto (some papers were manually
%re-evaluated)
RepeatFnamesInManual_RecDev = intersect(RecDevAutoFull_w_ReqFilesExcluded.FileName,RecDevManualFull_w_ReqFilesExcluded.FileName);
RecDevAutoFull_w_ReqFilesExcluded_ManualDupesRemoved =...
        RecDevAutoFull_w_ReqFilesExcluded(~contains(RecDevAutoFull_w_ReqFilesExcluded.FileName,RepeatFnamesInManual_RecDev),:);

%number check
TotObsExpt_RecDev = height(RecDevAutoFull_w_ReqFilesExcluded_ManualDupesRemoved) + height(RecDevManualFull_w_ReqFilesExcluded)

%put together tables
RecDev_PostProcessFull = [RecDevAutoFull_w_ReqFilesExcluded_ManualDupesRemoved; RecDevManualFull_w_ReqFilesExcluded];
%Note: Oller has methods in supp, and Galindo uses existing repo so the rec
%device could not be inferred

LENA_Rec_yes = height(RecDev_PostProcessFull.LENA(contains(RecDev_PostProcessFull.LENA,'yes','IgnoreCase',true),:))
LENA_Rec_no = height(RecDev_PostProcessFull.LENA(contains(RecDev_PostProcessFull.LENA,'no','IgnoreCase',true),:)) 




%software auto
SoftwareList = dir('Software_ExtractedSents*');
SoftwareAutoFull = readtable(SoftwareList(1).name);
SoftwareAutoFull = SoftwareAutoFull(:,{'FileName','LENA','Praat','OpenSMILE','ALICE','SpeechDetAlg'});
for i = 2:numel(SoftwareList)
    TempTab = readtable(SoftwareList(i).name);
    TempTab = TempTab(:,{'FileName','LENA','Praat','OpenSMILE','ALICE','SpeechDetAlg'});
    %TempTab.StatesUS_sents = [];
    SoftwareAutoFull = [SoftwareAutoFull; TempTab];
end

%software manual
SoftwareManualFull = readtable(NotExtractedPapersList(1).name,'Sheet','Software');
SoftwareManualFull = SoftwareManualFull(:,{'FileName','LENA','Praat','OpenSMILE','ALICE','SpeechDetAlg'});
for i = 2:numel(NotExtractedPapersList)
    if ~isempty(readtable(NotExtractedPapersList(i).name)) %the third one is empty cuz it's posthoc round 2
        tempExcTab = readtable(NotExtractedPapersList(i).name,'Sheet','Software');
        tempExcTab = tempExcTab(:,{'FileName','LENA','Praat','OpenSMILE','ALICE','SpeechDetAlg'});
        SoftwareManualFull = [SoftwareManualFull; tempExcTab];
    end
end

%auto and manual extractedSoftware after excluding papers that need exclusion
SoftwareAutoFull_w_ReqFilesExcluded = SoftwareAutoFull(~contains(SoftwareAutoFull.FileName,FnamesToExclude),:);
SoftwareManualFull_w_ReqFilesExcluded = SoftwareManualFull(~contains(SoftwareManualFull.FileName,FnamesToExclude),:);

%Remove file names in manual from auto (some papers were manually
%re-evaluated)
RepeatFnamesInManual_Software = intersect(SoftwareAutoFull_w_ReqFilesExcluded.FileName,SoftwareManualFull_w_ReqFilesExcluded.FileName);
SoftwareAutoFull_w_ReqFilesExcluded_ManualDupesRemoved =...
        SoftwareAutoFull_w_ReqFilesExcluded(~contains(SoftwareAutoFull_w_ReqFilesExcluded.FileName,RepeatFnamesInManual_Software),:);

%number check
TotObsExpt_Software = height(SoftwareAutoFull_w_ReqFilesExcluded_ManualDupesRemoved) + height(SoftwareManualFull_w_ReqFilesExcluded)

%put together tables
Software_PostProcessFull = [SoftwareAutoFull_w_ReqFilesExcluded_ManualDupesRemoved; SoftwareManualFull_w_ReqFilesExcluded];
%Note: Oller has methods in supp so software could not be inferred

LENA_Software_yes = height(Software_PostProcessFull.LENA(contains(Software_PostProcessFull.LENA,'yes','IgnoreCase',true),:))
LENA_Software_no = height(Software_PostProcessFull.LENA(contains(Software_PostProcessFull.LENA,'no','IgnoreCase',true),:))


LENA_Rec_SubTab = RecDev_PostProcessFull(:,{'FileName','LENA'});
LENA_Rec_SubTab = renamevars(LENA_Rec_SubTab,'LENA','LENA_Rec');
LENA_Software_SubTab = Software_PostProcessFull(:,{'FileName','LENA'});
LENA_Software_SubTab = renamevars(LENA_Software_SubTab,'LENA','LENA_Software');
LENA_RecSoftware_Combined = join(LENA_Software_SubTab,LENA_Rec_SubTab);
LENA_AsRecOrSoftware_SubTab = LENA_RecSoftware_Combined(contains(LENA_RecSoftware_Combined.LENA_Software,'yes','IgnoreCase',true) | ...
                                contains(LENA_RecSoftware_Combined.LENA_Rec,'yes','IgnoreCase',true),:);
LENA_AsRecAndSoftware_SubTab = LENA_RecSoftware_Combined(contains(LENA_RecSoftware_Combined.LENA_Software,'yes','IgnoreCase',true) & ...
                                contains(LENA_RecSoftware_Combined.LENA_Rec,'yes','IgnoreCase',true),:);
Num_LENA_RecORSoftware = height(LENA_AsRecOrSoftware_SubTab)
Num_LENA_RecANDSoftware = height(LENA_AsRecAndSoftware_SubTab)


%----
function [UnlistedVec] = GetUnlistedCtryLang(ColName,IpTab)

    ReqCol = IpTab.(ColName);
    for i = 1:numel(ReqCol)
        NewCtryLangCell{i} = strtrim(strsplit(ReqCol{i},';'));
    end
    
    Ctr = 0;
    for i = 1:numel(NewCtryLangCell)
        for j = 1:numel(NewCtryLangCell{i})
            Ctr = Ctr + 1;
            UnlistedVec{Ctr,1} = NewCtryLangCell{i}{j};
        end
    end
end





