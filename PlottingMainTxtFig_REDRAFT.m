clear; clc

PreHoc = readtable('DARCLE_extraction_2.16.26.xlsx','Sheet','Corrected_ReTyped_PostManualExt');
PostHoc = readtable('Post-hoc Paper Extractions_3.23.26.xlsx','Sheet','Corrected_ReTyped_PostManualExt');

CombinedTab = [PreHoc; PostHoc];
LFAndEngOnlyCombinedTab = CombinedTab(~contains(CombinedTab.Type,'Non-longform') & ...
    ~contains(CombinedTab.Type,'NonEnglishContent'),:); 

u_LFonlyType = unique(LFAndEngOnlyCombinedTab.Type);
clear NumOfType

for i = 1:numel(u_LFonlyType)
    NumOfType(i,1) = height(LFAndEngOnlyCombinedTab(contains(LFAndEngOnlyCombinedTab.Type,u_LFonlyType{i}),:));
end

table(u_LFonlyType,NumOfType,NumOfType/height(LFAndEngOnlyCombinedTab)*100)
height(LFAndEngOnlyCombinedTab)
min(LFAndEngOnlyCombinedTab.PublishedYear)
max(LFAndEngOnlyCombinedTab.PublishedYear)

uYear = unique(LFAndEngOnlyCombinedTab.PublishedYear);

u_LFonlyType_org = {'Observational','Experimental','Validation','Computational','Review',...
'Commentaries','Best Practice','Other'};
for i = 1:numel(uYear)
    NumPerYear(i,1) = height(LFAndEngOnlyCombinedTab(LFAndEngOnlyCombinedTab.PublishedYear == uYear(i),:));
    for j = 1:numel(u_LFonlyType_org)
        NumYrType(i,j) = height(LFAndEngOnlyCombinedTab(LFAndEngOnlyCombinedTab.PublishedYear == uYear(i)...
        & contains(LFAndEngOnlyCombinedTab.Type,u_LFonlyType_org{j}),:));
    end
end

OkabeItoMap_Rgb = hex2rgb(flip({'#E69F00', '#56B4E9','#009E73','#F0E442','#0072B2','#D55E00','#CC79A7','#000000'}));

Clrs = [0 0 146; %DB
0 78 162; %LB
0 146 150; %GB
87 185 134; %G
181 211 105; %LG
254 222 93; %Y
255 160 39; %O
229 82 0]/256; %R

%Clrs = OkabeItoMap_Rgb;


figure('Color',[1 1 1]); 
axes1 = subplot(2,2,1);
Aplot = area(uYear(3:end),NumYrType(3:end,:));
for i = 1:height(Clrs)
    Aplot(i).EdgeColor = 'none';
    Aplot(i).FaceColor = Clrs(i,:);
end
set(axes1,'FontSize',24);
legend(u_LFonlyType_org,'Direction','normal') %set direction to normal

axes2 = subplot(2,2,2); hold all
b = bar(uYear(1),NumYrType(1,:),'stacked');
for i = 1:height(Clrs)
b(i).EdgeColor = 'none';
b(i).FaceColor = Clrs(i,:);
end
b = bar(uYear(2),NumYrType(2,:),'stacked');
for i = 1:height(Clrs)
b(i).EdgeColor = 'none';
b(i).FaceColor = Clrs(i,:);
end
axis tight
set(axes2,'FontSize',24);

axes3 = subplot(2,2,3); 
set(axes3,'FontSize',24);

axes4 = subplot(2,2,4); bar(uYear,NumPerYear)
set(axes4,'FontSize',24);
axis tight