%%%%Sorting and counting spine morphology%%%

%Open file coloumn spine morphology%%

[status,sheets]= xlsfinfo('new fugw spines R156 summary.xlsx');
numOfSheets=numel(sheets);

for i=1:numOfSheets
    
    [num raw txt]= xlsread('new fugw spines R156 summary.xlsx',i,'K:K');%opens excel file sheet one and coloumn k as raw "text"%
    spine_morph = raw
    %%clear raw;
    %%clear txt;
    clear num;
    spine_morph(1:1)=[] %%Deletes first cell - looks like it always contains the title 'TYPE'??%%
%%%header = {'TYPE'}%%
%% find a word in a string
%%Mathworks recommends strcmp - this compares two strings and puts 1 if
%%correct or 0 is false%%
           mushroom = 'mushroom';%define string
           stubby = 'stubby';%define string
           thin = 'thin';%define string

                M = strcmp(spine_morph, mushroom);
                S = strcmp(spine_morph, stubby);
                T = strcmp(spine_morph, thin);

%Counting the number of 1's in an array%%
mushroom_total = sum(M);
stubby_total = sum(S);
thin_total = sum(T);
%% 
%%write excel document%%
xlswrite('new fugw spines R156 summary.xlsx', mushroom_total,i,'V2:V2');
xlswrite('new fugw spines R156 summary.xlsx', stubby_total,i,'W2:W2');
xlswrite('new fugw spines R156 summary.xlsx', thin_total,i,'X2:X2');


%%col headers and row headers%%
col_header= {'Mushroom','Stubby','Thin'};
row_header= {'Total'};
xlswrite('new fugw spines R156 summary.xlsx',col_header,i,'V1:X1')
xlswrite('new fugw spines R156 summary.xlsx',row_header,i,'U2:U2')

end



