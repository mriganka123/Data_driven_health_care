clear
close all
clc
%% Removing first row
dir_c = pwd;
file_name = [dir_c '\State_Under_65_Table.xlsx'];
row = 1;
excel = actxserver('Excel.Application');
workbook = excel.Workbooks.Open(file_name);
for i = 2007:2016
    sheet_name = ['State ' num2str(i)];
    worksheet = workbook.Worksheets.Item(sheet_name);
    worksheet.Rows.Item(row).Delete;
    workbook.Save;
end
excel.Quit;
%% Analysing 2016
Z = readtable(file_name,'Sheet','State 2016');
% Replacing '*' with average of the attributes
Z = Z(1:end-1,:);
size_Z = size(Z);
for i = 3:size_Z(2)
    idx = zeros(size_Z(1),1);
    idx = ismember(Z.(i),'*');
    idx = idx + ismember(Z.(i),'.');
    idx = idx + ismember(Z.(i),'NaN');
    II = [];
    if sum(idx)
        I = Z.(i)(not(idx));
        for j = 1:length(I)
            try
                idx_perc = strfind(I{j},'%');
                if idx_perc
                    II(j,1) = str2double(I{j}(1:end-1));
                else
                    II(j,1) = str2double(I{j});
                end
            catch
                II = I;
                break
            end
        end
        Z.(i)(logical(idx)) = {num2str(mean(II))};
    end
end
VarNames = Z.Properties.VariableNames;
% per_capita_cost = Z.StandardizedPerCapitaCosts;
% hist(per_capita_cost(2:end),length(per_capita_cost))
% xlabel('Dollar (\$)','Interpreter','latex')
Z_cost = Z(:,16:21);
Z_rel = Z(:,[1:15,22:end]);
% PCA
% Not considering the total data.
% Only focusing on standardized data
VarNames_rel = Z_rel.Properties.VariableNames;
idx = ones(1,length(VarNames_rel));
idx(1:2) = 0;
Rel_data = [];
for i = 3:length(VarNames_rel)
    temp = VarNames_rel{i};
    idx_str = strfind(temp,'Total');
    if idx_str
        idx(i) = 0;
    else
        I = Z_rel.(i);
        II = [];
        for j = 1:length(I)
            try
                idx_perc = strfind(I{j},'%');
                if idx_perc
                    II(j,1) = str2double(I{j}(1:end-1));
                else
                    II(j,1) = str2double(I{j});
                end
            catch
                II = I;
                break
            end
        end
        Rel_data = [Rel_data,(II-min(II))/max(II)];
    end
end
temp = Z_cost.(1);

for i = 1:length(temp)
    Y(i,1) = str2double(temp{i});
end

%% Analysing all year
Rel_data = [];
Y = [];
for i = 2007:2016
    sheet_name = ['State ' num2str(i)];
    temp_Z = readtable('State_Under_65_Table.xlsx','Sheet',sheet_name);
    % Replacing '*' with average of the attributes
    temp_Z = temp_Z(1:end-1,:);
    size_Z = size(temp_Z);
    for ii = 3:size_Z(2)
        idx = zeros(size_Z(1),1);
        idx = ismember(temp_Z.(ii),'*');
        idx = idx + ismember(temp_Z.(ii),'.');
        idx = idx + ismember(temp_Z.(ii),'NaN');
        II = [];
        if sum(idx)
            I = temp_Z.(ii)(not(idx));
            for j = 1:length(I)
                try
                    idx_perc = strfind(I{j},'%');
                    if idx_perc
                        II(j,1) = str2double(I{j}(1:end-1));
                    else
                        II(j,1) = str2double(I{j});
                    end
                catch
                    II = I;
                    break
                end
            end
            temp_Z.(ii)(logical(idx)) = {num2str(mean(II))};
        end
    end
    Z_all{i} = temp_Z;
    Z_cost = temp_Z(:,16:21);
    Z_rel = temp_Z(:,[1:15,22:end]);
    VarNames_rel = Z_rel.Properties.VariableNames;
    idx = ones(1,length(VarNames_rel));
    idx(1:2) = 0;
    Rel_data_temp = [];
    for i = 3:length(VarNames_rel)
        temp = VarNames_rel{i};
        idx_str = strfind(temp,'Total');
        if idx_str
            idx(i) = 0;
        else
            I = Z_rel.(i);
            II = [];
            for j = 1:length(I)
                try
                    idx_perc = strfind(I{j},'%');
                    if idx_perc
                        II(j,1) = str2double(I{j}(1:end-1));
                    else
                        II(j,1) = str2double(I{j});
                    end
                catch
                    II = I;
                    break
                end
            end
            Rel_data_temp = [Rel_data_temp,(II-min(II))/max(II)];
        end
    end
    Rel_data = [Rel_data;Rel_data_temp];
    temp = Z_cost.(1);

    for i = 1:length(temp)
        Y = [Y;str2double(temp{i})];
    end
end
% Linear Regression
[b1, bint] = regress(Y,Rel_data);
plot(bint)
[coeff,score,latent,tsquared,explained] = pca(Rel_data,'NumComponents',2);
plot(explained)
relevence = 0;
i = 1;
while relevence<90
    relevence = relevence + explained(i);
    i = i+1;
end
