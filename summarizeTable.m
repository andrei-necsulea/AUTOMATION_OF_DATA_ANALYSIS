function summary = summarizeTable(T)
    vars = T.Properties.VariableNames;
    rowNames = { ...
      'DataType','NonMissing','Missing', ...
      'Unique','MostFreq','Min','Max', ...
      'Mean','Median','Std','Skew','Kurt', ...
      'Range','PctMissing','IsConst'};

    n = numel(vars);
    data = cell(numel(rowNames),n);

    for j = 1:n
        col = T.(vars{j});
        nm  = class(col);
        data{1,j} = nm;

        % Counts
        ok = ~ismissing(col);
        data{2,j} = sum(ok);
        data{3,j} = sum(~ok);

        % Unique & MostFrequent
        nonmiss = col(ok);
        data{4,j} = numel(unique(nonmiss));
        if isempty(nonmiss)
            data{5,j} = "";
        else
            try
                data{5,j} = mode(nonmiss);
            catch
                data{5,j} = "";
            end
        end

        % Numeric stats
        if isnumeric(col)
            x = col(ok & isfinite(col));
            if ~isempty(x)
                data{6,j}  = min(x);
                data{7,j}  = max(x);
                data{8,j}  = mean(x);
                data{9,j}  = median(x);
                data{10,j} = std(x);
                data{11,j} = skewness(x);
                data{12,j} = kurtosis(x);
                data{13,j} = range(x);
                data{14,j} = 100*data{3,j}/height(T);
                data{15,j} = all(x==x(1));
            else
                data(6:15,j) = {NaN};
            end
        else
            % Non-numeric → only %missing + isConst
            data(6:13,j) = {NaN};
            data{14,j} = 100*data{3,j}/height(T);
            data{15,j} = isempty(nonmiss) || all(isequaln(nonmiss{1},nonmiss));
        end
    end

    summary = cell2table(data, ...
        'RowNames', rowNames, ...
        'VariableNames', vars);
end
