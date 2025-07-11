function generateCompleteReport(excelPath)
    import mlreportgen.report.*
    import mlreportgen.dom.*

    [file, folder] = uiputfile('*.pdf','Save Complete Report as');
    if isequal(file,0), return; end
    rpt = Report(fullfile(folder,file),'pdf');

    tp = TitlePage( ...
      'Title','Complete Data Report', ...
      'Author','AUTOMATION OF DATA ANALYSIS', ...
      'Subtitle',['Source file: ' extractFileName(excelPath)] ...
    );
    append(rpt,tp);
    append(rpt,TableOfContents);

    [~, sheets] = xlsfinfo(excelPath);

    for k = 1:numel(sheets)
        T = readtable(excelPath,'Sheet',sheets{k});
        chap = Chapter(['Sheet: ' sheets{k}]);
        add(chap,Paragraph(['Statistics for sheet "' sheets{k} '"']));

        summary = summarizeTable(T);
        tblData = ["Statistic" summary.Properties.VariableNames; summary.Properties.RowNames table2cell(summary)];
        tbl = FormalTable(tblData);
        tbl.Header.Style{end+1} = BackgroundColor('lightgray');
        tbl.Width = '100%';
        tbl.Border = 'solid';
        tbl.ColSep = 'solid';
        tbl.RowSep = 'solid';
        tbl.TableEntriesHAlign = 'center';
        add(chap,tbl);

        numCols = varfun(@isnumeric,T,'OutputFormat','uniform');
        names = T.Properties.VariableNames(numCols);
        if numel(names) >= 2
            x = T.(names{1}); y = T.(names{2});
            valid = isfinite(x)&isfinite(y);
            if sum(valid)>1
                f = figure('Visible','off','Position',[100 0 450, 400]);
                plot(x(valid),y(valid),'-o','LineWidth',1.5);
                xlabel(names{1}); ylabel(names{2});
                title([names{2} ' vs ' names{1}]);
                grid on;
                img = [tempname,'.png'];
                exportgraphics(f,img,'Resolution',150);
                close(f);
                add(chap,Image(img));
            else
                add(chap,Paragraph('⚠️ Not enough valid points to plot.'));
            end
        else
            add(chap,Paragraph('⚠️ Not enough numeric columns to plot.'));
        end

        append(rpt,chap);
    end

    concl = Chapter('Final Conclusions');
    txt = "The analyzed dataset contains multiple worksheets with structured information. Key observations include: ";
    for k = 1:numel(sheets)
        T = readtable(excelPath,'Sheet',sheets{k});
        [r,c] = size(T);
        txt = txt + sprintf("- Sheet '%s' includes %d rows and %d columns.", sheets{k}, r, c);

        % Check for high missing data
        missings = sum(ismissing(T));
        highMiss = find(missings > 0.3*r);
        if ~isempty(highMiss)
            for idx = highMiss
                txt = txt + sprintf("    Column '%s' has over 30%% missing values.", T.Properties.VariableNames{idx});
            end
        end

        % Look for constant columns
        constCols = varfun(@(x) numel(unique(x(~ismissing(x)))) == 1, T, 'OutputFormat','uniform');
        constNames = T.Properties.VariableNames(constCols);
        for name = constNames
            txt = txt + sprintf("    Column '%s' has constant values.\n", name{1});
        end
    end
    txt = txt + "Overall, the dataset is suitable for further statistical or predictive modeling. " + ...
        "Columns with high variance and minimal missing values are good candidates for in-depth analysis.";
    txt = txt + sprintf("" + ...
        "        Report generated on %s.", datestr(now));
    add(concl, Paragraph(txt))
    append(rpt,concl);

    close(rpt);
    rptview(rpt);
end

function name = extractFileName(filepath)
    [~,nm,ext] = fileparts(filepath);
    name = [nm,ext];
end

function summary = summarizeTable(T)
    vars = T.Properties.VariableNames;
    rowNames = { ...
      'DataType','NonMissing','Missing', ...
      'Unique','MostFreq','Min','Max', ...
      'Mean','Median','StdDev','Skewness','Kurtosis', ...
      'Range','PctMissing','IsConstant'};

    n = numel(vars);
    data = cell(numel(rowNames),n);

    for j = 1:n
        col = T.(vars{j});
        nm  = class(col);
        data{1,j} = nm;

        ok = ~ismissing(col);
        data{2,j} = sum(ok);
        data{3,j} = sum(~ok);

        nonmiss = col(ok);
        data{4,j} = numel(unique(nonmiss));
        try
            data{5,j} = mode(nonmiss);
        catch
            data{5,j} = "";
        end

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
                data{14,j} = 100 * data{3,j} / height(T);
                data{15,j} = all(x==x(1));
            else
                data(6:15,j) = {NaN};
            end
        else
            data(6:13,j) = {NaN};
            data{14,j} = 100 * data{3,j} / height(T);
            data{15,j} = isempty(nonmiss) || all(isequaln(nonmiss{1},nonmiss));
        end
    end

    summary = cell2table(data, 'RowNames', rowNames, 'VariableNames', vars);
end
