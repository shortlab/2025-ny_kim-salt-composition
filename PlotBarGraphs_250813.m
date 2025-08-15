% Clear workspace and command window
clear; clc;

%% --- Select and load the average concentration Excel file ---
[avgFile, avgPath] = uigetfile({'*.xlsx;*.xls','Excel Files (*.xlsx, *.xls)'}, ...
    'Select the Excel file for average concentration');
if isequal(avgFile, 0)
    error('No file selected for average concentration.');
end
avgFileFull = fullfile(avgPath, avgFile);
avgTable = readtable(avgFileFull);

% Extract analyte names from the table's variable names (skip the first column)
analyteNames = avgTable.Properties.VariableNames(2:end);
% Remove '_HeMode_', trim spaces, and remove a leading 'x'
for i = 1:length(analyteNames)
    analyteNames{i} = strrep(analyteNames{i}, '_HeMode_', '');
    analyteNames{i} = strtrim(analyteNames{i});
    if startsWith(analyteNames{i}, 'x')
        analyteNames{i} = strtrim(analyteNames{i}(2:end));
    end
    % Remove any digits from the name
    analyteNames{i} = regexprep(analyteNames{i}, '\d', '');
end

%% --- Select which analytes to display ---
% Prompt the user for analyte selection.
% Type "All" to display all analytes, or list one or more substrings separated by commas.
analyteInput = input(['Enter the analyte names (or parts of the names) to display separated ' ...
    'by commas (or type "All" for all): '], 's');
if strcmpi(strtrim(analyteInput), 'All')
    selectedAnalyteIndices = 1:length(analyteNames);
else
    selectedSubstrings = strsplit(analyteInput, ',');
    selectedSubstrings = strtrim(selectedSubstrings);
    selectedAnalyteIndices = [];
    % For each substring, find indices of analyteNames that contain it (case-insensitive)
    for j = 1:length(selectedSubstrings)
        matchInd = find(contains(lower(analyteNames), lower(selectedSubstrings{j})));
        selectedAnalyteIndices = union(selectedAnalyteIndices, matchInd);
    end
    if isempty(selectedAnalyteIndices)
        error('None of the specified analyte names were found.');
    end
end
% Filter the analyte names based on selection
analyteNames = analyteNames(selectedAnalyteIndices);

% Extract average concentration numeric values from columns 2:end
avgConcentration = table2array(avgTable(:, 2:end));

%% --- Select and load the sample standard deviation Excel file ---
[stdFile, stdPath] = uigetfile({'*.xlsx;*.xls','Excel Files (*.xlsx, *.xls)'}, ...
    'Select the Excel file for sample standard deviation');
if isequal(stdFile, 0)
    error('No file selected for sample standard deviation.');
end
stdFileFull = fullfile(stdPath, stdFile);
stdTable = readtable(stdFileFull);
stdDev = table2array(stdTable(:, 2:end));

%% --- Select and load the propagated error Excel file ---
[propFile, propPath] = uigetfile({'*.xlsx;*.xls','Excel Files (*.xlsx, *.xls)'}, ...
    'Select the Excel file for propagated error');
if isequal(propFile, 0)
    error('No file selected for propagated error.');
end
propFileFull = fullfile(propPath, propFile);
propTable = readtable(propFileFull);
propError = table2array(propTable(:, 2:end));

%% --- Choose the Rows to Plot ---
startRow = input('Enter the start row number for plotting (e.g., 1): ');
endRow = input('Enter the end row number for plotting: ');

[nTotalRows, ~] = size(avgConcentration);
if startRow < 1 || endRow > nTotalRows || startRow > endRow
    error('Invalid row selection. Ensure that 1 <= startRow <= endRow <= %d.', nTotalRows);
end

% Filter the data for the selected rows and the selected analytes (columns)
selectedAvgConcentration = avgConcentration(startRow:endRow, selectedAnalyteIndices);
selectedStdDev = stdDev(startRow:endRow, selectedAnalyteIndices);
selectedPropError = propError(startRow:endRow, selectedAnalyteIndices);

%% --- Define x-values and custom labels ---
% Determine the number of measurements (rows)
N = endRow - startRow + 1;
% Center the x-values
xValues = (1:N) - ((N+1)/2);

% Ask for custom x-axis tick labels
customXLabelsInput = input('Enter custom x-axis labels separated by commas (e.g., "Time1, Time2, Time3"): ', 's');
customXLabels = strsplit(customXLabelsInput, ',');
customXLabels = strtrim(customXLabels);
if length(customXLabels) ~= N
    error('The number of custom x-axis labels (%d) must match the number of measurements (%d).', length(customXLabels), N);
end

% Define a padding value for the x-axis (adjust as needed)
padding = 0.5;

%% --- Plot 1: Average Concentration with Standard Deviation Error Bars (Bar Graph) ---
figure;
b = bar(xValues, selectedAvgConcentration);  % Create grouped bar plot
hold on;
% Get number of groups and number of bars (analytes)
[ngroups, nbars] = size(selectedAvgConcentration);
% Compute the x coordinates for each bar group
xBar = nan(ngroups, nbars);
for i = 1:nbars
    xBar(:,i) = b(i).XEndPoints;  % requires R2019b or newer
end
% Add error bars at the computed x positions
errorbar(xBar, selectedAvgConcentration, selectedStdDev, 'k', 'linestyle', 'none');
hold off;

% Customize font sizes
xlabel('Digestion method', 'FontSize', 24);
ylabel('Concentration [wt. ppm]', 'FontSize', 24);
title('Average Concentration with Standard Deviation Error Bars', 'FontSize', 18);
legend(analyteNames, 'Location', 'Best', 'FontSize', 20);

grid on;
set(gca, 'XTick', xValues, 'XTickLabel', customXLabels, 'FontSize', 22);
xlim([min(xValues)-padding, max(xValues)+padding]);

%% --- Plot 2: Average Concentration with Propagated Error Bars (Bar Graph) ---
figure;
b = bar(xValues, selectedAvgConcentration);  % Create grouped bar plot
hold on;
[ngroups, nbars] = size(selectedAvgConcentration);
xBar = nan(ngroups, nbars);
for i = 1:nbars
    xBar(:,i) = b(i).XEndPoints;
end
errorbar(xBar, selectedAvgConcentration, selectedPropError, 'k', 'linestyle', 'none');
hold off;

% Customize font sizes
xlabel('Digestion method', 'FontSize', 24);
ylabel('Concentration [wt. ppm]', 'FontSize', 24);
title('Average Concentration with Propagated Error Bars', 'FontSize', 18);
legend(analyteNames, 'Location', 'Best', 'FontSize', 20);

grid on;
set(gca, 'XTick', xValues, 'XTickLabel', customXLabels, 'FontSize', 22);
xlim([min(xValues)-padding, max(xValues)+padding]);
