clc;
clear;
close all;

%% ✅ STEP 1. Extract dilution parameters 
[inputFile1, inputPath1] = uigetfile('*.xlsx', 'Select the Excel File for Dilution Data');
if isequal(inputFile1, 0) || isempty(inputFile1)
    error('❌ Error: No valid input file selected for dilution data.');
end
inputFilePath1 = fullfile(inputPath1, inputFile1);

% ✅ Extract "ICP ####" from the first filename
[~, fileName1, ~] = fileparts(inputFile1);
tokens1 = regexp(fileName1, 'ICP\s*\d+', 'match'); 
if ~isempty(tokens1)
    extractedICP = strtrim(tokens1{1}); % Extracted "ICP ####"
else
    extractedICP = 'ICP UNKNOWN'; % Default if no match
end

% ✅ Detect import options for dilution data
opts = detectImportOptions(inputFilePath1, 'PreserveVariableNames', true);
opts.VariableNamesRange = 'A11';  
opts.DataRange = 'A13';  
opts.SelectedVariableNames = opts.VariableNames;  
dataTable1 = readtable(inputFilePath1, opts);

% ✅ Define keywords for dilution data
keywords = {'Mass of salt', 'Volume of digested sol.', '[Tb expected]', 'C1C2', 'C1C3'};
normalizedColumnNames = lower(strrep(strrep(strtrim(dataTable1.Properties.VariableNames), " ", "_"), ".", ""));
normalizedKeywords = lower(strrep(strrep(strtrim(keywords), " ", "_"), ".", ""));
matchedColumns = contains(normalizedColumnNames, normalizedKeywords, 'IgnoreCase', true);

% ✅ Extract matched columns for dilution data
if any(matchedColumns)
    extractedData1 = dataTable1(:, matchedColumns);
    
    % ✅ Find valid rows (non-empty values)
    validRows1 = false(height(extractedData1), 1);
    for col = 1:width(extractedData1)
        colData = extractedData1{:, col};
        if isnumeric(colData) || islogical(colData)
            nonEmptyIdx = find(~isnan(colData));
        else
            nonEmptyIdx = find(~ismissing(colData) & colData ~= "");
        end
        if ~isempty(nonEmptyIdx)
            validRows1(min(nonEmptyIdx):max(nonEmptyIdx)) = true;
        end
    end
    extractedData1 = extractedData1(validRows1, :);

    % ✅ Convert the numerical table to a **cell array** for compatibility
    extractedData1_cell = table2cell(extractedData1);

    % ✅ Create a second row with specific labels
    labelsRow = repmat({''}, 1, width(extractedData1)); % Initialize as empty strings
    labelsRow(1:min(5, width(extractedData1))) = {'[mg]', '[mL]', '[ppm]', '[ppb->ppm]', '[ppb->ppm]'}; 

    % ✅ Combine the labelsRow and data into a **cell array**
    combinedData = [labelsRow; extractedData1_cell];

    % ✅ Convert the combined data back to a table with correct headers
    extractedData1 = cell2table(combinedData, 'VariableNames', extractedData1.Properties.VariableNames);

    % ✅ Store dilution data in MATLAB workspace
    assignin('base', 'DILUTION', extractedData1);

    % ✅ Save dilution data
    %outputFile1 = fullfile(inputPath1, [extractedICP, '_DILUTION.xlsx']);
    %writetable(extractedData1, outputFile1, 'Sheet', 'DILUTION', 'WriteVariableNames', true);
    %disp(['✅ Dilution data saved to: ', outputFile1]);
else
    disp('❌ No matching columns found in dilution data.');
end






%% ✅ STEP 2: Extract raw ICP-MS data
[inputFile2, inputPath2] = uigetfile('*.csv', 'Select the CSV File for Raw ICP-MS Data');
if isequal(inputFile2, 0)
    error('❌ Error: No valid input file selected for raw data.');
end
inputFilePath2 = fullfile(inputPath2, inputFile2);

% ✅ Extract filename for classification (Major/Trace)
[~, fileName2, ~] = fileparts(inputFile2);

% ✅ Read raw ICP-MS data
dataTable2 = readtable(inputFilePath2, 'PreserveVariableNames', true);

% ✅ Apply filtering to raw data (Remove rows with "Cal" in Col E, "blank" in Col G)
colE = 5;
colG = 7;
validE = ~contains(lower(string(dataTable2{:, colE})), "cal", 'IgnoreCase', true);
validG = ~contains(lower(string(dataTable2{:, colG})), "blank", 'IgnoreCase', true);
validRows2 = validE & validG;
filteredData2 = dataTable2(validRows2, :);

% ✅ Ensure consistent column selection for every 4th column starting from 8
startCol = 8;
step = 4;
numColumns = width(dataTable2);

selectedCols = [7]; % Always include column 7 (Sample ID)
validCols = startCol:step:numColumns;
selectedCols = [selectedCols, validCols(validCols <= numColumns)];

% ✅ Extract selected columns
filteredData2 = filteredData2(:, selectedCols);
row2Data = dataTable2(2, selectedCols);

% ✅ Determine the number of columns in filteredData2
numCols = width(filteredData2);

% ✅ Create a new second row with "[ppb]" for all columns except the first
ppbRow = repmat({''}, 1, numCols);
ppbRow(2:end) = {'[ppb]'};

% ✅ Convert all non-numeric values to numeric (0 if invalid)
for colIdx = 2:numCols  % Skip Sample ID
    colData = filteredData2{:, colIdx};

    if iscell(colData)
        colData = cellfun(@(x) parseToZero(x), colData, 'UniformOutput', false);
    elseif isnumeric(colData)
        colData = num2cell(colData);
        colData(cellfun(@isnan, colData)) = {0};
    end

    filteredData2.(colIdx) = colData;
end

% ✅ Convert Sample ID column to cellstr for consistency
filteredData2.(1) = cellstr(string(filteredData2.(1)));

% ✅ Store raw data in MATLAB workspace
assignin('base', 'RAWDATA_filt', filteredData2);
assignin('base', 'RAWDATA_row2', row2Data);

% ✅ Generate filename for raw data
newFileName = strcat(extractedICP, '_RAW ICP');
if contains(fileName2, "Major", 'IgnoreCase', true)
    newFileName = strcat(newFileName, '_Major');
elseif contains(fileName2, "Trace", 'IgnoreCase', true)
    newFileName = strcat(newFileName, '_Trace');
end

% ✅ Save raw ICP-MS data (headers + [ppb] row + data)
%RAWDATA = fullfile(inputPath1, [newFileName, '.xlsx']);
headers = filteredData2.Properties.VariableNames;
dataCell = table2cell(filteredData2);
saveCell = [headers; ppbRow; dataCell];
%writecell(saveCell, RAWDATA, 'Sheet', 'RAWDATA_filt');

% ✅ Save row 2 info as table
%writetable(row2Data, RAWDATA, 'Sheet', 'RAWDATA_row2', 'WriteVariableNames', true);

% ✅ Define helper function to convert text to numeric safely
function val = parseToZero(x)
    if isnumeric(x)
        if isnan(x)
            val = 0;
        else
            val = x;
        end
    elseif ischar(x) || isstring(x)
        num = str2double(x);
        if isnan(num)
            val = 0;
        else
            val = num;
        end
    else
        val = 0;
    end
end


%% ✅ STEP 3: Extract raw ICP-MS RSD data
% Reuse the same file as selected in STEP 2 (raw ICP‐MS data)
inputFile = inputFile2;       % From STEP 2
inputPath = inputPath2;       % From STEP 2
inputFilePath = fullfile(inputPath, inputFile);
[~, fileName2, ~] = fileparts(inputFile2);

% ✅ Read raw ICP-MS data (reuse dataTable2 from STEP 2 if available)
dataTable2 = readtable(inputFilePath, 'PreserveVariableNames', true);

% ✅ Apply filtering to raw data (Remove rows with "Cal" in Col E, "blank" in Col G)
colE = 5;
colG = 7;
validE = ~contains(lower(string(dataTable2{:, colE})), "cal", 'IgnoreCase', true);
validG = ~contains(lower(string(dataTable2{:, colG})), "blank", 'IgnoreCase', true);
validRows2 = validE & validG;
filteredData2 = dataTable2(validRows2, :);

% ✅ Ensure consistent column selection for every 4th column starting from 9
startCol = 10;  
step = 4;
numColumns = width(dataTable2);
selectedCols = [7]; % Always include column 7 (Sample ID)
validCols = startCol:step:numColumns; % Strictly select 9, 13, 17, 21, ...
selectedCols = [selectedCols, validCols(validCols <= numColumns)];

% ✅ Debugging output to verify selected columns
%disp('✅ Selected Column Indices:');
%disp(selectedCols);
%disp('✅ Selected Column Names:');
%disp(dataTable2.Properties.VariableNames(selectedCols));

% ✅ Extract selected columns
filteredData2 = filteredData2(:, selectedCols);
row2Data = dataTable2(2, selectedCols);

% ✅ Determine the number of columns in filteredData2
numCols = width(filteredData2); 

% ✅ Create a new second row with "[%]" for **all columns except the first**
ppbRow = repmat({''}, 1, numCols); % Initialize empty row
ppbRow(2:end) = {'[%]'}; % Insert "[%]" from column 2 onward

% ✅ Convert to table with the same variable names as filteredData2
ppbTable = cell2table(ppbRow, 'VariableNames', filteredData2.Properties.VariableNames);

% --- Create a new variable for RSD data ---
% Initialize filteredRSDData2 as a copy of filteredData2
filteredRSDData2 = filteredData2;

% Convert all non-numeric values in filteredData2 to 0 and store in filteredRSDData2
for colIdx = 2:numCols  % Skip first column (Sample ID)
    colData = filteredData2{:, colIdx}; % Extract column data
    
    % Replace non-numeric values with 0
    if iscell(colData)
        colData = cellfun(@(x) parseToZero(x), colData, 'UniformOutput', false);
    elseif isnumeric(colData)
        colData = num2cell(colData); % Convert numbers to cell format
        colData(cellfun(@isnan, colData)) = {0}; % Replace NaN with 0
    end
    filteredRSDData2.(colIdx) = colData; % Store back in filteredRSDData2
end

% ✅ Convert Sample ID column to cells for consistency
filteredRSDData2.(1) = cellstr(string(filteredRSDData2.(1)));

% ✅ Concatenate the "[%]" row with the filtered RSD data
filteredRSDData2 = [ppbTable; filteredRSDData2];

% ✅ Store raw RSD data in MATLAB workspace
assignin('base', 'RAWRSD_filt', filteredRSDData2);
assignin('base', 'RAWRSD_row2', row2Data);

% ✅ Generate filename for RSD data
newFileName = strcat(extractedICP, '_RAW RSD');
if contains(fileName2, "Major", 'IgnoreCase', true)
    newFileName = strcat(newFileName, '_Major');
elseif contains(fileName2, "Trace", 'IgnoreCase', true)
    newFileName = strcat(newFileName, '_Trace');
end

RAWRSD = fullfile(inputPath1, [newFileName, '.xlsx']);
%writetable(filteredRSDData2, RAWRSD, 'Sheet', 'RAWRSD_filt', 'WriteVariableNames', true);
%writetable(row2Data, RAWRSD, 'Sheet', 'RAWRSD_row2', 'WriteVariableNames', true);

%disp(['✅ Raw ICP-MS RSD data saved to: ', RAWRSD]);
%disp('✅ Extracted RSD data stored in MATLAB workspace as: RAWRSD_filt');






%% ✅ STEP 4: Apply Dilution Factors 
% ✅ Extract the column index based on file name
if contains(fileName2, "Trace", 'IgnoreCase', true)
    dilutionColIdx = 4; % Use Column 4 for Trace
elseif contains(fileName2, "Major", 'IgnoreCase', true)
    dilutionColIdx = 5; % Use Column 5 for Major
else
    error('❌ Error: File name must contain "Trace" or "Major" to define output file name.');
end

% ✅ Extract and clean dilution factor column
dilutionFactors = DILUTION{2:end, dilutionColIdx}; % Extract column

% ✅ Properly extract values from nested cells
if iscell(dilutionFactors)  
    dilutionFactors = cellfun(@(x) x, dilutionFactors); % Extract numeric values directly
end

% ✅ Extract numerical data from RAWDATA_filt
rawDataMatrix = RAWDATA_filt{1:end, 2:end}; % Extract numeric values

% ✅ Handle text values safely
if iscell(rawDataMatrix)
    rawDataMatrix = cellfun(@parseValue, rawDataMatrix, 'UniformOutput', false);
end

% ✅ Convert to numeric matrix safely
rawDataMatrix = cell2mat(rawDataMatrix);

% ✅ Ensure dimensions match before multiplication
if size(dilutionFactors, 1) ~= size(rawDataMatrix, 1)
    error('❌ Mismatch: dilutionFactors and rawDataMatrix have different row sizes.');
end

% ✅ Perform element-wise multiplication
correctedData = rawDataMatrix .* dilutionFactors; % Broadcasting effect (m x n)

% ✅ Convert back to table, retaining headers
correctedDataTable = array2table(correctedData, 'VariableNames', RAWDATA_filt.Properties.VariableNames(2:end));

% ✅ Convert all numeric columns to **cell format** for consistent table merging
for colIdx = 1:width(correctedDataTable)
    correctedDataTable.(colIdx) = num2cell(correctedDataTable{:, colIdx});
end

% ✅ Extract Sample ID column (First column remains unchanged)
sampleIDColumn = RAWDATA_filt(1:end, 1);

% ✅ Create `[ppm]` row, ensuring same number of columns as `correctedDataTable`
ppmRow = repmat({''}, 1, width(correctedDataTable) + 1); % First column empty
ppmRow(2:end) = {'[ppm]'}; % Apply "[ppm]" for numeric columns

% ✅ Convert `[ppm]` row to a table with correct column names
ppmTable = cell2table(ppmRow, 'VariableNames', [sampleIDColumn.Properties.VariableNames, correctedDataTable.Properties.VariableNames]);

% ✅ Concatenate `[ppm]` row with corrected data
correctedDataTable = [ppmTable; [sampleIDColumn, correctedDataTable]];

% ✅ Store the corrected data in MATLAB workspace as "DF_MULTIPLIED"
assignin('base', 'DF_MULTIPLIED', correctedDataTable);

% ✅ Generate filename based on ICP identifier and Major/Trace classification
if contains(fileName2, "Major", 'IgnoreCase', true)
    fileSuffix = "DF MULTIPLIED_Major";
elseif contains(fileName2, "Trace", 'IgnoreCase', true)
    fileSuffix = "DF MULTIPLIED_Trace";
else
    error('❌ Error: File name must contain "Major" or "Trace" to define output file name.');
end

% ✅ Extract "ICP ###" from the first input file name
icpMatch = regexp(inputFile1, 'ICP\s*\d+', 'match');
if ~isempty(icpMatch)
    icpIdentifier = strtrim(icpMatch{1}); % Extracted "ICP ###"
else
    icpIdentifier = "ICP_UNKNOWN"; % Default if no match
end

% ✅ Construct the new filename
%correctedFileName = sprintf('%s_%s.xlsx', icpIdentifier, fileSuffix);
%correctedFilePath = fullfile(inputPath1, correctedFileName);

% ✅ Save the corrected raw data
%writetable(correctedDataTable, correctedFilePath, 'Sheet', 'DF_MULTIPLIED', 'WriteVariableNames', true);

%disp(['✅ DF multiplied data saved to: ', correctedFilePath]);

% ✅ Define `parseValue` function at the end of the script
function val = parseValue(x)
    if isnumeric(x)  % If already numeric, return as is
        if isnan(x)  % Convert NaN values to 0
            val = 0;
        else
            val = x;
        end
    elseif ischar(x) || isstring(x)
        num = str2double(x);
        if isnan(num)  % If conversion fails (e.g., '<0.00000' or empty), set to 0
            val = 0;
        else
            val = num;
        end
    else
        val = 0; % Default case for unexpected values (also converts NaN)
    end
end


%% ✅ STEP 5: wt% concentration calculation

% ✅ Extract numerical data from DF_MULTIPLIED
numericalData = DF_MULTIPLIED{2:end, 2:end}; % Skip headers

% ✅ Handle text values safely
if iscell(numericalData)
    numericalData = cellfun(@parseValue, numericalData, 'UniformOutput', false);
end

% ✅ Convert to numeric matrix safely
numericalData = cell2mat(numericalData);

% ✅ Extract numerical values from DILUTION (Column 1 and Column 2)
dilutionFactorCol = DILUTION{2:end, 2}; % Column 2 numerical values
dilutionMassCol = DILUTION{2:end, 1}; % Column 1 numerical values

% ✅ Convert cell values to numeric if needed
if iscell(dilutionFactorCol)
    dilutionFactorCol = cellfun(@parseValue, dilutionFactorCol, 'UniformOutput', false);
end
if iscell(dilutionMassCol)
    dilutionMassCol = cellfun(@parseValue, dilutionMassCol, 'UniformOutput', false);
end

dilutionFactorCol = cell2mat(dilutionFactorCol);
dilutionMassCol = cell2mat(dilutionMassCol);

% ✅ Ensure dimensions match before performing element-wise operations
if size(numericalData, 1) ~= size(dilutionFactorCol, 1) || size(dilutionFactorCol, 1) ~= size(dilutionMassCol, 1)
    error('❌ Mismatch: Data dimensions do not match for calculations.');
end

% ✅ Perform element-wise multiplication and division
finalData = (numericalData .* dilutionFactorCol) ./ dilutionMassCol;

% ✅ Identify special element columns based on partial match
specialElements = {'Li', 'Na', 'K', 'Be', 'Tb', 'Si'};
columnNames = DF_MULTIPLIED.Properties.VariableNames(2:end);
unitRow = repmat({''}, 1, width(finalData) + 1); % First column empty

for colIdx = 1:length(columnNames)
    if any(contains(columnNames{colIdx}, specialElements, 'IgnoreCase', true))
        unitRow{colIdx + 1} = '[mg/g]';
    else
        unitRow{colIdx + 1} = '[ppm wt%]';
        finalData(:, colIdx) = finalData(:, colIdx) * 1000; % Convert ppm wt% values
    end
end

wtUnitRow = unitRow ;

% ✅ Convert back to table, retaining headers
finalDataTable = array2table(finalData, 'VariableNames', columnNames);

% ✅ Convert all numeric columns to **cell format** for consistent table merging
for colIdx = 1:width(finalDataTable)
    finalDataTable.(colIdx) = num2cell(finalDataTable{:, colIdx});
end

% ✅ Extract Sample ID column (First column remains unchanged)
sampleIDColumn = DF_MULTIPLIED(2:end, 1);

% ✅ Convert unit row to a table
unitTable = cell2table(unitRow, 'VariableNames', [sampleIDColumn.Properties.VariableNames, finalDataTable.Properties.VariableNames]);

% ✅ Concatenate unit row with final calculated data
finalDataTable = [unitTable; [sampleIDColumn, finalDataTable]];

% ✅ Store the final data in MATLAB workspace
assignin('base', 'CONC_wt', finalDataTable);

% ✅ Determine if file is Major or Trace
if contains(fileName2, "Major", 'IgnoreCase', true)
    fileSuffix = "Major";
elseif contains(fileName2, "Trace", 'IgnoreCase', true)
    fileSuffix = "Trace";
else
    error('❌ Error: File name must contain "Major" or "Trace" to define output file name.');
end

% ✅ Generate filename based on ICP identifier and Major/Trace classification
correctedFileName = sprintf('%s_CONC wt%%_%s.xlsx', extractedICP, fileSuffix);
correctedFilePath = fullfile(inputPath1, correctedFileName);

% ✅ Save the final calculated data
%writetable(finalDataTable, correctedFilePath, 'Sheet', 'CONC_wt', 'WriteVariableNames', true);

%disp(['✅ wt% concentration data saved to: ', correctedFilePath]);


%% ✅ STEP 6: Averaging wt% concentration values

% ✅ Extract Sample ID column
sampleIDColumn = finalDataTable{2:end, 1}; % Extract as array

% ✅ Extract numeric data and ensure proper conversion
numericData = finalDataTable{2:end, 2:end};
if iscell(numericData)
    numericData = cellfun(@parseValue, numericData, 'UniformOutput', false);
end
numericData = cell2mat(numericData);

% ✅ Extract row 2 from CONC_wt and ensure it's a 2D cell array
unitRowAvg = table2cell(finalDataTable(1, 2:end));
unitRowAvg = [repmat({''}, 1, 1), unitRowAvg]; % Add empty first column

% ✅ Identify rows with similar IDs differing by a number
uniqueIDs = regexprep(sampleIDColumn, '\d+$', ''); % Remove trailing numbers
[uniqueGroups, ~, groupIdx] = unique(uniqueIDs, 'stable');

% ✅ Initialize averaged data storage
averagedData = zeros(length(uniqueGroups), size(numericData, 2));

for i = 1:length(uniqueGroups)
    rowsToAverage = groupIdx == i;
    averagedData(i, :) = mean(numericData(rowsToAverage, :), 1, 'omitnan');
end

% ✅ Convert back to table format
averagedDataTable = array2table(averagedData, 'VariableNames', columnNames);

% ✅ Convert all numeric columns to **cell format** for consistent table merging
for colIdx = 1:width(averagedDataTable)
    averagedDataTable.(colIdx) = num2cell(averagedDataTable{:, colIdx});
end

% ✅ Assign new Sample ID column
sampleIDVarName = DF_MULTIPLIED.Properties.VariableNames{1}; % Use original variable name
averagedSampleID = table(uniqueGroups, 'VariableNames', {sampleIDVarName});
averagedDataTable = [averagedSampleID, averagedDataTable];

% ✅ Insert row 2 into the averaged table
unitTableAvg = cell2table(unitRowAvg, 'VariableNames', [sampleIDVarName, columnNames]);
averagedDataTable = [unitTableAvg; averagedDataTable];

% ✅ Store averaged data in MATLAB workspace
assignin('base', 'CONC_wt_avg', averagedDataTable);

% ✅ Save averaged data
averagedFileName = sprintf('%s_CONC wt%%_AVG_%s.xlsx', extractedICP, fileSuffix);
averagedFilePath = fullfile(inputPath1, averagedFileName);
writetable(averagedDataTable, averagedFilePath, 'Sheet', 'CONC_wt_avg', 'WriteVariableNames', true);

%disp(['✅ Averaged concentration data saved to: ', averagedFilePath]);



%% ✅ STEP 7: Calculating sample standard deviation of wt% concentration values

% ✅ Extract Sample ID column (reuse same extraction as before)
sampleIDColumn = finalDataTable{2:end, 1};  % Extract as array

% ✅ Extract numeric data and ensure proper conversion
numericData = finalDataTable{2:end, 2:end};
if iscell(numericData)
    numericData = cellfun(@parseValue, numericData, 'UniformOutput', false);
end
numericData = cell2mat(numericData);

% ✅ Identify rows with similar IDs differing by a number
uniqueIDs = regexprep(sampleIDColumn, '\d+$', ''); % Remove trailing numbers
[uniqueGroups, ~, groupIdx] = unique(uniqueIDs, 'stable');

% ✅ Initialize standard deviation data storage
stdevData = zeros(length(uniqueGroups), size(numericData, 2));

% ✅ Compute sample standard deviation for each group
for i = 1:length(uniqueGroups)
    rowsToStd = groupIdx == i;
    % Using std with 'omitnan' to handle missing values (0 flag computes sample std)
    stdevData(i, :) = std(numericData(rowsToStd, :), 0, 1, 'omitnan');
end

% ✅ Convert computed stdev values back to table format
stdevDataTable = array2table(stdevData, 'VariableNames', columnNames);

% ✅ Convert all numeric columns to cell format for consistent table merging
for colIdx = 1:width(stdevDataTable)
    stdevDataTable.(colIdx) = num2cell(stdevDataTable{:, colIdx});
end

% ✅ Assign new Sample ID column using the original variable name
sampleIDVarName = DF_MULTIPLIED.Properties.VariableNames{1};
stdevSampleID = table(uniqueGroups, 'VariableNames', {sampleIDVarName});
stdevDataTable = [stdevSampleID, stdevDataTable];

% ✅ Insert row 2 into the stdev table (using the same unit row from averaging)
unitTableStd = cell2table(unitRowAvg, 'VariableNames', [sampleIDVarName, columnNames]);
stdevDataTable = [unitTableStd; stdevDataTable];

% ✅ Save the standard deviation data to an Excel file
stdevFileName = sprintf('%s_CONC wt%%_STD_%s.xlsx', extractedICP, fileSuffix);
stdevFilePath = fullfile(inputPath1, stdevFileName);
writetable(stdevDataTable, stdevFilePath, 'Sheet', 'CONC_wt_std', 'WriteVariableNames', true);

%disp(['✅ Standard deviation concentration data saved to: ', stdevFilePath]);

% Save standard deviation table to MATLAB workspace variable
CONC_wt_STD = stdevDataTable;


%% ✅ STEP 12: Extracting values for propagated error calculation
% Reuse the same file as selected in STEP 1
inputFile = inputFile1;
inputPath = inputPath1;
inputFilePath = inputFilePath1;
fileName = fileName1;
extractedName = extractedICP;

% ✅ Extract "ICP ####" (with space if present)
[~, fileName, ~] = fileparts(inputFile);
tokens = regexp(fileName, 'ICP\s*\d+', 'match'); % Match "ICP" followed by optional spaces and numbers
if ~isempty(tokens)
    extractedName = strtrim(tokens{1});
else
    extractedName = 'ICP UNKNOWN';
end

% ✅ Detect import options and set the header row (Row 11) and data row (Row 12)
opts = detectImportOptions(inputFilePath, 'PreserveVariableNames', true);
opts.VariableNamesRange = 'A11';  % Row 11 as headers
opts.DataRange = 'A12';           % Data starts at row 12

opts.SelectedVariableNames = opts.VariableNames;  % Read all columns
dataTable = readtable(inputFilePath, opts);

% ✅ Treat all columns as text to avoid unwanted type conversion
opts = setvartype(opts, 'char');
opts.SelectedVariableNames = opts.VariableNames;
dataTable = readtable(inputFilePath, opts);

% ✅ Display detected column names
%fprintf('Total Columns Detected: %d\n', width(dataTable));
%disp('✅ Final Column Headers in Data Table (from Row 11):');
%disp(dataTable.Properties.VariableNames);

% ✅ Define the keywords to search for in column headers
% (Added 'DF(OES)' so that a header like "C1C2, DF(OES)" gets renamed to "DF(OES)"
% and now added 'DF(MS)' so that a header like "C1C3, DF(MS)" gets renamed to "DF(MS)")
keywords = {'m(salt)', 'm(tube)', 'm(tube+Tb)', 'm(tube+Tb+conc.)', 'V(conc.)', ...
    'V(Tb)', 'V(MD)', 'C(MD)', 'm/(tube)', 'm(tube+MD)', 'm(tube+MD+H2O)', ...
    'V/(MD)', 'V(H2O)', 'V(OES)', 'C(OES)', 'm//(tube)', 'm(tube+OES)', ...
    'm(tube+OES+5v% HNO3)', 'V/(OES)', 'V(5v% HNO3)', 'V(MS)', 'C(MS)', 'DF(OES)', 'DF(MS)'};

% ✅ Normalize column headers and keywords for matching
normalizedColumnNames = lower(strrep(strrep(strtrim(dataTable.Properties.VariableNames), " ", "_"), ".", ""));
normalizedKeywords = lower(strrep(strrep(strtrim(keywords), " ", "_"), ".", ""));
matchedColumns = contains(normalizedColumnNames, normalizedKeywords, 'IgnoreCase', true);

% ✅ Ensure column 3 is always included (if it exists)
if width(dataTable) >= 3
    matchedColumns(3) = true;
end

if ~any(matchedColumns)
    disp('❌ No matching columns found. Please check the displayed column names and modify the keyword list accordingly.');
    return;
end

% ✅ Extract the matched columns
extractedData = dataTable(:, matchedColumns);

% 🔄 Rename column headers to match the keywords (if found)
newVarNames = extractedData.Properties.VariableNames;
for i = 1:length(newVarNames)
    origHeader = newVarNames{i};
    for j = 1:length(keywords)
        % If the original header contains the keyword (ignoring case), replace it.
        if contains(origHeader, keywords{j}, 'IgnoreCase', true)
            newVarNames{i} = keywords{j};
            break; % Use the first matching keyword.
        end
    end
end
extractedData.Properties.VariableNames = newVarNames;

% ✂️ Trim to rows containing numerical values only (always keeping the first row)
dataCells = table2cell(extractedData);
lastNumericRow = height(extractedData);
for i = 2:size(dataCells,1)
    rowData = dataCells(i, :);
    if ~any(cellfun(@(x) ~isnan(str2double(x)), rowData))
        lastNumericRow = i - 1;
        break;
    end
end
if lastNumericRow < 1
    error('❌ No rows with numerical values found.');
end
extractedData = extractedData(1:lastNumericRow, :);


%% ✅ STEP 13: Calculating sigma values of dilution factors 
% 🧮 Calculate sigma_DF(OES)
% Formula:
% sigma_DF(OES) = DF(OES) * sqrt((0.0005365/V(OES))^2 + (0.0003408/V/(MD))^2)
%
DF_OES_vals = str2double(extractedData.("DF(OES)"));
V_OES_vals = str2double(extractedData.("V(OES)"));
V_2MD_vals  = str2double(extractedData.("V/(MD)"));

sigma_DF_OES = DF_OES_vals .* sqrt((0.00179 ./ V_OES_vals).^2 + (0.00114 ./ V_2MD_vals).^2);

% Append the calculated sigma_DF(OES) as a new column
extractedData.sigma_DF_OES = sigma_DF_OES;

% 🧮 Calculate sigma_DF(MS)
% Formula:
% sigma_DF(MS) = DF(MS) * sqrt((sigma_DF_OES/DF(OES))^2 + (0.0005859/V(MS))^2 + (0.0004143/V/(OES))^2)
%
DF_MS_vals    = str2double(extractedData.("DF(MS)"));
V_MS_vals     = str2double(extractedData.("V(MS)"));
V_2OES_vals   = str2double(extractedData.("V/(OES)"));

sigma_DF_MS = DF_MS_vals .* sqrt((sigma_DF_OES ./ DF_OES_vals).^2 + (0.00195 ./ V_MS_vals).^2 + (0.00138 ./ V_2OES_vals).^2);

% Append the calculated sigma_DF(MS) as a new column
extractedData.sigma_DF_MS = sigma_DF_MS;



% ✅ Save the extracted data (with calculations) and store in workspace
% Determine fileSuffix based on the raw data file name (fileName2)
if contains(fileName2, "Major", 'IgnoreCase', true)
    fileSuffix = "Major";
elseif contains(fileName2, "Trace", 'IgnoreCase', true)
    fileSuffix = "Trace";
else
    error('❌ Error: File name must contain "Major" or "Trace" to define output file name.');
end

% Generate filename based on the extracted ICP identifier and the file suffix.
% (Here, we assume that inputPath1 is the folder in which to save the output.)
correctedFileName = sprintf('%s_SIGMA_%s.xlsx', extractedICP, fileSuffix);
correctedFilePath = fullfile(inputPath1, correctedFileName);

assignin('base', 'SIGMA', extractedData);
%disp('✅ SIGMA has been stored in MATLAB workspace.');


% Save the propagated error data (sigma values) to the output Excel file.
%writetable(extractedData, correctedFilePath, 'Sheet', 'SIGMA', 'WriteVariableNames', true);
%disp(['✅ Propagated error data saved to: ', correctedFilePath]);
%disp('✅ Extracted data stored in MATLAB workspace as: SIGMA');



%% ✅ STEP 14: Calculating sigma_C_raw
% 🧮 Formula: RSD_C_raw = (RAWRSD values) / 100 (for each numeric column except Sample ID)

% ✅ Convert RAWRSD_filt to a cell array
rsdCell = table2cell(RAWRSD_filt);

% ✅ Extract the first row (unit row) and subsequent rows (data)
unitRow_RSD = rsdCell(1, :);  % Preserve the original unit row
dataRows_RSD = rsdCell(2:end, :);  % Data rows

% ✅ Create a new unit row with `[ppb]` labels instead of [%]
unitRow_ppb = repmat({''}, 1, size(dataRows_RSD, 2)); % Initialize empty row
unitRow_ppb(2:end) = {'[ppb]'};  % Assign "[ppb]" to all numeric columns

% ✅ Debug: Display first few rows BEFORE conversion
%disp('🔍 First few rows BEFORE conversion:');
%disp(dataRows_RSD(1:min(5, size(dataRows_RSD, 1)), :));

% ✅ Convert numeric values properly
computedData = dataRows_RSD;  % Initialize computed data storage
for col = 2:size(dataRows_RSD, 2)  % Skip first column (Sample ID)
    colData = dataRows_RSD(:, col);  % Extract column data

    % Convert text values to numeric safely, and divide by 100
    numericVals = cellfun(@(x) parseToZero(x) / 100, colData, 'UniformOutput', false);

    % Store the new values as cells
    computedData(:, col) = numericVals;
end

% ✅ Debug: Display first few rows AFTER conversion
%disp('🔍 First few rows AFTER conversion:');
%disp(computedData(1:min(5, size(computedData, 1)), :));

% ✅ Replace unit row values with "[fraction]" instead of "[%]"
unitRow_fraction = repmat({''}, 1, size(unitRow_RSD, 2)); % Create an empty row
unitRow_fraction(2:end) = {'[fraction]'}; % Set "[fraction]" for all numeric columns

% ✅ Reassemble the new table with updated unit row
finalComputedCell = [unitRow_fraction; computedData];


% ✅ Get original variable names from RAWRSD_filt.
origVarNames = RAWRSD_filt.Properties.VariableNames;

% ✅ Add prefix to variable names (excluding Sample ID column)
newVarNames = origVarNames;
for col = 2:length(newVarNames)
    newVarNames{col} = ['sigma_C_raw_', newVarNames{col}];
end

% ✅ Create the final computed table.
SIGMA_C_RAW = cell2table(finalComputedCell, 'VariableNames', newVarNames);

% ✅ Determine fileSuffix based on fileName2 (from STEP 2)
if contains(fileName2, "Major", 'IgnoreCase', true)
    fileSuffix = "Major";
elseif contains(fileName2, "Trace", 'IgnoreCase', true)
    fileSuffix = "Trace";
else
    error('❌ Error: File name must contain "Major" or "Trace" to define output file name.');
end

% ✅ Generate a new filename based on extracted ICP identifier
%sigmaFileName = sprintf('%s_SIGMA_C_RAW_%s.xlsx', extractedICP, fileSuffix);
%sigmaFilePath = fullfile(inputPath1, sigmaFileName);

% ✅ Save the computed sigma_C_raw values to a new Excel file.
%writetable(SIGMA_C_RAW, sigmaFilePath, 'Sheet', 'SIGMA_C_RAW', 'WriteVariableNames', true);
%disp(['✅ Sigma_C_RAW data saved to: ', sigmaFilePath]);

% ✅ Store the computed table in MATLAB workspace.
assignin('base', 'SIGMA_C_RAW', SIGMA_C_RAW);
%disp('✅ SIGMA_C_RAW stored in MATLAB workspace.');






%% ✅ STEP 15: Calculating sigma value of DF multiplied concentration data
% 🧮 Calculate sigma_C_DF
% Formula: sigma_C_DF = C_DF * sqrt((SIGMA_C_RAW)^2 + ((sigma_DF_OES/MS) / DF_2)^2)

% ✅ 1. Extract C_DF from DF_MULTIPLIED (skip headers)
C_DF = DF_MULTIPLIED{2:end, 2:end};  % Extract numeric values

% Convert to numeric if needed
if iscell(C_DF)
    C_DF = cellfun(@parseToZero, C_DF, 'UniformOutput', false);
    C_DF = cell2mat(C_DF);
end

% ✅ 2. Extract DF_MS/OES and sigma_DF_MS/OES based on Major/Trace condition
% ✅ Determine column headers based on file type
if contains(fileName2, "Major", 'IgnoreCase', true)
    DF_column = "DF(MS)";
    sigma_column = "sigma_DF_MS"; % Already in MATLAB space
    fileSuffix = "Major";
elseif contains(fileName2, "Trace", 'IgnoreCase', true)
    DF_column = "DF(OES)";
    sigma_column = "sigma_DF_OES"; % Already in MATLAB space
    fileSuffix = "Trace";
else
    error('❌ Error: File name must contain "Major" or "Trace" to define output file name.');
end

% ✅ Extract DF_MS/OES from SIGMA
DF_2 = SIGMA{2:end, strcmp(SIGMA.Properties.VariableNames, DF_column)};
sigma_DF_2 = evalin('base', sigma_column); % Fetch directly from MATLAB workspace

% ✅ Convert to numeric if needed
DF_2 = convertToNumeric(DF_2);
sigma_DF_2 = convertToNumeric(sigma_DF_2);

% ✅ Debugging Output
%disp(['✅ Extracted Column for DF: ', DF_column]);
%disp('🔍 First few values of DF_MS/OES:');
%disp(DF_2(1:min(5, end))); % Show first few values for debugging

%disp(['✅ Extracted Column for Sigma DF: ', sigma_column]);
%disp('🔍 First few values of sigma_DF_MS/OES:');
%disp(sigma_DF_2(1:min(5, end))); % Show first few values for debugging

% ✅ 3. Fetch SIGMA_C_RAW directly from MATLAB workspace (skip headers)
SIGMA_C_RAW = evalin('base', 'SIGMA_C_RAW'); % Get from workspace
RSD_Craw = SIGMA_C_RAW{2:end, 2:end}; % Extract numeric values

% Convert to numeric if needed
if iscell(RSD_Craw)
    RSD_Craw = cellfun(@parseToZero, RSD_Craw, 'UniformOutput', false);
    RSD_Craw = cell2mat(RSD_Craw);
end

% ✅ 4. Compute sigma_C_DF using the formula, selecting the correct sigma_DF

if contains(fileName2, "Major", 'IgnoreCase', true)
    sigma_DF_selected = evalin('base', 'sigma_DF_MS'); % Fetch sigma_DF_MS from MATLAB workspace
elseif contains(fileName2, "Trace", 'IgnoreCase', true)
    sigma_DF_selected = evalin('base', 'sigma_DF_OES'); % Fetch sigma_DF_OES from MATLAB workspace
else
    error('❌ Error: File name must contain "Major" or "Trace" to determine sigma_DF.');
end

% Remove first row if it contains NaN
if isnan(sigma_DF_selected(1))
    sigma_DF_selected = sigma_DF_selected(2:end, :);
end


% Convert to numeric if needed
sigma_DF_selected = convertToNumeric(sigma_DF_selected);

% ✅ Compute sigma_C_DF using the appropriate sigma_DF
sigma_C_DF = C_DF .* sqrt((RSD_Craw).^2 + (sigma_DF_selected ./ DF_2).^2);


% ✅ 5. Convert back to table format, retaining headers
sigma_C_DF_Table = array2table(sigma_C_DF, 'VariableNames', DF_MULTIPLIED.Properties.VariableNames(2:end));

% ✅ 6. Convert numeric columns to **cell format** for consistency
for colIdx = 1:width(sigma_C_DF_Table)
    sigma_C_DF_Table.(colIdx) = num2cell(sigma_C_DF_Table{:, colIdx});
end

% ✅ 7. Extract Sample ID column from DF_MULTIPLIED
sampleIDColumn = DF_MULTIPLIED(2:end, 1);

% ✅ 8. Create `[ppm]` row to match DF_MULTIPLIED format
ppmRow = repmat({''}, 1, width(sigma_C_DF_Table) + 1); % First column empty
ppmRow(2:end) = {'[ppm]'}; % Apply "[ppm]" for numeric columns

% ✅ 9. Convert `[ppm]` row to a table with correct column names
ppmTable = cell2table(ppmRow, 'VariableNames', [sampleIDColumn.Properties.VariableNames, sigma_C_DF_Table.Properties.VariableNames]);

% ✅ 10. Concatenate `[ppm]` row with sigma_C_DF data
sigma_C_DF_Table = [ppmTable; [sampleIDColumn, sigma_C_DF_Table]];

% ✅ 11. Store the final table in MATLAB workspace
assignin('base', 'SIGMA_C_DF', sigma_C_DF_Table);

% ✅ 12. Save as an Excel file
%sigmaCFileName = sprintf('%s_SIGMA_C_DF_%s.xlsx', extractedICP, fileSuffix);
%sigmaCFilePath = fullfile(inputPath1, sigmaCFileName);
%writetable(sigma_C_DF_Table, sigmaCFilePath, 'Sheet', 'SIGMA_C_DF', 'WriteVariableNames', true);

%disp(['✅ sigma_C_DF data saved to: ', sigmaCFilePath]);
%disp('✅ SIGMA_C_DF stored in MATLAB workspace.');

% ✅ Define Helper Function to Convert to Numeric
function numArray = convertToNumeric(cellArray)
    if iscell(cellArray)
        numArray = cellfun(@parseToZero, cellArray, 'UniformOutput', false);
        numArray = cell2mat(numArray);
    else
        numArray = cellArray; % Already numeric
    end
end





%% ✅ STEP 16: Calculating sigma value of wt% concentration data
% 🧮 Calculate sigma_C_wt
% Formula: sigma_C_wt = C_wt * sqrt((sigma_C_DF/C_DF)^2+(0.00171/V_MD)^2 + (1/m_salt)^2)

% 1. Extract V(MD) and m(salt) from SIGMA.
V_MD_vals = SIGMA{:, strcmp(SIGMA.Properties.VariableNames, 'V(MD)')};  
m_salt_vals = SIGMA{:, strcmp(SIGMA.Properties.VariableNames, 'm(salt)')};

% Ensure these are numeric and are column vectors.
if iscell(V_MD_vals)
    V_MD_vals = cellfun(@(x) parseToZero(x), V_MD_vals);
end
if iscell(m_salt_vals)
    m_salt_vals = cellfun(@(x) parseToZero(x), m_salt_vals);
end
V_MD_vals = V_MD_vals(2:end);  
m_salt_vals = m_salt_vals(2:end);

% 2. Compute the combined uncertainty ratio per sample/analyte.
% C_DF and sigma_C_DF are 49×17 numeric arrays.
% For V_MD_vals and m_salt_vals, replicate across 17 columns.
V_MD_mat = repmat(V_MD_vals, 1, size(C_DF, 2));   % 49×17
m_salt_mat = repmat(m_salt_vals, 1, size(C_DF, 2)); % 49×17

% Compute each term, omitting any term where the denominator is zero.
% We want to compute the terms for each sample and each analyte.
% Compute term1, term2, term3 using element-wise logical indexing:
% Compute term1, term2, and term3 elementwise.
[rows, cols] = size(C_DF);  % Both are 49×17
sqrtRatio = zeros(rows, cols);  % Preallocate result

for i = 1:rows
    for j = 1:cols
        % Term 1: use only if C_DF(i,j) is nonzero.
        if C_DF(i,j) == 0
            s1 = 0;
        else
            s1 = (sigma_C_DF(i,j) / C_DF(i,j))^2;
        end
        
        % Term 2: for sample i (same for all analytes), assume V_MD_vals(i) is nonzero.
        if V_MD_vals(i) == 0
            s2 = 0;
        else
            s2 = (0.00171 / V_MD_vals(i))^2;
        end
        
        % Term 3: omit if m_salt_vals(i) is zero.
        if m_salt_vals(i) == 0
            s3 = 0;
        else
            s3 = (1 / m_salt_vals(i))^2;
        end
        
        sqrtRatio(i,j) = sqrt(s1 + s2 + s3);
    end
end

% Debugging: For a specific sample/analyte, display intermediate values.
% For example, for sample 1, analyte 3:
%i_debug = 1; j_debug = 3;
%if C_DF(i_debug, j_debug) == 0
%    s1_debug = 0;
%else
%    s1_debug = (sigma_C_DF(i_debug, j_debug) / C_DF(i_debug, j_debug))^2;
%end
%if V_MD_vals(i_debug) == 0
%    s2_debug = 0;
%else
%    s2_debug = (0.0005133 / V_MD_vals(i_debug))^2;
%end
%if m_salt_vals(i_debug) == 0
%    s3_debug = 0;
%else
%    s3_debug = (0.3 / m_salt_vals(i_debug))^2;
%end
%fprintf('Sample %d, Analyte %d: C_DF = %f, sigma_C_DF = %f, V_MD = %f, m_salt = %f\n', ...
 %   i_debug, j_debug, C_DF(i_debug, j_debug), sigma_C_DF(i_debug, j_debug), V_MD_vals(i_debug), m_salt_vals(i_debug));
%fprintf('Computed terms: s1 = %e, s2 = %e, s3 = %e, sqrtRatio = %e\n', s1_debug, s2_debug, s3_debug, sqrtRatio(i_debug, j_debug));

% 4. Extract numeric DF-multiplied concentration data from CONC_wt.
% CONC_wt: row 1 is the unit row, column 1 is Sample ID, rows 2:end & columns 2:end are numeric.



C_wt_numeric = cell2mat(table2cell(CONC_wt(2:end, 2:end)));  % 49×17 numeric matrix.
sampleID = CONC_wt{2:end, 1};  % 49×1 cell array of Sample IDs.
unitRow = CONC_wt(1, 2:end);   % 1×17 (the original unit row for analyte columns)

% 5. Compute sigma_C_wt for each sample.
sigma_C_wt_numeric = C_wt_numeric .* sqrtRatio;  % 49×17 numeric matrix.

% 6. Rebuild the sigma_C_wt table as before
% Create the unit row as a cell array.
unitRow_cell = cell(1, width(CONC_wt));
unitRow_cell{1} = '';  % Leave Sample ID column blank.
unitRow_cell(2:end) = table2cell(CONC_wt(1, 2:end));  % Use original unit row text.

% Extract Sample ID as a cell array (using curly braces)
sampleID = CONC_wt{2:end, 1};  % 49×1 cell array

% Create data rows: combine Sample ID and computed sigma_C_wt values (converted to cells)
dataRows_cell = [sampleID, num2cell(sigma_C_wt_numeric)];  % 49×(1+17)

% Combine unit row and data rows.
finalSigmaCell = [unitRow_cell; dataRows_cell];

% Create a table with the same variable names as CONC_wt.
sigma_C_wt_table = cell2table(finalSigmaCell, 'VariableNames', CONC_wt.Properties.VariableNames);

% 7. Convert the table to a cell array of strings 
% This ensures that both the unit row and data rows are uniformly strings.
cellOut = table2cell(sigma_C_wt_table);
[rowsOut, colsOut] = size(cellOut);
for i = 1:rowsOut
    for j = 1:colsOut
        % If the cell is numeric, convert it to a string with 6 decimal places.
        if isnumeric(cellOut{i,j})
            cellOut{i,j} = num2str(cellOut{i,j}, '%.6f');
        else
            % Otherwise, leave it as is.
            cellOut{i,j} = char(cellOut{i,j});
        end
    end
end

% --- Save the cell array to Excel using writecell ---
% Determine fileSuffix based on fileName2.
if contains(fileName2, "Major", 'IgnoreCase', true)
    fileSuffix = "Major";
elseif contains(fileName2, "Trace", 'IgnoreCase', true)
    fileSuffix = "Trace";
else
    error('❌ File name must contain "Major" or "Trace" to define output file name.');
end

%outputFileName = sprintf('%s_sigma_C_wt_%s.xlsx', extractedICP, fileSuffix);
%outputFilePath = fullfile(inputPath1, outputFileName);
%writecell(cellOut, outputFilePath, 'Sheet', 'sigma_C_wt');
%disp(['✅ sigma_C_wt data saved to: ', outputFilePath]);

% --- Also store in MATLAB workspace ---
sigma_C_wt = cell2table(cellOut, 'VariableNames', CONC_wt.Properties.VariableNames);
assignin('base', 'sigma_C_wt', sigma_C_wt);




%% ✅ STEP 17: Calculating sigma value of average wt% concentration data

% 1. Extract the data rows (excluding the unit row)
dataTable = sigma_C_wt(2:end, :);  % Exclude the unit row.
% The first column is Sample ID; columns 2:end contain sigma_C_wt values as strings.

% 2. Extract Sample IDs and convert the numeric portion to a numeric matrix.
sampleIDs = dataTable{:,1};  % Extract Sample IDs as cell array of strings.
sampleIDs = cellfun(@strtrim, sampleIDs, 'UniformOutput', false);

% Convert columns 2:end to numeric matrix.
nRows = height(dataTable);
nCols = width(dataTable) - 1;  % Number of analyte columns.
sigmaNumeric = zeros(nRows, nCols);
for i = 1:nRows
    for j = 2:width(dataTable)
        sigmaNumeric(i, j-1) = str2double(dataTable{i,j});
    end
end

% 3. Group samples by their base ID (ignoring trailing digits)
% For instance, 'MDB1', 'MDB2', etc., become 'MDB'.
baseIDs = regexprep(sampleIDs, '\d+$', '');
[uniqueBaseIDs, ~, groupIndices] = unique(baseIDs, 'stable');
nGroups = numel(uniqueBaseIDs);
nAnalytes = nCols;

% 4. Compute the averaged sigma_C_wt for each group.
% For each unique base sample and each analyte:
%    sigma_avg = sqrt(sum(sigma^2)) / (number of samples in group)
avgSigma = zeros(nGroups, nAnalytes);
for g = 1:nGroups
    rowsInGroup = (groupIndices == g);
    nInGroup = sum(rowsInGroup);  % Number of samples in this group.
    % Sum squares across rows for each analyte, then take square root and divide by nInGroup.
    avgSigma(g,:) = sqrt(sum(sigmaNumeric(rowsInGroup, :).^2, 1)) / nInGroup;
end

% 5. Rebuild the averaged table.
% Use the same analyte headers as sigma_C_wt (columns 2:end).
analyteNames = sigma_C_wt.Properties.VariableNames(2:end);
avgDataTable = array2table(avgSigma, 'VariableNames', analyteNames);
% Prepend the unique base sample IDs as the first column.
avgDataTable = addvars(avgDataTable, uniqueBaseIDs, 'Before', 1, 'NewVariableNames', sigma_C_wt.Properties.VariableNames{1});

% 6. Optionally, add the original unit row from CONC_wt.
% Build a unit row cell array with the same number of columns as CONC_wt.
unitRow_cell = cell(1, width(CONC_wt));
unitRow_cell{1} = '';  % For the Sample ID column.
unitRow_cell(2:end) = table2cell(CONC_wt(1, 2:end));  % Use the original unit row for analytes.
% Convert the averaged table to a cell array.
avgDataCell = table2cell(avgDataTable);
% Combine unit row with the averaged data.
finalAvgCell = [unitRow_cell; avgDataCell];

% Convert the final cell array back to a table with the same headers as CONC_wt.
sigma_C_wt_avg_table = cell2table(finalAvgCell, 'VariableNames', CONC_wt.Properties.VariableNames);

% 7. Store and save the averaged table.
assignin('base', 'sigma_C_wt_avg', sigma_C_wt_avg_table);

% Determine fileSuffix based on fileName2.
if contains(fileName2, "Major", 'IgnoreCase', true)
    fileSuffix = "Major";
elseif contains(fileName2, "Trace", 'IgnoreCase', true)
    fileSuffix = "Trace";
else
    error('❌ Error: File name must contain "Major" or "Trace" to define output file name.');
end

outputFileName_avg = sprintf('%s_sigma_C_wt_avg_%s.xlsx', extractedICP, fileSuffix);
outputFilePath_avg = fullfile(inputPath1, outputFileName_avg);
writetable(sigma_C_wt_avg_table, outputFilePath_avg, 'Sheet', 'sigma_C_wt_avg', 'WriteVariableNames', true);
%disp(['✅ sigma_C_wt_avg data saved to: ', outputFilePath_avg]);




