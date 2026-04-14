%% --- (1) Step 1: File Selection ---
clear; clc; close all;

[fileName, pathName] = uigetfile('*.xlsx', 'Select the Excel data file');

% Exit if user cancels
if isequal(fileName, 0)
    disp('User cancelled file selection. Script terminated.');
    return;
end

fullFilePath = fullfile(pathName, fileName);
fprintf('File selected: %s\n', fullFilePath);


%% --- (2) Step 2: Read Sheet Names & Get User Selection ---
try
    sheetNames = sheetnames(fullFilePath);
catch ME
    errordlg(['Could not read sheet names from the file. Error: ' ME.message], 'File Error');
    return;
end

% Prompt user to select sheets for processing
[selectedIndex, isOK] = listdlg('PromptString', {'Select sheets to process:', '(These will be averaged together)'}, ...
                                'SelectionMode', 'multiple', ...
                                'ListString', sheetNames, ...
                                'InitialValue', 1:length(sheetNames), ...
                                'Name', 'Sheet Selection', ...
                                'ListSize', [300 300]);

% Exit if user cancels
if ~isOK
    disp('User cancelled sheet selection. Script terminated.');
    return;
end

selectedSheets = sheetNames(selectedIndex);


%% --- (3) Step 3: Initial Data Check & User Info ---
fprintf('Performing initial data check on sheet: "%s"\n', selectedSheets{1});

% Automated check to find the header row (looks for 'Hz')
max_check_rows = 50; % Check up to this many rows
header_row = 0;
for k = 1:max_check_rows
    % Read a single cell to check for 'Hz'
    cell_content = readcell(fullFilePath, 'Sheet', selectedSheets{1}, 'Range', ['A' num2str(k) ':A' num2str(k)]);
    if ~isempty(cell_content) && any(strcmpi(string(cell_content), "Hz"))
        header_row = k;
        break;
    end
end

if header_row == 0
    warning('Could not automatically find header row containing "Hz". Defaulting to previous setting.');
    header_lines = 38; % Fallback
else
    header_lines = header_row;
end

data_start_row = header_lines + 1;

% Inform the user about the findings
uiwait(msgbox(['Data check complete.' char(10) char(10) ...
               'The script has identified that your data starts on ROW ' num2str(data_start_row) '.' char(10) ...
               '(Header "Hz" found on row ' num2str(header_row) ').' char(10) char(10) ...
               'Click OK to proceed to processing options.'], ...
               'Data Check', 'help'));


%% --- (4) Step 4: Get Processing Options from User ---
% First, read data to get default frequency range
try
    first_sheet_data = readmatrix(fullFilePath, 'Sheet', selectedSheets{1}, 'Range', ['A' num2str(data_start_row)]);
    first_sheet_data = first_sheet_data(all(~isnan(first_sheet_data), 2), :);
    original_freq = first_sheet_data(:,1);
    min_freq_orig = min(original_freq);
    max_freq_orig = max(original_freq);
catch ME
    errordlg(['Could not read initial data for range check. Error: ' ME.message], 'Data Read Error');
    return;
end

% Create the dialog for user input
prompt = {['1. Downsample Step (Hz): (Leave empty = Original)'], ...
          ['2. Start Frequency (Hz): (Default = ' num2str(min_freq_orig) ')'], ...
          ['3. End Frequency (Hz): (Default = ' num2str(max_freq_orig) ')']};
dlgTitle = 'Processing Options';
dims = [1 60];
defInput = {'', num2str(min_freq_orig), num2str(max_freq_orig)};

% Show dialog and get user answers
answer = inputdlg(prompt, dlgTitle, dims, defInput);

% Exit if user cancels
if isempty(answer)
    disp('User cancelled processing options. Script terminated.');
    return;
end

% Interpret the user's answers
resample_step = str2double(answer{1});
freq_limit_start = str2double(answer{2});
freq_limit_end = str2double(answer{3});

% Validate inputs
if isnan(freq_limit_start) || isnan(freq_limit_end)
    errordlg('Frequency limits must be valid numbers.', 'Input Error');
    return;
end


%% --- (5) Step 5: Main Processing Engine ---
all_runs_sac = [];
frequency_vector = [];

fprintf('Starting main processing engine...\n');
for i = 1:length(selectedSheets)
    current_sheet = selectedSheets{i};
    fprintf(' -> Processing sheet: %s\n', current_sheet);
    
    data = readmatrix(fullFilePath, 'Sheet', current_sheet, 'Range', ['A' num2str(data_start_row)]);
    data = data(all(~isnan(data), 2), :);
    
    if isempty(frequency_vector)
        frequency_vector = data(:, 1);
    end
    all_runs_sac(:, i) = data(:, 2);
end

% --- Averaging ---
average_sac = mean(all_runs_sac, 2);
original_data = table(frequency_vector, average_sac);

% --- Apply Frequency Limits ---
% Keep only the data within the user-defined frequency range
rows_to_keep = original_data.frequency_vector >= freq_limit_start & original_data.frequency_vector <= freq_limit_end;
processed_data = original_data(rows_to_keep, :);
fprintf(' -> Applied frequency limits: %g Hz to %g Hz.\n', freq_limit_start, freq_limit_end);

% --- Apply Downsampling if requested ---
if ~isnan(resample_step) && resample_step > 0
    fprintf(' -> Downsampling data to %g Hz steps...\n', resample_step);
    
    new_freq_vector = (ceil(freq_limit_start) : resample_step : floor(freq_limit_end))';
    
    % Interpolate the SAC values at the new frequency points
    resampled_sac = interp1(processed_data.frequency_vector, processed_data.average_sac, new_freq_vector, 'linear');
    
    final_data_table = table(new_freq_vector, resampled_sac, ...
        'VariableNames', {'Frequency_Hz', 'Average_SAC'});
else
    fprintf(' -> No downsampling requested. Using original frequency steps.\n');
    final_data_table = table(processed_data.frequency_vector, processed_data.average_sac, ...
        'VariableNames', {'Frequency_Hz', 'Average_SAC'});
end

%% --- (6) Step 6: Save the Final Data ---
output_folder = 'Processed_Data';
if ~exist(output_folder, 'dir')
   mkdir(output_folder);
end

% Create a clean filename from the original Excel file name
[~, baseName, ~] = fileparts(fileName);
output_filename = fullfile(output_folder, ['Processed_', baseName, '.csv']);

writetable(final_data_table, output_filename);

fprintf('================================================\n');
fprintf('Processing complete!\n');

% Final confirmation message
uiwait(msgbox(['Successfully saved processed data to:' char(10) output_filename], 'Success!', 'help'));
