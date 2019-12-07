%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
% Normalize volumes of .wav sound files attached in a .pptx file          %
%                                                                         %
% Coded by D. Kitamura (d-kitamura@ieee.org)                              %
%                                                                         %
% --- HOW TO USE ---                                                      %
% 1. Put .pptx file to the current directory                              %
% 2. Type "pptxWavNormalize" to MATLAB command window                     %
% 3. Type .ppt file name                                                  %
% 4. Type a normalization value between 0 and 1 (0.95 is preferred)       %
% 5. Normalized .pptx file is generated in the same directory             %
% 6. Enjoy your presentation!                                             %
%                                                                         %
% See also:                                                               %
% http://d-kitamura.net                                                   %
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

clear;
close all;

% Read ppt file and set normalization value
dir = input('Enter name of .pptx file: ', 's');
if contains(dir, '.pptx')
    dir = strrep(dir, '.pptx', '');
end
normCoef = input('Enter normalization value (0 to 1, 0.95 is preferred): ');
if normCoef < 0 || normCoef > 1
    error('Normalization value must be in 0 to 1.\n');
end
fprintf('Processing...\n');

% Unzip ppt file and find audio files
zipName = sprintf('%s.pptx', dir);
inFileNames = unzip(zipName, dir);
ind = regexp(inFileNames, '.+\.wav', 'ignorecase'); % find cell that has '.wav' or '.WAV'
x = find(cellfun('isempty', ind)); % find empty cell elements
ind(x) = {0}; % replace empty cell to zero
wavFiles = inFileNames(find(cell2mat(ind)));
for i=1:size(wavFiles, 2)
    fileName = sprintf('./%s', wavFiles{1,i});
    info = audioinfo(fileName);
    wavSignal = audioread(fileName);
    wavSignal = (wavSignal./max(max(abs(wavSignal)))) * normCoef; % normalization
    audiowrite(fileName,wavSignal,info.SampleRate, 'BitsPerSample', info.BitsPerSample);
end

% Output normalized ppt file and delete zip file
outFileNames = strrep(inFileNames, sprintf('%s\\', dir), ''); % remove dir path from fileNames
cd(dir);
outFileName = sprintf('%s_normalized', dir);
zip(sprintf('../%s.zip', outFileName), outFileNames); % make zip file
cd('../');
movefile(sprintf('%s.zip', outFileName),sprintf('%s.pptx', outFileName), 'f'); % rename '.zip' to '.pptx'
rmdir(dir, 's'); % delete residual folder
fprintf('Output normalized file as "%s.pptx".\n', outFileName);
clear;
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%% EOF %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%