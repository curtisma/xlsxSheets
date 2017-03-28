function [SHEETS,STATUS] = xlsxSheets(filename)
%   Provides the sheet names contained in an XLSX file without using a com server,
%   even if you have Excel installed on your machine. This function runs 
%   much faster than xlsfinfo since it doesn't have to open a com server.
%
%   Note: The output parameters are different than xlsfinfo
%
%   [SHEETS,STATUS] = XLSFINFOXLSX(FILENAME) returns a cell array of strings
%   containing the names of each spreadsheet in the file. If XLSREAD cannot
%   read a particular worksheet, the corresponding cell contains an error
%   message. If XLSFINFO cannot read the file, SHEETS is a string
%   containing an error message.

% Based on a part of the code (xlsfinfoXLSX function) for xlsfinfo from Mathworks.
    % Validate filename data type
    if nargin < 1
        error(message('MATLAB:xlsfinfo:Nargin'));
    end
    if ~ischar(filename)
        error(message('MATLAB:xlsfinfo:InputClass'));
    end
    
    % Validate filename is not empty
    if isempty(filename)
        error(message('MATLAB:xlsfinfo:FileName'));
    end
    
    [~, ~, ext] = fileparts(filename);
    assert(any(strcmp(ext, matlab.io.internal.xlsreadSupportedExtensions('SupportedOfficeOpenXMLOnly'))))
    
    % Requires java to unzip xlsx files
    if ~usejava('jvm')
        error(message('MATLAB:xlsfinfo:noJVM'))
    end;
    
    % Unzip the XLSX file (a ZIP file) to a temporary location
    baseDir = tempname;
    mkdir(baseDir);
    cleanupBaseDir = onCleanup(@()rmdir(baseDir,'s'));
    unzip(filename, baseDir);
    
    docProps = fileread(fullfile(baseDir,'docProps','app.xml'));
    theMessage = '';
    matchMessage = regexp(docProps,'<Application>(?<message>Microsoft\s+(\w+\s+)?Excel)</Application>','names');
    if ~isempty(matchMessage)
        theMessage = [matchMessage.message ' Spreadsheet'];
    end
    
    workbook_xml_rels  = fileread(fullfile(baseDir, 'xl', '_rels', 'workbook.xml.rels')); 
    workbook_xml  = fileread(fullfile(baseDir, 'xl', 'workbook.xml')); 
    description = getSheetNames(workbook_xml_rels, workbook_xml);
    STATUS = theMessage;
    SHEETS = description;
    
    function sheets = getSheetNames(workbook_xml_rels, workbook_xml)
    % getSheetNames parses OpenXML to extract Worksheet names.
    %   sheets = getSheetNames(workbook_xml_rels, workbook_xml) parses the 
    %   Office Open XML code in the char array docProps to extract 
    %   Worksheet names into the cell array sheets.
    %
    %   See also xlsread, xlsfinfo
    
    % Copyright 2011-2014 The MathWorks, Inc.
    
    % Excel usually generates files with the 'Type' label first, but
    % Python usually generates file with the 'Target' label first.  We
    % account for both.
    sheetIDs = regexp(workbook_xml_rels, ...
                 ['<Relationship[^>]+Id="(?<rid>[^>]+?)"[^>]+(Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"[^>]+Target="worksheets/[^>]+?.xml"|' ...
                 'Target="worksheets/[^>]+?.xml"[^>]+Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet")[^>]*/>'], ...
                 'names' );
  
    match = regexp(workbook_xml, '<sheet[^>]+name="(?<sheetName>[^"]*)"[^>]*r:id="(?<rid>[^>]+?)"[^>]*/>|<sheet[^>]*r:id="(?<rid>[^>]+?)"[^>]*name="(?<sheetName>[^"]*)"[^>]*/>', 'names');
    
    validSheetIndices = zeros(size(sheetIDs));
    count = 1;
    
    % Match rIDs found in the header with rIDs for sheets in the file.
    % Only return the sheet names of sheet rIDs that are found in the
    % header.
    for i = 1:numel(sheetIDs)
       for j = 1:numel(match)
           if isequal(sheetIDs(i).rid, match(j).rid)
               validSheetIndices(count) = j;
               count = count + 1;
           end
       end
    end
    
    indices = sort(validSheetIndices);
    
    sheets = {match(indices).sheetName};
    
end
end

