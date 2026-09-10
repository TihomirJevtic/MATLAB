function savedFiles = outlook_save_attachments(folderPath, mailNumber, outputFolder, storeName)
% OUTLOOK_SAVE_ATTACHMENTS  Save a specific mail's attachments to disk.
%
%   savedFiles = OUTLOOK_SAVE_ATTACHMENTS(folderPath, mailNumber, outputFolder)
%   savedFiles = OUTLOOK_SAVE_ATTACHMENTS(folderPath, mailNumber, outputFolder, storeName)
%
% Reads the mail identified by folderPath/mailNumber (newest-first order)
% and copies its attachments to outputFolder. Read-only with respect to
% Outlook: the mail itself is never modified, moved, or deleted; only
% attachment copies are written to disk.
%
% INPUT
%   folderPath   - cell array of folder names from the store root, e.g.
%                  {'Inbox','HSE GROUP'}.
%   mailNumber   - 1-based index into the folder, sorted newest-first.
%   outputFolder - destination folder for saved attachments (created if
%                  it does not exist).
%   storeName    - optional display name of the Outlook store (account).
%                  Default: the first store returned by Outlook.
%
% OUTPUT
%   savedFiles - cell array of full paths to the files actually saved.
%
% EXAMPLE
%   savedFiles = outlook_save_attachments( ...
%       {'Inbox','HSE GROUP'}, 1, 'C:\BMV\MailAttachments');
%
% See also OUTLOOK_SEARCH, OUTLOOK_READ_MAIL, OUTLOOK_FOLDERS.

    %% ========================================
    %  PROVERA ARGUMENATA
    % =========================================

    if nargin < 3
        error('outlook_save_attachments:missingArgs', ...
            ['Potrebna su najmanje 3 argumenta: ' ...
             'folderPath, mailNumber, outputFolder']);
    end

    if ~iscell(folderPath)
        error('outlook_save_attachments:invalidFolderPath', ...
            'folderPath mora biti cell array, npr. {''Inbox'',''HSE GROUP''}');
    end

    %% ========================================
    %  OUTPUT FOLDER
    % =========================================

    if ~exist(outputFolder, 'dir')
        mkdir(outputFolder);
    end

    fprintf('\nOutput folder:\n%s\n\n', outputFolder);

    %% ========================================
    %  OUTLOOK
    % =========================================

    try
        outlook = actxserver('Outlook.Application');
    catch ME
        error('outlook_save_attachments:connectFailed', ...
            'Ne mogu da se povezem sa Outlookom: %s', ME.message);
    end

    % Garantovano oslobadjanje COM objekta, cak i ako dodje do greske.
    cleanupObj = onCleanup(@() safeRelease(outlook)); %#ok<NASGU>

    ns = outlook.GetNamespace('MAPI');

    if nargin < 4 || isempty(storeName)
        store = ns.Stores.Item(1);
    else
        store = findStoreByName(ns, storeName);
    end

    fprintf('Store: %s\n', safeChar(store, 'DisplayName'));

    %% ========================================
    %  FOLDER
    % =========================================

    folder = navigateToFolder(store.GetRootFolder(), folderPath);

    fprintf('Folder: %s\n', safeChar(folder, 'Name'));

    %% ========================================
    %  ITEMS
    % ========================================

    items = folder.Items;

    try
        items.Sort('[ReceivedTime]', true);
    catch
        warning('outlook_save_attachments:sortFailed', ...
            'Sortiranje po datumu nije uspelo.');
    end

    if mailNumber < 1 || mailNumber > items.Count
        error('outlook_save_attachments:mailNotFound', ...
            'Mail broj %d ne postoji. Folder ima %d stavki.', ...
            mailNumber, items.Count);
    end

    mail = items.Item(mailNumber);

    %% ========================================
    %  MAIL INFO
    % ========================================

    fprintf('\n========================================\n');
    fprintf('MAIL\n');
    fprintf('========================================\n');
    fprintf('Subject: %s\n', safeChar(mail, 'Subject'));
    fprintf('From: %s\n', safeChar(mail, 'SenderName'));
    fprintf('Date: %s\n', safeChar(mail, 'ReceivedTime'));

    %% ========================================
    %  ATTACHMENTS
    % ========================================

    try
        attachmentCount = mail.Attachments.Count;
    catch
        attachmentCount = 0;
    end

    fprintf('\nBroj priloga: %d\n', attachmentCount);

    savedFiles = {};

    if attachmentCount == 0
        fprintf('\nMail nema priloge.\n');
        return;
    end

    %% ========================================
    %  SAVE
    % ========================================

    for i = 1:attachmentCount

        try
            attachment = mail.Attachments.Item(i);
            filename = char(attachment.FileName);
        catch
            fprintf('\n%d. GRESKA: prilog nije dostupan.\n', i);
            continue;
        end

        % Uklanjanje problematicnih znakova iz imena fajla
        filename = strrep(filename, '/', '_');
        filename = strrep(filename, '\', '_');
        filename = strrep(filename, ':', '_');
        filename = strrep(filename, '*', '_');
        filename = strrep(filename, '?', '_');
        filename = strrep(filename, '"', '_');
        filename = strrep(filename, '<', '_');
        filename = strrep(filename, '>', '_');
        filename = strrep(filename, '|', '_');

        fullPath = fullfile(outputFolder, filename);

        % Ako fajl vec postoji, napravi novi naziv
        if exist(fullPath, 'file')

            [~, name, ext] = fileparts(filename);
            counter = 1;

            while true
                newName = sprintf('%s_%d%s', name, counter, ext);
                fullPath = fullfile(outputFolder, newName);

                if ~exist(fullPath, 'file')
                    break;
                end

                counter = counter + 1;
            end
        end

        fprintf('\n%d. %s\n', i, filename);
        fprintf('   Cuva se u:\n   %s\n', fullPath);

        try
            attachment.SaveAsFile(fullPath);

            if exist(fullPath, 'file')
                fprintf('   OK\n');
                savedFiles{end+1} = fullPath; %#ok<AGROW>
            else
                fprintf('   GRESKA: fajl nije napravljen.\n');
            end
        catch ME
            fprintf('   GRESKA: %s\n', ME.message);
        end

    end

    %% ========================================
    %  KRAJ
    % ========================================

    fprintf('\n========================================\n');
    fprintf('ZAVRSENO\n');
    fprintf('Sacuvano priloga: %d\n', length(savedFiles));
    fprintf('========================================\n');

end


%% =========================================
% POMOCNE FUNKCIJE
% =========================================

function value = safeChar(object, property)

    try
        value = char(object.(property));
    catch
        value = '';
    end

end


function store = findStoreByName(ns, storeName)

    store = [];

    for i = 1:ns.Stores.Count
        candidate = ns.Stores.Item(i);

        if strcmpi(safeChar(candidate, 'DisplayName'), storeName)
            store = candidate;
            return;
        end
    end

    error('outlook_save_attachments:storeNotFound', ...
        'Store nije pronadjen: %s', storeName);

end


function folder = navigateToFolder(folder, folderPath)

    for k = 1:length(folderPath)

        name = folderPath{k};
        found = false;

        for j = 1:folder.Folders.Count
            candidate = folder.Folders.Item(j);

            if strcmpi(safeChar(candidate, 'Name'), name)
                folder = candidate;
                found = true;
                break;
            end
        end

        if ~found
            error('outlook_save_attachments:folderNotFound', ...
                'Folder nije pronadjen: %s', name);
        end
    end

end


function safeRelease(outlook)
% Bezbedno oslobadjanje Outlook COM servera, poziva se uvek na kraju
% (uspesno zavrsen poziv, greska, ili Ctrl+C) preko onCleanup.

    try
        delete(outlook);
    catch
    end

end
