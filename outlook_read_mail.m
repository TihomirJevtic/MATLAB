function mailData = outlook_read_mail(folderPath, mailNumber, storeName)
% OUTLOOK_READ_MAIL  Read one specific Outlook mail (full body + attachment list).
%
%   mailData = OUTLOOK_READ_MAIL()
%   mailData = OUTLOOK_READ_MAIL(folderPath)
%   mailData = OUTLOOK_READ_MAIL(folderPath, mailNumber)
%   mailData = OUTLOOK_READ_MAIL(folderPath, mailNumber, storeName)
%
% Opens a specific folder (newest items first) and returns the full
% content of one mail, identified by its position in that sorted list.
% Intended to inspect a single mail after finding it with OUTLOOK_SEARCH.
% Read-only: does not modify, move, or delete the mail.
%
% INPUT
%   folderPath - cell array of folder names describing the path from the
%                store root, e.g. {'Inbox','HSE GROUP'}.
%                Default: {'Inbox','HSE GROUP'}.
%   mailNumber - 1-based index into the folder, sorted newest-first.
%                Default: 1 (the newest mail).
%   storeName  - display name of the Outlook store (account) to use.
%                Default: the first store returned by Outlook.
%
% OUTPUT
%   mailData - struct with fields: Subject, SenderName, SenderEmail,
%              ReceivedTime, To, CC, Body, Attachments (cell array of
%              file names).
%
% EXAMPLE
%   mailData = outlook_read_mail({'Inbox','HSE GROUP'}, 1);
%
% See also OUTLOOK_SEARCH, OUTLOOK_FOLDERS, OUTLOOK_SAVE_ATTACHMENTS.

    if nargin < 1 || isempty(folderPath)
        folderPath = {'Inbox', 'HSE GROUP'};
    end

    if ~iscell(folderPath)
        error('outlook_read_mail:invalidFolderPath', ...
            'folderPath mora biti cell array, npr. {''Inbox'',''HSE GROUP''}');
    end

    if nargin < 2 || isempty(mailNumber)
        mailNumber = 1;
    end

    %% Povezivanje sa Outlookom

    try
        outlook = actxserver('Outlook.Application');
    catch ME
        error('outlook_read_mail:connectFailed', ...
            'Ne mogu da se povezem sa Outlookom: %s', ME.message);
    end

    % Garantovano oslobadjanje COM objekta, cak i ako dodje do greske.
    cleanupObj = onCleanup(@() safeRelease(outlook)); %#ok<NASGU>

    ns = outlook.GetNamespace('MAPI');

    if nargin < 3 || isempty(storeName)
        store = ns.Stores.Item(1);
    else
        store = findStoreByName(ns, storeName);
    end

    fprintf('\nSTORE: %s\n', safeChar(store, 'DisplayName'));

    %% Pronadji folder

    folder = navigateToFolder(store.GetRootFolder(), folderPath);

    fprintf('FOLDER: %s\n', safeChar(folder, 'Name'));

    %% Outlook stavke

    items = folder.Items;

    try
        items.Sort('[ReceivedTime]', true);
    catch
        warning('outlook_read_mail:sortFailed', ...
            'Sortiranje po ReceivedTime nije uspelo.');
    end

    count = items.Count;

    fprintf('BROJ STAVKI: %d\n', count);

    if mailNumber < 1 || mailNumber > count
        error('outlook_read_mail:mailNotFound', ...
            'Trazeni broj poruke (%d) ne postoji. Folder ima %d stavki.', ...
            mailNumber, count);
    end

    %% Izabrani mail

    mail = items.Item(mailNumber);

    %% Osnovni podaci

    mailData = struct();

    mailData.Subject      = safeChar(mail, 'Subject');
    mailData.SenderName   = safeChar(mail, 'SenderName');
    mailData.SenderEmail  = getSenderEmail(mail);
    mailData.ReceivedTime = safeChar(mail, 'ReceivedTime');
    mailData.To           = safeChar(mail, 'To');
    mailData.CC           = safeChar(mail, 'CC');
    mailData.Body         = safeChar(mail, 'Body');
    mailData.Attachments  = {};

    %% Prikaz

    fprintf('\n========================================\n');
    fprintf('MAIL\n');
    fprintf('========================================\n');

    fprintf('From: %s\n', mailData.SenderName);
    fprintf('Email: %s\n', mailData.SenderEmail);
    fprintf('Date: %s\n', mailData.ReceivedTime);
    fprintf('To: %s\n', mailData.To);
    fprintf('CC: %s\n', mailData.CC);
    fprintf('Subject: %s\n', mailData.Subject);

    fprintf('\n----------------------------------------\n');
    fprintf('SADRZAJ MAILA\n');
    fprintf('----------------------------------------\n\n');

    fprintf('%s\n', mailData.Body);

    %% Prilozi

    fprintf('\n========================================\n');
    fprintf('PRILOZI\n');
    fprintf('========================================\n');

    try
        attachmentCount = mail.Attachments.Count;
    catch
        attachmentCount = 0;
    end

    fprintf('Broj priloga: %d\n\n', attachmentCount);

    for i = 1:attachmentCount
        try
            filename = char(mail.Attachments.Item(i).FileName);
            fprintf('%d. %s\n', i, filename);
            mailData.Attachments{i} = filename; %#ok<AGROW>
        catch
            % Preskoci prilog ako Outlook ne dozvoli pristup
        end
    end

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


function email = getSenderEmail(mail)

    try
        email = char(mail.SenderEmailAddress);

        % Exchange / MAPI adresa
        if startsWith(email, '/')
            email = char(mail.PropertyAccessor.GetProperty( ...
                'http://schemas.microsoft.com/mapi/proptag/0x39FE001E'));
        end
    catch
        email = '';
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

    error('outlook_read_mail:storeNotFound', ...
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
            error('outlook_read_mail:folderNotFound', ...
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
