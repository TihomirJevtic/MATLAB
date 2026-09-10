function results = outlook_search(searchTerm, maxResults)
% OUTLOOK_SEARCH  Search every Outlook store/folder for a text term.
%
%   results = OUTLOOK_SEARCH(searchTerm)
%   results = OUTLOOK_SEARCH(searchTerm, maxResults)
%
% Connects to the running (or a new) Outlook session and performs a
% single-pass, case-insensitive substring search across ALL Outlook
% stores (accounts) and ALL folders/subfolders reachable from each
% store's root folder. For every mail item it checks:
%
%   - Sender name and sender e-mail address
%   - Recipients: To and CC
%   - Subject / title
%   - Attachment file names
%   - E-mail body
%
% The search is read-only: it never moves, modifies, marks-as-read, or
% deletes any Outlook item. Folders/items Outlook refuses to enumerate
% (e.g. permission-restricted public folders) are skipped, not treated
% as fatal errors.
%
% INPUT
%   searchTerm  - text to search for (char/string, case-insensitive).
%   maxResults  - optional cap on number of matches to collect before
%                 stopping early (default: Inf, i.e. scan everything).
%
% OUTPUT
%   results - struct array, one entry per matching mail, with fields:
%       Store            - Outlook store (account) display name
%       FolderPath       - folder path within the store, e.g.
%                          'Inbox > HSE GROUP'
%       Subject          - mail subject
%       SenderName       - display name of the sender
%       SenderEmail      - resolved SMTP address of the sender
%       To               - recipients in the To field
%       CC               - recipients in the CC field
%       ReceivedTime     - received date/time as text
%       AttachmentCount  - number of attachments
%       Attachments      - cell array of attachment file names
%       Body             - full plain-text body of the mail
%
% EXAMPLES
%   matches = outlook_search('invoice 2024');
%   matches = outlook_search('Vetrozelena', 100);
%
% See also OUTLOOK_FOLDERS, OUTLOOK_READ_MAIL, OUTLOOK_SAVE_ATTACHMENTS.

    if nargin < 1 || isempty(searchTerm)
        error('outlook_search:missingSearchTerm', ...
            'Upotreba: outlook_search(searchTerm[, maxResults])');
    end

    if nargin < 2 || isempty(maxResults)
        maxResults = Inf;
    end

    searchTermLower = lower(string(searchTerm));

    %% ==============================
    %  POVEZIVANJE SA OUTLOOKOM
    % ==============================

    fprintf('\n============================================\n');
    fprintf(' OUTLOOK PRETRAGA\n');
    fprintf('============================================\n');
    fprintf('Trazeno: %s\n\n', searchTermLower);

    try
        outlook = actxserver('Outlook.Application');
    catch ME
        error('outlook_search:connectFailed', ...
            'Ne mogu da se povezem sa Outlookom: %s', ME.message);
    end

    % Garantovano oslobadjanje COM objekta, cak i ako dodje do greske
    % ili prekida (Ctrl+C).
    cleanupObj = onCleanup(@() safeRelease(outlook)); %#ok<NASGU>

    namespace = outlook.GetNamespace('MAPI');

    results = emptyResults();

    stores = namespace.Stores;
    storeCount = stores.Count;

    fprintf('Broj Store-ova: %d\n', storeCount);

    for s = 1:storeCount

        try
            store = stores.Item(s);
            storeName = safeChar(store, 'DisplayName');
        catch
            continue;
        end

        fprintf('\n--- STORE %d/%d: %s ---\n', s, storeCount, storeName);

        try
            root = store.GetRootFolder();
        catch ME
            fprintf('  Preskacem store (nije dostupan): %s\n', ME.message);
            continue;
        end

        results = searchFolder(root, storeName, '', searchTermLower, ...
            results, maxResults);

        if length(results) >= maxResults
            fprintf('\nDostignut maxResults limit (%d).\n', maxResults);
            break;
        end
    end

    fprintf('\n============================================\n');
    fprintf('Pronadjeno: %d poruka\n', length(results));
    fprintf('============================================\n');

end


%% ============================================================
%  REKURZIVNA PRETRAGA FOLDERA
% ============================================================

function results = searchFolder(folder, storeName, parentPath, ...
    searchTermLower, results, maxResults)

    if length(results) >= maxResults
        return;
    end

    try
        folderName = char(folder.Name);
    catch
        return; % Folder nedostupan - preskoci
    end

    if isempty(parentPath)
        currentPath = folderName;
    else
        currentPath = [parentPath ' > ' folderName];
    end

    try
        items = folder.Items;
        itemCount = items.Count;
    catch
        items = [];
        itemCount = 0;
    end

    fprintf('Folder: %-40s (%d stavki)\n', currentPath, itemCount);

    for i = 1:itemCount

        if length(results) >= maxResults
            return;
        end

        try
            item = items.Item(i);

            % Obuhvati SVE mail-tipa stavke (obicna posta, receipt/report
            % varijante, S/MIME poruke, itd). Numericki 'Class' property
            % (npr. 43) nije pouzdan - iste poruke mogu vratiti razlicite
            % vrednosti (npr. 111) u zavisnosti od Outlook verzije/tipa
            % veze. 'MessageClass' je stabilan string identifikator i
            % SVE prave mail poruke pocinju sa 'IPM.Note'.
            try
                msgClass = char(item.MessageClass);
            catch
                msgClass = '';
            end

            if ~startsWith(msgClass, 'IPM.Note')
                continue;
            end

            subject     = safeChar(item, 'Subject');
            senderName  = safeChar(item, 'SenderName');
            senderEmail = getSenderEmail(item);
            toField     = safeChar(item, 'To');
            ccField     = safeChar(item, 'CC');
            body        = safeChar(item, 'Body');
            attachments = getAttachmentNames(item);

            if matchesSearch(searchTermLower, subject, senderName, ...
                    senderEmail, toField, ccField, body, attachments)

                r = emptyResults();
                r(1).Store           = storeName;
                r(1).FolderPath      = currentPath;
                r(1).Subject         = subject;
                r(1).SenderName      = senderName;
                r(1).SenderEmail     = senderEmail;
                r(1).To              = toField;
                r(1).CC              = ccField;
                r(1).ReceivedTime    = safeChar(item, 'ReceivedTime');
                r(1).AttachmentCount = length(attachments);
                r(1).Attachments     = attachments;
                r(1).Body            = body;

                results(end+1) = r; %#ok<AGROW>

                fprintf('\n  [MATCH %d] %s\n', length(results), subject);
                fprintf('    Folder: %s\n', currentPath);
                fprintf('    From:   %s <%s>\n', senderName, senderEmail);
                fprintf('    Date:   %s\n', r(1).ReceivedTime);
                fprintf('    Attachments: %d\n\n', length(attachments));
            end

        catch
            % Preskoci pojedinacnu poruku ako Outlook vrati gresku
        end
    end

    % Rekurzija kroz podfoldere
    try
        subfolders = folder.Folders;
        subfolderCount = subfolders.Count;
    catch
        subfolderCount = 0;
    end

    for j = 1:subfolderCount

        if length(results) >= maxResults
            return;
        end

        try
            subfolder = subfolders.Item(j);
            results = searchFolder(subfolder, storeName, currentPath, ...
                searchTermLower, results, maxResults);
        catch
            % Preskoci folder ako Outlook ne dozvoli pristup
        end
    end

end


%% ============================================================
%  POMOCNE FUNKCIJE
% ============================================================

function tf = matchesSearch(searchTermLower, subject, senderName, ...
    senderEmail, toField, ccField, body, attachments)

    text = lower(string(subject)) + " " + lower(string(senderName)) + " " + ...
        lower(string(senderEmail)) + " " + lower(string(toField)) + " " + ...
        lower(string(ccField)) + " " + lower(string(body));

    for a = 1:length(attachments)
        text = text + " " + lower(string(attachments{a}));
    end

    tf = contains(text, searchTermLower);

end


function results = emptyResults()

    results = struct( ...
        'Store',           {}, ...
        'FolderPath',      {}, ...
        'Subject',         {}, ...
        'SenderName',      {}, ...
        'SenderEmail',     {}, ...
        'To',              {}, ...
        'CC',              {}, ...
        'ReceivedTime',    {}, ...
        'AttachmentCount', {}, ...
        'Attachments',     {}, ...
        'Body',            {});

end


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

        % Exchange moze vratiti EX/MAPI adresu umesto SMTP adrese.
        if startsWith(email, '/')
            email = char(mail.PropertyAccessor.GetProperty( ...
                'http://schemas.microsoft.com/mapi/proptag/0x39FE001E'));
        end
    catch
        email = '';
    end

end


function attachments = getAttachmentNames(mail)

    attachments = {};

    try
        n = mail.Attachments.Count;

        for i = 1:n
            attachments{end+1} = char(mail.Attachments.Item(i).FileName); %#ok<AGROW>
        end
    catch
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
