function outlook_folders()
% OUTLOOK_FOLDERS  List every Outlook store and its full folder tree.
%
%   OUTLOOK_FOLDERS()
%
% Prints all Outlook stores (accounts) and, for each store, recursively
% prints its complete folder/subfolder tree. Useful to discover the exact
% store display name and folder path arguments needed by OUTLOOK_SEARCH,
% OUTLOOK_READ_MAIL, and OUTLOOK_SAVE_ATTACHMENTS. Read-only: does not
% modify any Outlook item or folder.
%
% EXAMPLE
%   outlook_folders();
%
% See also OUTLOOK_SEARCH, OUTLOOK_READ_MAIL, OUTLOOK_SAVE_ATTACHMENTS.

    try
        outlook = actxserver('Outlook.Application');
    catch ME
        error('outlook_folders:connectFailed', ...
            'Ne mogu da se povezem sa Outlookom: %s', ME.message);
    end

    % Garantovano oslobadjanje COM objekta, cak i ako dodje do greske.
    cleanupObj = onCleanup(@() safeRelease(outlook)); %#ok<NASGU>

    namespace = outlook.GetNamespace('MAPI');

    fprintf('\n=== OUTLOOK FOLDERS ===\n\n');

    stores = namespace.Stores;

    for s = 1:stores.Count

        try
            store = stores.Item(s);
        catch
            continue;
        end

        fprintf('\nSTORE %d: %s\n', s, safeChar(store, 'DisplayName'));
        fprintf('----------------------------------------\n');

        try
            root = store.GetRootFolder();
            printFolder(root, 0);
        catch ME
            fprintf('  Store nije dostupan: %s\n', ME.message);
        end

    end

end


function printFolder(folder, level)

    indent = repmat('    ', 1, level);

    try
        fprintf('%s%s\n', indent, char(folder.Name));
    catch
        return; % Folder nedostupan - preskoci
    end

    try
        subfolders = folder.Folders;

        for i = 1:subfolders.Count
            try
                subfolder = subfolders.Item(i);
                printFolder(subfolder, level + 1);
            catch
                % Preskoci podfolder ako Outlook ne dozvoli pristup
            end
        end
    catch
        % Neki Outlook folderi mogu biti nedostupni
    end

end


function value = safeChar(object, property)

    try
        value = char(object.(property));
    catch
        value = '';
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
