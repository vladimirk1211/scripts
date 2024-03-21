# PowerShell script for renaming AppPool in IIS

If there are AppPool in IIS and this AppPool associated with some App that you can not use "rename" for AppPool in IIS.
For issuing this problem i created script on PowerShell.

I thing that this procedure correctly renames AppPool Original to AppPool New in IIS:
1. Checking existing AppPool Original and AppPool New.
1.1 AppPool Original must exist.
1.2 AppPool New must not exist.
2. Making a list of Apps that are associated with AppPool Original, if any there are.
3. Remember status of AppPool Original.
4. Exporting AppPool Original as xml to a temporary file.
5. Renaming AppPool Original in the temporary file to AppPool New.
6. Creating AppPool New by exporting a temporary file.
7. Stoping AppPool Original.
8. If there are Apps associated with AppPool Original, then reconfiguring these Apps to AppPool New.
9. Starting AppPool New.
10. Removing AppPool Original.

And script RenamePool.ps1 doing this.
