# MS-Access-RelativeTablePaths
Enable relative paths for MS Access linked tables and other data sources.

This module has a helper function called CollectLinkedTalbes which will provide a formatted string containing all of the databases linked tables that can be copied and pasted directly into the body of the function LinkedTablesCollection. The default output will have absolute paths to the linked tables, you will need to make them relative using ..\ (to move up a folder) or .\ (to use the database's current directory).

Verify all of your linked tables and datasources are in the LinkedTablesCollection before calling LinkAssociatedTables.

The module's main function is LinkAssociatedTables, which can be executed by adding it to a macro named "autoexec" with a "RunCode" step or calling it in the onOpen or onLoad event of the "Display Form" set in the current database's options. 
