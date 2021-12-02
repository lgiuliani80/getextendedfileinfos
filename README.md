# Get-ExtendedProps

Extracts the most common file metadata using Microsoft Windows Shell extended properties. Especially intended (but not limited to) for Microsoft Office documents.

Syntax:

    ./Get-ExtendedProps{-From-SortedFileList}.ps1 <utf8-encoded-file-containing-the-list-of-files-to-process>

NOTE: full path of files MUST be specified in the input file.

* `Get-ExtendedProps.ps1` :  
  The input file does not need to be sorted. Especially suitable for up to 500 files.

* `Get-ExtendedProps-From-SortedFileList.ps1` :  
  The input file should be sorted alphabetically so that all files belonging to the same folder appear consecutive one another.
  Suitable for massive extended attribute extraction scenarios, with thousands of files.
  
The output is a list of PSObject(s) containing the following properties:

* `Filename` : 
* `FileCreationDate` = Creation time, as returned by the filesystem
* `FileLastWriteDate` = Last write time, as returned by the filesystem
* `FileLastAccessDate` = Last access time, as returned by the filesystem
* `FileOwner` = File owner, as returned by the filesystem
* `ShellModifyDate` = Last modification date, as returned by Windows Shell object
* `DetailsOwner` = Owner, as returned by Windows Shell object
* `DetailsAuthors` = List of authors, separated by ', '
* `DetailsTitle` = Title of the document
* `DetailsAuthorLastSave` = Author responsible of the last save operation
* `DetailsLastPrint` = Last print time
* `DetailsLastSave` = Last save time (= last time the document was saved as an action of the user of as the result of autosave). This can differ from FileLastWriteDate.
* `DetailsCreationDate` = Creation time written in the document. This could differ from FileCreationDate.
