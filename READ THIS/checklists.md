Checklists
==========

Commit
------

* Save vb6 project (Menu->File->Save Project)
* Build project (Menu->File->Make Soldat PolyWorks.exe)
* Move any needed files to the installer directory (/Soldat PolyWorks.exe into /pwinstall)
* Compile installer and test it (open /pwinstall/pw.nsi with makensisw.exe)
* Test everything new/changed/fixed
* Commit files and changelog
* Done!


Release
-------

* Set version number in vb6 project
* Save vb6 project
* Build and move any needed files to the installer directory
* Open "Soldat PolyWorks.exe" in resource editor and replace the icon with "pwnew.ico"
* Set version number in readme
* Set update date in readme
* Add changes to readme
* Update help file
* Set version number in installer script
* Compile installer and install
* Compress installed files to zip file
* Uninstall
* Commit files and changelog
* Enter modify in Soldat PolyWorks topic
* Change version number in topic title
* Change version number in topic heading and installer and zip file link texts
* Remove changes from todo list and create a new changes block
* Upload new installer and zip file to original locations
* Save modification to PolyWorks topic
* Post update reply with version number and changes
* Done!
