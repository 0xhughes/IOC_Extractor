# IOC_Extractor
This Python script is used to generate Splunk/TAP/etc queries from Excel IOC definition files that match a certain format. I created this as a project to assist a colleague and felt maybe others would find it handy.

This script was written using Python 2.7, within Windows, and utilizes a number of modules. One you may need to download and acquire yourself is the xlrd module.

This script was originally distributed after being compiled to EXE using the py2exe module, then packaged into an installer (using NSI). Included in this repository is a "queryDefs.conf" file which is installed to a Program Files folder by default when used with the aforementioned NSI, installer. Be sure to place this in a location of your choosing and refer to it during execution as it will not reside within it's "default" location as no installation using the installer happens.
