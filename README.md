# EML2PDF

This is a simple CLI program that let's you convert an email (.eml file) to a PDF.

# Usage

Download the latest release. It should be a single file named `EML2PDF.zip`, which contains the program, as well as the settings file `appsettings.json`

There are two ways to use it:
1. Run the program and enter your eml file path when prompted
2. Provide the path as a cli argument

In the appsettings file, you can specify if you want your eml file to be deleted, or backed up.
Set `DeleteFileAfterProcessing` to "true" if you want your original file to be deleted, and "false" if you want it to be backed up.

The output PDF is created in the same folder your EML file is located in.