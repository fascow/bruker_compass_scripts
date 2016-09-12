# bruker_compass_scripts
Collection of scripts to support LC-MS data analysis with Bruker software

## DataAnalysis
The EIC_Import_and_FindAutoMSn script may be saved in a DataAnalysis method and can be used to automatically prepare LC-MS files for data analysis. Compounds can be generated and EIC values from an Excel table can be added to each analysis. The EIC_Export script is a GUI scripting bodge to save all EIC values from one analysis into an Excel file.

## LibraryEditor
The Library_Spectra_Export script is also a GUI scripting bodge that does the same as the Export function of the LibraryEditor, except it saves each compound in a separate file. These files may be combined into one *JSON* file with the functions in the Library_Spectra_Export_process_results IPython Notebook. There is one drawback - only the first spectrum of each compound gets exported. The Bruker_LibraryEditor_Export_Parser provides a better solution. It parses all spectra from a *.library* file exported from LibraryEditor and is able to reconstruct the compound structure of the Library. The result may be saved as a *JSON* file for further use.

## QuantAnalysis
This GUI scripting bodge adds a missing functionality to the Bruker QuantAnalysis 2.2 software. Method files can be created from a simple Excel table and the Work Table can be automatically filled with all files present in the working directory.
