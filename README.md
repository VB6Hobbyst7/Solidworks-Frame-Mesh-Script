# Solidworks-Frame-Mesh-Script
This is a VBA script to export 3D sketch data to an Excel file for further analysis.

# How to use
With a ZERO-REFERENCE 3D sketch selected (i.e. delete all references within the sketch), run Sketch3D_To_Excel.swp. This should create an output table in Excel or whatever tabular data sheet program your OS defaults too (e.g. OpenOffice). 

Failure to delete external references within the sketch can result in duplicate data, which can be handled in external processesing if you're patient and willing.

# Included Matlab visualization script
Included is a quick Matlab script to parse and visualize the data it reads from the returned data from the Sketch3D_To_Excel script. This script was written as a supplement to a Rose-GPE Chassis Torsional Stiffness FEA model that we are working on, and is used to import a chassis wireframe into the model.
