<div align="center">
<p align="center">




##  Quickly build a linelist from an excel designer

[![Download Latest version of setup file](https://github.com/epicentre-msf/outbreak-tools/raw/users/y-amevoin/src/imgs/setup_file.svg)](https://github.com/epicentre-msf/outbreak-tools-setup/raw/main/setup.xlsb)
[![Download Latest stable version of the designer](https://github.com/epicentre-msf/outbreak-tools/raw/users/y-amevoin/src/imgs/stable_designer.svg)](https://github.com/epicentre-msf/outbreak-tools/raw/main/designer.xlsb)
[![Download Latest development version of the designer](https://github.com/epicentre-msf/outbreak-tools/raw/users/y-amevoin/src/imgs/dev_designer.svg)](https://github.com/epicentre-msf/outbreak-tools/raw/dev/designer_dev.xlsb)
[![Documentation](https://github.com/epicentre-msf/outbreak-tools/raw/users/y-amevoin/src/imgs/docs.svg)](https://epicentre-msf.github.io/outbreak-tools/)
</p>
</div>

#### How it works?

In three steps:

1- Download the setup file and add the configurations of your linelist in it. The setup file is basically an excel file with sheets referring to differents configurations to take in account when bulding the linelist. Remember to check the setup for eventual errors before importing it in the designer.

2- Use a [geobase](https://reports.msf.net/secure/app/outbreak-tools-geoapp) related to your linelist. You can choose to generate a linelist without a geobase which is optional. You can also import a geobase in the created linelist if you want.

3- Feed the designer with a **valid**  setup file (a setup file without errors in it) with/without a geobase and it generates a linelist using the configurations you have defined.

**The linelist designer requires Excel >= Excel 2010** and works on both Windows and Mac operating sytems.
Automation of the work can be done on R (**only works on a windows machine**) using the provided [R script](https://github.com/epicentre-msf/outbreak-tools/raw/main/Rscripts/run_designer_on_windows.R) as example. It sends the required parameters for the designer to a [vbscript](https://github.com/epicentre-msf/outbreak-tools/raw/main/Rscripts/rundesigner.vbs) which in turn opens excel and runs the routines for linelist creation.

#### Structure of the repo

The repo contains codes and input files for a linelist builder written in VBA.

- `input` folder contains All the input files needed for building the linelist:

  - Setup files in the `setup` folder in `.xlsb` format.
  - Geobase files in the `geobase` folder: (Geobase data in excel `.xlsx` formats that could be loaded in the designer)
- `src` folder contains all the source codes in flat files used in the designer.
- `designer.xlsb` is the designer file.
