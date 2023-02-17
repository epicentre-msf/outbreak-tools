<div align="center">
<p align="center">




##  Quickly build a linelist from an excel designer

[![Download Latest version of setup file](https://github.com/epicentre-msf/outbreak-tools/raw/users/y-amevoin/src/imgs/setup_file.svg)](https://github.com/epicentre-msf/outbreak-tools-setup/raw/main/setup.xlsb)
[![Download Latest stable version of the designer](https://github.com/epicentre-msf/outbreak-tools/raw/users/y-amevoin/src/imgs/stable_designer.svg)](https://github.com/epicentre-msf/outbreak-tools/raw/main/linelist_designer.xlsb)
[![Download Latest development version of the designer](https://github.com/epicentre-msf/outbreak-tools/raw/users/y-amevoin/src/imgs/dev_designer.svg)](https://github.com/epicentre-msf/outbreak-tools/raw/dev/linelist_designer_dev.xlsb)

</p>
</div>

#### How it works?

In three steps:

1- Define the configurations of your linelist in a setup file. It is basically an excel file with sheets referring to differents configurations to take in account when bulding the linelist.

2- Use or not a [geobase](https://reports.msf.net/secure/app/outbreak-tools-geoapp) related to your linelist.

3- Feed the designer with a valid setup file with/without a valid geobase and it generates a linelist using the configurations you have defined.

The linelist designer requires Excel >= Excel 2010 and works on Windows and Mac Os.
Automation of the work can be done on R (**only works on a windows machine**) using the provided [R script](https://github.com/epicentre-msf/outbreak-tools/raw/main/Rscripts/run_designer_on_windows.R) as example. It sends the required parameters for the designer to a vbscript which in turn opens excel and runs the routines for linelist creation.



#### Cloning the repo on your local computer


This repo contains one sub-module, add it when cloning

```cmd
git clone --recurse-submodules https://github.com/epicentre-msf/outbreak-tools.git

```

Or after cloning :

```cmd

git clone https://github.com/epicentre-msf/outbreak-tools.git
git submodule init
git submodule update

```



The repo contains codes and input files for a linelist builder written in VBA.

- `input` folder contains All the input files needed for building the linelist:

  - Setup files in the `setup` folder : (Dictionary and parameters for building the linelist)
  - A geobase files in the `geobase` folder: (Geobase data in excel .xlsx formats that are loaded in the designer)
- `src` folder contains all the source codes in flat files used in the designer.
- `linelist_designer.xlsb` is the designer file.
