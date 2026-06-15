<div align="center">
<p align="center">

##  Quickly build a linelist from an excel designer :snail:

[![Download All](https://github.com/epicentre-msf/outbreak-tools/raw/main/src/imgs/download_all.svg)](https://github.com/epicentre-msf/outbreak-tools/releases/latest/download/OBT-main-latest.zip)
[![Download All Dev](https://github.com/epicentre-msf/outbreak-tools/raw/main/src/imgs/download_all_dev.svg)](https://github.com/epicentre-msf/outbreak-tools/releases/download/dev-latest/OBT-dev-latest.zip)
[![Download the master setup](https://github.com/epicentre-msf/outbreak-tools/raw/main/src/imgs/setup_file.svg)](https://github.com/epicentre-msf/outbreak-tools-setup/raw/main/releases/latest/disease_setup-latest.xlsb)
[![All Releases](https://hits.sh/epicentre-msf.github.io/outbreak-tools/releases.svg?label=All%20Releases&color=8250df)](https://epicentre-msf.github.io/outbreak-tools/releases.html)
</p>
</div>


#### How it works?

In three steps:

1- Download the setup file and add the configurations of your linelist in it. The setup file is basically an excel file with sheets referring to differents configurations to take in account when bulding the linelist. Remember to check the setup for eventual errors before importing it in the designer.

2- Use a [geobase](https://reports.msf.net/secure/app/outbreak-tools-geoapp) related to your linelist. You can choose to generate a linelist without a geobase which is optional. You can also import a geobase in the created linelist if you want.

3- Feed the designer with a **valid**  setup file (a setup file without errors in it) with/without a geobase and it generates a linelist using the configurations you have defined. 

For more informations about the setup, please [browse]((https://github.com/epicentre-msf/outbreak-tools-setup)) elements of the setup repository, read the [setup wiki](https://github.com/epicentre-msf/outbreak-tools-setup/wiki) or [browse the outbreak-tools showcase repo](https://github.com/epicentre-msf/outbreak-tools-demo).

**The linelist designer requires Excel >= Excel 2010** and works on both Windows and Mac operating sytems.

#### Automation

Automation of the work can be done on R (**only works on a windows machine**) using the provided [R script](https://github.com/epicentre-msf/outbreak-tools/raw/main/automate/codes/run_designer_on_windows.R) as example. It sends the required parameters for the designer to a [vbscript](https://github.com/epicentre-msf/outbreak-tools/raw/main/automate/codes/rundesigner.vbs) which in turn opens excel and runs the routines for linelist creation.

#### Structure of the repo

- `automate`: scripts for automating linelist creation, the release workflow, and development
- `docs`: Documentation website
- `src`: Source codes — binaries are **not** in git; they live in the GitHub Release asset store (see [RELEASING.md](RELEASING.md))

Releases are published as **GitHub Releases** (no `releases/` folder in the repo). See [RELEASING.md](RELEASING.md) for how releases work, and the [releases page](https://epicentre-msf.github.io/outbreak-tools/releases.html).

#### Limitations

Outbreak tool is limited by Excel's limitations. Using Excel 2010, here are your limitations:

- Maximum Number of variables in HList: 16384 (including hidden columns for geo variables)
- Maximum number of dropdowns you can use : 8184 (including at least one geo variable)
- Maximum length of data validation messages: 255


