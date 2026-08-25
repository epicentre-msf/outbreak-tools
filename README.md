<div align="center">
<p align="center">

##  Quickly build a linelist from an excel designer :snail:

[OBT main](https://github.com/epicentre-msf/outbreak-tools/releases/latest/download/OBT-main-latest.zip) ·
[OBT Dev](https://github.com/epicentre-msf/outbreak-tools/releases/download/dev-latest/OBT-dev-latest.zip) ·
[Master Setup](https://github.com/epicentre-msf/outbreak-tools-setup/raw/main/releases/latest/disease_setup-latest.xlsb)
</p>
</div>


#### How it works?

In three steps:

1- Download the OBT folder and add the configurations of your linelist in the setup file. The setup file is basically an excel file with sheets referring to differents configurations to take in account when bulding the linelist. Remember to check the setup for eventual errors before importing it in the designer.

2- Use a [geobase](https://reports.msf.net/secure/app/outbreak-tools-geoapp) related to your linelist. You can choose to generate a linelist without a geobase which is optional. You can also import a geobase in the generated linelist.

3- Feed the designer with a **valid**  setup file (a setup file without errors in it) with/without a geobase and it generates a linelist using the configurations you have defined in the setup. 

For more informations about the setup, please [browse the outbreak-tools showcase repo](https://github.com/epicentre-msf/outbreak-tools-demo).

**The linelist designer requires Excel >= Excel 2010** and works on both Windows and Mac operating sytems.

#### Automation

Use the [obt](https://github.com/epicentre-msf/obt) R package.

#### Structure of the repo

- `scripts`: scripts for automating linelist creation, the release workflow, and development
- `docs`: Documentation website
- `src`: Source codes — binaries are **not** in git; they live in the GitHub Release asset store.

#### Limitations

Outbreak tool is limited by Excel's limitations. Using Excel 2010, here are your limitations:

- Maximum Number of variables in HList: 16384 (including hidden columns for geo variables)
- Maximum number of dropdowns you can use : 8184 (including at least one geo variable)
- Maximum length of data validation messages: 255


