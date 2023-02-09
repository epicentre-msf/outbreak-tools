## outbreak-tools : Quickly build linelist from an excel designer

### Download Files

[![Download Latest version of setup file](https://github.com/epicentre-msf/outbreak-tools-setup/raw/main/src/imgs/download_setup.svg)](https://github.com/epicentre-msf/outbreak-tools-setup/raw/main/setup.xlsb)

[![Download Latest stable version of the designer](https://github.com/epicentre-msf/outbreak-tools/raw/users/y-amevoin/src/imgs/stable_designer.svg)](https://github.com/epicentre-msf/outbreak-tools/raw/main/linelist_designer.xlsb)

[![Download Latest development version of the designer](https://github.com/epicentre-msf/outbreak-tools/raw/users/y-amevoin/src/imgs/dev_designer.svg)](https://github.com/epicentre-msf/outbreak-tools/raw/dev/linelist_designer.xlsb)



### Cloning the repo on your local computer


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
