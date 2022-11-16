# Intelligent Contrator Analysis (ICM)

Tool to match contractor list provided by hiring clients to business units in CBX

## Prerequisite

1. Have Docker Desktop installed on your computer. To know that docker works, run the command "docker --version" from Windows Powershell should work and show you the installed version. (https://www.docker.com/products/docker-desktop/)

__** Please note that virtualization must be enabled in your BIOS, you need to have installed/enabled WSL2 and install a Linux kernel (Ex: Ubuntu 18.04) **__

> See the folowing instruction for help: https://docs.docker.com/desktop/windows/install/

2. Have git installed on your computer. to make sure git is installed properly run the command "git --version". (https://git-scm.com/)

The hardest is done...

3. Create a github account (free) https://github.com/signup and ask R&D to give you access to the repository
4. create a personal token that you will use to access the repository https://github.com/settings/tokens and name it docker access (store your token securly)

## Work to do in preparation to the analysis

1. Create a folder (let's say on your Desktop) where you will do your analysis (With Windows Explorer)
2. Connect to redash (with your browser, you need an account...) and run the "SBL - Business Unit Extractor" query (don't forget to update the dataset by clicking the button in the bottom right corner)
3. It can take quite sometime to run the query be patient...
4. Download the query result in CSV format (by clicking the ... button at the bottom of the results)
5. Rename the downloaded file into something short and friendly, Ex: db-jan.csv
6. Move the file into the analysis folder created in step 1
7. Copy the hiring client list file into the analysis folder created in step 1

## Do the analysis

From Windows Powershell use the following (requires Docker)
```
cd < path to the analysis folder >

$env:token = '<your personal github token to access the repository>'

docker run --rm -it -v ${pwd}:/home/script/data $(docker build -t icm -q https://${env:token}:@github.com/cognibox/icm.git) <cbx_contractor_db_dump.csv> <hc_list.xlsx> <results.xlsx>
```

To see the command line tool help use the following:

```
docker run --rm -it -v  ${pwd}:/home/script/data $(docker build -t icm -q https://${env:token}:@github.com/cognibox/icm.git) -h
```

__** Please note that the script doesn't actually support "paths" to the input/output files since it uses a "hack" to map the files into the docker container. Only use filename and make sure they are located where the script is ran from.__


See the analysis [procedure documentation](ProcedureToProcessList.docx) and the hiring client Excel input file [template](hiring_client_input_template.xlsx).

