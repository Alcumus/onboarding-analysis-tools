# Intelligent Contrator Analysis (ICM)

Tool to match contractor list provided by hiring clients to business units in CBX


From Windows Powershell use the following (requires Docker)

> cd <to folder where your input files iare located> 

> docker run --rm -it -v ${pwd}:/home/script/data $(docker build -t icm -q https://github.com/iguzu/icm.git) <cbx_contractor_db_dump.csv> <hc_list.xlsx> <results.xlsx>

To see the command line tool help use the following:

> docker run --rm -it -v ${pwd}:/home/script/data $(docker build -t icm -q https://github.com/iguzu/icm.git) -h


__** Please note that the script doesn't actually support "paths" to the input/output files since it uses a "hack" to map the files into the docker container. The file must be located where the script is located.__


See the analsysis procedure documentation at https://github.com/iguzu/icm/blob/master/ProcedureToProcessList.docx and the hiring client Excel input file template at https://github.com/iguzu/icm/blob/master/hring_client_input_template.xlsx.
  
  [Template](hiring_client_input_template.xlsx)

