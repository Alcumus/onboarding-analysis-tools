# Intelligent Contrator Analysis

Tool to match contractor list provided by hiring clients to business units in CBX

From Microsoft Powershell use the following (requires Docker)
> docker run --rm -it -v ${pwd}:/home/script/data $(docker build -t icm -q https://github.com/iguzu/icm.git) <hc_list.csv> <cbx db dump.csv> <Result.csv>

To see the command line tool help use the following:

> docker run --rm -it -v ${pwd}:/home/script/data $(docker build -t icm -q https://github.com/iguzu/icm.git) -h
