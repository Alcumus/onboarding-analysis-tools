# Onboarding Analysis Tools (Formerly Intelligent Contrator Analysis (ICM))

Tool to match contractor list provided by hiring clients to business units in CBX

## Prerequisite

1. Have Docker Desktop installed on your computer. To know that docker works, run the command "docker --version" from Windows Powershell should work and show you the installed version. (https://www.docker.com/products/docker-desktop/)

__** Please note that virtualization must be enabled in your BIOS, you need to have installed/enabled WSL2 and install a Linux kernel (Ex: Ubuntu 18.04) **__

> See the folowing instruction for help: https://docs.docker.com/desktop/windows/install/

2. Have git installed on your computer. to make sure git is installed properly run the command "git --version". (https://git-scm.com/)

3. Have Python 3 installed and available in your PATH. To check, run `python3 --version` (WSL/macOS) or `python --version` (Windows/PowerShell). Download from https://www.python.org/downloads/
   Install the required Python packages:
   ```bash
   pip install pandas openpyxl
   ```
   (Run in your shell or PowerShell)

The hardest is done...

4. Create a github account (free) https://github.com/signup and ask R&D to give you access to the repository
5. create a personal token that you will use to access the repository https://github.com/settings/tokens and name it docker access (store your token securly)

## Work to do in preparation to the analysis

1. Create a folder (let's say on your Desktop) where you will do your analysis (With Windows Explorer)
2. Connect to redash (with your browser, you need an account...) and run the "2024 Business Units Extractor" query (don't forget to update the dataset by clicking the button in the bottom right corner)
3. It can take quite sometime to run the query be patient...
4. Download the query result in CSV format (by clicking the ... button at the bottom of the results)
5. Rename the downloaded file into something short and friendly, Ex: db-jan.csv
6. Move the file into the analysis folder created in step 1
7. Copy the hiring client list file into the analysis folder created in step 1

## Do the analysis

From Windows Powershell use the following (requires Docker)
```bash

cd < path to the analysis folder >

# Set your GitHub token as an environment variable
$env:token = '<your personal github token to access the repository>'

# For bash/zsh (WSL/macOS), use:
# export GITHUB_TOKEN='<your personal github token to access the repository>'

docker run --rm -it -v ${pwd}:/home/script/data $(docker build -t icm -q https://${env:token}:@github.com/Alcumus/onboarding-analysis-tools.git) <cbx_contractor_db_dump.csv> <hc_list.xlsx> <results.xlsx>
```

To see the command line tool help use the following:

```bash
docker run --rm -it -v  ${pwd}:/home/script/data $(docker build -t icm -q https://${env:token}:@github.com/Alcumus/onboarding-analysis-tools.git) -h
```

__** Please note that the script doesn't actually support "paths" to the input/output files since it uses a "hack" to map the files into the docker container. Only use filename and make sure they are located where the script is ran from.__

## Parallel Analysis Scripts


For large datasets, you can use the provided scripts to automate splitting, parallel processing, merging, and formatting. These scripts now support both remote (GitHub Docker) and local (pre-built Docker image) execution modes. Set the mode using the optional `--local` or `--remote` flag as the first argument. If omitted, remote mode is used by default. Your GitHub token must be set as an environment variable:


### Shell Script (WSL/macOS/Linux)

Set your token:
```bash
export token='<your personal github token>'
```

Run the script (remote mode, default):
```bash
chmod +x run_parallel_analysis.sh
./run_parallel_analysis.sh <input_xlsx> <chunk_size> <csv_file> <output_file>
```
Run the script (local mode):
```bash
chmod +x run_parallel_analysis.sh
./run_parallel_analysis.sh --local <input_xlsx> <chunk_size> <csv_file> <output_file>
```
Example:
```bash
chmod +x run_parallel_analysis.sh
./run_parallel_analysis.sh --local OCWAwave2.xlsx 50 OCT16.csv output_remote_master_formatted.xlsx
```

### PowerShell Script (Windows)

Set your token:
```powershell
$env:token = '<your personal github token to access the repository>'
```

Run the script (remote mode, default):
```powershell
./run_parallel_analysis.ps1 <input_xlsx> <chunk_size> <csv_file> <output_file>
```
Run the script (local mode):
```powershell
./run_parallel_analysis.ps1 --local <input_xlsx> <chunk_size> <csv_file> <output_file>
```
Example:
```powershell
./run_parallel_analysis.ps1 --local OCWAwave2.xlsx 50 OCT16.csv output_remote_master_formatted.xlsx
```

Both scripts will:

Both scripts will:
- Split the input Excel file into chunks of the specified size
- Run parallel Docker containers for each chunk (using either remote or local Docker mode)
- Merge the output chunk files into a single Excel file
- Format the final output file for analysis

**Note:** The scripts require Docker, Python 3, pandas, and openpyxl to be installed. The GITHUB_TOKEN environment variable must be set before running.


See the analysis [procedure documentation](ProcedureToProcessList.docx) and the hiring client Excel input file [template](hiring_client_input_template.xlsx).

