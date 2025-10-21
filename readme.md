# Onboarding Analysis Tools (Formerly Intelligent Contrator Analysis (ICM))

Tool to match contractor list provided by hiring clients to business units in CBX

## Prerequisite

1. Have Docker Desktop installed on your computer. To know that docker works, run the command "docker --version" from Windows Powershell should work and show you the installed version. (https://www.docker.com/products/docker-desktop/)

__** Please note that virtualization must be enabled in your BIOS, you need to have installed/enabled WSL2 and install a Linux kernel (Ex: Ubuntu 18.04) **__

> See the folowing instruction for help: https://docs.docker.com/desktop/windows/install/

2. Have git installed on your computer. to make sure git is installed properly run the command "git --version". (https://git-scm.com/)

The hardest is done...

3. Create a github account (free) https://github.com/signup and ask R&D to give you access to the repository
4. create a personal token that you will use to access the repository https://github.com/settings/tokens and name it docker access (store your token securly)

## Setup and File Preparation

### Typical Usage (Recommended for Most Users)

Most users run the analysis directly from their own folder without cloning the repository. Here's the typical workflow:

1. **Create your analysis folder** (anywhere on your computer):
   ```bash
   # Example locations:
   # Windows: C:\Users\YourName\Desktop\MyAnalysis
   # Mac/Linux: /Users/YourName/Desktop/MyAnalysis
   mkdir MyAnalysis
   cd MyAnalysis
   ```

2. **Get CBX contractor database**:
   - Connect to Redash and run the "2024 Business Units Extractor" query
   - Update the dataset by clicking the refresh button
   - Download the query result in CSV format
   - Rename to something short and friendly (e.g., `OCT16.csv`, `db-jan.csv`)
   - **Place in your analysis folder**

3. **Prepare hiring client list**:
   - Ensure your hiring client Excel file follows the [template format](hiring_client_input_template.xlsx)
   - **Place in your analysis folder**

4. **Your folder structure should look like**:
   ```
   MyAnalysis/
   ├── OCT16.csv           # Your CBX database export
   ├── OCWAwave2.xlsx      # Your hiring client list
   └── (results will be generated here)
   ```

### Advanced Usage (For Developers/Power Users)

If you need to modify the code, use parallel processing, or contribute to development:

**Project Structure Setup** (after cloning):
```
onboarding-analysis-tools/
├── main.py                    # Core processing logic
├── format_excel.py           # Excel formatting
├── parallel_process.sh       # Linux/Mac parallel processing
├── parallel_process.ps1      # Windows parallel processing  
├── Dockerfile               # Container configuration
├── requirements.txt         # Python dependencies
├── readme.md               # This file
├── data/                   # Your input/output files go here
│   ├── OCT16.csv          # CBX contractor database (you provide)
│   ├── OCWAwave2.xlsx     # Hiring client list (you provide)
│   └── results.xlsx       # Generated output files
└── .gitignore             # Excludes data/ folder from git
```

**Setup Steps:**
1. Clone the repository: `git clone https://github.com/Alcumus/onboarding-analysis-tools.git`
2. Navigate to project: `cd onboarding-analysis-tools`
3. Create data folder: `mkdir data`
4. Place input files in `data/` folder
5. Use parallel processing or local development commands

## Do the analysis

### Option 1: Direct Docker Run (Recommended for Most Users)

**This is the typical way most users run the analysis** - directly from their analysis folder without cloning the repository.

#### Windows PowerShell:
```powershell
# Navigate to your analysis folder (where you placed your CSV and Excel files)
cd "C:\path\to\your\analysis\folder"

# Set your GitHub token
$env:token = '<your personal github token to access the repository>'

# Run the analysis
docker run --rm -it -v ${pwd}:/home/script/data $(docker build -t icm -q https://${env:token}:@github.com/Alcumus/onboarding-analysis-tools.git) <cbx_contractor_db_dump.csv> <hc_list.xlsx> <results.xlsx>
```

#### Mac/Linux Terminal:
```bash
# Navigate to your analysis folder (where you placed your CSV and Excel files)
cd /path/to/your/analysis/folder

# Set your GitHub token
export token='<your personal github token to access the repository>'

# Run the analysis
docker run --rm -it -v $(pwd):/home/script/data $(docker build -t icm -q https://${token}:@github.com/Alcumus/onboarding-analysis-tools.git) <cbx_contractor_db_dump.csv> <hc_list.xlsx> <results.xlsx>
```

**Example Commands:**
```powershell
# Windows Example
cd "C:\Users\John\Desktop\MyAnalysis"
$env:token = 'ghp_xxxxxxxxxxxxxxxxxxxx'
docker run --rm -it -v ${pwd}:/home/script/data $(docker build -t icm -q https://${env:token}:@github.com/Alcumus/onboarding-analysis-tools.git) OCT16.csv OCWAwave2.xlsx results.xlsx
```

```bash
# Mac/Linux Example  
cd /Users/john/Desktop/MyAnalysis
export token='ghp_xxxxxxxxxxxxxxxxxxxx'
docker run --rm -it -v $(pwd):/home/script/data $(docker build -t icm -q https://${token}:@github.com/Alcumus/onboarding-analysis-tools.git) OCT16.csv OCWAwave2.xlsx results.xlsx
```

### Option 2: Local Development (For Developers Only)

**Only use this if you've cloned the repository locally for development or modifications.**

To run it locally for debugging (On Mac)
```bash
docker build -t onboarding-analysis-tools . && docker run --rm -v $(pwd)/data:/home/script/data onboarding-analysis-tools <cbx_contractor_db_dump.csv> <hc_list.xlsx> <results.xlsx>
```

To run it locally for debugging (On Windows)
```bash
docker build -t onboarding-analysis-tools . && docker run --rm -v ${pwd}/data:/home/script/data onboarding-analysis-tools <cbx_contractor_db_dump.csv> <hc_list.xlsx> <results.xlsx>
```

### Option 3: Parallel Processing with Docker (For Large Datasets)

**For datasets with >500 records, you can run multiple Docker containers in parallel for much faster processing.**

This approach splits your hiring client file into chunks and processes each chunk in a separate Docker container simultaneously.

#### Manual Parallel Processing (Any Platform):

**Step 1: Prepare your files in your analysis folder**
```
MyAnalysis/
├── OCT16.csv           # Your CBX database export  
├── OCWAwave2.xlsx      # Your hiring client list (large dataset)
└── (chunks and results will be created here)
```

**Step 2: Split your Excel file into chunks** (example using Python):
```python
# Create a simple script to split your Excel file
# Save this as split_excel.py in your analysis folder

import pandas as pd
import math

# Configuration
input_file = "OCWAwave2.xlsx"  # Your hiring client file
chunk_size = 50                # Records per chunk
output_prefix = "chunk_"       # Prefix for chunk files

# Read and split the Excel file
df = pd.read_excel(input_file)
total_records = len(df)
num_chunks = math.ceil(total_records / chunk_size)

print(f"Splitting {total_records} records into {num_chunks} chunks of {chunk_size} records each...")

for i in range(num_chunks):
    start_idx = i * chunk_size
    end_idx = min((i + 1) * chunk_size, total_records)
    chunk_df = df.iloc[start_idx:end_idx]
    
    chunk_filename = f"{output_prefix}{i+1}.xlsx"
    chunk_df.to_excel(chunk_filename, index=False)
    print(f"Created {chunk_filename} with {len(chunk_df)} records")

print(f"Total chunks created: {num_chunks}")
```

**Step 3: Run parallel Docker containers**

#### Windows PowerShell:
```powershell
# Set your GitHub token
$env:token = '<your personal github token>'

# Run containers in parallel (example for 4 chunks)
$jobs = @()
$jobs += Start-Job -ScriptBlock { docker run --rm -v ${using:PWD}:/home/script/data $(docker build -t icm-1 -q https://${using:env:token}:@github.com/Alcumus/onboarding-analysis-tools.git) OCT16.csv chunk_1.xlsx output_chunk_1.xlsx }
$jobs += Start-Job -ScriptBlock { docker run --rm -v ${using:PWD}:/home/script/data $(docker build -t icm-2 -q https://${using:env:token}:@github.com/Alcumus/onboarding-analysis-tools.git) OCT16.csv chunk_2.xlsx output_chunk_2.xlsx }
$jobs += Start-Job -ScriptBlock { docker run --rm -v ${using:PWD}:/home/script/data $(docker build -t icm-3 -q https://${using:env:token}:@github.com/Alcumus/onboarding-analysis-tools.git) OCT16.csv chunk_3.xlsx output_chunk_3.xlsx }
$jobs += Start-Job -ScriptBlock { docker run --rm -v ${using:PWD}:/home/script/data $(docker build -t icm-4 -q https://${using:env:token}:@github.com/Alcumus/onboarding-analysis-tools.git) OCT16.csv chunk_4.xlsx output_chunk_4.xlsx }

# Wait for all jobs to complete
$jobs | Wait-Job
$jobs | Receive-Job
$jobs | Remove-Job

Write-Host "All chunks processed!"
```

#### Mac/Linux Bash:
```bash
# Set your GitHub token
export token='<your personal github token>'

# Run containers in parallel (example for 4 chunks)
docker run --rm -v $(pwd):/home/script/data $(docker build -t icm-1 -q https://${token}:@github.com/Alcumus/onboarding-analysis-tools.git) OCT16.csv chunk_1.xlsx output_chunk_1.xlsx &
docker run --rm -v $(pwd):/home/script/data $(docker build -t icm-2 -q https://${token}:@github.com/Alcumus/onboarding-analysis-tools.git) OCT16.csv chunk_2.xlsx output_chunk_2.xlsx &  
docker run --rm -v $(pwd):/home/script/data $(docker build -t icm-3 -q https://${token}:@github.com/Alcumus/onboarding-analysis-tools.git) OCT16.csv chunk_3.xlsx output_chunk_3.xlsx &
docker run --rm -v $(pwd):/home/script/data $(docker build -t icm-4 -q https://${token}:@github.com/Alcumus/onboarding-analysis-tools.git) OCT16.csv chunk_4.xlsx output_chunk_4.xlsx &

# Wait for all background jobs to complete
wait

echo "All chunks processed!"
```

**Step 4: Combine results** (example using Python):
```python
# Create merge_results.py in your analysis folder

import pandas as pd
import glob

# Find all output chunk files
chunk_files = sorted(glob.glob("output_chunk_*.xlsx"))
combined_df = pd.DataFrame()

print(f"Found {len(chunk_files)} chunk result files")

# Combine all chunks
for file in chunk_files:
    print(f"Processing {file}...")
    chunk_df = pd.read_excel(file)
    combined_df = pd.concat([combined_df, chunk_df], ignore_index=True)

# Save combined results
output_file = "combined_results.xlsx"
combined_df.to_excel(output_file, index=False)
print(f"Combined {len(combined_df)} total records into {output_file}")

# Cleanup chunk files (optional)
# for file in chunk_files + glob.glob("chunk_*.xlsx"):
#     os.remove(file)
```

#### Automated Parallel Processing (Advanced - Requires Local Clone):

If you frequently process large datasets, clone the repository once to use the automated parallel processing scripts:

```bash
# One-time setup
git clone https://github.com/Alcumus/onboarding-analysis-tools.git
cd onboarding-analysis-tools
mkdir data

# Copy your files to data/ folder
cp /path/to/your/files/* data/

# Run automated parallel processing
bash parallel_process.sh OCT16.csv OCWAwave2.xlsx results_parallel.xlsx 50
```

#### Parameter Explanations:

1. **`<cbx_contractor_db.csv>`**: 
   - CBX contractor database export (CSV format)
   - File should be in `data/` folder
   - Example: `OCT16.csv`, `db-jan.csv`
   - This contains ~66,000+ contractor records from CBX system

2. **`<hiring_clients.xlsx>`**: 
   - Hiring client list to process (Excel format)
   - File should be in `data/` folder  
   - Example: `OCWAwave2.xlsx`, `QSLMAUDE1.xlsx`
   - Contains contractor names from hiring client to match against CBX

3. **`<output_results.xlsx>`**: 
   - Name for the output file (will be created in `data/` folder)
   - Contains 15 formatted sheets with categorized results
   - Example: `results_parallel.xlsx`, `my_analysis.xlsx`

4. **`[chunk_size]`** (Optional):
   - Number of records to process per container (default: 50)
   - Smaller chunks = more containers = faster processing (up to CPU limit)
   - Larger chunks = fewer containers = less overhead
   - Recommended: 50-100 records per chunk

#### Performance Benefits:

- **Speed**: Process large datasets 10-15x faster than single container
- **Parallel Processing**: Automatically splits work across multiple containers
- **Resource Utilization**: Uses available CPU cores efficiently  
- **Progress Tracking**: Shows real-time progress of each chunk
- **Memory Efficiency**: Each container uses ~2GB RAM limit

**Example Processing Times:**
- 684 records: ~14 minutes (vs 2+ hours single container)
- 201 records: ~6 minutes (vs 45+ minutes single container)

## Troubleshooting

### Common Issues:

1. **"File not found" errors**:
   - Ensure input files are in the `data/` folder
   - Use only filenames, not full paths
   - Check file extensions (.csv for CBX data, .xlsx for hiring client lists)

2. **Docker permission errors**:
   - Ensure Docker Desktop is running
   - Check that virtualization is enabled in BIOS (Windows)
   - Verify WSL2 is installed and working (Windows)

3. **Parallel processing fails**:
   - Check available system memory (each container uses ~2GB)
   - Reduce number of parallel containers or chunk size if running out of memory
   - Ensure Docker Desktop has sufficient resources allocated
   - For manual parallel processing, monitor system resources during execution
   - Start with fewer containers (2-3) and scale up based on your system capacity

4. **Slow performance**:
   - Use parallel processing for datasets >100 records
   - Adjust chunk size based on your system specs
   - Close other applications to free up resources

### Getting Help:

To see the command line tool help:

**Direct Docker mode (typical usage):**
```bash
# Windows PowerShell
$env:token = '<your github token>'
docker run --rm -it -v ${pwd}:/home/script/data $(docker build -t icm -q https://${env:token}:@github.com/Alcumus/onboarding-analysis-tools.git) -h

# Mac/Linux
export token='<your github token>'
docker run --rm -it -v $(pwd):/home/script/data $(docker build -t icm -q https://${token}:@github.com/Alcumus/onboarding-analysis-tools.git) -h
```

**Parallel processing mode (requires local clone):**
```bash
# Linux/Mac
bash parallel_process.sh

# Windows  
powershell -ExecutionPolicy Bypass -File parallel_process.ps1
```

### Performance Guidelines:

| Dataset Size | Recommended Mode | Expected Time | Chunk Size |
|-------------|------------------|---------------|------------|
| < 100 records | Single container | 5-15 minutes | N/A |
| 100-500 records | Parallel | 10-20 minutes | 50 |
| 500-1000 records | Parallel | 15-30 minutes | 50-100 |
| 1000+ records | Parallel | 30+ minutes | 100 |

## Documentation

- See the analysis [procedure documentation](ProcedureToProcessList.docx) 
- Hiring client Excel input file [template](hiring_client_input_template.xlsx)
- All input files must be placed in the `data/` folder
- Results are automatically formatted into 15 Excel sheets for easy import

