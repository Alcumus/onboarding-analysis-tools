# Usage: .\run_parallel_analysis.ps1 <input_xlsx> <chunk_size> <csv_file> <output_file> [--local|--remote]
# Ensure Python 3 is installed (user must do this manually)
# Install required Python packages
py -3.12 -m pip install --upgrade pip
py -3.12 -m pip install pandas openpyxl

if ($args.Count -lt 4) {
    Write-Host "Usage: .\run_parallel_analysis.ps1 <input_xlsx> <chunk_size> <csv_file> <output_file> [--local|--remote]"
    exit 1
}

$input_xlsx = $args[0]
$chunk_size = [int]$args[1]
$csv_file = $args[2]
$output_file = $args[3]

# Parse mode from last optional parameter
$mode = "remote" # default
if ($args.Count -gt 4) {
    if ($args[4] -eq "--local") {
        $mode = "local"
    } elseif ($args[4] -eq "--remote") {
        $mode = "remote"
    }
}

if ($mode -eq "remote") {
    Write-Host "[INFO] Running in REMOTE mode (GitHub Docker build)"
} else {
    Write-Host "[INFO] Running in LOCAL mode (local Docker image)"
}

# Step 1: Split input file
Write-Host "Splitting $input_xlsx into chunks of $chunk_size rows..."
py -3.12 -c "
import pandas as pd, sys
input_file = sys.argv[1]
chunk_size = int(sys.argv[2])
df = pd.read_excel(input_file)
num_chunks = (len(df) + chunk_size - 1) // chunk_size
for i in range(num_chunks):
    start_idx = i * chunk_size
    end_idx = min(start_idx + chunk_size, len(df))
    chunk_df = df.iloc[start_idx:end_idx]
    chunk_df.to_excel(f'chunk_{i+1}.xlsx', index=False)
    print(f'Created chunk_{i+1}.xlsx with {len(chunk_df)} records')
print(f'✅ Created {num_chunks} chunks')
with open('num_chunks.txt', 'w') as f:
    f.write(str(num_chunks))
" $input_xlsx $chunk_size

 $num_chunks = Get-Content num_chunks.txt
Remove-Item num_chunks.txt

# Step 2: Run parallel analysis
if ($mode -eq "remote") {
    Write-Host "[INFO] Running in REMOTE mode (GitHub Docker build)"
    Write-Host "Running parallel analysis for $num_chunks chunks..."
    if (-not $env:token) {
        Write-Host "Error: GITHUB_TOKEN environment variable is not set."
        exit 1
    }
    $current_path = $PWD.Path
    $jobs = @()
    for ($i = 1; $i -le $num_chunks; $i++) {
        Write-Host "Starting chunk $i..."
        $jobs += Start-Job -ScriptBlock {
            param($i, $csv_file, $token, $pwd)
            docker run --rm `
                -v "${pwd}:/home/script/data" `
                $(docker build -t icm-$i -q "https://${token}:@github.com/Alcumus/onboarding-analysis-tools.git") `
                $csv_file "chunk_${i}.xlsx" "output_chunk_${i}.xlsx"
        } -ArgumentList $i, $csv_file, $env:token, $current_path
    }
    Write-Host "Waiting for all jobs to complete..."
    $jobs | Wait-Job | Receive-Job
    $jobs | Remove-Job
} else {
    Write-Host "[INFO] Running in LOCAL mode (local Docker image)"
    Write-Host "Building Docker image..."
    docker build -t onboarding-analysis-tools .
    Write-Host "Running parallel analysis for $num_chunks chunks..."
    $current_path = $PWD.Path
    $jobs = @()
    for ($i = 1; $i -le $num_chunks; $i++) {
        Write-Host "Starting chunk $i..."
        $jobs += Start-Job -ScriptBlock {
            param($i, $csv_file, $pwd)
            docker run --rm `
                -v "${pwd}:/home/script/data" `
                onboarding-analysis-tools $csv_file "chunk_${i}.xlsx" "output_chunk_${i}.xlsx"
        } -ArgumentList $i, $csv_file, $current_path
    }
    Write-Host "Waiting for all jobs to complete..."
    $jobs | Wait-Job | Receive-Job
    $jobs | Remove-Job
}
Write-Host "✅ All containers completed!"

# Step 3: Merge results
Write-Host "Merging chunk outputs into output_remote_master.xlsx..."
py -3.12 -c "
import pandas as pd, glob
chunks = sorted(glob.glob('output_chunk_*.xlsx'))
sheet_names = [
    'all', 'onboarding', 'association_fee', 're_onboarding', 'subscription_upgrade',
    'ambiguous_onboarding', 'restore_suspended', 'activation_link', 'already_qualified',
    'add_questionnaire', 'missing_info', 'follow_up_qualification',
    'Data to import', 'Existing Contractors', 'Data for HS'
]
merged_sheets = {}
for sheet_name in sheet_names:
    sheet_dfs = []
    for chunk_file in chunks:
        try:
            df = pd.read_excel(chunk_file, sheet_name=sheet_name)
            if len(df) > 0:
                sheet_dfs.append(df)
        except Exception:
            pass
    if sheet_dfs:
        merged_sheets[sheet_name] = pd.concat(sheet_dfs, ignore_index=True)
    else:
        merged_sheets[sheet_name] = pd.DataFrame()
with pd.ExcelWriter('output_remote_master.xlsx') as writer:
    for sheet_name in sheet_names:
        merged_sheets[sheet_name].to_excel(writer, sheet_name=sheet_name, index=False)
"

# Step 4: Format output
Write-Host "Formatting merged output..."
py -3.12 format_excel.py output_remote_master.xlsx $output_file
Write-Host "✅ All steps completed. Final output: $output_file"

# Step 5: Cleanup intermediate files
Write-Host "Cleaning up intermediate files..."
Remove-Item -Force -ErrorAction SilentlyContinue chunk_*.xlsx
Remove-Item -Force -ErrorAction SilentlyContinue output_chunk_*.xlsx
Remove-Item -Force -ErrorAction SilentlyContinue output_remote_master.xlsx
Write-Host "Cleanup complete. Only $output_file retained."
