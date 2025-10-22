# Usage: .\run_parallel_analysis.ps1 <input_xlsx> <chunk_size> <csv_file> <output_file>
# Ensure Python 3 is installed (user must do this manually)
# Install required Python packages
python -m pip install --upgrade pip
python -m pip install pandas openpyxl

param(
    [string]$input_xlsx,
    [int]$chunk_size,
    [string]$csv_file,
    [string]$output_file
)

if (-not $input_xlsx -or -not $chunk_size -or -not $csv_file -or -not $output_file) {
    Write-Host "Usage: .\run_parallel_analysis.ps1 <input_xlsx> <chunk_size> <csv_file> <output_file>"
    exit 1
}

# Step 1: Split input file
Write-Host "Splitting $input_xlsx into chunks of $chunk_size rows..."
python -c "
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
Write-Host "Running parallel analysis for $num_chunks chunks..."
if (-not $env:token) {
    Write-Host "Error: token environment variable is not set."
    exit 1
}
$jobs = @()
for ($i = 1; $i -le $num_chunks; $i++) {
    $jobs += Start-Job -ScriptBlock {
        docker run --rm `
            -v "$PWD:/home/script/data" `
            $(docker build -t icm-$using:i -q https://${env:token}:@github.com/Alcumus/onboarding-analysis-tools.git) `
            $using:csv_file "chunk_${using:i}.xlsx" "output_chunk_${using:i}.xlsx"
    }
}
$jobs | Wait-Job | Out-Null
Write-Host "✅ All containers completed!"

# Step 3: Merge results
Write-Host "Merging chunk outputs into output_remote_master.xlsx..."
python -c "
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
python format_excel.py output_remote_master.xlsx $output_file
Write-Host "✅ All steps completed. Final output: $output_file"
