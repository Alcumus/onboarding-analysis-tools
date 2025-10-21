# parallel_process.ps1
# Usage: .\parallel_process.ps1 <cbx_file> <input_file> <output_file> [chunk_size]

param(
    [Parameter(Mandatory=$true)][string]$CbxFile,
    [Parameter(Mandatory=$true)][string]$InputFile,
    [Parameter(Mandatory=$true)][string]$OutputFile,
    [int]$ChunkSize = 50
)

# Validate parameters
if (-not $CbxFile -or -not $InputFile -or -not $OutputFile) {
    Write-Host "Usage: .\parallel_process.ps1 <cbx_file> <input_file> <output_file> [chunk_size]"
    Write-Host "Example: .\parallel_process.ps1 Oct3.csv ONCW_250.xlsx output_parallel.xlsx 50"
    exit 1
}

Write-Host "=== Starting Parallel Processing ==="
Write-Host "CBX File: $CbxFile"
Write-Host "Input: $InputFile"
Write-Host "Output: $OutputFile"
Write-Host "Chunk size: $ChunkSize rows"
Write-Host "Start time: $(Get-Date)"

# Step 1: Split data
Write-Host "Splitting data into chunks..."
$numChunks = python -c @"
import pandas as pd
import numpy as np

df = pd.read_excel('data/$InputFile')
chunk_size = $ChunkSize
num_chunks = len(df) // chunk_size + (1 if len(df) % chunk_size > 0 else 0)

print(f'Splitting {len(df)} rows into {num_chunks} chunks')

for i in range(num_chunks):
    start_idx = i * chunk_size
    end_idx = min((i + 1) * chunk_size, len(df))
    chunk_df = df.iloc[start_idx:end_idx]
    
    output_file = f'data/chunk_{i+1}.xlsx'
    chunk_df.to_excel(output_file, index=False)
    print(f'Created chunk_{i+1}.xlsx with {len(chunk_df)} rows')

print(num_chunks)
"@ | Select-Object -Last 1

# Step 2: Build Docker image
Write-Host "Building Docker image..."
docker build -t onboarding-analysis-tools .

# Step 3: Run parallel containers
Write-Host "Starting $numChunks parallel containers..."
$jobs = @()

for ($i = 1; $i -le $numChunks; $i++) {
    Write-Host "Starting container for chunk $i..."
    $jobs += Start-Job -ScriptBlock {
        param($cbx, $chunk, $output, $chunkNum)
        docker run --memory=2g --cpus=1 --rm -v ${pwd}/data:/home/script/data onboarding-analysis-tools $cbx $chunk $output
    } -ArgumentList $CbxFile, "chunk_$i.xlsx", "output_chunk_$i.xlsx", $i
}

# Step 4: Wait for completion
Write-Host "Waiting for all containers to complete..."
$jobs | Wait-Job | ForEach-Object {
    $jobResult = Receive-Job $_
    Write-Host "Chunk completed: $jobResult"
    Remove-Job $_
}

# Step 5: Merge results
Write-Host "Merging results..."
python -c @"
import pandas as pd
import glob
import os

chunk_files = sorted(glob.glob('data/output_chunk_*.xlsx'))
combined_df = pd.DataFrame()

for file in chunk_files:
    chunk_df = pd.read_excel(file)
    combined_df = pd.concat([combined_df, chunk_df], ignore_index=True)

combined_df.to_excel('data/$OutputFile', index=False)
print(f'Merged {len(combined_df)} total rows into $OutputFile')

# Cleanup
for file in chunk_files + glob.glob('data/chunk_*.xlsx'):
    os.remove(file)
"@

Write-Host "=== Parallel Processing Complete ==="
Write-Host "Output: $OutputFile"
Write-Host "End time: $(Get-Date)"