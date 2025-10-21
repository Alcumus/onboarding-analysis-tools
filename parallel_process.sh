#!/bin/bash
# parallel_process.sh
# Usage: ./parallel_process.sh <cbx_file> <input_file> <output_file> [chunk_size]

set -e

# Check arguments
if [ "$#" -lt 3 ]; then
    echo "Usage: $0 <cbx_file> <input_file> <output_file> [chunk_size]"
    echo "Example: $0 Oct3.csv ONCW_250.xlsx output_parallel.xlsx 50"
    exit 1
fi

CBX_FILE="$1"
INPUT_FILE="$2"
OUTPUT_FILE="$3"
CHUNK_SIZE="${4:-50}"  # Default to 50 if not provided

echo "=== Starting Parallel Processing ==="
echo "Input: $INPUT_FILE"
echo "Chunk size: $CHUNK_SIZE rows"
echo "Start time: $(date)"

# Step 1: Split data
echo "Splitting data into chunks..."
python3 -c "
import pandas as pd
import numpy as np

df = pd.read_excel('data/$INPUT_FILE')
chunk_size = $CHUNK_SIZE
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
" > /tmp/num_chunks.txt

NUM_CHUNKS=$(cat /tmp/num_chunks.txt | tail -1)

# Step 2: Build Docker image
echo "Building Docker image..."
docker build -t enhanced-algorithm .

# Step 3: Run parallel containers
echo "Starting $NUM_CHUNKS parallel containers..."
PIDS=()

for i in $(seq 1 $NUM_CHUNKS); do
    echo "Starting container for chunk $i..."
    docker run --memory=2g --cpus=1 --rm \
        -v $(pwd)/data:/home/script/data \
        enhanced-algorithm $CBX_FILE chunk_$i.xlsx output_chunk_$i.xlsx &
    PIDS+=($!)
done

# Step 4: Wait for completion
echo "Waiting for all containers to complete..."
for i in "${!PIDS[@]}"; do
    wait ${PIDS[$i]}
    echo "Chunk $((i+1)) completed"
done

# Step 5: Merge results
echo "Merging results..."
python3 -c "
import pandas as pd
import glob
import os

chunk_files = sorted(glob.glob('data/output_chunk_*.xlsx'))
combined_df = pd.DataFrame()

for file in chunk_files:
    chunk_df = pd.read_excel(file)
    combined_df = pd.concat([combined_df, chunk_df], ignore_index=True)

# Save temporary unformatted file
temp_file = 'data/temp_$OUTPUT_FILE'
combined_df.to_excel(temp_file, index=False)
print(f'Merged {len(combined_df)} total rows into temporary file')

# Cleanup chunk files
for file in chunk_files + glob.glob('data/chunk_*.xlsx'):
    os.remove(file)
"

# Step 6: Apply Excel formatting
echo "Applying Excel formatting..."
python3 format_excel.py "data/temp_$OUTPUT_FILE" "data/$OUTPUT_FILE"

# Cleanup temporary file
rm -f "data/temp_$OUTPUT_FILE"

echo "=== Parallel Processing Complete ==="
echo "Output: $OUTPUT_FILE"
echo "End time: $(date)"