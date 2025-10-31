#!/bin/bash
# Usage: ./run_parallel_analysis.sh <input_xlsx> <chunk_size> <csv_file> <output_file>
# Ensure Python 3 is installed (user must do this manually)
# Install required Python packages
python3 -m pip install --upgrade pip
python3 -m pip install pandas openpyxl

set -e
input_xlsx="$1"
chunk_size="$2"
csv_file="$3"
output_file="$4"

mode="local" # default
if [[ "$1" == "--local" ]]; then
  mode="local"
  shift
elif [[ "$1" == "--remote" ]]; then
  mode="remote"
  shift
fi

if [[ -z "$input_xlsx" || -z "$chunk_size" || -z "$csv_file" || -z "$output_file" ]]; then
  echo "Usage: $0 [--local|--remote] <input_xlsx> <chunk_size> <csv_file> <output_file>"
  exit 1
fi

# Step 1: Split input file
echo "Splitting $input_xlsx into chunks of $chunk_size rows..."
python3 -c "
import pandas as pd
import sys
import os

# Change to script directory to ensure we're working in the right place
script_dir = os.path.dirname(os.path.abspath(__file__)) if '__file__' in dir() else os.getcwd()
os.chdir(script_dir)

input_file = sys.argv[1]
chunk_size = int(sys.argv[2])
output_dir = os.getcwd()

print(f'Script directory: {script_dir}')
print(f'Working directory: {output_dir}')
print(f'Directory is writable: {os.access(output_dir, os.W_OK)}')

# Test simple file write first
try:
    with open('test_write.txt', 'w') as f:
        f.write('test')
    os.remove('test_write.txt')
    print('✅ Simple write test successful')
except Exception as e:
    print(f'❌ Simple write test failed: {e}')
    import traceback
    traceback.print_exc()
    sys.exit(1)

print(f'Reading input file: {input_file}')
try:
    df = pd.read_excel(input_file)
    print(f'✅ Loaded {len(df)} rows from input file')
except Exception as e:
    print(f'❌ Failed to read input file: {e}')
    import traceback
    traceback.print_exc()
    sys.exit(1)

num_chunks = (len(df) + chunk_size - 1) // chunk_size
print(f'Will create {num_chunks} chunks')

for i in range(num_chunks):
    start_idx = i * chunk_size
    end_idx = min(start_idx + chunk_size, len(df))
    chunk_df = df.iloc[start_idx:end_idx]
    filename = f'chunk_{i+1}.xlsx'
    print(f'Attempting to write: {filename}')
    try:
        chunk_df.to_excel(filename, index=False)
        if os.path.exists(filename):
            print(f'✅ Created {filename} with {len(chunk_df)} records')
        else:
            print(f'❌ File was not created: {filename}')
    except Exception as e:
        print(f'❌ Error creating {filename}: {e}')
        import traceback
        traceback.print_exc()
        sys.exit(1)

print(f'✅ Successfully created {num_chunks} chunks')
with open('num_chunks.txt', 'w') as f:
    f.write(str(num_chunks))
print('✅ Wrote num_chunks.txt')
" "$input_xlsx" "$chunk_size"

num_chunks=$(cat num_chunks.txt)
rm num_chunks.txt

# Step 2: Run parallel analysis
if [[ "$mode" == "remote" ]]; then
    echo "[INFO] Running in REMOTE mode (GitHub Docker build)"
    echo "Running parallel analysis for $num_chunks chunks..."
  if [[ -z "$token" ]]; then
      echo "Error: GITHUB_TOKEN environment variable is not set."
      exit 1
  fi
  for i in $(seq 1 $num_chunks); do
      docker run --rm \
          -v $(pwd):/home/script/data \
          $(docker build -t icm-$i -q https://${token}:@github.com/Alcumus/onboarding-analysis-tools.git) \
          "$csv_file" "chunk_${i}.xlsx" "output_chunk_${i}.xlsx" &
  done
else
    echo "[INFO] Running in LOCAL mode (local Docker image)"
    echo "Building Docker image..."
    docker build -t onboarding-analysis-tools .
    echo "Running parallel analysis for $num_chunks chunks..."
    for i in $(seq 1 $num_chunks); do
            docker run --rm \
                    -v $(pwd):/home/script/data \
                    onboarding-analysis-tools "$csv_file" "chunk_${i}.xlsx" "output_chunk_${i}.xlsx" &
    done
fi
wait
echo "✅ All containers completed!"

# Step 3: Merge results
echo "Merging chunk outputs into output_remote_master.xlsx..."
python3 << 'PYEOF'
import pandas as pd, glob, os
output_dir = os.getcwd()
print(f'Working directory: {output_dir}')
chunks = sorted(glob.glob(os.path.join(output_dir, 'output_chunk_*.xlsx')))
print(f'Found {len(chunks)} chunk files to merge')
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
output_path = os.path.join(output_dir, "output_remote_master.xlsx")
with pd.ExcelWriter(output_path) as writer:
    for sheet_name in sheet_names:
        merged_sheets[sheet_name].to_excel(writer, sheet_name=sheet_name, index=False)
print(f'✅ Merged output saved to {output_path}')
PYEOF

# Step 4: Format output
echo "Formatting merged output..."
python3 format_excel.py output_remote_master.xlsx "$output_file"
echo "✅ All steps completed. Final output: $output_file"
echo "Cleaning up intermediate files..."
rm -f chunk_*.xlsx
rm -f output_chunk_*.xlsx
rm -f output_remote_master.xlsx
echo "Cleanup complete. Only $output_file retained."
