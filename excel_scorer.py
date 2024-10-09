import pandas as pd
from readability import Readability
import os
import warnings

# suppress UserWarnings (for SMOG warnings)
warnings.filterwarnings("ignore", category=UserWarning)

def process_single_file(input_file, output_file):

    df = pd.read_excel(input_file)

    fres_scores = []
    fkgl_scores = []
    gf_scores = []
    smog_scores = []
    cli_scores = []
    ari_scores = []

    for text in df.iloc[:, 0]:
        r = Readability(text, 5)  # text must contain >= 5 words
        
        fres_scores.append(r.flesch().score)
        fkgl_scores.append(r.flesch_kincaid().score)
        gf_scores.append(r.gunning_fog().score)
        smog_scores.append(r.smog(all_sentences=True, ignore_length=True).score)
        cli_scores.append(r.coleman_liau().score)
        ari_scores.append(r.ari().score)

    df['FRES'] = fres_scores
    df['FKGL'] = fkgl_scores
    df['GF'] = gf_scores
    df['SMOG'] = smog_scores
    df['CLI'] = cli_scores
    df['ARI'] = ari_scores

    df.to_excel(output_file, index=False)


input_folder = 'readability_eval/'
output_folder = 'readability_eval_results/'

if not os.path.exists(output_folder):
    os.makedirs(output_folder)

for dirpath, dirnames, filenames in os.walk(input_folder):
    
    for file_name in filenames:

        # skip temporary or hidden files (starting with `~` or `.`)
        if file_name.startswith('~') or file_name.startswith('.'):
            print(f"Skipping temporary file: {file_name}")
            continue

        if file_name.endswith('.xlsx'):
            input_file = os.path.join(dirpath, file_name)
            # input_file_base = os.path.splitext(os.path.basename(input_file))[0]
            # output_file = os.path.join(output_folder, f'{input_file_base}_readability.xlsx')
            output_file_name = file_name.replace('_Empty', '')

            relative_path = os.path.relpath(dirpath, input_folder)
            output_dir = os.path.join(output_folder, relative_path)

            if not os.path.exists(output_dir):
                os.makedirs(output_dir)

            output_file = os.path.join(output_dir, output_file_name)

            print(f"Processing {input_file}...")

            try:
                process_single_file(input_file, output_file)
                print(f"    Succeeded. Saved to {output_file}")
            except Exception as e:
                print(f"    Failed. {e}")

print("Processing completed for all files.")