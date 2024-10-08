import pandas as pd
from readability import Readability
import os


def process_single_file(input_file):
    df = pd.read_excel(input_file)
    input_file_base = os.path.splitext(os.path.basename(input_file))[0]
    output_file = os.path.join(output_folder, f'{input_file_base}_readability_scores.xlsx')

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

    print(f"Readability scores saved to {output_file}")


input_folder = 'input/'
output_folder = 'output/'

if not os.path.exists(output_folder):
    os.makedirs(output_folder)

for file_name in os.listdir(input_folder):
    if file_name.endswith('.xlsx'):
        input_file = os.path.join(input_folder, file_name)
        process_single_file(input_file)

print("Processing completed for all files.")