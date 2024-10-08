import pandas as pd
from readability import Readability

df = pd.read_excel('test.xlsx')

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

output_file = 'test_w_readability.xlsx'
df.to_excel(output_file, index=False)

print(f"Readability scores saved to {output_file}")