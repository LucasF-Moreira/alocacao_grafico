data_cols_existentes = [c for c in pivot.columns if c not in ["Pessoa", "Processo", "Etapa"]]
ordered_cols = ["Pessoa", "Processo", "Etapa"] + data_cols_existentes
pivot = pivot[ordered_cols]



