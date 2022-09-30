import pandas as pd

df = pd.read_csv("./ok.csv")

data = df[['Issue Key', 'Issue summary']]

data.to_csv("./file.csv", index=False)