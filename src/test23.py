#normalize data with pandas
import pandas as pd

df = pd.DataFrame({'ZONE': ['A', 'B', 'C', 'D', 'E'],
                   'usCASH': [18, 22, 19, 14, 14],
                   'bsCASH': [5, 7, 7, 9, 12],
                   'usCHECK': [11, 8, 10, 6, 6],
                   'bsCHECK': [3, 4, 4, 3, 3],},
                   )
df_unpivoted = pd.DataFrame(columns=['Currency', 'Payment Type', 'Amount', 'ZONE'])

# Add rows to the new DataFrame with the desired structure
for index, row in df.iterrows():
    zone = row['ZONE']
    us_cash = row['usCASH']
    bs_cash = row['bsCASH']
    us_check = row['usCHECK']
    bs_check = row['bsCHECK']
    df_unpivoted = df_unpivoted.concat({'Currency': 'us', 'Payment Type': 'CASH', 'Amount': us_cash, 'ZONE': zone}, ignore_index=True)
    df_unpivoted = df_unpivoted.concat({'Currency': 'bs', 'Payment Type': 'CASH', 'Amount': bs_cash, 'ZONE': zone}, ignore_index=True)
    df_unpivoted = df_unpivoted.concat({'Currency': 'us', 'Payment Type': 'CHECK', 'Amount': us_check, 'ZONE': zone}, ignore_index=True)
    df_unpivoted = df_unpivoted.concat({'Currency': 'bs', 'Payment Type': 'CHECK', 'Amount': bs_check, 'ZONE': zone}, ignore_index=True)

print(df_unpivoted)