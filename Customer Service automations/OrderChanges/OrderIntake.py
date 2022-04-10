import pandas as pd
from os import walk


f = []
for (dirpath, dirnames, filenames) in walk(""):
    f.extend(filenames)
    break

for filename in filenames:
    final_df = pd.DataFrame()
    print('file loaded')
    df = pd.read_excel(dirpath + "\\" + filename)
    #df = df.append(df_temp,  ignore_index=True)

    #drop columns that i will not need
    columns_to_drop = ["Date.1", "Time.1","User.1","Doc.Number.1","Unnamed: 11","Unnamed: 12","Unnamed: 13","Unnamed: 14","Doc.Number", "Table"]
    for col in columns_to_drop:
        try:
            df = df.drop(col, axis = 1)
        except:
            continue

    # this is needed in order to merge Date and Time columns
    df['Date'] = df['Date'].astype('str')
    df['DateAndTime'] = df['Date'].str.cat(df['Time'],sep=" ")
    df['DateAndTime'] = pd.to_datetime(df['DateAndTime'], format="%Y-%m-%d %H:%M:%S")
    df = df.sort_values("DateAndTime")

    # make columns with orders Str, so we can use filter
    df['Obj. Value'] = df['Obj. Value'].astype('str')
    # key column that i wan't to get from this code
    df['Before/After'] = ""
    orders = set(df['Obj. Value'])

    iter = 0
    for order in orders:
        iter += 1
        if iter%1000 == 0: print(iter, "/", len(orders))
        #print('NEXT ORDER')
        temp_df = df.loc[df['Obj. Value'] == order]
        if len(temp_df[temp_df['Field Name'].str.contains('ZZ_WADAT')]) > 0:
            fasd_bool = False
            for row in range(0, len(temp_df)):
                #print(row, " out of ", len(temp_df))
                # if FASD field was found, next lines will be marked as 'After', otherwise 'Before'
                if temp_df.iloc[row, 2] == 'ZZ_WADAT':
                    temp_df.iloc[row, -1] = 'Before'
                    fasd_bool = True
                    #print("FASD found")
                    continue

                if fasd_bool:
                    temp_df.iloc[row, -1] = 'After'
                    #print("After")
                else:
                    #print("Before")
                    temp_df.iloc[row, -1] = 'Before'

        # if there is no FASD then it was entered before saving the doc, and all changes happend after it.
        else:
            #print('not FASD in', order)
            temp_df['Before/After'] = "After"

        final_df = final_df.append(temp_df,  ignore_index=True)

    final_df.to_excel(r"C:\Users\d4an\Downloads\Order Changes\final_result" + filename)